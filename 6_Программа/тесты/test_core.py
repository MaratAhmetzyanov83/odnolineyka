# -*- coding: utf-8 -*-
"""Быстрые проверки ядра и правил — без Excel (объект собирается в коде).

    py -m unittest discover -s тесты -v
"""
import os
import sys
import unittest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from odn import checks  # noqa: E402
from odn.core import changes, serialize  # noqa: E402
from odn.core.model import Document, FeederSeen, Line, Panel, Project, Ref, Supply  # noqa: E402
from odn.core.revision import panel_revision  # noqa: E402
from odn.services import linking  # noqa: E402
from odn.services.project import recheck  # noqa: E402


def line(pid, number, in_a=16.0, length=20.0, group="Быт", name=""):
    ln = Line(number=number, panel=pid, name=name or number)
    ln.scheme.in_a, ln.input.length, ln.input.group_ks = in_a, length, group
    ln.refs["input"] = Ref("x.xlsx", "Исходные данные", "D3")
    return ln


def demo():
    """ГРЩ (фидер Щ.1 25 А на ЩР-1) -> ЩР-1 (отходящие 16 А, Iр 30 А)."""
    grsh = Panel(id="ГРЩ.xlsx / ГРЩ", name="ГРЩ", doc="ГРЩ.xlsx", supply=Supply(manual="ТП"))
    grsh.lines = [line(grsh.id, "Щ.1", 25, 30, "Щит", "Питание ЩР-1")]
    grsh.feeders_seen = [FeederSeen(line="Щ.1", doc="ЩР-1.xlsx", p_inst=3.0)]
    shr = Panel(id="ЩР-1.xlsx / ЩР-1", name="ЩР-1", doc="ЩР-1.xlsx",
                supply=Supply(from_doc="ГРЩ.xlsx / ГРЩ", in_a=20, length=30))
    shr.result.p_inst, shr.result.i_calc = 5.5, 30.0
    shr.lines = [line(shr.id, "R.1"), line(shr.id, "R.2")]
    prj = Project(name="t", folder="/obj",
                  documents=[Document("ГРЩ.xlsx", panels=[grsh]), Document("ЩР-1.xlsx", panels=[shr])])
    return recheck(prj)


def rules_hit(prj):
    return sorted({(i.rule, i.panel.split(" / ")[1], i.line) for i in prj.issues})


class TestLinks(unittest.TestCase):
    def test_link(self):
        prj = demo()
        shr = prj.panel("ЩР-1.xlsx / ЩР-1")
        self.assertEqual((shr.parent, shr.parent_line), ("ГРЩ.xlsx / ГРЩ", "Щ.1"))
        self.assertEqual(prj.line("ГРЩ.xlsx / ГРЩ", "Щ.1").feeds, shr.id)
        self.assertEqual([p.name for p in prj.children("ГРЩ.xlsx / ГРЩ")], ["ЩР-1"])

    def test_rules(self):
        self.assertEqual(rules_hit(demo()), [
            ("links.feeder_small", "ГРЩ", "Щ.1"),        # 25 А < Iр 30 А
            ("links.feeder_stale", "ЩР-1", ""),         # ЩР-1 знает 20 А, у ГРЩ 25 А
            ("links.load_stale", "ГРЩ", "Щ.1"),          # ГРЩ знает Руст 3, сейчас 5,5
            ("links.selectivity", "ГРЩ", "Щ.1"),         # 25 < 1,6 × 16
        ])

    def test_fixed(self):
        prj = demo()
        shr, grsh = prj.panel("ЩР-1.xlsx / ЩР-1"), prj.panel("ГРЩ.xlsx / ГРЩ")
        shr.supply.in_a, shr.result.i_calc = 32, 20
        grsh.feeders_seen[0].p_inst = 5.5
        grsh.lines[0].scheme.in_a = 32
        self.assertEqual(rules_hit(recheck(prj)), [])

    def test_missing_parent(self):
        prj = demo()
        prj.panel("ЩР-1.xlsx / ЩР-1").supply.from_doc = "..\\Нет.xlsx"
        hit = rules_hit(recheck(prj))
        self.assertIn(("links.parent_missing", "ЩР-1", ""), hit)
        self.assertIn(("links.feeder_no_book", "ГРЩ", "Щ.1"), hit)

    def test_path_variants(self):
        self.assertEqual(linking._abs("/obj", "Корпус 2/ЩР-2.xlsx", "..\\ГРЩ.xlsx"), linking.norm("/obj/ГРЩ.xlsx"))


class TestCore(unittest.TestCase):
    def test_roundtrip(self):
        prj = demo()
        d = serialize.to_dict(prj)
        back = serialize.from_dict(Project, d).reindex()
        self.assertEqual(serialize.to_dict(back), d)
        self.assertIsInstance(back.panels()[0].lines[0].refs["input"], Ref)

    def test_old_json(self):
        """Лишние и недостающие поля не ломают чтение (кэш другой версии)."""
        ln = serialize.from_dict(Line, {"number": "1", "лишнее": 5, "input": {"p_kw": 1.0, "нет": 1}})
        self.assertEqual((ln.number, ln.input.p_kw, ln.input.length), ("1", 1.0, None))

    def test_revision(self):
        prj = demo()
        p = prj.panel("ЩР-1.xlsx / ЩР-1")
        r1 = panel_revision(p)
        p.issues_dummy = 1                      # то, что не в чертеже, не влияет
        p.result.errors = 7
        self.assertEqual(panel_revision(p), r1)
        p.lines[0].input.length = 99
        self.assertNotEqual(panel_revision(p), r1)

    def test_change(self):
        prj = demo()
        ch = changes.apply(prj, changes.Change("ЩР-1.xlsx / ЩР-1", "input.length_plan", 45.0, line="R.1"))
        self.assertEqual((ch.old, prj.line("ЩР-1.xlsx / ЩР-1", "R.1").input.length_plan), (None, 45.0))
        with self.assertRaises(PermissionError):
            changes.apply(prj, changes.Change("ЩР-1.xlsx / ЩР-1", "result.i_calc", 1, line="R.1"))
        with self.assertRaises(KeyError):
            changes.apply(prj, changes.Change("ЩР-1.xlsx / ЩР-1", "input.p_kw", 1, line="нет"))

    def test_broken_rule(self):
        """Упавшее правило даёт замечание, остальные работают."""
        @checks.rule("test.broken", "Сломанное", "Тест")
        def broken(prj):
            raise ValueError("сбой")
        try:
            prj = demo()
            got = [i for i in prj.issues if i.rule == "test.broken"]
            self.assertEqual(len(got), 1)
            self.assertIn("не выполнена", got[0].text)
            self.assertTrue(any(i.rule == "links.feeder_small" for i in prj.issues))
        finally:
            checks.registry._RULES.pop("test.broken")


if __name__ == "__main__":
    unittest.main()
