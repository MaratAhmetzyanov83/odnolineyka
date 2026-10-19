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

    def test_selectivity_input_qs(self):
        """Автомат на вводе ЩР-1 32 А — база 32 (фидер 32 < 51,2). QS на вводе — база отходящий 16 (32 ≥ 25,6)."""
        prj = demo()
        shr, grsh = prj.panel("ЩР-1.xlsx / ЩР-1"), prj.panel("ГРЩ.xlsx / ГРЩ")
        shr.result.i_calc, grsh.lines[0].scheme.in_a = 20, 32
        inp = line(shr.id, "Ввод.1", 32)
        inp.scheme.breaker = "3P C32А"
        shr.lines.insert(0, inp)
        self.assertIn(("links.selectivity", "ГРЩ", "Щ.1"), rules_hit(recheck(prj)))
        inp.scheme.breaker = "QS 3P 32А"
        self.assertNotIn(("links.selectivity", "ГРЩ", "Щ.1"), rules_hit(recheck(prj)))

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
        P = "ЩР-1.xlsx / ЩР-1"
        ch = changes.apply(prj, changes.Change(P, "input.length_plan", "45,5", line="R.1"))
        self.assertEqual((ch.old, prj.line(P, "R.1").input.length_plan), (None, 45.5))
        ch = changes.apply(prj, changes.Change(P, "input.cores", "5", line="R.1"))
        self.assertEqual(prj.line(P, "R.1").input.cores, 5)
        ch = changes.apply(prj, changes.Change(P, "input.p_kw", "", line="R.1"))     # пусто — очистить
        self.assertIsNone(prj.line(P, "R.1").input.p_kw)
        for bad in [changes.Change(P, "result.i_calc", 1, line="R.1"),          # расчёт не правится
                    changes.Change(P, "input.p_kw", 1, line="нет"),             # нет линии
                    changes.Change(P, "input.p_kw", "abc", line="R.1"),         # не число
                    changes.Change(P, "input.cores", "2,5", line="R.1"),        # не целое
                    changes.Change(P, "input.p_kw", -1, line="R.1"),            # меньше нуля
                    ]:
            with self.assertRaises(changes.EditError, msg=bad.path):
                changes.apply(prj, bad)
        ln = prj.line(P, "R.1")
        ln.scheme.qf, ln.scheme.group_head = "QF7", False          # продолжение чужой группы — нельзя
        with self.assertRaises(changes.EditError):
            changes.apply(prj, changes.Change(P, "scheme.breaker_set", "авто", line="R.1"))
        ln.scheme.qf = ""                                           # строка без группы (ввод) — можно
        changes.apply(prj, changes.Change(P, "scheme.breaker_set", "авто", line="R.1"))
        prj.line(P, "R.1").scheme.group_head = True
        changes.apply(prj, changes.Change(P, "scheme.breaker_set", "1P C10А", line="R.1"))
        self.assertEqual(prj.line(P, "R.1").scheme.breaker_set, "1P C10А")

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


class TestDuLength(unittest.TestCase):
    """ΔUдоп + 0,005 %/м сверх 100 м от ВРУ, не больше 0,5 % (ГОСТ Р 50571.5.52, прил. G)."""

    def test_extra(self):
        from odn.calc.du import extra_for
        self.assertEqual([extra_for(x) for x in (None, 50, 100, 130, 200, 300, 500)], [0, 0, 0, 0.15, 0.5, 0.5, 0.5])

    def test_chain_and_status(self):
        prj = demo()                                   # ГРЩ: Щ.1 L = 30 м -> ЩР-1: R.1, R.2 L = 20 м
        shr = prj.panel("ЩР-1.xlsx / ЩР-1")
        r1, r2 = shr.lines
        r1.input.length, r1.result.du_max, r1.result.du_sum = 100.0, 4.0, 4.3   # 30 + 100 = 130 м -> +0,15 %
        r1.result.status, r1.result.level = "потери 4,3 % > 4 %", 1
        r2.input.length, r2.result.du_max, r2.result.du_sum = 100.0, 4.0, 4.1
        r2.result.status, r2.result.level = "потери 4,1 % > 4 %; нет мощности", 1
        recheck(prj)
        self.assertEqual((r1.result.du_path_len, r1.result.du_extra, r1.result.du_max_eff), (130.0, 0.15, 4.15))
        st = {i.line: i.text for i in prj.issues if i.rule == "source.row_status"}
        self.assertEqual(st.get("R.1"), "потери 4,3 % > 4 %")      # 4,3 > 4,15 — остаётся
        self.assertEqual(st.get("R.2"), "нет мощности")             # 4,1 ≤ 4,15 — потери сняты
        r1.input.u_v = 24.0                                         # до 50 В — без надбавки
        recheck(prj)
        self.assertIsNone(r1.result.du_extra)


class TestVru(unittest.TestCase):
    """ВРУ здания: потери и длина — от его шин."""

    def setUp(self):
        from odn.calc import du
        self.du = du
        self.prj = demo()                                            # ГРЩ: Щ.1 L = 30 м -> ЩР-1
        self.shr = self.prj.panel("ЩР-1.xlsx / ЩР-1")
        self.shr.result.du_bus = 1.2                                 # книга: ΔU до шин ЩР-1 от ГРЩ
        self.r1 = self.shr.lines[0]
        self.r1.input.length, self.r1.result.du_max, self.r1.result.du_sum = 90.0, 4.0, 4.6
        self.r1.result.status, self.r1.result.level = "потери 4,6 % > 4 %", 1

    def test_program_vru(self):
        self.du.run(self.prj, {self.shr.id: True})
        r = self.r1.result
        self.assertEqual((self.shr.vru, self.shr.vru_src, r.du_from), (True, "программа", self.shr.id))
        self.assertEqual((r.du_sum_eff, r.du_path_len, r.du_extra), (3.4, 90.0, 0.0))   # 4,6 − 1,2; длина без фидера
        from odn import checks
        st = [i.text for i in checks.run(self.prj) if i.line == "R.1" and i.rule.startswith("source.")]
        self.assertFalse(any("потери" in t for t in st), st)                           # 3,4 ≤ 4 — снято

    def test_book_vru(self):
        self.shr.scheme.is_vru = "да"                                 # книга уже считает от его шин
        self.du.run(self.prj, {})
        self.assertEqual((self.shr.vru_src, self.r1.result.du_sum_eff), ("книга", 4.6))  # вычитать нечего

    def test_no_vru(self):
        self.du.run(self.prj, {})
        r = self.r1.result
        self.assertEqual((self.shr.vru, r.du_sum_eff, r.du_path_len, r.du_extra), (False, 4.6, 120.0, 0.1))

    def test_over_from_vru(self):
        """ВРУ только в программе, а потери от него всё равно больше допуска — своё замечание."""
        self.r1.result.du_sum, self.r1.result.status, self.r1.result.level = 5.5, "", 0
        self.du.run(self.prj, {self.shr.id: True})
        from odn import checks
        hits = [i.text for i in checks.run(self.prj) if i.rule == "source.du_over"]
        self.assertEqual(hits, ["потери 4,30 % > 4,00 % (от ВРУ здания)"])
