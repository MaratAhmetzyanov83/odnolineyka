# -*- coding: utf-8 -*-
"""Чтение книг v3.12 на тестовом объекте (make_demo.py: ГРЩ -> ЩР-1, ЩР-2 в подпапке, копия в _резерв).

Объект строится один раз (≈30 с) в папку ODN_DEMO или во временную; повторно — берётся готовый.
"""
import os
import sys
import tempfile
import unittest

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.dirname(HERE))
sys.path.insert(0, HERE)

from odn.services.project import load_project  # noqa: E402

FOLDER = os.environ.get("ODN_DEMO") or os.path.join(tempfile.gettempdir(), "odn_demo")


class TestExcelV312(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        if not os.path.isfile(os.path.join(FOLDER, "Корпус 2", "ЩР-2.xlsx")):
            import make_demo
            make_demo.main(FOLDER)
        cls.prj = load_project(FOLDER, progress=lambda s: None, use_cache=False)
        cls.p = {p.name: p for p in cls.prj.panels()}

    def test_panels(self):
        self.assertEqual(sorted(self.p), ["ГРЩ", "ЩР-1", "ЩР-2"], "копия в _резерв не должна попасть")

    def test_links(self):
        self.assertEqual((self.p["ЩР-1"].parent, self.p["ЩР-1"].parent_line), (self.p["ГРЩ"].id, "Щ.1"))
        self.assertEqual((self.p["ЩР-2"].parent, self.p["ЩР-2"].parent_line), (self.p["ГРЩ"].id, "Щ.2"))

    def test_line_fields(self):
        ln = next(x for x in self.p["ЩР-1"].lines if x.number == "R.1.01.1")
        self.assertEqual((ln.scheme.place, ln.scheme.qf, ln.scheme.breaker, ln.scheme.in_a), (19, "QFD1", "2P С16А 30мА", 16.0))
        self.assertEqual((ln.input.p_kw, ln.input.length, ln.result.in_by_current, ln.result.by), (0.06, 10.0, 6.0, "excel"))
        self.assertEqual(str(ln.refs["status"]), "Однолинейка!AE21")
        # строка группы без своего автомата берёт автомат группы
        ln2 = next(x for x in self.p["ЩР-1"].lines if x.number == "R.1.02.1")
        self.assertEqual((ln2.scheme.qf, ln2.scheme.breaker), ("QFD1", "2P С16А 30мА"))

    def test_link_issues(self):
        got = sorted((i.rule, next(p.name for p in self.prj.panels() if p.id == i.panel), i.line)
                     for i in self.prj.issues if i.group == "Связи щитов")
        self.assertEqual(got, [
            ("links.feeder_no_book", "ГРЩ", "Щ.3"),
            ("links.feeder_stale", "ЩР-2", ""),
            ("links.load_stale", "ГРЩ", "Щ.1"),
            ("links.selectivity", "ГРЩ", "Щ.1"),
            ("links.selectivity", "ГРЩ", "Щ.2"),
        ])

    def test_excel_counts(self):
        errs = [i for i in self.prj.issues if i.rule == "source.row_status" and i.level == 2 and i.panel == self.p["ЩР-1"].id]
        self.assertEqual(len(errs), self.p["ЩР-1"].result.errors)
        self.assertEqual(len(errs), 6)

    def test_cache(self):
        a = load_project(FOLDER, progress=lambda s: None)
        b = load_project(FOLDER, progress=self.fail)          # всё из кэша — ничего не читается
        self.assertEqual(len(a.issues), len(b.issues))
        self.assertEqual([p.revision for p in a.panels()], [p.revision for p in b.panels()])


if __name__ == "__main__":
    unittest.main()
