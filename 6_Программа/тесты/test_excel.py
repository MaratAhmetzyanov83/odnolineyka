# -*- coding: utf-8 -*-
"""Чтение книг v3.12 на тестовом объекте (make_demo.py: ГРЩ -> ЩР-1, ЩР-2 в подпапке, копия в _резерв).

Объект строится один раз (≈30 с) в папку ODN_DEMO или во временную; повторно — берётся готовый.
"""
import os
import shutil
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


class TestEdit(unittest.TestCase):
    """Запись правок, журнал, резерв, отмена — на копии объекта (тестовая запись openpyxl, без Excel)."""

    @classmethod
    def setUpClass(cls):
        TestExcelV312.setUpClass()
        cls.dir = tempfile.mkdtemp(prefix="odn_edit_")
        cls.folder = os.path.join(cls.dir, "объект")
        shutil.copytree(FOLDER, cls.folder, ignore=shutil.ignore_patterns("_окно"))
        os.environ["ODN_TEST_WRITER"] = "1"

    @classmethod
    def tearDownClass(cls):
        os.environ.pop("ODN_TEST_WRITER", None)
        shutil.rmtree(cls.dir, ignore_errors=True)

    def load(self):
        return load_project(self.folder, progress=lambda s: None)

    def test_write_undo(self):
        from odn.services import edit
        prj = self.load()
        pid = next(p.id for p in prj.panels() if p.name == "ЩР-1")
        r = edit.apply_changes(prj, [
            {"panel": pid, "line": "R.1.01.1", "path": "input.length_plan", "new": "12,5"},
            {"panel": pid, "line": "R.1.01.1", "path": "scheme.breaker_set", "new": "1P C10А"},
            {"panel": pid, "line": "R.1.02.1", "path": "input.cores", "new": 5},
        ])
        self.assertEqual((r["written"], r["docs"], r["errors"]), (3, ["ЩР-1.xlsx"], []))
        prj = self.load()
        self.assertEqual(prj.line(pid, "R.1.01.1").input.length_plan, 12.5)
        self.assertEqual(prj.line(pid, "R.1.01.1").scheme.breaker_set, "1P C10А")
        self.assertEqual(prj.line(pid, "R.1.02.1").input.cores, 5)
        j = edit.read_journal(prj)
        self.assertEqual(len(j), 1)
        self.assertEqual(j[0]["changes"][0]["cell"], "Исходные данные!I21")
        self.assertTrue(os.path.isfile(os.path.join(self.folder, j[0]["backup"])))
        # ничего не изменилось — ничего не пишется
        self.assertEqual(edit.apply_changes(prj, [{"panel": pid, "line": "R.1.01.1", "path": "input.length_plan", "new": 12.5}])["written"], 0)
        # отмена
        r = edit.undo_last(prj)
        self.assertEqual(r["written"], 3)
        prj = self.load()
        self.assertEqual((prj.line(pid, "R.1.01.1").input.length_plan, prj.line(pid, "R.1.01.1").scheme.breaker_set), (8.0, "авто"))
        self.assertEqual(edit.undo_last(prj)["errors"], ["нечего отменять"])
        # служебная папка не читается как книга объекта
        self.assertEqual(sorted(p.name for p in prj.panels()), ["ГРЩ", "ЩР-1", "ЩР-2"])

    def test_conflict(self):
        from odn.services import edit
        prj = self.load()
        pid = next(p.id for p in prj.panels() if p.name == "ЩР-2")
        path = os.path.join(self.folder, prj.panel(pid).doc)
        os.utime(path, (os.path.getatime(path), os.path.getmtime(path) + 10))     # книгу изменили после чтения
        r = edit.apply_changes(prj, [{"panel": pid, "line": "R.1.01.1", "path": "input.p_kw", "new": 1}])
        self.assertEqual(r["written"], 0)
        self.assertIn("изменилась после чтения", r["errors"][0])

    def test_locked(self):
        from odn.services import edit
        prj = self.load()
        pid = next(p.id for p in prj.panels() if p.name == "ГРЩ")
        lock = os.path.join(self.folder, "~$ГРЩ.xlsx")
        open(lock, "w").close()
        try:
            r = edit.apply_changes(prj, [{"panel": pid, "line": "Щ.1", "path": "input.p_kw", "new": 9}])
        finally:
            os.remove(lock)
        self.assertIn("открыта в Excel", r["errors"][0])

    def test_export(self):
        from odn.services import export
        import io
        import openpyxl
        data = export.to_xlsx("Линии", [{"t": "Линия"}, {"t": "P", "n": True}], [["R.1", 0.5], ["R.2", None]])
        ws = openpyxl.load_workbook(io.BytesIO(data)).active
        self.assertEqual([[c.value for c in r] for r in ws.iter_rows()], [["Линия", "P"], ["R.1", 0.5], ["R.2", None]])


if __name__ == "__main__":
    unittest.main()
