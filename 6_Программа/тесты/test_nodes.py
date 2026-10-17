# -*- coding: utf-8 -*-
"""Узлы каскада: расчёт (calc.nodes), операции структуры (services.structure), проверки «Узлы». Без Excel."""
import os
import shutil
import sys
import tempfile
import unittest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from odn.core.changes import EditError  # noqa: E402
from odn.core.model import Document, Line, Panel, Project  # noqa: E402
from odn.services import structure  # noqa: E402
from odn.services.project import recheck  # noqa: E402
from odn.calc.nodes import text_in, std_in  # noqa: E402

P = "ЩР.xlsx / ЩР"


def line(num, i, phase="L1", qf="", brk="1P C16А", in_a=16.0):
    ln = Line(number=num, panel=P, name=num)
    ln.result.i_calc, ln.scheme.phase, ln.scheme.qf, ln.scheme.breaker, ln.scheme.in_a = i, phase, qf, brk, in_a
    ln.input.phases = 3 if "L3" in phase and "," in phase else 1
    return ln


class TestNodes(unittest.TestCase):
    def setUp(self):
        self.dir = tempfile.mkdtemp(prefix="odn_nodes_")
        p = Panel(id=P, name="ЩР", doc="ЩР.xlsx")
        p.lines = [line("R.1", 7, "L1", "QFD1", "2P C16А 30мА"), line("R.2", 3, "L2", "QF2"), line("L.1", 2, "L3", "QF3"),
                   line("L.2", 1.5, "L3", "QF4"), line("P.1", 10, "L1,L2,L3", "QF5", "3P C16А"), line("R.9", 1, "L1", "QF38")]
        self.prj = recheck(Project(name="t", folder=self.dir, documents=[Document("ЩР.xlsx", panels=[p])]))

    def tearDown(self):
        shutil.rmtree(self.dir, ignore_errors=True)

    def op(self, **kw):
        r = structure.operate(self.prj, {"panel": P, **kw})
        recheck(self.prj)
        return r

    def nodes(self):
        return {n.id: n for n in self.prj.panel(P).nodes}

    def test_cascade(self):
        self.assertEqual(self.op(op="add_node", kind="УЗО", lines=["R.1", "R.2"])["node"], "У1")
        self.op(op="add_node", kind="Автомат", parent="У1", lines=["L.1", "L.2"])
        n = self.nodes()
        self.assertEqual((n["У2"].level, n["У2"].i_calc, n["У2"].in_a, n["У2"].poles), (2, 3.5, 6, "1P"))
        self.assertEqual(n["У1"].phase_i, {"L1": 7.0, "L2": 3.0, "L3": 3.5})
        self.assertEqual((n["У1"].i_calc, n["У1"].in_a, n["У1"].text), (7.0, 10, "УЗО 2P 10А 30мА"))
        self.assertEqual(sorted(n["У1"].lines_all), ["L.1", "L.2", "R.1", "R.2"])
        self.assertEqual((n["У1"].designation, n["У2"].designation), ("QD39", "QF40"))     # после QF38
        rcd = {ln.number: ln.rcd_node for ln in self.prj.panel(P).lines}
        self.assertEqual(rcd, {"R.1": "У1", "R.2": "У1", "L.1": "У1", "L.2": "У1", "P.1": "", "R.9": ""})
        hits = {(i.rule, i.line) for i in self.prj.issues}
        self.assertIn(("nodes.rcd_protection", "У1"), hits)       # 16+16+6 > 10
        self.assertIn(("nodes.selectivity", "У2"), hits)           # 6 < 1,6 × 16
        self.assertIn(("nodes.extra_rcd", "R.1"), hits)            # свой QFD под УЗО
        # трёхфазная линия под АВДТ -> 4P
        self.op(op="add_node", kind="АВДТ", device="авто", lines=["P.1"])
        self.assertEqual((self.nodes()["У3"].poles, self.nodes()["У3"].text), ("4P", "4P C10А 30мА"))

    def test_edit_delete_undo(self):
        self.op(op="add_node", kind="УЗО", lines=["R.1", "R.2"])
        self.op(op="add_node", kind="Автомат", parent="У1", lines=["L.1"])
        self.op(op="update_node", id="У1", fields={"device": "УЗО 2P 40А 30мА", "kind": "УЗО"})
        self.assertEqual((self.nodes()["У1"].in_a, self.nodes()["У1"].text), (40, "УЗО 2P 40А 30мА"))
        self.assertNotIn(("nodes.rcd_protection", "У1"), {(i.rule, i.line) for i in self.prj.issues})
        with self.assertRaises(EditError):                          # цикл
            self.op(op="update_node", id="У1", fields={"parent": "У2"})
        self.op(op="delete_node", id="У1")                         # У2 и линии — на шину
        self.assertEqual(self.nodes()["У2"].parent, "")
        self.assertEqual(self.prj.line(P, "R.1").node, "")
        self.op(op="move_lines", lines=["R.9"], node="У2")
        self.assertEqual(self.prj.line(P, "R.9").node, "У2")
        structure.undo_last(self.prj); recheck(self.prj)           # отмена переноса
        self.assertEqual(self.prj.line(P, "R.9").node, "")
        structure.undo_last(self.prj); recheck(self.prj)           # отмена удаления
        self.assertEqual(self.prj.line(P, "R.1").node, "У1")
        with self.assertRaises(EditError):
            self.op(op="add_node", kind="Предохранитель", lines=["R.1"])
        with self.assertRaises(EditError):
            self.op(op="move_lines", lines=["нет"], node="")

    def test_no_lines_and_lost(self):
        self.op(op="add_node", kind="Рубильник", device="QS 3P 63А")
        hits = {(i.rule, i.line) for i in self.prj.issues}
        self.assertIn(("nodes.structure", "У1"), hits)
        self.assertEqual(self.nodes()["У1"].in_a, 63)
        self.op(op="move_lines", lines=["R.1"], node="У1")
        self.prj.panel(P).lines = [ln for ln in self.prj.panel(P).lines if ln.number != "R.1"]   # линию убрали в книге
        recheck(self.prj)
        self.assertIn(("nodes.lost_lines", "R.1"), {(i.rule, i.line) for i in self.prj.issues})

    def test_drawing_export(self):
        from odn.services import drawing_export as D
        self.op(op="add_node", kind="УЗО", lines=["R.1", "R.2"])
        self.op(op="add_node", kind="Автомат", parent="У1", lines=["L.1"])
        path = os.path.join(self.dir, "ЩР.nodes.txt")
        self.assertTrue(os.path.isfile(path))
        with open(path, encoding="ascii") as f:
            text = f.read()
        rows = text.strip().splitlines()
        self.assertTrue(rows[0].startswith(";; odn nodes"))
        self.assertIn('(NODE (1065 1056) (1059 49) "RCD" (81 68 51 57)', text)          # ЩР, У1, QD39
        self.assertIn("(LINE (1065 1056) (76 46 49) (1059 49))", text)                     # L.1 -> корень У1
        self.assertEqual(sum(r.startswith("(LINE") for r in rows), 3)
        hits = {(i.rule, i.line) for i in self.prj.issues}
        self.assertIn(("nodes.drawing", "У2"), hits)                                       # уровень 2 — плоско
        self.op(op="delete_node", id="У2"); self.op(op="delete_node", id="У1")
        self.assertFalse(os.path.isfile(path))                                             # узлов нет — файла нет

    def test_contiguous(self):
        for i, ln in enumerate(self.prj.panel(P).lines):
            ln.scheme.place, ln.scheme.sheet = i + 1, 1
        self.op(op="add_node", kind="УЗО", lines=["R.1", "L.1"])                          # места 1 и 3, между — R.2
        msg = [i.text for i in self.prj.issues if i.rule == "nodes.drawing"]
        self.assertTrue(any("R.2" in m for m in msg), msg)

    def test_parse(self):
        self.assertEqual([text_in(t) for t in ("УЗО 2P 40А 30мА", "1P C16А", "QS 3P 63A", "2P С25А 30мА", "x")],
                         [40, 16, 63, 25, None])
        self.assertEqual((std_in(7), std_in(10), std_in(0)), (10, 10, 6))


if __name__ == "__main__":
    unittest.main()
