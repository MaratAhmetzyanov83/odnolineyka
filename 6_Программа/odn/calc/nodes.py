# -*- coding: utf-8 -*-
"""Расчёт узлов каскада (первый кусок расчёта, который живёт в Python, а не в Excel).

Для каждого узла щита:
  * линии под узлом — все уровни (узел под узлом);
  * Iр узла — сумма Iр линий по фазам (трёхфазная — на все три), берётся наиболее нагруженная фаза, без Кс (в запас);
  * фаз — 3, если под узлом есть трёхфазная линия; полюса: УЗО / АВДТ — 2P / 4P, автомат — 1P / 3P, рубильник — 2P / 3P;
  * In «авто» — ближайший стандартный не меньше Iр узла; текст аппарата — свой номинал из него;
  * обозначение: УЗО — QD, АВДТ — QFD, автомат — QF, рубильник — QS; номер — после наибольшего номера аппаратов щита;
  * линии под УЗО / АВДТ получают rcd_node — 30 мА им уже дано.
"""
from __future__ import annotations

import re

from ..core.model import Panel, Project

SERIES = [6, 10, 16, 20, 25, 32, 40, 50, 63, 80, 100, 125, 160, 200, 250, 315, 400, 500, 630, 800, 1000]
PREFIX = {"УЗО": "QD", "АВДТ": "QFD", "Автомат": "QF", "Рубильник": "QS"}
RCD = ("УЗО", "АВДТ")


def std_in(i):
    if i is None:
        return None
    return next((x for x in SERIES if x >= i - 1e-9), SERIES[-1])


def text_in(text: str):
    """Номинал из текста аппарата: «УЗО 2P 40А 30мА» -> 40, «1P C16А» -> 16."""
    t = (text or "").replace("A", "А").replace("С", "C")
    m = re.search(r"(?:C|B|D|\s)(\d+(?:[.,]\d+)?)\s*А(?!\w)", " " + t)
    return float(m.group(1).replace(",", ".")) if m else None


def _phases_of(ln):
    ph = [x for x in re.split(r"[,\s]+", ln.scheme.phase or "") if x in ("L1", "L2", "L3")]
    if len(ph) >= 3 or (ln.input.phases or 1) >= 3:
        return ["L1", "L2", "L3"]
    return ph or ["L1"]


def panel_nodes(p: Panel):
    nodes = {n.id: n for n in p.nodes}
    kids = {}
    for n in p.nodes:
        kids.setdefault(n.parent if n.parent in nodes else "", []).append(n)
    direct = {}
    for ln in p.lines:
        if ln.node in nodes:
            direct.setdefault(ln.node, []).append(ln)

    def subtree(nid, seen):
        if nid in seen:                       # цикл — не идём дальше (проверка сообщит)
            return []
        seen.add(nid)
        out = list(direct.get(nid, []))
        for k in kids.get(nid, []):
            out += subtree(k.id, seen)
        return out

    def level(n, seen=()):
        if not n.parent or n.parent not in nodes or n.id in seen:
            return 1
        return 1 + level(nodes[n.parent], (*seen, n.id))

    # номер: после наибольшего номера аппаратов щита (QF12, QFD38 -> 39)
    nums = [int(m.group(1)) for ln in p.lines for m in [re.search(r"(\d+)$", ln.scheme.qf or "")] if m]
    nxt = max(nums, default=0) + 1
    for n in p.nodes:
        lines = subtree(n.id, set())
        n.lines_all = [ln.number for ln in lines]
        n.level = level(n)
        acc = {"L1": 0.0, "L2": 0.0, "L3": 0.0}
        for ln in lines:
            i = ln.result.i_calc or 0.0
            for ph in _phases_of(ln):
                acc[ph] += i
        n.phase_i = {k: round(v, 3) for k, v in acc.items()}
        n.i_calc = round(max(acc.values()), 3) if lines else 0.0
        n.phases = 3 if any(len(_phases_of(ln)) == 3 for ln in lines) else 1
        rcd = n.kind in RCD
        n.poles = {("УЗО", 1): "2P", ("УЗО", 3): "4P", ("АВДТ", 1): "2P", ("АВДТ", 3): "4P",
                   ("Автомат", 1): "1P", ("Автомат", 3): "3P", ("Рубильник", 1): "2P", ("Рубильник", 3): "3P"
                   }.get((n.kind, n.phases), "")
        manual = n.device and n.device.strip().lower() not in ("авто", "+")
        n.in_a = text_in(n.device) if manual else std_in(n.i_calc)
        ma = n.rcd_ma or 30
        if manual:
            n.text = n.device
        elif n.kind == "УЗО":
            n.text = f"УЗО {n.poles} {n.in_a:g}А {ma}мА"
        elif n.kind == "АВДТ":
            n.text = f"{n.poles} C{n.in_a:g}А {ma}мА"
        elif n.kind == "Рубильник":
            n.text = f"QS {n.poles} {n.in_a:g}А"
        else:
            n.text = f"{n.poles} C{n.in_a:g}А"
        n.designation = f"{PREFIX.get(n.kind, 'Q')}{nxt}"
        nxt += 1

    # 30 мА от узла выше
    for ln in p.lines:
        ln.rcd_node = ""
        nid, seen = ln.node, set()
        while nid in nodes and nid not in seen:
            seen.add(nid)
            if nodes[nid].kind in RCD:
                ln.rcd_node = nid
                break
            nid = nodes[nid].parent


def run(prj: Project):
    for p in prj.panels():
        if p.nodes:
            panel_nodes(p)
        else:
            for ln in p.lines:
                ln.rcd_node = ""
