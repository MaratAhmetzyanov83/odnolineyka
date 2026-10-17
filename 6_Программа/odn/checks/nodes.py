# -*- coding: utf-8 -*-
"""Узлы каскада (структура программы): расчёт — calc.nodes."""
from ..calc.nodes import RCD
from .registry import Hit, rule
from .util import f

G = "Узлы"
SEL_K = 1.6


def _nodes(prj):
    for p in prj.panels():
        ids = {n.id: n for n in p.nodes}
        for n in p.nodes:
            yield p, n, ids


def _children_in(p, n, ids):
    """In того, что стоит прямо под узлом: автоматы линий и узлы."""
    out = [ln.in_a for ln in p.lines if ln.node == n.id and ln.in_a]
    out += [k.in_a for k in ids.values() if k.parent == n.id and k.in_a]
    return out


@rule("nodes.structure", "Узел: структура", G)
def structure(prj):
    for p, n, ids in _nodes(prj):
        if n.parent and n.parent not in ids:
            yield Hit(2, f"узел {n.id}: нет узла выше {n.parent}", p.id, n.id)
        seen, x = set(), n
        while x.parent in ids:
            if x.id in seen:
                yield Hit(2, f"узел {n.id}: цикл в цепочке узлов", p.id, n.id)
                break
            seen.add(x.id)
            x = ids[x.parent]
        if not n.lines_all and not any(k.parent == n.id for k in ids.values()):
            yield Hit(1, f"узел {n.id}: под ним нет линий", p.id, n.id)


@rule("nodes.overload", "Узел: перегруз", G)
def overload(prj):
    for p, n, ids in _nodes(prj):
        if n.in_a and n.i_calc and n.in_a < n.i_calc:
            yield Hit(2, f"узел {n.id} {n.designation}: In {f(n.in_a, 0)} А < Iр {f(n.i_calc)} А", p.id, n.id)


@rule("nodes.selectivity", "Узел: селективность", G)
def selectivity(prj):
    for p, n, ids in _nodes(prj):
        if n.kind in ("Автомат", "АВДТ") and n.in_a:
            ch = _children_in(p, n, ids)
            if ch and n.in_a < SEL_K * max(ch):
                yield Hit(1, f"узел {n.id} {n.designation}: неселективно — In {f(n.in_a, 0)} А < 1,6 × {f(max(ch), 0)} А",
                          p.id, n.id)


@rule("nodes.rcd_protection", "УЗО: защита от перегрузки", G)
def rcd_protection(prj):
    """УЗО не защищает от сверхтока: сумма In автоматов под ним не больше In УЗО, иначе — автомат перед УЗО."""
    for p, n, ids in _nodes(prj):
        if n.kind == "УЗО" and n.in_a:
            s = sum(_children_in(p, n, ids))
            up = ids.get(n.parent)
            protected = up is not None and up.kind in ("Автомат", "АВДТ") and up.in_a and up.in_a <= n.in_a
            if s > n.in_a and not protected:
                yield Hit(1, f"узел {n.id} {n.designation}: УЗО {f(n.in_a, 0)} А, сумма автоматов под ним {f(s, 0)} А — "
                             f"нужен автомат перед УЗО не больше {f(n.in_a, 0)} А", p.id, n.id)


@rule("nodes.extra_rcd", "Лишнее 30 мА под УЗО", G)
def extra_rcd(prj):
    for p in prj.panels():
        for ln in p.lines:
            if ln.rcd_node and ("мА" in (ln.scheme.breaker or "") or (ln.scheme.qf or "").startswith("QFD")):
                yield Hit(1, f"под узлом {ln.rcd_node} (30 мА) у линии свой дифавтомат {ln.scheme.qf} — можно QF",
                          p.id, ln.number, ln.refs.get("scheme"))


@rule("nodes.lost_lines", "Узел: линия не найдена", G)
def lost_lines(prj):
    from ..services.structure import load
    data = load(prj.folder)
    for p in prj.panels():
        have = {ln.number for ln in p.lines}
        for num, nid in data["panels"].get(p.id, {}).get("lines", {}).items():
            if num not in have:
                yield Hit(1, f"линия {num} (узел {nid}) — её нет в книге (переименована или удалена)", p.id, num)
