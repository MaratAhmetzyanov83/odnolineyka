# -*- coding: utf-8 -*-
"""«Исправить ошибки»: предложения правок по замечаниям. Сами ничего не пишут — окно показывает список «было → стало»,
пользователь отмечает нужные, отмеченные идут обычным путём: правки книг — в «несохранённые» (Ctrl+S), узлы —
операцией структуры, связи — «Синхронизировать книги». Отмена — как у любой правки.

Предложение: {id, rule, kind, panel, line, text, was, now, path?, op?, checked}
    kind  edit       — правка поля линии (panel, line, path, now) -> pending
          structure  — операция структуры (op) -> /api/structure
          sync       — «Синхронизировать книги»
          manual     — только подсказка (DWG, КЖ — перестроить в AutoCAD / Excel)
    checked — отметить сразу (однозначная правка); False — проверить глазами.
Новое исправление — функция в FIXERS: (prj) -> предложения.
"""
from __future__ import annotations

import re

from ..checks.links import SEL_K, _feeders, _is_input
from ..checks.util import f

SERIES = [6, 10, 16, 20, 25, 32, 40, 50, 63, 80, 100, 125, 160, 200, 250, 315, 400, 500, 630, 800, 1000, 1250, 1600]


def pick_in(need):
    return next((x for x in SERIES if x >= need - 1e-9), SERIES[-1])


def feeder_need(child):
    """In фидера: ≥ 1,6 × In ввода щита (если есть), иначе × наибольшего отходящего; и ≥ Iр щита (как окно «Фидер»)."""
    inp = next((x for x in child.lines if _is_input(x) and x.in_a), None)
    outs = [x.in_a for x in child.lines if x.in_a and not _is_input(x)]
    base = inp.in_a if inp else (max(outs) if outs else None)
    return max(SEL_K * base if base else 0, child.result.i_calc or 0), base


def fix_feeders(prj):
    for child, par, ln in _feeders(prj):
        need, base = feeder_need(child)
        have = ln.in_a
        if not need or (have is not None and have >= need - 1e-9):
            continue
        now = pick_in(need)
        why = []
        if have is not None and child.result.i_calc and have < child.result.i_calc:
            why.append(f"In {f(have)} А < Iр щита {f(child.result.i_calc)} А")
        if base and (have is None or have < SEL_K * base):
            why.append(f"неселективно: нужно ≥ 1,6 × {f(base)} А")
        yield {"rule": "links.feeder", "kind": "edit", "panel": par.id, "line": ln.number, "path": "input.in_manual",
               "text": f"фидер {child.name}: {'; '.join(why) or 'In мал'} — In ручной {now} А",
               "was": f"{f(have)} А" if have is not None else "—", "now": str(now), "checked": True}


def fix_sync(prj):
    """Книги «не знают» друг о друге (нагрузка, фидер) — пересчитать связанные книги вместе."""
    n = 0
    for child, par, ln in _feeders(prj):
        seen = next((e for e in par.feeders_seen if e.line == ln.number), None)
        if (seen and seen.p_inst is not None and child.result.p_inst is not None and abs(seen.p_inst - child.result.p_inst) > 1e-3) \
                or (ln.in_a is not None and child.supply.in_a is not None and abs(ln.in_a - child.supply.in_a) > 1e-3):
            n += 1
    if n:
        yield {"rule": "links.sync", "kind": "sync", "panel": "", "line": "",
               "text": f"связей с устаревшими данными: {n} — пересчитать связанные книги вместе и сохранить",
               "was": f"{n} связ.", "now": "синхронно", "checked": True}


def fix_nodes(prj):
    for p in prj.panels():
        for n in p.nodes:
            if n.in_a and n.i_calc and n.in_a < n.i_calc and (n.device or "авто") != "авто":
                yield {"rule": "nodes.overload", "kind": "structure", "panel": p.id, "line": n.id,
                       "op": {"op": "update_node", "panel": p.id, "id": n.id, "fields": {"device": "авто"}},
                       "text": f"узел {n.id} {n.designation}: In {f(n.in_a, 0)} А < Iр {f(n.i_calc)} А — аппарат подобрать по току",
                       "was": n.device, "now": "авто", "checked": True}


def _no_rcd(text):
    """«2P C16А 30мА» -> «1P C16А»: дифзащита уже у узла. Однофазная линия — 1P, трёхфазная — 3P."""
    t = re.sub(r"\s*\d+\s*мА", "", text or "", flags=re.I).strip()
    return re.sub(r"^\s*2P", "1P", re.sub(r"^\s*4P", "3P", t))


def fix_extra_rcd(prj):
    for p in prj.panels():
        for ln in p.lines:
            if not ln.rcd_node or not ("мА" in (ln.scheme.breaker or "") or (ln.scheme.qf or "").startswith("QFD")):
                continue
            if not (ln.scheme.group_head or not ln.scheme.qf):
                continue                        # автомат правится только у первой строки группы
            now = _no_rcd(ln.scheme.breaker)
            if now and now != ln.scheme.breaker:
                yield {"rule": "nodes.extra_rcd", "kind": "edit", "panel": p.id, "line": ln.number, "path": "scheme.breaker_set",
                       "text": f"под узлом {ln.rcd_node} (30 мА) свой дифавтомат не нужен — автомат без утечки",
                       "was": ln.scheme.breaker, "now": now, "checked": False}


def fix_manual(prj):
    for p in prj.panels():
        if any(d["state"] == "stale" for d in p.drawings):
            yield {"rule": "drawings.stale", "kind": "manual", "panel": p.id, "line": "",
                   "text": "однолинейка устарела — в чертеже SXUPDATE", "was": "устарела", "now": "SXUPDATE", "checked": False}
    for d in prj.documents:
        if d.kj.get("state") == "stale" and d.panels:
            yield {"rule": "drawings.kj_stale", "kind": "manual", "panel": d.panels[0].id, "line": "",
                   "text": f"КЖ устарел — в книге {d.file} «КЖ в DWG»", "was": "устарел", "now": "КЖ в DWG", "checked": False}


FIXERS = [fix_feeders, fix_sync, fix_nodes, fix_extra_rcd, fix_manual]


def propose(prj) -> list:
    out = []
    for fn in FIXERS:
        for i, x in enumerate(fn(prj)):
            x["id"] = f"{fn.__name__}:{x['panel']}:{x['line']}:{i}"
            p = prj.panel(x["panel"]) if x["panel"] else None
            x["panel_name"] = p.name if p else ""
            out.append(x)
    return out
