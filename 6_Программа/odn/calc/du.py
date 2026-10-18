# -*- coding: utf-8 -*-
"""Допустимая потеря напряжения с надбавкой за длину (ГОСТ Р 50571.5.52-2011, прил. G, примечание к табл. G.52.1):
для проводки длиннее 100 м допустимую ΔU можно увеличить на 0,005 % на каждый метр сверх 100 м, но не больше 0,5 %.

Длина — от ВРУ (начала установки) до потребителя: длина линии (L итог) + длины фидеров всех щитов выше по цепочке.
Линии до 50 В (свой источник, ΔUдоп 5 %) — без надбавки. ΔUдоп без надбавки — из книги (BF: 3 % свет / 4 %).
"""
from __future__ import annotations

from ..core.model import Project

PER_M, MAX_EXTRA, FREE_M = 0.005, 0.5, 100.0


def extra_for(length):
    if length is None or length <= FREE_M:
        return 0.0
    return round(min(MAX_EXTRA, (length - FREE_M) * PER_M), 4)


def upstream_len(prj: Project, panel, seen=None) -> float:
    """Сумма длин фидеров от ВРУ до шин щита."""
    seen = seen or set()
    if not panel or not panel.parent or panel.id in seen:
        return 0.0
    seen.add(panel.id)
    par = prj.panel(panel.parent)
    ln = prj.line(panel.parent, panel.parent_line) if panel.parent_line else None
    own = (ln.input.length or 0.0) if ln else (panel.supply.length or 0.0)
    return own + upstream_len(prj, par, seen)


def run(prj: Project):
    for p in prj.panels():
        up = upstream_len(prj, p)
        for ln in p.lines:
            r = ln.result
            low = ln.input.u_v is not None and 0 < ln.input.u_v <= 50
            if ln.input.length is None or r.du_max is None or low:
                r.du_path_len, r.du_extra, r.du_max_eff = None, None, r.du_max
                continue
            r.du_path_len = round(ln.input.length + up, 1)
            r.du_extra = extra_for(r.du_path_len)
            r.du_max_eff = round(r.du_max + r.du_extra, 4)
