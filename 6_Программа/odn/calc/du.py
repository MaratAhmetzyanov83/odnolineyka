# -*- coding: utf-8 -*-
"""Допустимая потеря напряжения с надбавкой за длину (ГОСТ Р 50571.5.52-2011, прил. G, примечание к табл. G.52.1):
для проводки длиннее 100 м допустимую ΔU можно увеличить на 0,005 % на каждый метр сверх 100 м, но не больше 0,5 %.

Длина — от ВРУ (начала установки) до потребителя: длина линии (L итог) + длины фидеров всех щитов выше по цепочке.
Линии до 50 В (свой источник, ΔUдоп 5 %) — без надбавки. ΔUдоп без надбавки — из книги (BF: 3 % свет / 4 %).

ВРУ здания (несколько зданий в объекте): щит-ВРУ — начало отсчёта для своих линий и всех щитов ниже: длина — от его шин,
ΔU — от его шин. Книга это умеет («Сеть_щита» C26 «да»: ΔU до шин = 0). Если ВРУ задан только в программе —
из ΔU книги вычитается ΔU до шин ВРУ (то, что книга прибавила от сети выше).
"""
from __future__ import annotations

from ..core.model import Project

PER_M, MAX_EXTRA, FREE_M = 0.005, 0.5, 100.0


def extra_for(length):
    if length is None or length <= FREE_M:
        return 0.0
    return round(min(MAX_EXTRA, (length - FREE_M) * PER_M), 4)


def upstream_len(prj: Project, panel, seen=None) -> float:
    """Сумма длин фидеров от ВРУ до шин щита (ВРУ здания — начало отсчёта)."""
    seen = seen or set()
    if not panel or panel.vru or not panel.parent or panel.id in seen:
        return 0.0
    seen.add(panel.id)
    par = prj.panel(panel.parent)
    ln = prj.line(panel.parent, panel.parent_line) if panel.parent_line else None
    own = (ln.input.length or 0.0) if ln else (panel.supply.length or 0.0)
    return own + upstream_len(prj, par, seen)


def apply_vru(prj: Project, flags: dict):
    for p in prj.panels():
        if p.id in flags:
            p.vru, p.vru_src = bool(flags[p.id]), "программа"
        else:
            p.vru, p.vru_src = p.scheme.is_vru == "да", ("книга" if p.scheme.is_vru == "да" else "")


def vru_of(prj: Project, panel):
    """Ближайший ВРУ здания: сам щит или выше по цепочке."""
    seen = set()
    while panel is not None and panel.id not in seen:
        if panel.vru:
            return panel
        seen.add(panel.id)
        panel = prj.panel(panel.parent) if panel.parent else None
    return None


def run(prj: Project, flags: dict | None = None):
    apply_vru(prj, flags or {})
    for p in prj.panels():
        up = upstream_len(prj, p)
        v = vru_of(prj, p)
        # книга не знает этот ВРУ — вычесть ΔU до его шин, которую книга прибавила от сети выше
        corr = (v.result.du_bus or 0.0) if v is not None and v.scheme.is_vru != "да" else 0.0
        for ln in p.lines:
            r = ln.result
            r.du_from = v.id if v is not None else ""
            r.du_sum_eff = round(r.du_sum - corr, 4) if r.du_sum is not None else None
        for ln in p.lines:
            r = ln.result
            low = ln.input.u_v is not None and 0 < ln.input.u_v <= 50
            if low:
                r.du_sum_eff = r.du_sum                       # до 50 В — от своего источника
            if ln.input.length is None or r.du_max is None or low:
                r.du_path_len, r.du_extra, r.du_max_eff = None, None, r.du_max
                continue
            r.du_path_len = round(ln.input.length + up, 1)
            r.du_extra = extra_for(r.du_path_len)
            r.du_max_eff = round(r.du_max + r.du_extra, 4)
