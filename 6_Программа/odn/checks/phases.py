# -*- coding: utf-8 -*-
"""Перекос фаз щита — с учётом нижестоящих щитов (calc.phases)."""
from .registry import Hit, rule
from .util import f

G = "Фазы"
SKEW_MAX = 15.0          # %, как подсветка в окне
I_MIN = 5.0              # А: малые токи не проверять (перекос 1 А — не перекос)


@rule("phases.skew", "Перекос фаз (с нижестоящими)", G)
def skew(prj):
    for p in prj.panels():
        v = p.phase_i or {}
        if max(v.values() or [0]) >= I_MIN and p.phase_skew > SKEW_MAX:
            fed = any(ln.phase_i for ln in p.lines)
            yield Hit(1, f"перекос фаз {f(p.phase_skew, 0)} %: " + " · ".join(f"{k} {f(v[k])} А" for k in ("L1", "L2", "L3"))
                      + (" (с нижестоящими щитами)" if fed else ""), p.id)


@rule("phases.unknown", "Линия без фазы", G)
def unknown(prj):
    for p in prj.panels():
        if p.phase_unknown:
            yield Hit(1, f"линий с током, но без фазы: {p.phase_unknown} — в расчёт перекоса не вошли", p.id)
