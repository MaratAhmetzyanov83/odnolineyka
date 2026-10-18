# -*- coding: utf-8 -*-
"""Нагрузка по фазам с учётом нижестоящих щитов.

Ток линии (Iр, без Кс) ложится на её фазы («Однолинейка»: L1 / L2,L3 / L1,L2,L3; трёхфазная — на все три).
Линия-фидер несёт не равные токи, а фазные токи щита, который она питает (по его вводу: линии его секции),
поэтому перекос нижних щитов виден в вышестоящих. Фазы фидера — как у щита (L1 -> L1).
Щит с вводами 2–3: фидер ввода k — линии секции k («Исходные данные» BB, пусто = 1), резерв по АВР не добавляется.
Перекос = (max − min) / max, как в окне. Строки вводов не считаются; линия без фазы — отдельно (unknown).
"""
from __future__ import annotations

import re

from ..core.model import Project

PH = ("L1", "L2", "L3")


def _is_input(ln):
    return ln.number.strip().lower().startswith("ввод")


def phases_of(ln):
    ph = [x for x in re.split(r"[,\s]+", ln.scheme.phase or "") if x in PH]
    if len(ph) >= 3 or (ln.input.phases or 1) >= 3:
        return list(PH)
    return ph


def skew(acc: dict):
    v = [acc.get(k, 0.0) for k in PH]
    return (max(v) - min(v)) / max(v) * 100 if max(v) > 0 else 0.0


def run(prj: Project):
    memo, busy = {}, set()

    def load(pid, section=None):
        """Фазные токи щита (или его секции — для ввода 2–3)."""
        key = (pid, section)
        if key in memo:
            return memo[key]
        p = prj.panel(pid)
        if p is None or key in busy:                 # цикл связей — не считать дважды
            return {k: 0.0 for k in PH}
        busy.add(key)
        acc = {k: 0.0 for k in PH}
        for ln in p.lines:
            if _is_input(ln):
                continue
            if section is not None and (ln.input.section or 1) != section:
                continue
            if ln.feeds and prj.panel(ln.feeds):
                child = prj.panel(ln.feeds)
                # у щита с вводами 2–3 фидер ввода k несёт линии секции k (обычный режим, без АВР)
                sec = (ln.feeds_input or 1) if child.inputs_more else None
                sub = load(child.id, sec)
                ln.phase_i = {k: round(v, 3) for k, v in sub.items()}
                for k in PH:
                    acc[k] += sub[k]
                continue
            i = ln.result.i_calc or 0.0
            for k in phases_of(ln):
                acc[k] += i
        busy.discard(key)
        memo[key] = acc
        return acc

    for p in prj.panels():
        acc = load(p.id)
        p.phase_i = {k: round(v, 3) for k, v in acc.items()}
        p.phase_skew = round(skew(acc), 1)
        own = {k: 0.0 for k in PH}
        for ln in p.lines:
            if not _is_input(ln) and not (ln.feeds and prj.panel(ln.feeds)):
                for k in phases_of(ln):
                    own[k] += ln.result.i_calc or 0.0
        p.phase_own = {k: round(v, 3) for k, v in own.items()}
        p.phase_unknown = sum(1 for ln in p.lines if not _is_input(ln) and not phases_of(ln) and ln.result.i_calc)
