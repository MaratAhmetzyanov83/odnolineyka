# -*- coding: utf-8 -*-
"""Фидер в двух книгах: строка линии вышестоящего щита и строка ввода нижестоящего («Ввод…» в «Исходных данных»).

Если у нижестоящего есть строка ввода, кабель в обеих книгах должен быть один: марка, жилы, сечение, длина,
прокладка (лоток, гофра ПВХ, гофра ПНД). Расхождение — Line.mismatch обеих строк {поле: значение на той стороне};
окно подсвечивает ячейки, проверка links.cable_mismatch пишет замечание.
"""
from __future__ import annotations

from ..core.model import Project

FIELDS = [  # (поле, подпись, как взять)
    ("input.cable_code", "марка", lambda l: (l.input.cable_code or "").strip().lower()),
    ("input.cores", "жил", lambda l: l.input.cores),
    ("result.s_mm2", "сечение", lambda l: l.result.s_mm2),
    ("input.length", "L", lambda l: l.input.length),
    ("input.tray", "лоток", lambda l: l.input.tray or 0),
    ("input.pvc", "гофра ПВХ", lambda l: l.input.pvc or 0),
    ("input.pnd", "гофра ПНД", lambda l: l.input.pnd or 0),
]
LABEL = {k: t for k, t, _ in FIELDS}


def _is_input(ln):
    return ln.number.strip().lower().startswith("ввод")


def input_line(child, k):
    """Строка ввода k щита: «Ввод 2» / «Ввод2» для k > 1; для ввода 1 — «Ввод», «Ввод 1» или единственная строка ввода."""
    rows = [ln for ln in child.lines if _is_input(ln)]
    digit = lambda ln: "".join(c for c in ln.number if c.isdigit())  # noqa: E731
    return next((ln for ln in rows if digit(ln) == str(k)), None) or \
        (next((ln for ln in rows if digit(ln) in ("", "1")), None) if k == 1 else None)


def _same(a, b):
    if a in (None, "") or b in (None, ""):
        return True                              # пусто на одной стороне — нечего сверять
    if isinstance(a, (int, float)) and isinstance(b, (int, float)):
        return abs(a - b) <= 0.5
    return a == b


def run(prj: Project):
    for p in prj.panels():
        for ln in p.lines:
            ln.mismatch = {}
    for par in prj.panels():
        for ln in par.lines:
            child = prj.panel(ln.feeds) if ln.feeds else None
            inp = input_line(child, ln.feeds_input or 1) if child else None
            if inp is None:
                continue
            for k, _, get in FIELDS:
                a, b = get(ln), get(inp)
                if not _same(a, b):
                    ln.mismatch[k] = b
                    inp.mismatch[k] = a
