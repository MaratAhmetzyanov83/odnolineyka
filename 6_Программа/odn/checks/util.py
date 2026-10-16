# -*- coding: utf-8 -*-
"""Общее для правил: сравнение, формат чисел."""

TOL = 0.01           # 1 % — значения считаются одинаковыми


def ok_text(t: str) -> bool:
    t = (t or "").strip().lower()
    return t == "" or t.startswith(("ok", "ок", "—", "-", "да", "обеспечена"))


def differ(a, b, tol=TOL) -> bool:
    if a is None or b is None:
        return False
    return abs(a - b) > tol * max(abs(a), abs(b), 1e-9)


def f(x, d=1):
    if x is None:
        return "—"
    s = f"{x:.{d}f}".rstrip("0").rstrip(".") if d else f"{x:.0f}"
    return s.replace(".", ",")
