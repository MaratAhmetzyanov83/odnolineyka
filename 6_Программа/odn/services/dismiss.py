# -*- coding: utf-8 -*-
"""Замечания, снятые вручную (пользователь проверил и согласен): _окно/снятые.json — данные программы, книги не меняются.

Снятие — по правилу, щиту, линии и тексту замечания: текст поменялся (другое значение, новая причина) —
замечание снова в списке. Снятые не идут в счёт ошибок и в цвет строки; окно показывает их отдельно и может вернуть.
"""
from __future__ import annotations

import json
import os
import time

from ..core.changes import EditError
from ..core.model import Project

FILE = "снятые.json"


def _path(prj: Project) -> str:
    from .edit import SERVICE_DIR
    return os.path.join(prj.folder, SERVICE_DIR, FILE)


def key(i) -> tuple:
    g = i.get if isinstance(i, dict) else lambda k: getattr(i, k)
    return (g("rule") or "", g("panel") or "", str(g("line") or ""), g("text") or "")


def load(prj: Project) -> list:
    try:
        with open(_path(prj), encoding="utf-8") as f:
            rows = json.load(f)
    except (OSError, ValueError):
        return []
    return [r for r in rows if isinstance(r, dict)]


def _save(prj: Project, rows: list):
    from .edit import _dir
    _dir(prj)
    tmp = _path(prj) + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(rows, f, ensure_ascii=False, indent=1)
    os.replace(tmp, _path(prj))


def apply(prj: Project):
    """prj.issues -> активные; снятые — в prj.dismissed с отметкой «когда, кто: причина»."""
    rows = {key(r): r for r in load(prj)}
    prj.dismissed = []
    if not rows:
        return
    keep = []
    for i in prj.issues:
        r = rows.get(key(i))
        if r is None:
            keep.append(i)
        else:
            i.dismissed = f"{r.get('at', '')} {r.get('user', '')}".strip() + (f": {r['note']}" if r.get("note") else "")
            prj.dismissed.append(i)
    prj.issues = keep


def dismiss(prj: Project, items: list, note: str = "") -> dict:
    """items: [{rule, panel, line, text}] — только из текущих замечаний."""
    from .edit import _user
    cur = {key(i): i for i in prj.issues}
    rows = load(prj)
    have = {key(r) for r in rows}
    n = 0
    for it in items:
        k = key(it)
        i = cur.get(k)
        if i is None:
            raise EditError("замечание уже изменилось — нажмите «Обновить» и повторите")
        if k in have:
            continue
        p = prj.panel(i.panel) if i.panel else None
        rows.append({"rule": k[0], "panel": k[1], "line": k[2], "text": k[3], "panel_name": p.name if p else "",
                     "note": (note or "").strip(), "at": time.strftime("%d.%m.%Y %H:%M"), "user": _user()})
        have.add(k)
        n += 1
    _save(prj, rows)
    return {"ok": True, "n": n}


def restore(prj: Project, items: list) -> dict:
    drop = {key(it) for it in items}
    rows = load(prj)
    left = [r for r in rows if key(r) not in drop]
    _save(prj, left)
    return {"ok": True, "n": len(rows) - len(left)}
