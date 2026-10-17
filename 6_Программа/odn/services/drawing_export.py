# -*- coding: utf-8 -*-
"""Узлы для чертежа: файл рядом с книгой «<книга>.nodes.txt» — его читает SXDRAW (v3.25+).

Только ASCII: строки — формы LISP, русские тексты — списки кодов символов (в LISP — sx:ru), так нет проблем с кодировкой.
    (NODE щит id вид обозн текст родитель уровень фаз)   вид: QF / RCD / RCBO / QS
    (LINE щит линия корень)                              линия под узлом верхнего уровня «корень»
Файл пишется, только если содержимое изменилось; книга без узлов — свой старый файл удаляется.
"""
from __future__ import annotations

import os

from ..core.model import Project

HEADER = ";; odn nodes v1 - program Objekt (6_Programma). Do not edit: rewritten on every change."
KIND = {"УЗО": "RCD", "АВДТ": "RCBO", "Автомат": "QF", "Рубильник": "QS"}


def codes(s: str) -> str:
    return "(" + " ".join(str(ord(c)) for c in (s or "")) + ")"


def _root(n, ids):
    seen = set()
    while n.parent in ids and n.id not in seen:
        seen.add(n.id)
        n = ids[n.parent]
    return n


def content(doc) -> str:
    rows = []
    for p in doc.panels:
        if not p.nodes:
            continue
        ids = {n.id: n for n in p.nodes}
        for n in p.nodes:
            rows.append(f'(NODE {codes(p.name)} {codes(n.id)} "{KIND.get(n.kind, "QF")}" {codes(n.designation)} '
                        f'{codes(n.text)} {codes(n.parent if n.parent in ids else "")} {n.level} {n.phases})')
        for ln in p.lines:
            if ln.node in ids:
                rows.append(f"(LINE {codes(p.name)} {codes(ln.number)} {codes(_root(ids[ln.node], ids).id)})")
    return (HEADER + "\n" + "\n".join(rows) + "\n") if rows else ""


def path_for(prj: Project, doc) -> str:
    base, _ = os.path.splitext(os.path.join(prj.folder, doc.file))
    return base + ".nodes.txt"


def export(prj: Project) -> list:
    """Записать файлы узлов книг объекта; вернуть список записанных."""
    done = []
    for d in prj.documents:
        if d.error:
            continue
        path, text = path_for(prj, d), content(d)
        try:
            old = ""
            if os.path.isfile(path):
                with open(path, encoding="ascii") as f:
                    old = f.read()
        except (OSError, UnicodeDecodeError):
            old = None
        if text == old:
            continue
        if not text:
            if old and old.startswith(";; odn nodes"):
                os.remove(path)
            continue
        with open(path, "w", encoding="ascii", newline="\r\n") as f:
            f.write(text)
        done.append(path)
    return done
