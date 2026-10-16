# -*- coding: utf-8 -*-
"""Связи щитов: из записи «питается от» (Supply) -> parent / parent_line; у фидера — feeds."""
from __future__ import annotations

import os

from ..core.model import Project


def norm(path: str) -> str:
    return os.path.normcase(os.path.normpath(path.replace("\\", "/"))).lower()


def _abs(folder, base_doc, p):
    p = p.replace("\\", "/")
    if os.path.isabs(p) or p[1:3] == ":/":
        return norm(p)
    return norm(os.path.join(folder, os.path.dirname(base_doc), p))


def link(prj: Project) -> None:
    by_doc = {}
    for p in prj.panels():
        by_doc.setdefault(norm(os.path.join(prj.folder, p.doc)), []).append(p)
    for p in prj.panels():
        p.parent = p.parent_line = p.link_error = ""
        for ln in p.lines:
            ln.feeds = ""
    for p in prj.panels():
        s = p.supply
        if not s.from_doc:
            continue
        fpath, _, pname = (x.strip() for x in s.from_doc.partition(" / "))
        pname = pname or s.from_panel
        cands = by_doc.get(_abs(prj.folder, p.doc, fpath))
        if not cands:                                 # книга переехала — ищем по имени файла
            fname = os.path.basename(fpath.replace("\\", "/")).lower()
            cands = [q for k, v in by_doc.items() if os.path.basename(k) == fname for q in v]
        if not cands:
            p.link_error = f"вышестоящая книга не найдена в папке объекта: {fpath}"
            continue
        par = next((q for q in cands if q.name == pname), cands[0])
        p.parent = par.id
        me = norm(os.path.join(prj.folder, p.doc))
        # линия-фидер: 1) раздел Б вышестоящего; 2) фидер «Щит» с именем щита; 3) QF фидера
        p.parent_line = next((fs.line for fs in par.feeders_seen
                              if fs.doc and _abs(prj.folder, par.doc, fs.doc.partition(" / ")[0].strip()) == me), "")
        if not p.parent_line:
            p.parent_line = next((ln.number for ln in par.lines
                                  if ln.input.group_ks == "Щит" and p.name and p.name in ln.name), "")
        if not p.parent_line and s.qf:
            p.parent_line = next((ln.number for ln in par.lines if ln.scheme.qf == s.qf), "")
        ln = next((x for x in par.lines if x.number == p.parent_line), None)
        if ln:
            ln.feeds = p.id
    prj.reindex()
