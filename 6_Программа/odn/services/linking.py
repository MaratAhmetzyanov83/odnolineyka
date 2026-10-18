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


def suggest_feeder(par, child):
    """Линия-фидер вышестоящего к щиту child: номер линии = имя щита; фидер «Щит» с именем щита; имя щита в наименовании."""
    name = (child.name or "").strip().lower()
    if not name:
        return ""
    for test in (lambda ln: ln.number.strip().lower() == name,
                 lambda ln: ln.input.group_ks == "Щит" and name in ln.name.lower(),
                 lambda ln: name in ln.name.lower()):
        ln = next((x for x in par.lines if test(x)), None)
        if ln:
            return ln.number
    return ""


def _book_link(prj, p, by_doc):
    """Связь, как она записана в книге щита: (id вышестоящего, линия) или ошибка."""
    s = p.supply
    fpath, _, pname = (x.strip() for x in s.from_doc.partition(" / "))
    pname = pname or s.from_panel
    cands = by_doc.get(_abs(prj.folder, p.doc, fpath))
    if not cands:                                 # книга переехала — ищем по имени файла
        fname = os.path.basename(fpath.replace("\\", "/")).lower()
        cands = [q for k, v in by_doc.items() if os.path.basename(k) == fname for q in v]
    if not cands:
        return "", "", f"вышестоящая книга не найдена в папке объекта: {fpath}"
    # после « / » надстройка пишет номер линии-фидера (старые книги — имя щита)
    by_line = next((q for q in cands if pname and any(ln.number == pname for ln in q.lines)), None)
    par = by_line or next((q for q in cands if q.name == pname), cands[0])
    me = norm(os.path.join(prj.folder, p.doc))
    numbers = {ln.number for ln in par.lines}
    # линия-фидер: 0) из G6; 1) раздел Б вышестоящего; 2) по имени щита; 3) QF фидера
    line = pname if by_line else ""
    line = line or next((fs.line for fs in par.feeders_seen if fs.line in numbers and fs.doc
                         and _abs(prj.folder, par.doc, fs.doc.partition(" / ")[0].strip()) == me), "")
    line = line or suggest_feeder(par, p)
    if not line and s.qf:
        line = next((ln.number for ln in par.lines if ln.scheme.qf == s.qf), "")
    return par.id, line, ""


def link(prj: Project, overrides: dict | None = None) -> None:
    """Связи щитов: из книг; поверх — связи программы (overrides: {id щита: {parent, line, input}})."""
    overrides = overrides or {}
    by_doc = {}
    for p in prj.panels():
        by_doc.setdefault(norm(os.path.join(prj.folder, p.doc)), []).append(p)
    ids = {p.id for p in prj.panels()}
    for p in prj.panels():
        p.parent = p.parent_line = p.link_error = p.book_parent = p.book_line = p.link_src = ""
        for ln in p.lines:
            ln.feeds = ""
    for p in prj.panels():
        if p.supply.from_doc:
            p.book_parent, p.book_line, p.link_error = _book_link(prj, p, by_doc)
        ov = overrides.get(p.id)
        if ov and ov.get("parent") in ids:
            p.parent, p.parent_line, p.link_src = ov["parent"], ov.get("line", ""), "программа"
            p.link_error = ""
        elif p.book_parent:
            p.parent, p.parent_line, p.link_src = p.book_parent, p.book_line, "книга"
        if p.parent:
            ln = next((x for x in prj.panel_raw(p.parent).lines if x.number == p.parent_line), None)
            if ln:
                ln.feeds = p.id
    prj.reindex()
