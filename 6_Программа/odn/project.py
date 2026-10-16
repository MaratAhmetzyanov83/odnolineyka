# -*- coding: utf-8 -*-
"""Объект = папка с книгами щитов. Чтение всех книг (с кэшем), связи щитов, замечания."""
from __future__ import annotations

import json
import os
import time
from dataclasses import asdict

from .checks import run_checks
from .model import FORMAT, Book, Issue, Line, Panel, Project, Supply
from .xlsx_import import read_book

SKIP_DIRS = ("архив", "резерв", "_резерв", "резервные копии")


def _skip_dir(name: str) -> bool:
    n = name.lower()
    return n.startswith((".", "_")) or any(n.startswith(s) for s in SKIP_DIRS)


def find_books(folder: str):
    """Книги .xlsx / .xlsm в папке объекта и подпапках (без архивов, резервов, временных ~$ и ПЗ_)."""
    out = []
    for root, dirs, files in os.walk(folder):
        dirs[:] = sorted(d for d in dirs if not _skip_dir(d))
        for f in sorted(files):
            if f.lower().endswith((".xlsx", ".xlsm")) and not f.startswith(("~$", "ПЗ_")):
                out.append(os.path.join(root, f))
    return out


def norm(path: str) -> str:
    return os.path.normcase(os.path.normpath(path.replace("\\", "/"))).lower()


# ---------------------------------------------------------------- кэш (книга не менялась — не читаем заново)
def _book_from_dict(d: dict) -> Book:
    b = Book(file=d["file"], mtime=d["mtime"], error=d.get("error", ""))
    for pd in d.get("panels", []):
        pd = dict(pd)
        sup = Supply(**pd.pop("supply"))
        lines = [Line(**ld) for ld in pd.pop("lines")]
        p = Panel(**pd, supply=sup)
        p.lines = lines
        b.panels.append(p)
    return b


def _cache_path(folder: str) -> str:
    base = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "кэш")
    os.makedirs(base, exist_ok=True)
    key = "".join(ch if ch.isalnum() else "_" for ch in os.path.abspath(folder))[-120:]
    return os.path.join(base, key + ".json")


def _load_cache(folder: str) -> dict:
    try:
        with open(_cache_path(folder), encoding="utf-8") as f:
            d = json.load(f)
        if d.get("format") != FORMAT:
            return {}
        return {b["file"]: b for b in d.get("books", [])}
    except (OSError, ValueError, KeyError):
        return {}


# ---------------------------------------------------------------- загрузка объекта
def load_project(folder: str, progress=print, use_cache=True) -> Project:
    folder = os.path.abspath(folder)
    prj = Project(name=os.path.basename(folder.rstrip("\\/")), folder=folder,
                  imported=time.strftime("%d.%m.%Y %H:%M"))
    cache = _load_cache(folder) if use_cache else {}
    paths = find_books(folder)
    for i, path in enumerate(paths, 1):
        rel = os.path.relpath(path, folder)
        c = cache.get(rel)
        if c and abs(c.get("mtime", 0) - os.path.getmtime(path)) < 1e-6 and not c.get("error"):
            try:
                prj.books.append(_book_from_dict(c))
                continue
            except TypeError:
                pass                       # кэш от другой версии модели — читаем заново
        progress(f"[{i}/{len(paths)}] {rel}")
        b = read_book(path, rel)
        if b.error == "не книга «Однолинейка»":
            continue
        prj.books.append(b)
    link_panels(prj)
    run_checks(prj)
    try:
        with open(_cache_path(folder), "w", encoding="utf-8") as f:
            json.dump({"format": FORMAT, "books": [asdict(b) for b in prj.books if not b.error]}, f, ensure_ascii=False)
    except OSError:
        pass
    return prj


def link_panels(prj: Project):
    """Связи: «Сеть_щита» G6 нижестоящего (файл вышестоящего [/ щит]) -> id вышестоящего; линия-фидер у него."""
    by_file = {}
    for p in prj.panels():
        by_file.setdefault(norm(os.path.join(prj.folder, p.file)), []).append(p)
    for p in prj.panels():
        s = p.supply
        if not s.from_file:
            continue
        part = s.from_file.split(" / ")
        fpath = part[0].strip()
        pname = (part[1].strip() if len(part) > 1 else "") or s.from_panel
        book_dir = os.path.dirname(os.path.join(prj.folder, p.file))
        target = norm(fpath if os.path.isabs(fpath) or fpath[1:3] == ":\\" else os.path.join(book_dir, fpath))
        cands = by_file.get(target)
        if not cands:                       # книга переехала — ищем по имени файла
            fname = os.path.basename(fpath.replace("\\", "/")).lower()
            cands = [q for k, v in by_file.items() if os.path.basename(k) == fname for q in v]
        if not cands:
            prj.issues.append(Issue(2, f"вышестоящая книга не найдена в папке объекта: {fpath}", p.id,
                                    source="объект", where="Сеть_щита!G6"))
            continue
        par = next((q for q in cands if q.name == pname), cands[0])
        p.parent = par.id
        # линия-фидер: раздел Б вышестоящего (файл нижестоящего), иначе — QF фидера / имя щита в наименовании
        me = norm(os.path.join(prj.folder, p.file))
        par_dir = os.path.dirname(os.path.join(prj.folder, par.file))
        for ent in par.feeders_seen:
            f = ent["file"].split(" / ")[0].strip()
            if f and norm(os.path.join(par_dir, f)) == me:
                p.parent_line = ent["line"]
                break
        if not p.parent_line:
            for ln in par.lines:
                if ln.group_ks == "Щит" and p.name and p.name in ln.name:
                    p.parent_line = ln.number
                    break
        if not p.parent_line and s.qf:
            p.parent_line = next((ln.number for ln in par.lines if ln.qf == s.qf), "")
        for ln in par.lines:
            if ln.number == p.parent_line:
                ln.feeds = p.id
