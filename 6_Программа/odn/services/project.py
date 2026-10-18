# -*- coding: utf-8 -*-
"""Загрузка объекта — конвейер шагов. Новый шаг (расчёт, отметки DWG…) — функция (prj, ctx) в STEPS.

    read      — документы папки (адаптеры odn.sources, кэш по дате файла)
    link      — связи щитов
    calc      — расчёт в Python (шаг 3 плана; пока пусто — результаты из Excel)
    revision  — отпечатки щитов (для актуальности DWG / КЖ)
    checks    — замечания (odn.checks)
"""
from __future__ import annotations

import os
import time

from .. import checks, sources
from ..core.model import Project
from ..core.revision import panel_revision
from . import cache, linking, structure

SKIP_DIRS = ("архив", "резерв", "резервные копии")
SKIP_WORDS = ("резерв", "копия", "backup", "~$")                # в имени файла — резервная копия, не читаем
EXCLUDE_FILE = os.path.join("_окно", "исключить.txt")           # свой список объекта: по строке — путь или маска


def _skip_dir(name: str) -> bool:
    n = name.lower()
    return n.startswith((".", "_")) or n.startswith(SKIP_DIRS)


def exclusions(folder: str) -> list:
    try:
        with open(os.path.join(folder, EXCLUDE_FILE), encoding="utf-8") as f:
            return [x.strip().replace("\\", "/") for x in f if x.strip() and not x.startswith("#")]
    except OSError:
        return []


def add_exclusion(folder: str, rel: str):
    os.makedirs(os.path.join(folder, "_окно"), exist_ok=True)
    path = os.path.join(folder, EXCLUDE_FILE)
    new = not os.path.isfile(path)
    with open(path, "a", encoding="utf-8") as f:
        if new:
            f.write("# Книги, которые программа не читает (путь от папки объекта или маска *, ?). Правьте свободно.\n")
        f.write(rel.replace("\\", "/") + "\n")


def skipped(rel: str, masks) -> bool:
    import fnmatch
    r = rel.replace("\\", "/")
    name = os.path.basename(r).lower()
    if any(w in name for w in SKIP_WORDS):
        return True
    return any(fnmatch.fnmatch(r.lower(), m.lower()) or fnmatch.fnmatch(name, m.lower()) for m in masks)


def find_files(folder: str):
    """Файлы папки объекта, которые может прочитать какой-нибудь источник (без архивов, резервов и исключённых)."""
    masks = exclusions(folder)
    for root, dirs, files in os.walk(folder):
        dirs[:] = sorted(d for d in dirs if not _skip_dir(d))
        for f in sorted(files):
            path = os.path.join(root, f)
            if skipped(os.path.relpath(path, folder), masks):
                continue
            src = sources.source_for(path)
            if src:
                yield path, src


# ---------------------------------------------------------------- шаги
def step_read(prj: Project, ctx):
    old = cache.load(prj.folder) if ctx.get("cache", True) else {}
    paths = list(find_files(prj.folder))
    for i, (path, src) in enumerate(paths, 1):
        rel = os.path.relpath(path, prj.folder)
        c = old.get(rel)
        if c and c.kind == src.kind and abs(c.mtime - os.path.getmtime(path)) < 1e-6:
            prj.documents.append(c)
            continue
        ctx["progress"](f"[{i}/{len(paths)}] {rel}")
        d = src.read(path, rel)
        if d.error == "не книга «Однолинейка»":
            continue                                # чужой Excel в папке объекта — молча мимо
        prj.documents.append(d)
    prj.reindex()
    cache.save(prj.folder, prj.documents)


def step_structure(prj, ctx):
    structure.apply(prj)                            # узлы каскада — данные программы (_окно/структура.json)


def step_link(prj, ctx):
    linking.link(prj)


def step_calc(prj, ctx):
    from ..calc import nodes
    nodes.run(prj)                                  # узлы каскада; остальной расчёт — пока из Excel


def step_drawing(prj, ctx):
    from . import drawing_export
    try:
        drawing_export.export(prj)                  # узлы для SXDRAW: «<книга>.nodes.txt»
    except OSError as e:
        ctx["progress"](f"файл узлов для чертежа не записан: {e}")


def step_revision(prj, ctx):
    for p in prj.panels():
        p.revision = panel_revision(p)


def step_checks(prj, ctx):
    prj.issues = checks.run(prj, ctx.get("disabled", ()))


STEPS = [("read", step_read), ("structure", step_structure), ("link", step_link), ("calc", step_calc), ("drawing", step_drawing),
         ("revision", step_revision), ("checks", step_checks)]


def load_project(folder: str, progress=print, use_cache=True, disabled=()) -> Project:
    folder = os.path.abspath(folder)
    prj = Project(name=os.path.basename(folder.rstrip("\\/")), folder=folder, loaded=time.strftime("%d.%m.%Y %H:%M"))
    ctx = {"progress": progress, "cache": use_cache, "disabled": set(disabled)}
    for _, fn in STEPS:
        fn(prj, ctx)
    return prj


def recheck(prj: Project, disabled=()) -> Project:
    """После правок в памяти: связи, расчёт, отпечатки, проверки — без чтения файлов."""
    ctx = {"progress": lambda s: None, "disabled": set(disabled)}
    prj.reindex()
    for name, fn in STEPS:
        if name != "read":
            fn(prj, ctx)
    return prj


def open_at(prj: Project, ref) -> str:
    """Открыть место в источнике (книга на ячейке)."""
    doc = next((d for d in prj.documents if d.file == ref.file), None)
    if not doc:
        return "нет такого документа в объекте"
    src = sources.base.by_kind(doc.kind)
    return src.open_at(os.path.join(prj.folder, doc.file), ref)
