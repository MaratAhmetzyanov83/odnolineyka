# -*- coding: utf-8 -*-
"""Правки из окна: проверка -> резерв -> запись в книги -> журнал. Отмена — обратные правки последней записи.

Всё служебное лежит в папке объекта, в «_окно» (программа такие папки не читает):
  _окно/журнал.jsonl          — одна строка = одна запись (пакет правок одной книги)
  _окно/резерв/<время>/<книга> — копия книги до записи
"""
from __future__ import annotations

import dataclasses
import getpass
import json
import os
import shutil
import time
import uuid

from .. import sources
from ..core import changes as C
from ..core.model import Project

SERVICE_DIR = "_окно"


def _dir(prj, *parts):
    d = os.path.join(prj.folder, SERVICE_DIR, *parts)
    os.makedirs(d, exist_ok=True)
    return d


def _journal(prj):
    return os.path.join(_dir(prj), "журнал.jsonl")


def _user():
    try:
        return getpass.getuser()
    except Exception:  # noqa: BLE001
        return ""


def read_journal(prj: Project, limit=200):
    try:
        with open(_journal(prj), encoding="utf-8") as f:
            rows = [json.loads(x) for x in f if x.strip()]
    except OSError:
        return []
    undone = {r["undo"] for r in rows if r.get("undo")}
    for r in rows:
        r["undone"] = r.get("batch") in undone
    return rows[-limit:]


def related_docs(prj: Project, docs) -> list:
    """Книги, связанные с docs формулами Excel (связь «Связать с вышестоящим»): нижестоящие получают питание,
    вышестоящие — нагрузку. Их надо пересчитать вместе с изменённой книгой."""
    docs, rel = set(docs), set()
    for p in prj.panels():
        bp = prj.panel(p.book_parent) if p.book_parent else None
        if not bp or bp.doc == p.doc:
            continue
        if bp.doc in docs and p.doc not in docs:
            rel.add(p.doc)
        if p.doc in docs and bp.doc not in docs:
            rel.add(bp.doc)
    return sorted(rel)


def _fresh(prj, docname):
    doc = next(d for d in prj.documents if d.file == docname)
    return abs(os.path.getmtime(os.path.join(prj.folder, docname)) - doc.mtime) <= 1e-6


def apply_changes(prj: Project, raw: list, kind="правка", undo_of="") -> dict:
    """raw: [{panel, line, path, new}]. Пишет по книгам одной сессией Excel и пересчитывает связанные книги
    (вышестоящую и нижестоящие), чтобы правка была видна с обеих сторон связи."""
    from ..sources.excel_v312.writer import is_locked
    chs = []
    for r in raw:
        ch = C.Change(r["panel"], r["path"], r.get("new"), r.get("line", ""))
        C.apply(prj, ch)                              # тип, разрешено ли поле, old
        if ch.old != ch.new:
            chs.append(ch)
    by_doc = {}
    for ch in chs:
        by_doc.setdefault(prj.panel(ch.panel).doc, []).append(ch)
    errors, warnings, cells_by_doc = [], [], {}
    for docname, items in by_doc.items():
        doc = next(d for d in prj.documents if d.file == docname)
        src = sources.base.by_kind(doc.kind)
        try:
            if not _fresh(prj, docname):
                raise C.EditError(f"{docname} изменилась после чтения — нажмите «Обновить» и повторите")
            if is_locked(os.path.join(prj.folder, docname)):
                raise C.EditError(f"{docname} открыта в Excel — закройте книгу и повторите")
            cells_by_doc[docname] = (items, src.cells_for(prj, items))
        except (C.EditError, NotImplementedError) as e:
            errors.append(str(e))
    related = []
    for d in related_docs(prj, cells_by_doc):
        if is_locked(os.path.join(prj.folder, d)):
            warnings.append(f"{d} открыта в Excel — не пересчитана, связь обновится, когда её откроете")
        elif not _fresh(prj, d):
            warnings.append(f"{d} изменилась после чтения — не пересчитана («Обновить», затем «Синхронизировать книги»)")
        else:
            related.append(d)
    if not cells_by_doc:
        return {"ok": not errors, "written": 0, "docs": [], "related": [], "errors": errors, "warnings": warnings,
                "skipped": len(raw) - len(chs)}
    stamp = time.strftime("%Y-%m-%d_%H%M%S")
    backups = {}
    for d in list(cells_by_doc) + related:
        bak = os.path.join(_dir(prj, "резерв", stamp), d)
        os.makedirs(os.path.dirname(bak), exist_ok=True)
        shutil.copy2(os.path.join(prj.folder, d), bak)
        backups[d] = os.path.relpath(bak, prj.folder)
    from ..sources.excel_v312.writer import write_batch
    try:
        write_batch({os.path.join(prj.folder, d): c for d, (_, c) in cells_by_doc.items()},
                    [os.path.join(prj.folder, d) for d in related])
    except (C.EditError, Exception) as e:  # noqa: BLE001 — Excel: книги вернуть из копий
        for d, bak in backups.items():
            shutil.copy2(os.path.join(prj.folder, bak), os.path.join(prj.folder, d))
        return {"ok": False, "written": 0, "docs": [], "related": [], "errors": errors + [str(e)], "warnings": warnings,
                "skipped": len(raw) - len(chs)}
    done = []
    for docname, (items, cells) in cells_by_doc.items():
        entry = {"batch": f"{stamp}|{docname}|{uuid.uuid4().hex[:8]}", "at": time.strftime("%d.%m.%Y %H:%M:%S"),
                 "user": _user(), "kind": kind, "doc": docname, "backup": backups[docname], "undo_of": undo_of,
                 "related": {d: backups[d] for d in related},
                 "changes": [dict(dataclasses.asdict(ch), cell=f"{c[0]}!{c[1]}") for ch, c in zip(items, cells)]}
        with open(_journal(prj), "a", encoding="utf-8") as f:
            f.write(json.dumps(entry, ensure_ascii=False, default=str) + "\n")
        done.append(entry)
    return {"ok": not errors, "written": sum(len(e["changes"]) for e in done), "docs": [e["doc"] for e in done],
            "related": related, "errors": errors, "warnings": warnings, "skipped": len(raw) - len(chs)}


def sync_books(prj: Project) -> dict:
    """«Синхронизировать книги»: все книги со связями — в одну копию Excel, пересчёт, сохранение (копии — до)."""
    from ..sources.excel_v312.writer import is_locked, write_batch
    docs = set()
    for p in prj.panels():
        if p.book_parent and prj.panel(p.book_parent) and prj.panel(p.book_parent).doc != p.doc:
            docs |= {p.doc, prj.panel(p.book_parent).doc}
    ok, skipped = [], []
    for d in sorted(docs):
        if is_locked(os.path.join(prj.folder, d)):
            skipped.append(f"{d}: открыта в Excel")
        elif not _fresh(prj, d):
            skipped.append(f"{d}: изменилась после чтения — «Обновить»")
        else:
            ok.append(d)
    if not ok:
        return {"ok": False, "errors": ["нет книг со связями для пересчёта"] + skipped}
    stamp = time.strftime("%Y-%m-%d_%H%M%S")
    backups = {}
    for d in ok:
        bak = os.path.join(_dir(prj, "резерв", stamp), d)
        os.makedirs(os.path.dirname(bak), exist_ok=True)
        shutil.copy2(os.path.join(prj.folder, d), bak)
        backups[d] = os.path.relpath(bak, prj.folder)
    try:
        write_batch({}, [os.path.join(prj.folder, d) for d in ok])
    except Exception as e:  # noqa: BLE001
        for d, bak in backups.items():
            shutil.copy2(os.path.join(prj.folder, bak), os.path.join(prj.folder, d))
        return {"ok": False, "errors": [str(e)]}
    entry = {"batch": f"{stamp}|синхронизация|{uuid.uuid4().hex[:8]}", "at": time.strftime("%d.%m.%Y %H:%M:%S"),
             "user": _user(), "kind": "синхронизация", "changes": [], "related": backups,
             "what": f"синхронизация связей: пересчитаны и сохранены {len(ok)} книг"}
    with open(_journal(prj), "a", encoding="utf-8") as f:
        f.write(json.dumps(entry, ensure_ascii=False) + "\n")
    return {"ok": True, "docs": ok, "warnings": skipped}


def undo_last(prj: Project) -> dict:
    """Отменить последнюю не отменённую запись (одну книгу)."""
    rows = [r for r in read_journal(prj, 10_000)
            if r.get("batch") and not r["undone"] and r.get("kind") not in ("отмена", "структура", "связь", "синхронизация")]
    if not rows:
        return {"ok": False, "errors": ["нечего отменять"]}
    last = rows[-1]
    raw = [{"panel": c["panel"], "line": c["line"], "path": c["path"], "new": c["old"]} for c in last["changes"]]
    res = apply_changes(prj, raw, kind="отмена", undo_of=last["batch"])
    if res["docs"]:
        with open(_journal(prj), "a", encoding="utf-8") as f:
            f.write(json.dumps({"undo": last["batch"], "at": time.strftime("%d.%m.%Y %H:%M:%S")}, ensure_ascii=False) + "\n")
    res["undone"] = last["batch"]
    return res


def options(prj: Project) -> dict:
    """Списки для выбора в ячейке: стандартные значения + то, что уже встречается в объекте."""
    seen = {p: set() for p in C.FIELDS}
    for pn in prj.panels():
        for ln in pn.lines:
            seen["input.cable_code"].add(ln.input.cable_code)
            seen["scheme.breaker_set"].add(ln.scheme.breaker_set)
            seen["scheme.bus"].add(ln.scheme.bus)
        seen["supply.manual"].add(pn.supply.manual)
    out = {}
    for path, spec in C.FIELDS.items():
        vals = [str(v) for v in spec.get("options", [])]
        vals += sorted(str(v) for v in seen[path] if v not in (None, "") and str(v) not in vals)
        out[path] = {"title": spec["title"], "type": spec["type"], "options": vals, "when": spec.get("when", "")}
    return out
