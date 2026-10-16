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


def apply_changes(prj: Project, raw: list, kind="правка", undo_of="") -> dict:
    """raw: [{panel, line, path, new}]. Пишет по книгам; ошибка в одной книге не мешает другим."""
    chs = []
    for r in raw:
        ch = C.Change(r["panel"], r["path"], r.get("new"), r.get("line", ""))
        C.apply(prj, ch)                              # тип, разрешено ли поле, old
        if ch.old != ch.new:
            chs.append(ch)
    by_doc = {}
    for ch in chs:
        by_doc.setdefault(prj.panel(ch.panel).doc, []).append(ch)
    stamp = time.strftime("%Y-%m-%d_%H%M%S")
    done, errors = [], []
    for docname, items in by_doc.items():
        doc = next(d for d in prj.documents if d.file == docname)
        path = os.path.join(prj.folder, docname)
        src = sources.base.by_kind(doc.kind)
        try:
            if abs(os.path.getmtime(path) - doc.mtime) > 1e-6:
                raise C.EditError(f"{docname} изменилась после чтения — нажмите «Обновить» и повторите")
            cells = src.cells_for(prj, items)
            bak = os.path.join(_dir(prj, "резерв", stamp), docname)
            os.makedirs(os.path.dirname(bak), exist_ok=True)
            shutil.copy2(path, bak)
            src.write(path, cells)
        except (C.EditError, NotImplementedError) as e:
            errors.append(str(e))
            continue
        entry = {"batch": f"{stamp}|{docname}", "at": time.strftime("%d.%m.%Y %H:%M:%S"), "user": _user(),
                 "kind": kind, "doc": docname, "backup": os.path.relpath(bak, prj.folder), "undo_of": undo_of,
                 "changes": [dict(dataclasses.asdict(ch), cell=f"{c[0]}!{c[1]}") for ch, c in zip(items, cells)]}
        with open(_journal(prj), "a", encoding="utf-8") as f:
            f.write(json.dumps(entry, ensure_ascii=False, default=str) + "\n")
        done.append(entry)
    return {"ok": not errors, "written": sum(len(e["changes"]) for e in done), "docs": [e["doc"] for e in done],
            "errors": errors, "skipped": len(raw) - len(chs)}


def undo_last(prj: Project) -> dict:
    """Отменить последнюю не отменённую запись (одну книгу)."""
    rows = [r for r in read_journal(prj, 10_000) if r.get("batch") and not r["undone"] and r.get("kind") != "отмена"]
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
