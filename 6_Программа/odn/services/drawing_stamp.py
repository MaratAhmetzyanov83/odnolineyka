# -*- coding: utf-8 -*-
"""Актуальность однолинейных схем (DWG): по отметкам построения, которые пишет SXDRAW (v3.26+).

Отметка — файл рядом с книгой «<книга>.drawn.txt», строка на (щит, чертёж) — последнее полное построение
(SXDRAW / SXUPDATE; частичное SXUPDSEL отметку не меняет). Только ASCII, формы LISP, тексты — коды символов:
    (DRAWN щит dwg "время книги" "отпечаток узлов" "время построения" сохранена)
        время книги — дата изменения файла книги в момент построения (ГГГГ-ММ-ДД ЧЧ:ММ:СС);
        отпечаток узлов — строка REV щита из «<книга>.nodes.txt» («-» — узлов нет);
        сохранена — T, если книга в Excel была сохранена (чертёж — из того, что на диске).

Сверка. Программа при каждом чтении запоминает отпечаток щита (core.revision) для даты изменения книги:
«_окно/ревизии.json». По времени книги из отметки находится отпечаток, по которому строили:
    тот же, что сейчас → схема актуальна (книгу меняли, но не то, что на схеме);
    другой → схема устарела;
    не найден (книгу меняли, а программа ту версию не читала) → «проверить».
"""
from __future__ import annotations

import json
import os
import re
import time

from ..core.model import Project
from . import drawing_export

EXT = ".drawn.txt"
HISTORY = "ревизии.json"
KEEP = 60                        # версий книги в истории


def stamp_path(prj: Project, doc) -> str:
    base, _ = os.path.splitext(os.path.join(prj.folder, doc.file))
    return base + EXT


def book_time(mtime: float) -> str:
    return time.strftime("%Y-%m-%d %H:%M:%S", time.localtime(mtime))


# ---------------------------------------------------------------- чтение отметок
_TOK = re.compile(r'\s*(?:(\()|(\))|"((?:[^"\\]|\\.)*)"|([^\s()"]+))')


def _parse(s):
    """Форма LISP -> списки Python (строки, числа, символы-строки)."""
    stack, cur, pos = [], None, 0
    while pos < len(s):
        m = _TOK.match(s, pos)
        if not m or m.end() == pos:
            break
        pos = m.end()
        if m.group(1):
            stack.append([])
        elif m.group(2):
            if not stack:
                return None
            cur = stack.pop()
            if not stack:
                return cur
            stack[-1].append(cur)
        elif stack:
            if m.group(3) is not None:
                stack[-1].append(m.group(3).replace('\\"', '"').replace("\\\\", "\\"))
            else:
                t = m.group(4)
                try:
                    stack[-1].append(int(t))
                except ValueError:
                    stack[-1].append(t)
    return None


def text(codes) -> str:
    """Коды символов из LISP: Unicode (LISPSYS 1+) или байты cp1251 (старый AutoLISP)."""
    if not isinstance(codes, list):
        return str(codes or "")
    codes = [c for c in codes if isinstance(c, int)]
    if any(c > 255 for c in codes):
        return "".join(chr(c) for c in codes)
    return bytes(codes).decode("cp1251", errors="replace")


def read_stamps(path) -> list:
    out = []
    try:
        with open(path, encoding="ascii", errors="replace") as f:
            rows = f.read().splitlines()
    except OSError:
        return out
    for ln in rows:
        x = _parse(ln) if ln.startswith("(") else None
        if not x or x[0] != "DRAWN" or len(x) < 7:
            continue
        out.append({"panel": text(x[1]), "dwg": text(x[2]), "book_time": str(x[3]), "nodes_rev": str(x[4]),
                    "at": str(x[5]), "saved": x[6] == "T"})
    return out


# ---------------------------------------------------------------- история отпечатков
def _hist_path(prj):
    return os.path.join(prj.folder, "_окно", HISTORY)


def load_history(prj) -> dict:
    try:
        with open(_hist_path(prj), encoding="utf-8") as f:
            return json.load(f)
    except (OSError, ValueError):
        return {}


def record_history(prj: Project, hist: dict) -> dict:
    """Запомнить отпечатки щитов для текущей даты изменения каждой книги."""
    changed = False
    for d in prj.documents:
        if d.error or not d.mtime:
            continue
        h = hist.setdefault(d.file, {})
        key, now = book_time(d.mtime), {p.name: p.revision for p in d.panels}
        if h.get(key) != now:
            h[key] = now
            changed = True
        if len(h) > KEEP:
            for k in sorted(h)[:-KEEP]:
                del h[k]
    if changed:
        try:
            os.makedirs(os.path.dirname(_hist_path(prj)), exist_ok=True)
            tmp = _hist_path(prj) + ".tmp"
            with open(tmp, "w", encoding="utf-8") as f:
                json.dump(hist, f, ensure_ascii=False, indent=0)
            os.replace(tmp, _hist_path(prj))
        except OSError:
            pass
    return hist


# ---------------------------------------------------------------- сверка
def state_of(st: dict, p, doc, hist: dict) -> tuple:
    """(состояние, почему): ok / stale / check."""
    why = []
    state = "ok"
    now_time = book_time(doc.mtime)
    if st["book_time"] != now_time:
        old = hist.get(doc.file, {}).get(st["book_time"], {}).get(p.name)
        if old is None:
            state, why = "check", [f"книга изменена после построения ({now_time}), ту версию программа не читала"]
        elif old != p.revision:
            state, why = "stale", [f"книга изменена после построения ({now_time}): данные схемы другие"]
    if st["nodes_rev"] != drawing_export.panel_rev(p):
        state = "stale"
        why.append("узлы (вкладка «Структура») изменены после построения")
    if not st["saved"] and state == "ok":
        state = "check"
        why.append("при построении книга в Excel была не сохранена")
    return state, "; ".join(why)


def apply(prj: Project, record: bool):
    """Panel.drawings — построения щита по чертежам (последнее на каждый DWG) с состоянием."""
    hist = load_history(prj)
    if record:
        hist = record_history(prj, hist)
    for d in prj.documents:
        stamps = read_stamps(stamp_path(prj, d)) if not d.error else []
        for p in d.panels:
            last = {}
            for st in stamps:
                if st["panel"] == p.name and st["at"] >= last.get(st["dwg"].lower(), {}).get("at", ""):
                    last[st["dwg"].lower()] = st
            p.drawings = []
            for st in sorted(last.values(), key=lambda s: s["at"], reverse=True):
                state, why = state_of(st, p, d, hist)
                p.drawings.append({"dwg": st["dwg"], "at": st["at"], "book_time": st["book_time"],
                                   "state": state, "why": why})
