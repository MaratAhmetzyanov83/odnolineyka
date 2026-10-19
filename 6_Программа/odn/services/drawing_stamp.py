# -*- coding: utf-8 -*-
"""Актуальность однолинейных схем (DWG) — по отметкам SXDRAW (v3.26+), и кабельных журналов (КЖ в DWG) — по отметкам
надстройки (SX_KJ, кнопка «КЖ в DWG»): «<книга>.kj.txt» — графы 1–8 в том виде, в каком ушли в DWG; программа сравнивает
их с листом «Кабельный_журнал» книги по маркировке (изменённые, новые, убранные кабели).

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


# ---------------------------------------------------------------- КЖ в DWG (надстройка SX_KJ, кнопка «КЖ в DWG»)
KJ_EXT = ".kj.txt"
KJ_NAMES = 8                     # графы 1–8: по проекту (9–13 заполняют в DWG вручную — не сверяются)


def kj_path(prj: Project, doc) -> str:
    base, _ = os.path.splitext(os.path.join(prj.folder, doc.file))
    return base + KJ_EXT


KJ_SKIP_EXT = ".kj.skip.txt"     # кабели, убранные из КЖ: маркировка в строке (пишет окно, читает и надстройка)


def kj_skip_path(prj: Project, doc) -> str:
    base, _ = os.path.splitext(os.path.join(prj.folder, doc.file))
    return base + KJ_SKIP_EXT


def read_kj_skip(prj: Project, doc) -> list:
    try:
        with open(kj_skip_path(prj, doc), encoding="utf-8-sig", errors="replace") as f:
            rows = f.read().splitlines()
    except OSError:
        return []
    return [_norm(x) for x in rows if _norm(x) and not x.lstrip().startswith(";")]


def kj_rows(prj: Project, doc) -> list:
    """Строки КЖ книги без убранных из КЖ кабелей."""
    skip = set(read_kj_skip(prj, doc))
    return [r for r in doc.cable_journal if _norm(r[0]) not in skip] if skip else doc.cable_journal


def read_kj(path):
    """Отметка КЖ: UTF-8, строки «ключ<Tab>значения»: DWG, AT, BOOK, SAVED, ROW графа1…графа8 (как видно в Excel)."""
    try:
        with open(path, encoding="utf-8-sig", errors="replace") as f:
            rows = f.read().splitlines()
    except OSError:
        return None
    st = {"dwg": "", "at": "", "book_time": "", "saved": True, "rows": []}
    for ln in rows:
        k, _, v = ln.partition("\t")
        if k == "ROW":
            st["rows"].append((v.split("\t") + [""] * KJ_NAMES)[:KJ_NAMES])
        elif k in ("DWG", "AT", "BOOK"):
            st[{"DWG": "dwg", "AT": "at", "BOOK": "book_time"}[k]] = v.strip()
        elif k == "SAVED":
            st["saved"] = v.strip() != "0"
    return st


def _norm(s) -> str:
    return " ".join(str(s).replace("\r", " ").replace("\n", " ").replace("\t", " ").split())


def same_cell(txt: str, val) -> bool:
    """Текст ячейки (как в Excel, с форматом) == значение ячейки (openpyxl)? Числа — с точностью показа."""
    t = _norm(txt)
    if val is None or val == "":
        return t == ""
    if isinstance(val, str):
        v = _norm(val)
        return t == v or (t.startswith("#") and v.startswith("#"))      # #Н/Д == #N/A
    if isinstance(val, bool):
        return t.upper() in (("ИСТИНА", "TRUE") if val else ("ЛОЖЬ", "FALSE"))
    if t.startswith("#"):
        return False
    x = t.replace("\u00a0", "").replace(" ", "").replace(",", ".")
    pct = x.endswith("%")
    x = x.rstrip("%")
    try:
        f = float(x)
    except ValueError:
        return t == _norm(val)
    dec = len(x.partition(".")[2])
    v = val * 100 if pct else val
    return abs(v - f) <= 0.5 * 10 ** -dec + 1e-9


def kj_diff(stamp_rows, cur_rows) -> dict:
    """{added, removed, changed: [(маркировка, графа, было, стало)]} — по маркировке кабеля."""
    old = {_norm(r[0]): r for r in stamp_rows}
    new = {_norm(r[0]): r for r in cur_rows}
    changed = []
    for mk, r in new.items():
        o = old.get(mk)
        if o is None:
            continue
        for k in range(1, KJ_NAMES):
            v = r[k] if k < len(r) else None
            if not same_cell(o[k], v):
                changed.append((mk, k + 1, _norm(o[k]), "" if v is None else _norm(v)))
    return {"added": [m for m in new if m not in old], "removed": [m for m in old if m not in new], "changed": changed}


def kj_state(prj: Project, doc) -> dict:
    st = read_kj(kj_path(prj, doc))
    if st is None:
        return {}
    rows = kj_rows(prj, doc)
    d = kj_diff(st["rows"], rows)
    why = []
    if d["changed"]:
        cab = sorted({c[0] for c in d["changed"]})
        why.append(f"изменено кабелей {len(cab)}: " + "; ".join(
            f"{m} (гр. {g}: {a or '—'} → {b or '—'})" for m, g, a, b in
            [next(c for c in d["changed"] if c[0] == m) for m in cab[:5]]) + (" …" if len(cab) > 5 else ""))
    if d["added"]:
        why.append(f"новые {len(d['added'])}: " + ", ".join(d["added"][:5]) + (" …" if len(d["added"]) > 5 else ""))
    if d["removed"]:
        why.append(f"убраны {len(d['removed'])}: " + ", ".join(d["removed"][:5]) + (" …" if len(d["removed"]) > 5 else ""))
    return {"dwg": st["dwg"], "at": st["at"], "state": "stale" if why else "ok", "why": "; ".join(why),
            "cables": len(rows), "saved": st["saved"]}


def apply(prj: Project, record: bool):
    """Panel.drawings — построения щита по чертежам (последнее на каждый DWG) с состоянием."""
    hist = load_history(prj)
    if record:
        hist = record_history(prj, hist)
    for d in prj.documents:
        d.kj = kj_state(prj, d) if not d.error else {}
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
