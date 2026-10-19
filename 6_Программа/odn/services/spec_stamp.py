# -*- coding: utf-8 -*-
"""Актуальность спецификации книги (лист «Спецификация»): не забыть поправить её после правок аппаратов.

Лист «Спецификация» в объектах ведут вручную (корпус, пластроны, провод…), черновик «Спецификация_авто» неполный —
сравнивать листы нельзя. Поэтому программа помнит, какие аппараты были в книге, когда спецификацию считали верной:
«_окно/спецификация.json» — {книга: {at, how, items, sheet}}
    items — отпечаток аппаратов: «щит|линия» -> «автомат | аппараты | клемма», «щит|*» -> «вводы | АВР | Icu»;
    sheet — отпечаток значений листа «Спецификация» B4:J160 (Document.spec_sheet, читает reader);
    how   — auto: первая встреча книги, принято как есть; user — «Спецификация верна» в окне.

Сверка (Document.spec):
    аппараты те же                           → ok (лист меняли — запоминаем новый отпечаток листа);
    аппараты другие, лист тот же              → stale: что изменилось;
    аппараты другие, лист тоже меняли         → check: проверить и нажать «Спецификация верна».
"""
from __future__ import annotations

import json
import os
import time

from ..core.model import Project

STORE = "спецификация.json"
SHOW = 8                         # изменений в подсказке


def _path(prj):
    return os.path.join(prj.folder, "_окно", STORE)


def load(prj) -> dict:
    try:
        with open(_path(prj), encoding="utf-8") as f:
            return json.load(f)
    except (OSError, ValueError):
        return {}


def save(prj, data: dict):
    try:
        os.makedirs(os.path.dirname(_path(prj)), exist_ok=True)
        tmp = _path(prj) + ".tmp"
        with open(tmp, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=0)
        os.replace(tmp, _path(prj))
    except OSError:
        pass


def _v(x) -> str:
    if x is None:
        return ""
    if isinstance(x, float) and x.is_integer():
        x = int(x)
    return " ".join(str(x).split())


def items(doc) -> dict:
    """Отпечаток аппаратов книги: то, что попадает в спецификацию щита."""
    out = {}
    for p in doc.panels:
        s = p.scheme
        out[f"{p.name}|*"] = " | ".join(_v(x) for x in (s.inputs, s.avr, s.icu))
        for ln in p.lines:
            out[f"{p.name}|{ln.number}"] = " | ".join(
                (_v(ln.scheme.breaker), ", ".join(_v(d) for d in ln.scheme.devices), _v(ln.scheme.terminal)))
    return out


def _part(k, v):
    """Текст изменения: «автомат С16А», «аппараты …», «клемма X1»."""
    names = ("вводы", "АВР", "Icu") if k.endswith("|*") else ("автомат", "аппараты", "клемма")
    vals = (v.split(" | ") + ["", "", ""])[:3]
    return dict(zip(names, vals))


def diff(old: dict, cur: dict) -> list:
    out = []
    for k in cur:
        if k not in old:
            if k.endswith("|*"):
                continue
            a = _part(k, cur[k])
            out.append(f"новая линия {k.replace('|', ' ')}" + (f" ({a['автомат']})" if a["автомат"] else ""))
        elif old[k] != cur[k]:
            a, b = _part(k, old[k]), _part(k, cur[k])
            what = "; ".join(f"{n} {a[n] or '—'} → {b[n] or '—'}" for n in a if a[n] != b[n])
            out.append(f"{k.replace('|*', '').replace('|', ' ')}: {what}")
    for k in old:
        if k not in cur and not k.endswith("|*"):
            out.append(f"убрана линия {k.replace('|', ' ')}")
    return out


def state(rec: dict, doc) -> dict:
    cur = items(doc)
    ch = diff(rec.get("items", {}), cur)
    base = {"at": rec.get("at", ""), "how": rec.get("how", ""), "changes": len(ch)}
    if not ch:
        return dict(base, state="ok", why="")
    more = f" … и ещё {len(ch) - SHOW}" if len(ch) > SHOW else ""
    why = "; ".join(ch[:SHOW]) + more
    if doc.spec_sheet != rec.get("sheet"):
        return dict(base, state="check", why="лист «Спецификация» меняли после изменения аппаратов — проверьте: " + why)
    return dict(base, state="stale", why=why)


def _record(doc, how) -> dict:
    return {"at": time.strftime("%d.%m.%Y %H:%M"), "how": how, "items": items(doc), "sheet": doc.spec_sheet}


def apply(prj: Project, record: bool):
    """Document.spec для книг с листом «Спецификация». record — при чтении книг: новые книги принять как есть."""
    data, changed = load(prj), False
    for d in prj.documents:
        d.spec = {}
        if d.error or not d.spec_sheet or not d.panels:
            continue
        rec = data.get(d.file)
        if rec is None:
            if not record:
                continue
            rec = data[d.file] = _record(d, "auto")
            changed = True
        d.spec = state(rec, d)
        if record and d.spec["state"] == "ok" and rec.get("sheet") != d.spec_sheet:
            rec["sheet"] = d.spec_sheet                    # лист поправили, аппараты те же — помнить новый лист
            changed = True
    if changed:
        save(prj, data)


def accept(prj: Project, docname: str) -> dict:
    """«Спецификация верна»: запомнить аппараты и лист книги как есть."""
    doc = next((d for d in prj.documents if d.file == docname), None)
    if doc is None or not doc.spec_sheet:
        return {"ok": False, "errors": ["нет книги с листом «Спецификация»"]}
    data = load(prj)
    data[docname] = _record(doc, "user")
    save(prj, data)
    doc.spec = state(data[docname], doc)
    return {"ok": True}
