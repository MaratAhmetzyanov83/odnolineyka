# -*- coding: utf-8 -*-
"""Связь щитов — запись в книги. Делает то же, что кнопка надстройки «Связать с вышестоящим»: вызывает её функцию
SX_LinkUpstreamTo (Однолинейка.xlam) в своей скрытой копии Excel. Логика связи — одна, в надстройке.

Перед записью — копии обеих книг в «_окно/резерв/<время>/». Журнал: kind «связь», отмена — книги из копий
(если после записи их не меняли).
Автотесты (ODN_TEST_WRITER=1): без Excel — openpyxl пишет то, что читает программа (G6 нижестоящего, раздел Б вышестоящего).
"""
from __future__ import annotations

import json
import os
import shutil
import signal
import time
import uuid

from ..core.changes import EditError
from ..core.model import Project
from ..sources.excel_v312 import layout as X
from ..sources.excel_v312.writer import is_locked

HERE = os.path.dirname(os.path.abspath(__file__))
ADDIN = os.environ.get("SX_ADDIN") or os.path.normpath(os.path.join(HERE, "..", "..", "..", "3_Макросы_Excel",
                                                                     "Однолинейка.xlam"))


def _rel(base_dir, target):
    return os.path.relpath(target, base_dir)


def write_link(prj: Project, child_id: str, parent_id: str, feeder: str, inp: int = 1) -> dict:
    from .edit import _dir, _journal, _user
    child, par = prj.panel(child_id), prj.panel(parent_id)
    if child is None or par is None:
        raise EditError("нет щита")
    if not feeder:
        raise EditError("выберите линию-фидер")
    if child.doc == par.doc:
        raise EditError("щиты в одной книге — связь между книгами не нужна")
    cpath, ppath = (os.path.join(prj.folder, x) for x in (child.doc, par.doc))
    for d, path in ((child.doc, cpath), (par.doc, ppath)):
        doc = next(x for x in prj.documents if x.file == d)
        if abs(os.path.getmtime(path) - doc.mtime) > 1e-6:
            raise EditError(f"{d} изменилась после чтения — нажмите «Обновить» и повторите")
        if is_locked(path):
            raise EditError(f"{d} открыта в Excel — закройте книгу и повторите")
    stamp = time.strftime("%Y-%m-%d_%H%M%S")
    backups = {}
    for d, path in ((child.doc, cpath), (par.doc, ppath)):
        bak = os.path.join(_dir(prj, "резерв", stamp), d)
        os.makedirs(os.path.dirname(bak), exist_ok=True)
        shutil.copy2(path, bak)
        backups[d] = os.path.relpath(bak, prj.folder)
    idx = next(i for i, p in enumerate(next(x for x in prj.documents if x.file == child.doc).panels) if p.id == child_id) + 1
    try:
        if os.environ.get("ODN_TEST_WRITER") == "1":
            msg = _write_openpyxl(cpath, ppath, child, par, feeder, idx)
        else:
            msg = _write_com(cpath, ppath, feeder, idx, inp)
    except Exception:
        for d, path in ((child.doc, cpath), (par.doc, ppath)):      # что-то записалось наполовину — вернуть
            shutil.copy2(os.path.join(prj.folder, backups[d]), path)
        raise
    entry = {"batch": f"{stamp}|связь|{uuid.uuid4().hex[:8]}", "at": time.strftime("%d.%m.%Y %H:%M:%S"),
             "user": _user(), "kind": "связь", "doc": child.doc, "changes": [],
             "what": f"связь в книгах: {child.name} <- {par.name} / {feeder} (ввод {inp}). {msg.splitlines()[0] if msg else ''}",
             "backups": backups, "mtimes": {d: os.path.getmtime(os.path.join(prj.folder, d)) for d in backups}}
    with open(_journal(prj), "a", encoding="utf-8") as f:
        f.write(json.dumps(entry, ensure_ascii=False) + "\n")
    return {"ok": True, "msg": msg, "docs": list(backups)}


def undo_last(prj: Project) -> dict:
    from .edit import read_journal, _journal
    rows = [r for r in read_journal(prj, 10_000) if r.get("kind") == "связь" and not r["undone"]]
    if not rows:
        return {"ok": False, "errors": ["нечего отменять в связях"]}
    last = rows[-1]
    for d, mt in last["mtimes"].items():
        if abs(os.path.getmtime(os.path.join(prj.folder, d)) - mt) > 1e-6:
            return {"ok": False, "errors": [f"{d} менялась после записи связи — верните её из копии вручную: {last['backups'][d]}"]}
        if is_locked(os.path.join(prj.folder, d)):
            return {"ok": False, "errors": [f"{d} открыта в Excel — закройте книгу"]}
    for d, bak in last["backups"].items():
        shutil.copy2(os.path.join(prj.folder, bak), os.path.join(prj.folder, d))
    with open(_journal(prj), "a", encoding="utf-8") as f:
        f.write(json.dumps({"undo": last["batch"], "at": time.strftime("%d.%m.%Y %H:%M:%S")}, ensure_ascii=False) + "\n")
    return {"ok": True, "undone": last["what"], "docs": list(last["backups"])}


def _write_com(cpath, ppath, feeder, idx, inp):
    try:
        import win32com.client
        import win32process
    except ImportError:
        raise EditError("для записи в книги нужен pywin32: py -m pip install pywin32") from None
    if not os.path.isfile(ADDIN):
        raise EditError(f"нет надстройки {ADDIN}")
    app = win32com.client.DispatchEx("Excel.Application")
    pid = win32process.GetWindowThreadProcessId(app.Hwnd)[1]
    try:
        app.Visible = False
        app.DisplayAlerts = False
        app.AskToUpdateLinks = False
        app.AutomationSecurity = 1                   # макросы своей надстройки — только в этой копии Excel
        ad = app.Workbooks.Open(ADDIN)
        pre = "'" + ad.Name + "'!"
        try:
            app.Run(pre + "SX_SetQuiet", True)       # без окон сообщений
        except Exception:  # noqa: BLE001 — старая сборка надстройки
            pass
        up = app.Workbooks.Open(os.path.abspath(ppath), 0, False)
        ch = app.Workbooks.Open(os.path.abspath(cpath), 0, False)
        if up.ReadOnly or ch.ReadOnly:
            raise EditError("книга открыта у другого пользователя — запись не сделана")
        ch.Activate()
        msg = str(app.Run(pre + "SX_LinkUpstreamTo", os.path.abspath(ppath), feeder, idx, inp) or "")
        if not msg.startswith("Связано"):
            raise EditError(msg or "надстройка не вернула ответ")
        app.CalculateFull()
        ch.Save()
        up.Save()
        ch.Close(False)
        up.Close(False)
        return msg
    finally:
        try:
            app.Quit()
        except Exception:  # noqa: BLE001
            pass
        try:
            os.kill(pid, signal.SIGTERM)
        except Exception:  # noqa: BLE001
            pass


def _write_openpyxl(cpath, ppath, child, par, feeder, idx):
    """Только для автотестов: значения, которые читает программа (формулы связи не пишутся)."""
    import openpyxl
    wb = openpyxl.load_workbook(cpath)
    net = wb[X.NET]
    r = X.NET_PANEL_ROWS[idx - 1]
    net[f"{X.NET_SUPPLY['from_doc']}{r}"] = f"{_rel(os.path.dirname(cpath), ppath)} / {feeder}"
    net[f"{X.NET_SUPPLY['from_panel']}{r}"] = par.name
    wb.save(cpath)
    wb = openpyxl.load_workbook(ppath)
    net = wb[X.NET]
    rows = range(X.NET_SUB_ROW, X.NET_SUB_ROW + X.NET_SUB_COUNT)
    r = next((r for r in rows if net[f"B{r}"].value == feeder), None) or next(r for r in rows if not net[f"B{r}"].value)
    net[f"B{r}"] = feeder
    net[f"C{r}"] = _rel(os.path.dirname(ppath), cpath)
    wb.save(ppath)
    return f"Связано: щит «{child.name}» ← «{os.path.basename(ppath)}», линия {feeder}."
