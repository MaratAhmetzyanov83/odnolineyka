# -*- coding: utf-8 -*-
"""Кабельный журнал в DWG из окна: выпуск (новый DWG по КЖ_шаблон.dwg) и обновление (DWG из отметки .kj.txt).

Делает то же, что кнопка надстройки «КЖ в DWG»: вызывает её функцию SX_KJTo (Однолинейка.xlam) в своей скрытой
копии Excel. Книга открывается только для чтения. DWG пишет надстройка через AutoCAD (ObjectDBX, в фоне);
прежний DWG — рядом *.bak, ручные графы 9–13 сохраняются. Отметка <книга>.kj.txt — от надстройки.
Автотесты (ODN_TEST_WRITER=1): без Excel и AutoCAD — пишется только отметка .kj.txt.
"""
from __future__ import annotations

import json
import os
import re
import signal
import time
import uuid

from ..core.changes import EditError
from ..core.model import Project
from .drawing_stamp import kj_path
from .link_writer import ADDIN

OK_PREFIX = "Кабельный журнал обновлён"


def default_dwg(doc) -> str:
    """КЖ_<щит>.dwg рядом с книгой (как предлагает надстройка)."""
    s = os.path.splitext(os.path.basename(doc.file))[0]
    if s.lower().startswith("однолинейка"):
        s = s[len("однолинейка"):]
    s = s.lstrip(" _-") or "щит"
    return os.path.join(os.path.dirname(doc.file), f"КЖ_{s}.dwg")


def _target(prj: Project, doc, dwg: str) -> str:
    """Полный путь DWG: заданный (относительный — от папки книги), иначе из отметки, иначе по имени книги."""
    dwg = (dwg or "").strip().strip('"')
    if not dwg:
        dwg = (getattr(doc, "kj", None) or {}).get("dwg") or ""
    if not dwg:
        return os.path.abspath(os.path.join(prj.folder, default_dwg(doc)))
    if not dwg.lower().endswith(".dwg"):
        dwg += ".dwg"
    if not (os.path.isabs(dwg) or dwg.startswith("\\\\")):
        dwg = os.path.join(prj.folder, os.path.dirname(doc.file), dwg)
    return os.path.abspath(dwg)


def issue(prj: Project, items: list) -> dict:
    """items: [{"doc": книга, "dwg": путь или ""}]. Один Excel на все книги. Ошибка одной книги не мешает другим."""
    from .edit import _journal, _user
    jobs, errors = [], []
    for it in items:
        doc = next((d for d in prj.documents if d.file == it.get("doc")), None)
        if doc is None:
            errors.append(f"{it.get('doc')}: нет книги")
            continue
        if not doc.cable_journal:
            errors.append(f"{doc.file}: на листе «Кабельный_журнал» нет кабелей")
            continue
        dwg = _target(prj, doc, it.get("dwg", ""))
        if not os.path.isdir(os.path.dirname(dwg)):
            errors.append(f"{doc.file}: нет папки {os.path.dirname(dwg)}")
            continue
        jobs.append((doc, dwg))
    if not jobs:
        return {"ok": False, "errors": errors or ["нечего выпускать"]}
    run = _run_test if os.environ.get("ODN_TEST_WRITER") == "1" else _run_com
    done = []
    for (doc, dwg), (ok, msg) in zip(jobs, run(prj, jobs)):
        if ok:
            done.append({"doc": doc.file, "dwg": dwg, "msg": msg})
        else:
            errors.append(f"{doc.file}: {msg}")
    if done:
        with open(_journal(prj), "a", encoding="utf-8") as f:
            for d in done:
                f.write(json.dumps({"batch": f"{time.strftime('%Y-%m-%d_%H%M%S')}|КЖ|{uuid.uuid4().hex[:8]}",
                                    "at": time.strftime("%d.%m.%Y %H:%M:%S"), "user": _user(), "kind": "КЖ",
                                    "doc": d["doc"], "changes": [],
                                    "what": f"КЖ в DWG: {d['dwg']}. " + " ".join(d["msg"].splitlines()[2:])},
                                   ensure_ascii=False) + "\n")
    return {"ok": not errors, "done": done, "errors": errors}


def _run_com(prj, jobs):
    try:
        import pywintypes
        import win32com.client
        import win32process
    except ImportError:
        raise EditError("для выпуска КЖ нужен pywin32: py -m pip install pywin32") from None
    if not os.path.isfile(ADDIN):
        raise EditError(f"нет надстройки {ADDIN}")
    app = win32com.client.DispatchEx("Excel.Application")
    pid = win32process.GetWindowThreadProcessId(app.Hwnd)[1]
    res = []
    try:
        app.Visible = False
        app.DisplayAlerts = False
        app.AskToUpdateLinks = False
        app.AutomationSecurity = 1                   # макросы своей надстройки — только в этой копии Excel
        ad = app.Workbooks.Open(ADDIN)
        pre = "'" + ad.Name + "'!"
        app.Run(pre + "SX_SetQuiet", True)           # без окон: ни вопросов, ни сообщений
        for doc, dwg in jobs:
            wb = None
            try:
                wb = app.Workbooks.Open(os.path.abspath(os.path.join(prj.folder, doc.file)), 0, True)
                wb.Activate()
                try:
                    msg = str(app.Run(pre + "SX_KJTo", dwg) or "")
                except pywintypes.com_error as e:
                    if "SX_KJTo" in str(e) or "-2146827284" in str(e):
                        raise EditError("в надстройке нет SX_KJTo — пересоберите Однолинейка.xlam") from None
                    raise
                res.append((msg.startswith(OK_PREFIX), msg or "надстройка не вернула ответ"))
            except EditError as e:
                res.append((False, str(e)))
            except Exception as e:  # noqa: BLE001 — Excel / AutoCAD / COM
                res.append((False, f"Excel: {e}"))
            finally:
                if wb is not None:
                    try:
                        wb.Close(False)
                    except Exception:  # noqa: BLE001
                        pass
        return res
    finally:
        try:
            app.Quit()
        except Exception:  # noqa: BLE001
            pass
        try:
            os.kill(pid, signal.SIGTERM)
        except Exception:  # noqa: BLE001
            pass


def _clean(s):
    return re.sub(r"[\t\r\n]", " ", "" if s is None else str(s))


def _run_test(prj, jobs):
    """Только для автотестов: отметка .kj.txt, как пишет надстройка (графы 1–8 — значения книги)."""
    res = []
    for doc, dwg in jobs:
        with open(kj_path(prj, doc), "w", encoding="utf-8") as f:
            f.write(";; odn kj v1 - test\n")
            f.write(f"DWG\t{dwg}\nAT\t{time.strftime('%Y-%m-%d %H:%M:%S')}\nBOOK\t\nSAVED\t1\n")
            for r in doc.cable_journal:
                f.write("ROW\t" + "\t".join(_clean(v) for v in list(r)[:8]) + "\n")
        res.append((True, f"{OK_PREFIX}.\n{dwg}\nКабелей: {len(doc.cable_journal)} (тест)."))
    return res
