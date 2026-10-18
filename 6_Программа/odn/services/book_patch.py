# -*- coding: utf-8 -*-
"""Доработки книги из окна: запуск скриптов 5_Проверка/patch_*.py на одной книге (свой Excel, закрывает только себя).

Перед запуском — копия книги в «_окно/резерв/<время>/», запись в журнал (kind «доработка», отмена — из копии вручную).
Автотесты (ODN_TEST_WRITER=1): без Excel — openpyxl ставит то, что читает программа.
"""
from __future__ import annotations

import json
import os
import shutil
import subprocess
import sys
import time
import uuid

from ..core.changes import EditError
from ..core.model import Project
from ..sources.excel_v312 import layout as X
from ..sources.excel_v312.writer import lock_reason

HERE = os.path.dirname(os.path.abspath(__file__))
PATCHES = os.path.normpath(os.path.join(HERE, "..", "..", "..", "5_Проверка"))
KNOWN = {"lowcurrent": ("patch_lowcurrent.py", "слаботочные линии (признак по коду, столбец «Слаб.»)")}


def run_patch(prj: Project, docname: str, name: str) -> dict:
    from .edit import _fresh, _journal, _user, backup_path
    if name not in KNOWN:
        raise EditError(f"нет доработки {name}")
    script, what = KNOWN[name]
    doc = next((d for d in prj.documents if d.file == docname), None)
    if doc is None:
        raise EditError("нет книги")
    path = os.path.join(prj.folder, docname)
    if not _fresh(prj, docname):
        raise EditError(f"{docname} изменилась после чтения — нажмите «Обновить» и повторите")
    why = lock_reason(path)
    if why:
        raise EditError(why)
    stamp = time.strftime("%Y-%m-%d_%H%M%S")
    bak = backup_path(prj, stamp, docname)
    shutil.copy2(path, bak)
    try:
        if os.environ.get("ODN_TEST_WRITER") == "1":
            out = _test_lowcurrent(path)
        else:
            r = subprocess.run([sys.executable, os.path.join(PATCHES, script), os.path.abspath(path)],
                               capture_output=True, text=True, encoding="utf-8", errors="replace", timeout=900)
            out = (r.stdout or "").strip()
            if r.returncode != 0:
                raise EditError(f"{script}: " + ((r.stderr or "").strip().splitlines() or ["ошибка"])[-1])
    except Exception:
        shutil.copy2(bak, path)                     # наполовину доработанную книгу — вернуть
        raise
    with open(_journal(prj), "a", encoding="utf-8") as f:
        f.write(json.dumps({"batch": f"{stamp}|доработка|{uuid.uuid4().hex[:8]}", "at": time.strftime("%d.%m.%Y %H:%M:%S"),
                            "user": _user(), "kind": "доработка", "doc": docname, "changes": [],
                            "what": f"{what}: {out}", "backup": os.path.relpath(bak, prj.folder)}, ensure_ascii=False) + "\n")
    return {"ok": True, "msg": out or "готово", "backup": os.path.relpath(bak, prj.folder)}


def _test_lowcurrent(path):
    import openpyxl
    wb = openpyxl.load_workbook(path)
    wb[X.SRC][f"{X.SRC_FIELDS['input.low_current']}2"] = X.LOW_HDR
    wb.save(path)
    return "тест: столбец слаботочных"
