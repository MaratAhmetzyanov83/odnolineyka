# -*- coding: utf-8 -*-
"""Запись правок в книгу v3.12.

Как в скриптах 5_Проверка: свой скрытый Excel (DispatchEx), книга открыта у другого — не трогаем,
защиту листа снимаем и ставим обратно с теми же настройками, полный пересчёт, сохранение, закрываем только свой Excel.
COM — только позиционные аргументы.

Для автотестов (без Excel): ODN_TEST_WRITER=1 — запись через openpyxl. Только для тестовых книг без формул:
openpyxl теряет часть оформления и кэш формул, на рабочих книгах его не включать.
"""
from __future__ import annotations

import os
import signal

from ...core.changes import EditError
from . import layout as X

PWD = "sx"


def cells_for(prj, changes):
    """Правки -> [(лист, адрес, значение)] по layout.py и местам линий в книге (refs)."""
    out = []
    for ch in changes:
        if ch.line:
            ln = prj.line(ch.panel, ch.line)
            if ch.path in X.SRC_EDIT:
                ref = ln.refs.get("input")
                col, sheet = X.SRC_EDIT[ch.path], X.SRC
            elif ch.path in X.ODN_EDIT:
                ref = ln.refs.get("scheme")
                col, sheet = X.ODN_EDIT[ch.path], X.ODN
            else:
                raise EditError(f"{ch.path}: нет адреса в книге")
            if ref is None:
                raise EditError(f"линия {ch.line}: нет строки на листе «{sheet}»")
            row = "".join(c for c in ref.cell if c.isdigit())
            out.append((sheet, f"{col}{row}", ch.new))
        else:
            if ch.path not in X.NET_EDIT:
                raise EditError(f"{ch.path}: нет адреса в книге")
            r, c = X.NET_EDIT[ch.path]
            out.append((X.NET, f"{c}{r}", ch.new))
    return out


def is_locked(path: str) -> bool:
    """Книга открыта у кого-то: есть «~$имя» рядом или файл не открывается на запись."""
    d, n = os.path.split(path)
    if os.path.exists(os.path.join(d, "~$" + n[2:])) or os.path.exists(os.path.join(d, "~$" + n)):
        return True
    try:
        with open(path, "r+b"):
            return False
    except OSError:
        return True


def write_cells(path: str, cells) -> None:
    if is_locked(path):
        raise EditError(f"{os.path.basename(path)} открыта в Excel — закройте книгу и повторите")
    if os.environ.get("ODN_TEST_WRITER") == "1":
        return _write_openpyxl(path, cells)
    return _write_com(path, cells)


def _write_openpyxl(path, cells):
    import openpyxl
    wb = openpyxl.load_workbook(path)
    for sheet, addr, v in cells:
        wb[sheet][addr] = v
    wb.save(path)


def _protection(ws):
    p = ws.Protection
    return (True, True, True, False, p.AllowFormattingCells, p.AllowFormattingColumns, p.AllowFormattingRows,
            p.AllowInsertingColumns, p.AllowInsertingRows, p.AllowInsertingHyperlinks, p.AllowDeletingColumns,
            p.AllowDeletingRows, p.AllowSorting, p.AllowFiltering, p.AllowUsingPivotTables)


def _write_com(path, cells):
    try:
        import win32com.client
        import win32process
    except ImportError:
        raise EditError("для записи в книги нужен pywin32: py -m pip install pywin32") from None
    app = win32com.client.DispatchEx("Excel.Application")
    pid = win32process.GetWindowThreadProcessId(app.Hwnd)[1]
    try:
        app.Visible = False
        app.DisplayAlerts = False
        app.AskToUpdateLinks = False
        app.AutomationSecurity = 3
        wb = app.Workbooks.Open(os.path.abspath(path), 0, False)
        if wb.ReadOnly:
            wb.Close(False)
            raise EditError(f"{os.path.basename(path)} открыта у другого пользователя — запись не сделана")
        by_sheet = {}
        for sheet, addr, v in cells:
            by_sheet.setdefault(sheet, []).append((addr, v))
        for sheet, items in by_sheet.items():
            ws = wb.Worksheets(sheet)
            opts = None
            if ws.ProtectContents:
                opts = _protection(ws)
                ws.Unprotect(PWD)
            try:
                for addr, v in items:
                    ws.Range(addr).Value = v
            finally:
                if opts:
                    ws.Protect(PWD, *opts)
        app.CalculateFull()
        wb.Save()
        wb.Close(False)
    finally:
        try:
            app.Quit()
        except Exception:  # noqa: BLE001
            pass
        try:
            os.kill(pid, signal.SIGTERM)
        except Exception:  # noqa: BLE001
            pass


def write_batch(cells_by_path: dict, related: list) -> None:
    """Одна сессия Excel: связанные книги открыты вместе с изменяемыми — формулы связей между ними считаются вживую.
    cells_by_path: {путь: [(лист, адрес, значение)]}; related — книги только пересчитать и сохранить."""
    if os.environ.get("ODN_TEST_WRITER") == "1":
        for path, cells in cells_by_path.items():
            _write_openpyxl(path, cells)               # тест: без пересчёта связанных
        return
    try:
        import win32com.client
        import win32process
    except ImportError:
        raise EditError("для записи в книги нужен pywin32: py -m pip install pywin32") from None
    app = win32com.client.DispatchEx("Excel.Application")
    pid = win32process.GetWindowThreadProcessId(app.Hwnd)[1]
    try:
        app.Visible = False
        app.DisplayAlerts = False
        app.AskToUpdateLinks = False
        app.AutomationSecurity = 3
        books = {}
        for path in list(related) + list(cells_by_path):
            wb = app.Workbooks.Open(os.path.abspath(path), 0, False)
            if wb.ReadOnly:
                raise EditError(f"{os.path.basename(path)} открыта у другого пользователя — запись не сделана")
            books[path] = wb
        for path, cells in cells_by_path.items():
            wb = books[path]
            by_sheet = {}
            for sheet, addr, v in cells:
                by_sheet.setdefault(sheet, []).append((addr, v))
            for sheet, items in by_sheet.items():
                ws = wb.Worksheets(sheet)
                opts = None
                if ws.ProtectContents:
                    opts = _protection(ws)
                    ws.Unprotect(PWD)
                try:
                    for addr, v in items:
                        ws.Range(addr).Value = v
                finally:
                    if opts:
                        ws.Protect(PWD, *opts)
        for _ in range(2):                             # вниз и вверх по связям — до устойчивых значений
            app.CalculateFull()
        for wb in books.values():
            wb.Save()
        for wb in books.values():
            wb.Close(False)
    finally:
        try:
            app.Quit()
        except Exception:  # noqa: BLE001
            pass
        try:
            os.kill(pid, signal.SIGTERM)
        except Exception:  # noqa: BLE001
            pass
