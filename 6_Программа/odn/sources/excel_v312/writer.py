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
        if ch.addr:                                   # производная правка: адрес готов
            sheet, _, a = ch.addr.partition("!")
            out.append((sheet, a, ch.new))
        elif ch.line:
            ln = prj.line(ch.panel, ch.new if ch.path == "number" else ch.line)
            if ch.path in X.SRC_EDIT:
                ref = ln.refs.get("input")
                col, sheet = X.SRC_EDIT[ch.path], X.SRC
            elif ch.path in X.ODN_EDIT:
                ref = ln.refs.get("scheme")
                col, sheet = X.ODN_EDIT[ch.path], X.ODN
            elif ch.path in X.LOAD_EDIT:
                ref = ln.refs.get("load")
                col, sheet = X.LOAD_EDIT[ch.path], X.LOAD
            else:
                raise EditError(f"{ch.path}: нет адреса в книге")
            if ref is None:
                raise EditError(f"линия {ch.line}: нет строки на листе «{sheet}»")
            row = "".join(c for c in ref.cell if c.isdigit())
            v = ch.new
            if ch.path == "input.low_current_set":     # да / нет — число; по коду (и пусто — отмена) — формула
                v = X.LOW_VALUE.get(v, X.LOW_FORMULA.format(r=row))
                doc = next((d for d in prj.documents if d.file == prj.panel(ch.panel).doc), None)
                col = (doc.low_col if doc else "") or col
            if ch.path == "scheme.phase_set" and v in ("1", "2", "3"):
                v = int(v)                              # формула G: «L»&E и E=0 — число, как вводят вручную
            out.append((sheet, f"{col}{row}", v))
        elif ch.path == "scheme.is_vru":              # столбец щита: C / D / E по порядку щитов книги
            p = prj.panel(ch.panel)
            k = next(i for i, q in enumerate(next(d for d in prj.documents if d.file == p.doc).panels) if q.id == p.id)
            out.append((X.NET, f"{X.NET_PANEL_COLS[min(k, 2)]}{X.NET_VRU_ROW}", ch.new))
        else:
            if ch.path not in X.NET_EDIT:
                raise EditError(f"{ch.path}: нет адреса в книге")
            r, c = X.NET_EDIT[ch.path]
            out.append((X.NET, f"{c}{r}", ch.new))
    return out


def _writable(path: str) -> bool:
    """Файл можно открыть на запись. Excel держит открытую книгу с запретом записи — тогда нельзя."""
    try:
        with open(path, "r+b"):
            return True
    except OSError:
        return False


def lock_owner(lock: str) -> str:
    """Кто открыл книгу — из файла-метки Excel «~$книга»: длина + имя (ANSI), с байта 55 — длина + имя (UTF-16)."""
    try:
        with open(lock, "rb") as f:
            b = f.read(200)
    except OSError:
        return ""
    try:
        m = b[55] if len(b) > 56 else 0
        if m and len(b) >= 57 + 2 * m:
            return b[57:57 + 2 * m].decode("utf-16-le").strip("\x00 ")
        return b[1:1 + b[0]].decode("cp1251").strip("\x00 ") if b else ""
    except (UnicodeDecodeError, IndexError):
        return ""


def lock_reason(path: str) -> str:
    """Почему книгу нельзя писать; пусто — можно. Старая метка «~$» без занятого файла (после сбоя Excel) не мешает."""
    if _writable(path):
        return ""
    name = os.path.basename(path)
    if not os.access(path, os.W_OK):
        return f"{name}: у файла атрибут «только чтение» (Свойства файла) или нет прав на запись в папку"
    d = os.path.dirname(path)
    lock = next((os.path.join(d, x) for x in ("~$" + name, "~$" + name[2:]) if os.path.exists(os.path.join(d, x))), None)
    if lock:
        who = lock_owner(lock)
        return f"{name} открыта в Excel" + (f" у пользователя «{who}»" if who else "") + " — закройте книгу и повторите"
    return (f"{name} занята другим процессом: чаще всего — Excel без окна (Диспетчер задач → «Microsoft Excel» "
            f"в фоновых процессах → «Снять задачу»), иначе синхронизация (OneDrive) или антивирус. Повторите через минуту")


def is_locked(path: str) -> bool:
    return bool(lock_reason(path))


def write_cells(path: str, cells) -> None:
    why = lock_reason(path)
    if why:
        raise EditError(why)
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


XL_MANUAL, XL_AUTO = -4135, -4105


def _excel():
    """Свой скрытый Excel в ручном пересчёте -> (app, pid)."""
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
        app.ScreenUpdating = False
        app.EnableEvents = False
        tmp = app.Workbooks.Add()                      # режим пересчёта ставится, когда открыта хоть одна книга
        app.Calculation = XL_MANUAL                    # не пересчитывать на каждой записанной ячейке и при открытии
        tmp.Close(False)
    except Exception:
        _quit(app, pid)
        raise
    return app, pid


def _quit(app, pid):
    """Закрыть только свой Excel."""
    try:
        app.Quit()
    except Exception:  # noqa: BLE001
        pass
    try:
        os.kill(pid, signal.SIGTERM)
    except Exception:  # noqa: BLE001
        pass


def _open(app, path, links=0):
    """links: 0 — значения связей как сохранены в книге; 3 — обновить связи из сохранённых файлов других книг."""
    wb = app.Workbooks.Open(os.path.abspath(path), links, False)
    if wb.ReadOnly:
        raise EditError(f"{os.path.basename(path)} открыта у другого пользователя — запись не сделана")
    return wb


def write_batch(cells_by_path: dict, related: list, full: bool = False) -> None:
    """Одна сессия Excel: связанные книги открыты вместе с изменяемыми — формулы связей между ними считаются вживую.
    cells_by_path: {путь: [(лист, адрес, значение)]}; related — книги только пересчитать и сохранить.
    Пересчёт: обычный (изменённые ячейки и зависимые — быстро); full — полный (синхронизация книг)."""
    if os.environ.get("ODN_TEST_WRITER") == "1":
        for path, cells in cells_by_path.items():
            _write_openpyxl(path, cells)               # тест: без пересчёта связанных
        return
    app, pid = _excel()
    try:
        books = {path: _open(app, path) for path in list(related) + list(cells_by_path)}
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
            app.CalculateFull() if full else app.Calculate()
        app.Calculation = XL_AUTO                      # книги сохраняются с автопересчётом, как были
        for wb in books.values():
            wb.Save()
        for wb in books.values():
            wb.Close(False)
    finally:
        _quit(app, pid)


def _refresh_group(paths, done):
    """Книги одной группы — в своём Excel: открыть с обновлением связей, пересчитать, сохранить."""
    import pythoncom
    pythoncom.CoInitialize()
    try:
        app, pid = _excel()
        try:
            books = [_open(app, p, 3) for p in paths]
            for _ in range(2):
                app.Calculate()
            app.Calculation = XL_AUTO
            for wb in books:
                wb.Save()
                done(1)
            for wb in books:
                wb.Close(False)
        finally:
            _quit(app, pid)
    finally:
        pythoncom.CoUninitialize()


def refresh_books(groups: list, workers: int = 4, done=lambda n: None) -> dict:
    """Пересчитать и сохранить книги, которые берут значения из уже записанных книг (связи «Связать с вышестоящим»).
    groups — [[путь, …]]: книги со связями между собой — в одной группе (одна сессия Excel), группы — параллельно,
    не больше workers Excel сразу. Связи обновляются из сохранённых файлов. -> {путь: ошибка} для несделанных."""
    if os.environ.get("ODN_TEST_WRITER") == "1" or not groups:
        for g in groups:
            done(len(g))
        return {}
    from concurrent.futures import ThreadPoolExecutor
    errors = {}
    with ThreadPoolExecutor(max(1, min(workers, len(groups)))) as ex:
        futs = {ex.submit(_refresh_group, g, done): g for g in groups}
        for f, g in futs.items():
            try:
                f.result()
            except Exception as e:  # noqa: BLE001 — Excel / COM
                for p in g:
                    errors[p] = str(e)
    return errors
