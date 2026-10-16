# -*- coding: utf-8 -*-
"""ДСУП (дополнительная система уравнивания потенциалов): линии РЕ и КУП — от шины PE щита.

    py patch_dsup.py <книга.xlsx> [...]      — книга закрыта; повторный запуск ничего не портит

Образец — «Однолинейка ЩР-2» (098, Том 3 ХБК), лист 5: проводник ПуГВ от шины PE, без автомата и клемм;
внизу — КУП (блок «_КУП») или точка заземления (вывод); точки РЕ сразу после КУП — отводом от линии КУП.
Как пользоваться в книге («Однолинейка»):
  * Линия ДСУП — код в «Справочная» DF (РЕ, PE, КУП, ШДУП, ДСУП) и пустой E (автомат). Тогда CJ = 1:
    нет группы QF, нет клеммы XT, фаза «PE», нет «нет мощности». E «авто» — обычная линия с автоматом.
  * CK — от какой КУП идёт точка РЕ: сама — от КУП выше по порядку (строки ДСУП подряд), иначе пусто =
    от шины PE. Можно вписать имя КУП вручную или очистить ячейку (тогда — от шины).
SXDRAW v3.23 рисует ДСУП на уровне общей шины PE (sx:vPE).
Свой скрытый экземпляр Excel; закрывается только он.
"""
import os
import signal
import sys

PWD = "sx"
ODN = "Однолинейка"
SPR = "Справочная"
TITLE = "ДСУП: коды линий от шины PE"
CODES = ["РЕ", "PE", "КУП", "ШДУП", "ДСУП"]
ROWS = 10                                        # DF4:DF13 — место под свои коды
NAME = "ДСУП_коды"
NOTES = [
    "Линия с таким кодом и пустым E («Однолинейка») — ДСУП: от шины PE, без автомата и клемм (CJ = 1).",
    "КУП, ШДУП — коробка уравнивания (блок «_КУП»); точки РЕ после неё — отвод от её линии (CK).",
    "E «авто» у такой линии — обычная линия с автоматом.",
]
COLS = {   # столбец: (заголовок, формула для строки 3)
    "CJ": ("ДСУП\n(1 = от шины PE,\nбез автомата)",
           '=IF(OR($R3="",$E3<>""),0,IF(ISNUMBER(MATCH($AL3,ДСУП_коды,0)),1,0))'),
    "CK": ("ДСУП: от КУП\n(пусто = от шины PE;\nможно вписать)",
           '=IF(N($CJ3)=0,"",IF(ISNUMBER(MATCH($AL3,{"КУП","ШДУП"},0)),"",'
           'IF(AND($A3=$A2,N($CJ2)=1),IF(ISNUMBER(MATCH($AL2,{"КУП","ШДУП"},0)),$R2,$CK2),"")))'),
}
# правки формул (R1C1): столбец -> (метка «уже сделано», старый кусок, новый кусок) или (метка, None, обёртка)
EDITS = {
    # группа QF: строка ДСУП не продолжает группу (CJ = столбец 88)
    "AF": ("N(RC[56])=1", 'RC[-7]="вручную")', 'RC[-7]="вручную",N(RC[56])=1)'),
    "AG": ("N(RC[55])=1", 'RC[-8]="вручную")', 'RC[-8]="вручную",N(RC[55])=1)'),
    # клеммы XT нет
    "AY": ("N(RC[37])=1", '=IF(RC[-37]="б/к",0,', '=IF(OR(RC[-37]="б/к",N(RC[37])=1),0,'),
    # «нет мощности» у ДСУП не пишем
    "AE": ("N(RC[57])=0", 'N(RC[-11])=0', 'N(RC[-11])=0,N(RC[57])=0'),
    # фаза «PE»; цепь одной строкой: «шина PE» / «PE → КУП…»
    "V": ("N(RC[66])=1", None, '=IF(N(RC[66])=1,"PE",{})'),
    "AD": ("N(RC[58])=1", None, '=IF(N(RC[58])=1,IF(RC[59]="","шина PE","PE → "&RC[59]),{})'),
}


def excel():
    import win32com.client
    import win32process
    app = win32com.client.DispatchEx("Excel.Application")
    app.Visible = False
    app.DisplayAlerts = False
    app.AskToUpdateLinks = False
    app.AutomationSecurity = 3
    return app, win32process.GetWindowThreadProcessId(app.Hwnd)[1]


def unprotect(ws):
    if not ws.ProtectContents:
        return None
    p = ws.Protection
    opts = (True, True, True, False, p.AllowFormattingCells, p.AllowFormattingColumns, p.AllowFormattingRows,
            p.AllowInsertingColumns, p.AllowInsertingRows, p.AllowInsertingHyperlinks, p.AllowDeletingColumns,
            p.AllowDeletingRows, p.AllowSorting, p.AllowFiltering, p.AllowUsingPivotTables)
    ws.Unprotect(PWD)
    return opts


def runs(ws, col, r0, r1):
    fs = [x[0] for x in ws.Range(f"{col}{r0}:{col}{r1}").Formula2R1C1]
    out, start = [], r0
    for i in range(1, len(fs) + 1):
        if i == len(fs) or fs[i] != fs[i - 1]:
            out.append((start, r0 + i - 1, fs[i - 1]))
            start = r0 + i
    return out


def fmt(ws, src, dst):
    ws.Range(src).Copy(ws.Range(dst))
    ws.Range(dst).ClearContents()


def spr(wb, log):
    ws = wb.Worksheets(SPR)
    if str(ws.Range("DF2").Value or "").strip() == TITLE:
        return
    if ws.Range("DF1:DF40").Find("*") is not None:
        raise SystemExit(f"«{SPR}» DF занят — книга не изменена")
    opts = unprotect(ws)
    fmt(ws, "CC1", "DF1")
    ws.Range("DF1").Value = "ДСУП. Жёлтые — можно менять и дополнять."
    fmt(ws, "CC2", "DF2")
    ws.Range("DF2").Value = TITLE
    fmt(ws, "BP3", "DF3")
    ws.Range("DF3").Value = "Код линии"
    fmt(ws, "BQ4", f"DF4:DF{3 + ROWS}")
    ws.Range(f"DF4:DF{3 + len(CODES)}").Value = tuple((c,) for c in CODES)
    fmt(ws, "BR14", f"DF{6 + ROWS}:DF{5 + ROWS + len(NOTES)}")
    for i, s in enumerate(NOTES):
        ws.Range(f"DF{6 + ROWS + i}").Value = s
    ws.Columns("DF").ColumnWidth = 14
    try:
        wb.Names(NAME).Delete()
    except Exception:
        pass
    wb.Names.Add(NAME, f"='{SPR}'!$DF$4:$DF${3 + ROWS}")
    if opts:
        ws.Protect(PWD, *opts)
    log.append(f"«{SPR}» DF — коды ДСУП ({', '.join(CODES)}), имя {NAME}")


def put(ws, rng, f):
    """Formula2R1C1 на диапазон; не вышло (формулы массива) — по ячейкам."""
    try:
        ws.Range(rng).Formula2R1C1 = f
    except Exception:
        for c in ws.Range(rng):
            c.Formula2R1C1 = f


def odn(wb, log):
    ws = wb.Worksheets(ODN)
    opts = unprotect(ws)
    last = max(ws.Cells(ws.Rows.Count, c).End(-4162).Row for c in ("A", "H", "AF", "BH"))
    for col, (hdr, f) in COLS.items():
        h = ws.Range(f"{col}2")
        if str(h.Value or "").strip() not in ("", hdr):
            raise SystemExit(f"«{ODN}» {col}2 занят: «{h.Value}» — книга не изменена")
        if h.Value != hdr:
            ws.Range("CB2").Copy(h)
            h.Value = hdr
            ws.Range("CB3").Copy(ws.Range(f"{col}3:{col}{last}"))
            ws.Range(f"{col}3:{col}{last}").ClearContents()
            ws.Columns(col).ColumnWidth = 12
            log.append(f"«{ODN}» {col}2 — «{hdr.replace(chr(10), ' ')}»")
        if ws.Range(f"{col}3").Formula2 != f or ws.Range(f"{col}{last}").Formula2 == "":
            vals = [x[0] for x in ws.Range(f"{col}3:{col}{last}").Formula2]
            # CK — вписанные вручную КУП (текст без «=») не затираем
            keep = {i: v for i, v in enumerate(vals) if v not in ("", None) and not str(v).startswith("=")} \
                if col == "CK" else {}
            ws.Range(f"{col}3:{col}{last}").Formula2 = f
            for i, v in keep.items():
                ws.Range(f"{col}{3 + i}").Value = v
            log.append(f"«{ODN}» {col}3:{col}{last} — формула")
    if not str(ws.Range("CJ1").Value or ""):
        ws.Range("CJ1").Value = "ДСУП"
    for col, (mark, old, new) in EDITS.items():
        n = bad = 0
        for a, b, f in runs(ws, col, 3, last):
            if not (isinstance(f, str) and f.startswith("=")) or mark in f:
                continue
            if old is None:
                g = new.format(f[1:])
            elif col == "AE" and '"нет мощности"' in f:
                k = f.find('"нет мощности"')         # условие «нет мощности» — ближайшее перед текстом
                j = f.rfind(old, 0, k)
                if j < 0:
                    bad += b - a + 1
                    continue
                g = f[:j] + new + f[j + len(old):]
            elif col != "AE" and old in f:
                g = f.replace(old, new, 1)
            else:
                bad += b - a + 1
                continue
            put(ws, f"{col}{a}:{col}{b}", g)
            n += b - a + 1
        if n:
            log.append(f"«{ODN}» {col} — ДСУП учтён: {n} строк")
        if bad:
            log.append(f"«{ODN}» {col} — формула не узнана в {bad} строках, не изменена")
    if opts:
        ws.Protect(PWD, *opts)


def run(path):
    path = os.path.abspath(path)
    app, pid = excel()
    log = []
    try:
        wb = app.Workbooks.Open(path, 0, False)
        if wb.ReadOnly:                              # открыта у другого — не трогаем
            wb.Close(False)
            raise SystemExit(f"{os.path.basename(path)}: открыта у другого пользователя — пропущена")
        app.Calculation = -4135
        spr(wb, log)
        odn(wb, log)
        app.Calculation = -4105
        app.CalculateFull()
        if log:
            wb.Save()
        wb.Close(False)
    finally:
        try:
            app.Quit()
        except Exception:
            pass
        try:
            os.kill(pid, signal.SIGTERM)
        except Exception:
            pass
    print(os.path.basename(path))
    for s in log or ["изменений нет"]:
        print("  " + s)


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(1)
    for p in sys.argv[1:]:
        run(p)
