# -*- coding: utf-8 -*-
"""Характеристика автомата B / C / D — сама, по Iкз (ПУЭ 1.7.79: отключение ≤ 0,4 с).

    py patch_char.py <книга.xlsx> [...]      — книга закрыта; повторный запуск ничего не портит

Было: «Однолинейка» H при E = «авто» — всегда « С», у двигателей « D». Проверка КЗ («Исходные данные» BK)
уже понимала B (5·In), C (10·In), D (20·In), но выбора не было: B вписывали в E вручную.
Стало (E = «авто» или «+»; текст в E — как есть):
  двигатель                                    -> D (как было);
  в группе УЗО 30 мА                           -> C;
  мин. Iкз.min линий группы / In ≥ 10          -> C;
  5 ≤ мин. Iкз.min / In < 10                    -> B;
  < 5 или Iкз не посчитан                       -> C (проверка покажет «мало» / «не проверено»).
Новые служебные столбцы «Однолинейки» (справа, после CB):
  CC — Iкз.min в конце линии, А («Исходные данные» BI по линии R);
  CD — кратность группы: мин. CC по группе (AF, щит A) / номинал AQ (номинал до текста H — без цикла).
Свой скрытый экземпляр Excel; закрывается только он.
"""
import os
import signal
import sys

PWD = "sx"
ODN = "Однолинейка"
COLS = {   # столбец: (заголовок, формула для строки 3)
    "CC": ("Iкз min\nв конце линии, А",
           "=IF($R3=\"\",\"\",IFERROR(--INDEX('Исходные данные'!$BI$1:$BI$1000,"
           "MATCH($R3,'Исходные данные'!$D$1:$D$1000,0)),\"\"))"),
    "CD": ("Кратность\nIкз/In группы\n(B или C)",
           "=IF(OR($E3=\"\",N($AQ3)=0),\"\",IFERROR(MINIFS($CC$3:$CC$1000,$AF$3:$AF$1000,$AF3,"
           "$A$3:$A$1000,$A3)/$AQ3,\"\"))"),
}
OLD = ',FALSE)," D"," С")'
NEW = ',FALSE)," D",IF(AND(N(RC45)<>1,N(RC82)>=5,N(RC82)<10)," B"," С"))'   # RC45 = AS (30 мА), RC82 = CD


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


def patch(wb, log):
    ws = wb.Worksheets(ODN)
    opts = unprotect(ws)
    last = ws.Cells(ws.Rows.Count, "H").End(-4162).Row
    last = max(last, ws.Cells(ws.Rows.Count, "AF").End(-4162).Row)
    # служебные столбцы
    for col, (hdr, f) in COLS.items():
        h = ws.Range(f"{col}2")
        if str(h.Value or "").strip() not in ("", hdr):
            raise SystemExit(f"{ODN} {col}2 занят: «{h.Value}» — книга не изменена")
        if h.Value != hdr:
            ws.Range("CB2").Copy()
            h.PasteSpecial(-4122)                               # формат заголовка как у CB
            h.Value = hdr
            ws.Range("CB3").Copy()
            ws.Range(f"{col}3:{col}{last}").PasteSpecial(-4122)
            ws.Columns(col).ColumnWidth = 11
            log.append(f"«{ODN}» {col}2 — «{hdr.replace(chr(10), ' ')}»")
        rg = ws.Range(f"{col}3:{col}{last}")
        cur = ws.Range(f"{col}3").Formula2
        if cur != f or ws.Range(f"{col}{last}").Formula2 == "":
            rg.Formula2 = f
            log.append(f"«{ODN}» {col}3:{col}{last} — формула")
    # H: выбор B
    n = 0
    for a, b, f in runs(ws, "H", 3, last):
        if isinstance(f, str) and OLD in f and "RC82" not in f:
            ws.Range(f"H{a}:H{b}").Formula2R1C1 = f.replace(OLD, NEW, 1)
            n += b - a + 1
    if n:
        log.append(f"«{ODN}» H — характеристика B по кратности CD: {n} строк")
    elif not any("RC82" in str(f) for _, _, f in runs(ws, "H", 3, last)):
        log.append(f"«{ODN}» H — формула не узнана, выбор B не вставлен")
    if opts:
        ws.Protect(PWD, *opts)


def run(path):
    path = os.path.abspath(path)
    app, pid = excel()
    log = []
    try:
        wb = app.Workbooks.Open(path, 0, False)
        app.Calculation = -4135
        patch(wb, log)
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
