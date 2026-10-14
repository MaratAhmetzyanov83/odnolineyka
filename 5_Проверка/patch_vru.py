# -*- coding: utf-8 -*-
"""«Этот щит — ВРУ здания»: потери 4 % (СП 256) для линий считаются от его шин, даже если щит связан с вышестоящим.

    py patch_vru.py <книга.xlsx> [...]            — книга закрыта; повторный запуск ничего не портит
    py patch_vru.py <книга.xlsx> --да [щит 1–3]   — то же и включить для щита (ГРЩ / ВРУ отдельного здания)

Было: связанный щит всегда прибавлял к линиям ΔU от вышестоящего («Сеть_щита» C6) — ВРУ здания, питающийся
от ГРЩ другого здания (кабель 153 м), получал ΔU внешней сети в норму 4 % своих линий. Вариант «ТП» в C12
работал только без связи.
«Сеть_щита» строка 26: B26 — вопрос, C26:E26 — «да» / пусто по щиту 1–3 (ввод).
C33:E33 (ΔU от ВРУ до шин → в линии) и служебная AH6:AH14 (ВЩ_dU, по вводам) — 0 при «да».
Zп и Iкз (C34, C36, AI, AJ) не меняются — КЗ по-прежнему от вышестоящего по связи.
Внешний кабель проверяется своей линией в книге вышестоящего.
Свой скрытый экземпляр Excel; закрывается только он. COM — только позиционные аргументы.
"""
import os
import re
import signal
import sys

PWD = "sx"
NET = "Сеть_щита"
ASK = "Этот щит — ВРУ здания: ΔU линий (4 %) от его шин (да / пусто)"
SCH = "CDE"


def excel():
    import win32com.client
    import win32process
    app = win32com.client.DispatchEx("Excel.Application")
    app.Visible = False
    app.DisplayAlerts = False
    app.AskToUpdateLinks = False
    app.AutomationSecurity = 3
    return app, win32process.GetWindowThreadProcessId(app.Hwnd)[1]


def flag(c):
    return f'LOWER(TRIM(${c}$26&""))="да"'


def build(ws, log):
    if ws.Range("B26").Value == ASK:
        return
    if ws.Application.WorksheetFunction.CountA(ws.Range("B26:E26")) > 0:
        raise RuntimeError(f"«{NET}» B26:E26 занято")
    ws.Range("B26").Value = ASK
    ws.Range("B26").WrapText = True
    for c in SCH:
        cell = ws.Range(f"{c}26")
        cell.Locked = False
        cell.Interior.Color = ws.Range(f"{c}12").Interior.Color     # как ввод раздела 2
        v = ws.Range(f"{c}26").Validation
        v.Delete()
        v.Add(3, 1, 1, "да")
    n = 0
    for c in SCH:
        f = ws.Range(f"{c}33").Formula
        if "$26" not in f and f.startswith("="):
            ws.Range(f"{c}33").Formula = f'=IF({flag(c)},0,{f[1:]})'
            n += 1
    m = 0
    for r in range(6, 15):
        f = ws.Range(f"AH{r}").Formula
        cols = re.findall(r"([C-E])\$33", f)
        if f.startswith("=") and cols and "$26" not in f:
            c = cols[0]
            ws.Range(f"AH{r}").Formula = f'=IF({flag(c)},IF(AF{r}="","",0),{f[1:]})'
            m += 1
    log.append(f"«{NET}» B26:E26 — «ВРУ здания»; C33:E33 — {n}, AH6:AH14 — {m}")


def run(path, on=None):
    path = os.path.abspath(path)
    app, pid = excel()
    log = []
    try:
        wb = app.Workbooks.Open(path, 0, False)
        ws = wb.Worksheets(NET)
        p = ws.Protection
        opts = (True, True, True, False, p.AllowFormattingCells, p.AllowFormattingColumns, p.AllowFormattingRows,
                p.AllowInsertingColumns, p.AllowInsertingRows, p.AllowInsertingHyperlinks, p.AllowDeletingColumns,
                p.AllowDeletingRows, p.AllowSorting, p.AllowFiltering, p.AllowUsingPivotTables) if ws.ProtectContents else None
        if opts:
            ws.Unprotect(PWD)
        build(ws, log)
        if on:
            ws.Range(f"{SCH[on - 1]}26").Value = "да"
            log.append(f"щит {on} («{ws.Range(f'{SCH[on - 1]}28').Text}») — ВРУ здания: да")
        if opts:
            ws.Protect(PWD, *opts)
        app.CalculateFull()
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
    args = sys.argv[1:]
    if not args:
        print(__doc__)
        sys.exit(1)
    on = None
    if "--да" in args:
        i = args.index("--да")
        on = int(args[i + 1]) if i + 1 < len(args) and args[i + 1].isdigit() else 1
        args = [a for j, a in enumerate(args) if j not in (i, i + 1) or (j == i + 1 and not a.isdigit())]
        args = [a for a in args if a != "--да"]
    for p in args:
        run(p, on)
