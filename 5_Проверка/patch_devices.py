# -*- coding: utf-8 -*-
"""Коды аппаратов «Однолинейки» (BH:BJ по I:K): + БП KNX («PS-KNX…», «ZPSU…», «БП KNX…») = PK,
диммер («DIM…», «Диммер…») = DIM. SXDRAW v3.20 рисует их значками SX_PSK / SX_DIM4.

    py patch_devices.py <книга.xlsx> [...]   — книга закрыта; повторный запуск ничего не портит
"""
import os
import subprocess
import sys

PWD = "sx"


def excel():
    import win32com.client
    import win32process
    app = win32com.client.DispatchEx("Excel.Application")
    app.Visible = False
    app.DisplayAlerts = False
    app.AskToUpdateLinks = False
    app.AutomationSecurity = 3
    return app, win32process.GetWindowThreadProcessId(app.Hwnd)[1]


def patch(wb):
    ws = wb.Worksheets("Однолинейка")
    prot = ws.ProtectContents
    if prot:
        p = ws.Protection
        opts = dict(DrawingObjects=True, Contents=True, Scenarios=True,
                    AllowFormattingCells=p.AllowFormattingCells, AllowFormattingColumns=p.AllowFormattingColumns,
                    AllowFormattingRows=p.AllowFormattingRows, AllowSorting=p.AllowSorting,
                    AllowFiltering=p.AllowFiltering, AllowUsingPivotTables=p.AllowUsingPivotTables)
        ws.Unprotect(PWD)
    n = 0
    for col, src in (("BH", "I"), ("BI", "J"), ("BJ", "K")):
        fs = [x[0] for x in ws.Range(f"{col}3:{col}1000").Formula]
        for i, f in enumerate(fs):
            r = 3 + i
            old = f'IF(LEFT(UPPER({src}{r}),2)="БП","PS",'
            if not (isinstance(f, str) and old in f) or '"PK"' in f:
                continue
            new = (f'IF(OR(LEFT(UPPER({src}{r}),6)="PS-KNX",LEFT(UPPER({src}{r}),4)="ZPSU",'
                   f'LEFT(UPPER({src}{r}),6)="БП KNX"),"PK",'
                   f'IF(OR(LEFT(UPPER({src}{r}),3)="DIM",LEFT(UPPER({src}{r}),6)="ДИММЕР"),"DIM",') + old
            ws.Range(f"{col}{r}").Formula = f.replace(old, new) + "))"
            n += 1
    if prot:
        ws.Protect(PWD, **opts)
    return n


def run(paths):
    app, pid = excel()
    try:
        for path in paths:
            wb = app.Workbooks.Open(os.path.abspath(path), 0, False)
            try:
                app.Calculation = -4135
                n = patch(wb)
                app.Calculation = -4105
                app.CalculateFull()
                wb.Save()
                print(f"{os.path.basename(path)}: формул кодов аппаратов обновлено {n}")
            finally:
                wb.Close(False)
    finally:
        try:
            app.Quit()
        except Exception:
            pass
        subprocess.run(["taskkill", "/PID", str(pid), "/F"], capture_output=True)


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    run(sys.argv[1:])
