# -*- coding: utf-8 -*-
"""Исправить сдвиг формул фазы в «Нагрузка_щитов» (ошибка шаблона v3.12).

    py fix_phase_rows.py <книга.xlsx> [...]        — исправить (книга должна быть закрыта)
    py fix_phase_rows.py --check <книга.xlsx> [...] — только показать

Ошибка: в строках 29–127 формулы G («Фаза (Итог)») и M ссылаются на ручную фазу E СЛЕДУЮЩЕЙ строки (G29 = …E30…),
поэтому с 27-й линии щита фаза берётся от соседней линии. Скрипт копирует в такие строки формулу из ближайшей
правильной строки выше. Свой скрытый экземпляр Excel; строки и столбцы не вставляются; повторный запуск ничего не меняет.
"""
import os
import re
import subprocess
import sys

SH = "Нагрузка_щитов"


def bad_rows(ws, col):
    out = []
    for i, (f,) in enumerate(ws.Range(f"{col}3:{col}1000").Formula):
        m = re.search(r"\$?E\$?(\d+)", f or "")
        if m and int(m.group(1)) != i + 3:
            out.append(i + 3)
    return out


def excel():
    import win32com.client
    import win32process
    app = win32com.client.DispatchEx("Excel.Application")
    app.Visible = False
    app.DisplayAlerts = False
    app.AutomationSecurity = 3
    return app, win32process.GetWindowThreadProcessId(app.Hwnd)[1]


def run(paths, check):
    app, pid = excel()
    try:
        for path in paths:
            wb = app.Workbooks.Open(os.path.abspath(path), 0, check)
            try:
                ws = wb.Worksheets(SH)
                prot = ws.ProtectContents
                if prot and not check:
                    ws.Unprotect("sx")
                total = 0
                for col in ("G", "M"):
                    rows = bad_rows(ws, col)
                    total += len(rows)
                    if check:
                        continue
                    for r in rows:
                        src = r - 1
                        while src in rows:
                            src -= 1
                        ws.Range(f"{col}{src}").Copy(ws.Range(f"{col}{r}"))
                if check:
                    print(f"{os.path.basename(path)}: неверных формул {total}")
                    continue
                if prot:
                    ws.Protect("sx", DrawingObjects=True, Contents=True, Scenarios=True, AllowFormattingColumns=True)
                app.CalculateFull()
                wb.Save()
                print(f"{os.path.basename(path)}: исправлено формул {total}")
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
    args = sys.argv[1:]
    chk = "--check" in args
    run([a for a in args if a != "--check"], chk)
