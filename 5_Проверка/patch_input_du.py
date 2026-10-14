# -*- coding: utf-8 -*-
"""ΔU от ВРУ в строке ввода щита — без двойного учёта кабеля ввода.

    py patch_input_du.py <книга.xlsx> [...]   — книга закрыта; повторный запуск ничего не портит

«Исходные данные» BG (ΔU от ВРУ, %) = ΔU линии (AB) + ΔU от ВРУ до шин щита («Сеть_щита», ВЩ_dU).
Для строки ввода (AH = «Ввод») кабель ввода — та же питающая линия, что уже вошла в ΔU до шин
(«Связать с вышестоящим» / «Сеть_щита» стр. 33 = ΔU вышестоящего + ΔU фидера). Было: ввод считался дважды
(пример: ЩУОВ 3,21 % в ГРЩ + 3,09 % ввод = 6,3 %, а на шинах ЩУОВ 3,21 %).
Стало: строка ввода — ΔU до шин щита; если до шин не задано (0) — ΔU самого кабеля ввода (как раньше).
Линии щита не меняются: AB + ΔU до шин.
Свой скрытый экземпляр Excel; закрывается только он. COM — только позиционные аргументы.
"""
import os
import re
import signal
import sys

SRC = "Исходные данные"
PWD = "sx"
R1, R2 = 3, 1000


def excel():
    import win32com.client
    import win32process
    app = win32com.client.DispatchEx("Excel.Application")
    app.Visible = False
    app.DisplayAlerts = False
    app.AskToUpdateLinks = False
    app.AutomationSecurity = 3
    return app, win32process.GetWindowThreadProcessId(app.Hwnd)[1]


def new_formula(f, r):
    """=IF(ISNUMBER(ABr),ABr+X,"") -> =IF(ISNUMBER(ABr),IF(AND($AHr="Ввод",N(X)>0),X,ABr+X),"")"""
    head, tail = f"=IF(ISNUMBER(AB{r}),AB{r}+", ',"")'
    if not (f.startswith(head) and f.endswith(tail)) or f'$AH{r}="Ввод"' in f:
        return None
    x = f[len(head):-len(tail)]
    return f'=IF(ISNUMBER(AB{r}),IF(AND($AH{r}="Ввод",N({x})>0),{x},AB{r}+{x}),"")'


def patch(wb, log):
    ws = wb.Worksheets(SRC)
    rng = ws.Range(f"BG{R1}:BG{R2}")
    old = rng.Formula
    out, n = [], 0
    for i, (f,) in enumerate(old):
        r = R1 + i
        g = new_formula(f, r) if isinstance(f, str) else None
        if g:
            n += 1
        out.append((g or f,))
    if n:
        rng.Formula = tuple(out)
    log.append(f"«{SRC}» BG{R1}:BG{R2} — строка ввода без двойного учёта: {n}")


def run(path):
    path = os.path.abspath(path)
    app, pid = excel()
    log = []
    try:
        wb = app.Workbooks.Open(path, 0, False)
        ws = wb.Worksheets(SRC)
        prot = ws.ProtectContents
        if prot:
            ws.Unprotect(PWD)
        patch(wb, log)
        if prot:
            ws.Protect(PWD, True, True, True)
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
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(1)
    for p in sys.argv[1:]:
        run(p)
