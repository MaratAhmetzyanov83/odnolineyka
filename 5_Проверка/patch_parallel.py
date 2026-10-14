# -*- coding: utf-8 -*-
"""Параллельные кабели линии: «Исходные данные», новый столбец «Кабелей параллельно n» (пусто = 1).

    py patch_parallel.py <книга.xlsx> [...]                      — книга закрыта; повторный запуск ничего не портит
    py patch_parallel.py <книга.xlsx> --линия Щ.0.01.1=2 [...]   — то же и вписать n для линий

Было: линию на 400–600 А одним кабелем не подобрать — сечение «Не исп», нет ΔU, Zп и Iкз у нижестоящих щитов.
Параллельные кабели были только у питания от ТП («Сеть_щита», раздел 2, строка 21).
Стало (n — число кабелей, ток делится поровну):
  V  сечение — по току In / n и по потерям Iр / n;
  AA потеря напряжения — / n;  BL Zп в конце линии — 2ρL / (S·n);  BM Iкз max — ρL / (S·n);
  AE/AO марка — «2хВВГнг(А)-LS (5х150)»;
  «Однолинейка» AB макс. автомат по кабелю — Iдоп × n;
  «Сеть_щита» раздел Б, проверка фидера — термостойкость по сечению S·n.
Новый столбец — справа от последнего заголовка «Исходных данных» (ищется по заголовку).
Свой скрытый экземпляр Excel; закрывается только он. COM — только позиционные аргументы.
"""
import os
import re
import signal
import sys

PWD = "sx"
ISH = "Исходные данные"
HDR = "Кабелей параллельно n, шт. (пусто = 1)"


def excel():
    import win32com.client
    import win32process
    app = win32com.client.DispatchEx("Excel.Application")
    app.Visible = False
    app.DisplayAlerts = False
    app.AskToUpdateLinks = False
    app.AutomationSecurity = 3
    return app, win32process.GetWindowThreadProcessId(app.Hwnd)[1]


def colname(n):
    s = ""
    while n:
        n, k = divmod(n - 1, 26)
        s = chr(65 + k) + s
    return s


def unprotect(ws):
    if not ws.ProtectContents:
        return None
    p = ws.Protection
    opts = (True, True, True, False, p.AllowFormattingCells, p.AllowFormattingColumns, p.AllowFormattingRows,
            p.AllowInsertingColumns, p.AllowInsertingRows, p.AllowInsertingHyperlinks, p.AllowDeletingColumns,
            p.AllowDeletingRows, p.AllowSorting, p.AllowFiltering, p.AllowUsingPivotTables)
    ws.Unprotect(PWD)
    return opts


def column(ws):
    """Столбец n: найти по заголовку или создать справа от последнего заголовка."""
    last = 0
    for c in range(1, 200):
        v = ws.Cells(2, c).Value
        if v is not None and str(v).strip() != "":
            last = c
            if str(v).strip() == HDR:
                return colname(c), False
    c = last + 1
    L = colname(c)
    ws.Cells(2, c).Value = HDR
    ws.Range("Q2").Copy()
    ws.Cells(2, c).PasteSpecial(-4122)                      # формат заголовка как у Q
    ws.Cells(2, c).Value = HDR
    ws.Cells(1, c).Value = c
    return L, True


def patch_rows(ws, L, log):
    """Формулы V, AA, AE, BL, BM во всех строках, где они есть."""
    last = ws.Cells(ws.Rows.Count, "V").End(-4162).Row
    n = {k: 0 for k in ("V", "AA", "AE", "BL", "BM")}
    for r in range(3, last + 1):
        P = f"MAX(1,N(${L}{r}))"
        tag = f"${L}{r}"
        # V — сечение
        c = ws.Range(f"V{r}")
        f = c.Formula2
        if f.startswith("=") and tag not in f:
            g = f.replace(f"(U{r}/(", f"(U{r}/{P}/(")
            g = g.replace("*i_c/u_f*100", f"*i_c/{P}/u_f*100")
            if g != f:
                c.Formula2 = g
                n["V"] += 1
        # AA — ΔU, В
        c = ws.Range(f"AA{r}")
        f = c.Formula2
        head = f"=IF($CB{r}=1,0,"
        if f.startswith(head) and tag not in f:
            c.Formula2 = head + f"1/{P}*" + f[len(head):]
            n["AA"] += 1
        # AE — марка
        c = ws.Range(f"AE{r}")
        f = c.Formula2
        if f.startswith("=") and tag not in f and "Кабель_обозначение" in f:
            c.Formula2 = f'=IF({P}>1,{P}&"х","")&' + f[1:]
            n["AE"] += 1
        # BL — Zп в конце линии
        c = ws.Range(f"BL{r}")
        f = c.Formula2
        old = f"2*ρкз_Cu*M{r}/V{r})"
        if old in f and tag not in f:
            c.Formula2 = f.replace(old, f"2*ρкз_Cu*M{r}/V{r}/{P})")
            n["BL"] += 1
        # BM — Iкз max
        c = ws.Range(f"BM{r}")
        f = c.Formula2
        old = f"0.0175*N(M{r})/V{r})"
        if old in f and tag not in f:
            c.Formula2 = f.replace(old, f"0.0175*N(M{r})/V{r}/{P})")
            n["BM"] += 1
    log.append(f"«{ISH}» строки 3–{last}: " + ", ".join(f"{k} {v}" for k, v in n.items()))


def patch_od(ws, L, log):
    last = ws.Cells(ws.Rows.Count, "AB").End(-4162).Row
    k = 0
    for r in range(3, last + 1):
        c = ws.Range(f"AB{r}")
        f = c.Formula2
        key = f"MATCH(VLOOKUP(R{r},Исходные_данные,22,FALSE),Iдоп_S,0))"
        tag = f"'Исходные данные'!${L}$1"
        if key in f and tag not in f:
            add = f"*MAX(1,N(INDEX('Исходные данные'!${L}$1:${L}$1000,MATCH(R{r},'Исходные данные'!$D$1:$D$1000,0))))"
            c.Formula2 = f.replace(key, key + add, 1)
            k += 1
    log.append(f"«Однолинейка» AB (Iдоп × n): {k}")


def patch_net(ws, L, log):
    k = 0
    for r in range(90, 200):
        c = ws.Range(f"O{r}")
        f = c.Formula2
        key = f"115*INDEX('Исходные данные'!$V$1:$V$1000,MATCH($B{r},'Исходные данные'!$D$1:$D$1000,0))"
        if key in f and f"${L}$1" not in f:
            add = f"*MAX(1,N(INDEX('Исходные данные'!${L}$1:${L}$1000,MATCH($B{r},'Исходные данные'!$D$1:$D$1000,0))))"
            c.Formula2 = f.replace(key, key + add)
            k += 1
    log.append(f"«Сеть_щита» раздел Б, проверка фидера (S × n): {k}")


def run(path, lines=None):
    path = os.path.abspath(path)
    app, pid = excel()
    log = []
    try:
        wb = app.Workbooks.Open(path, 0, False)
        app.Calculation = -4135
        ws = wb.Worksheets(ISH)
        prot = {}
        for name in (ISH, "Однолинейка", "Сеть_щита"):
            prot[name] = unprotect(wb.Worksheets(name))
        L, new = column(ws)
        if new:
            last = ws.Cells(ws.Rows.Count, "V").End(-4162).Row
            rg = ws.Range(f"{L}3:{L}{last}")
            rg.Locked = False
            rg.HorizontalAlignment = -4108
            v = rg.Validation
            v.Delete()
            v.Add(1, 1, 1, "1", "8")                            # целое 1…8
            ws.Columns(L).ColumnWidth = 12
            log.append(f"«{ISH}» {L}2 — «{HDR}»")
        patch_rows(ws, L, log)
        patch_od(wb.Worksheets("Однолинейка"), L, log)
        patch_net(wb.Worksheets("Сеть_щита"), L, log)
        for name, val in (lines or {}).items():
            rr = next((r for r in range(3, 1001) if str(ws.Cells(r, 4).Value or "").strip() == name), None)
            if rr is None:
                log.append(f"линии «{name}» нет")
                continue
            ws.Range(f"{L}{rr}").Value = val
            log.append(f"{name}: кабелей параллельно {val}")
        for name, opts in prot.items():
            if opts:
                wb.Worksheets(name).Protect(PWD, *opts)
        app.Calculation = -4105
        app.CalculateFull()
        for name in (lines or {}):
            rr = next((r for r in range(3, 1001) if str(ws.Cells(r, 4).Value or "").strip() == name), None)
            if rr:
                log.append(f"  {name}: In {ws.Range(f'U{rr}').Text} А, {ws.Range(f'AO{rr}').Text.replace(chr(10), ' ')}, "
                           f"Zп {ws.Range(f'BL{rr}').Text}")
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
    lines = {}
    if "--линия" in args:
        i = args.index("--линия")
        for a in args[i + 1:]:
            m = re.match(r"(.+)=(\d+)$", a)
            if m:
                lines[m.group(1)] = int(m.group(2))
        args = [a for a in args[:i]]
    for p in args:
        run(p, lines)
