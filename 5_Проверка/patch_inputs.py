# -*- coding: utf-8 -*-
"""Питание по вводам: у Ввода 2 и Ввода 3 свой ΔU, Zп, Iкз, In (книги v3.12 после merge_network.py).

    py patch_inputs.py <книга.xlsx> [...]   — книга закрыта; повторный запуск ничего не портит

«Сеть_щита», справа от раздела А:
  U4:AC11 — «Вводы 2 и 3»: по щиту 1–3 строки «Ввод 2», «Ввод 3»: ввод W ΔU до шин, X Zп на шинах, Y Iкз max, Z In
            фидера, AA файл / фидер (пишет «Связать с вышестоящим» или вручную); итог AB Zп (+ переходные), AC Iкз min.
            Ввод 1 — как раньше: строки 6–8 (C:F) и разделы 2–3.
  AE4:AJ14 — служебная таблица «для линий» (щит | секция): ΔU, Zп, Iкз max, которые берёт линия.
            Один ввод — Ввод 1. Два-три ввода без АВР — ввод своей секции («Исходные данные» BB).
            С АВР (C66 = да) — худший из вводов: больший Zп и ΔU (линия может остаться на любом вводе), больший Iкз max.
            Нет данных Ввода 2 / 3 — берётся Ввод 1 (как было).
«Исходные данные» BG (ΔU от ВРУ), BL (Zп в конце), BM (Iкз max), V (сечение по остатку ΔU) — от этой таблицы.
Свой скрытый экземпляр Excel; закрывается только он. COM — только позиционные аргументы.
"""
import os
import re
import signal
import sys

PWD = "sx"
NET, SRC = "Сеть_щита", "Исходные данные"
MARK = "ВВОДЫ 2 И 3"
EXP_MARK = "ДЛЯ ВЫШЕСТОЯЩЕГО ПО ВВОДАМ"
EXP_ROW = 46                                   # щит 1 Ввод 1; строка = 46 + (щит-1)*2 + (ввод-1)
COLS = "CDE"                                   # щит 1 / 2 / 3 в разделе А
IN_ROW = 6                                     # первая строка таблицы вводов (щит 1, Ввод 2)
EFF_ROW = 6                                    # первая строка таблицы «для линий»
# R1C1-шаблон поиска щита линии (как в формулах v3.12); группа 2 — выражение «щит линии»
LOOK = re.compile(r"INDEX\(ПЩ_(dU|Z|Iкз3),MATCH\((INDEX\(Разбивка_по_щитам!R1C6:R1000C6,"
                  r"MATCH\((RC\[-?\d+\]|RC\d+),Разбивка_по_щитам!R1C8:R1000C8,0\)\)),ПЩ_щиты,0\)\)")
SEC = 'MAX(1,MIN(3,IFERROR(--RC54,1)))'        # секция линии: «Исходные данные» BB (пусто = 1)


def excel():
    import win32com.client
    import win32process
    app = win32com.client.DispatchEx("Excel.Application")
    app.Visible = False
    app.DisplayAlerts = False
    app.AskToUpdateLinks = False
    app.AutomationSecurity = 3
    return app, win32process.GetWindowThreadProcessId(app.Hwnd)[1]


def rgb(t):
    return t[0] + t[1] * 256 + t[2] * 65536


def in_row(s, v):
    """строка таблицы вводов: щит s = 1..3, ввод v = 2..3"""
    return IN_ROW + (s - 1) * 2 + (v - 2)


def build_table(wb, log):
    ws = wb.Worksheets(NET)
    if ws.Range("U4").Value == MARK:
        return False
    if wb.Application.WorksheetFunction.CountA(ws.Range("U4:AJ14")) > 0:
        raise RuntimeError(f"«{NET}» U4:AJ14 занято")
    ws.Range("U4").Value = MARK
    ws.Range("U4").Font.Bold = True
    ws.Range("V4").Value = "(Ввод 1 — строки 6–8 и разделы 2–3; пусто = как Ввод 1; «Связать с вышестоящим» заполнит сама)"
    ws.Range("V4").Font.Italic = True
    hdr = ["Щит", "Ввод", "ΔU от ВРУ до шин, %", "Zп ф–0 на шинах, Ом", "Iкз max (3ф), кА", "In фидера, А",
           "Файл / фидер", "Zп итог, Ом", "Iкз min, А"]
    ws.Range("U5:AC5").Value = hdr
    for s in (1, 2, 3):
        c = COLS[s - 1]
        for v in (2, 3):
            r = in_row(s, v)
            ws.Range(f"U{r}").Formula = f"=IF($B${5 + s}=\"\",\"\",$B${5 + s})"
            ws.Range(f"V{r}").Value = v
            ws.Range(f"AB{r}").Formula = f'=IF(OR(U{r}="",N(X{r})<=0),"",X{r}+R_пер)'
            ws.Range(f"AC{r}").Formula = f'=IF(ISNUMBER(AB{r}),c_min*(Uн_лин*1000/SQRT(3))/AB{r},"")'
    inp = ws.Range(f"W{IN_ROW}:AA{IN_ROW + 5}")
    inp.Interior.Color = rgb((255, 242, 204))                    # жёлтое — ввод
    ws.Range("U5:AC5").Font.Bold = True
    ws.Range("U5:AC5").WrapText = True
    ws.Range("U5:AC5").Interior.Color = rgb((221, 235, 247))
    ws.Range(f"U5:AC{IN_ROW + 5}").Borders.LineStyle = 1
    # служебная: что берёт линия щита s в секции k
    ws.Range("AE4").Value = "ДЛЯ ЛИНИЙ (служебное): щит | секция -> ΔU, Zп, Iкз max"
    ws.Range("AE4").Font.Italic = True
    ws.Range("AE5:AJ5").Value = ["ключ", "щит", "секция", "ΔU, %", "Zп, Ом", "Iкз max, кА"]
    r = EFF_ROW
    for s in (1, 2, 3):
        c = COLS[s - 1]
        r2, r3 = in_row(s, 2), in_row(s, 3)
        for k in (1, 2, 3):
            ws.Range(f"AE{r}").Formula = f'=IF(AF{r}="","",AF{r}&"|"&AG{r})'
            ws.Range(f"AF{r}").Formula = f'=IF($B${5 + s}="","",$B${5 + s})'
            ws.Range(f"AG{r}").Value = k
            # n_ — вводов, a_ — АВР; v_1..v_3 — значения вводов (пусто у 2/3 -> как ввод 1); имена не как адреса ячеек
            for col, v1, v2, v3, worst in (("AH", f"{c}$33", f"$W${r2}", f"$W${r3}", "MAX"),
                                          ("AI", f"{c}$34", f"$AB${r2}", f"$AB${r3}", "MAX"),
                                          ("AJ", f"{c}$36", f"$Y${r2}", f"$Y${r3}", "MAX")):
                ws.Range(f"{col}{r}").Formula2 = (
                    f'=IF(AF{r}="","",LET(n_,MAX(1,N({c}$65)),a_,LOWER(TRIM({c}$66&""))="да",'
                    f'v_1,IF(ISNUMBER({v1}),{v1},""),v_2,IF(ISNUMBER({v2}),{v2},v_1),v_3,IF(ISNUMBER({v3}),{v3},v_1),'
                    f'k_,MIN(AG{r},n_),'
                    f'IF(n_<2,v_1,IF(a_,IF(COUNT(v_1,v_2,IF(n_>=3,v_3,""))=0,"",'
                    f'{worst}(N(v_1),N(v_2),IF(n_>=3,N(v_3),0))),CHOOSE(k_,v_1,v_2,v_3)))))')
            r += 1
    ws.Range(f"AE5:AJ{r - 1}").Font.Color = rgb((128, 128, 128))
    for col, w in zip(("U", "V", "W", "X", "Y", "Z", "AA", "AB", "AC"), (10, 6, 11, 11, 10, 9, 26, 10, 10)):
        ws.Columns(col).ColumnWidth = w
    names = {"ВЩ_ключ": "AE", "ВЩ_dU": "AH", "ВЩ_Z": "AI", "ВЩ_Iкз3": "AJ"}
    for n, col in names.items():
        wb.Names.Add(n, f"='{NET}'!${col}${EFF_ROW}:${col}${EFF_ROW + 8}")
    log.append(f"«{NET}»: таблица «Вводы 2 и 3» U4:AC11, служебная AE4:AJ14, имена ВЩ_*")
    return True


def build_export(wb, log):
    """Что отдать вышестоящему по каждому вводу: в норме — своя секция («Нагрузка_щитов» AG / AH),
    ток фидера — весь щит при АВР (аварийный режим), без АВР — своя секция. Один ввод — весь щит."""
    ws = wb.Worksheets(NET)
    if ws.Range("U44").Value == EXP_MARK:
        return False
    if wb.Application.WorksheetFunction.CountA(ws.Range("U44:AC51")) > 0:
        raise RuntimeError(f"«{NET}» U44:AC51 занято")
    ws.Range("U44").Value = EXP_MARK
    ws.Range("U44").Font.Bold = True
    ws.Range("V44").Value = "(«Связать с вышестоящим» берёт отсюда; в норме — своя секция, ток — аварийный при АВР)"
    ws.Range("V44").Font.Italic = True
    ws.Range("U45:AC45").Value = ["Щит", "Ввод", "Руст, кВт", "Руст быт, кВт", "Руст сил, кВт", "Рр сил, кВт",
                                  "Руст пост, кВт", "cos φ", "Iр фидера, А"]
    N = "Нагрузка_щитов!"
    for s_ in (1, 2, 3):
        c = COLS[s_ - 1]
        b = (s_ - 1) * 9                         # блоки «Нагрузка_щитов»: строки 4–12, 13–21, 22–30
        for v in (1, 2):
            r = EXP_ROW + (s_ - 1) * 2 + (v - 1)
            sec = "AG" if v == 1 else "AH"
            ws.Range(f"U{r}").Formula = f'=IF($B${5 + s_}="","",$B${5 + s_})'
            ws.Range(f"V{r}").Value = v

            def pick(row):
                # один ввод: Ввод 1 = весь щит, Ввод 2 пусто; два ввода: своя секция
                one = f'N({N}$AF${row + b})' if v == 1 else '""'
                return (f'=IF(OR(U{r}="",ISERROR({N}$R${4 + b})),"",'
                        f'IF(MAX(1,N({c}$65))<2,{one},N({N}${sec}${row + b})))')
            ws.Range(f"W{r}").Formula = pick(5)
            ws.Range(f"X{r}").Formula = pick(6)
            ws.Range(f"Z{r}").Formula = pick(8)
            ws.Range(f"AA{r}").Formula = pick(9)
            ws.Range(f"Y{r}").Formula = f'=IF(ISNUMBER(W{r}),W{r}-N(X{r})-N(AA{r}),"")'
            ws.Range(f"AB{r}").Formula = f'=IF(ISNUMBER(W{r}),N({N}$V${8 + b}),"")'
            ws.Range(f"AC{r}").Formula = (
                f'=IF(NOT(ISNUMBER(W{r})),"",IF(OR(MAX(1,N({c}$65))<2,LOWER(TRIM({c}$66&""))="да"),'
                f'N({N}$AF${12 + b}),N({N}${sec}${12 + b})))')
    ws.Range("U45:AC45").Font.Bold = True
    ws.Range("U45:AC45").WrapText = True
    ws.Range("U45:AC45").Interior.Color = rgb((221, 235, 247))
    ws.Range(f"U45:AC{EXP_ROW + 5}").Borders.LineStyle = 1
    log.append(f"«{NET}»: таблица «{EXP_MARK}» U44:AC51")
    return True


def runs(ws, col, r0=3, r1=1000):
    fs = [x[0] for x in ws.Range(f"{col}{r0}:{col}{r1}").Formula2R1C1]
    out, start = [], r0
    for i in range(1, len(fs) + 1):
        if i == len(fs) or fs[i] != fs[i - 1]:
            out.append((start, r0 + i - 1, fs[i - 1]))
            start = r0 + i
    return out


def rule(f):
    if not isinstance(f, str) or "ПЩ_" not in f or "ВЩ_" in f:
        return None
    g, n = LOOK.subn(lambda m: f'INDEX(ВЩ_{m.group(1)},MATCH({m.group(2)}&"|"&{SEC},ВЩ_ключ,0))', f)
    return g if n else None


def patch_lines(wb, log):
    ws = wb.Worksheets(SRC)
    for col in ("V", "BG", "BL", "BM"):
        n = 0
        for a, b, f in runs(ws, col):
            g = rule(f)
            if g:
                ws.Range(f"{col}{a}:{col}{b}").Formula2R1C1 = g
                n += b - a + 1
        if n:
            log.append(f"«{SRC}» {col}: {n} строк — от ввода секции")


def run(path):
    path = os.path.abspath(path)
    app, pid = excel()
    log = []
    try:
        wb = app.Workbooks.Open(path, 0, False)
        app.Calculation = -4135
        try:
            wb.Worksheets(NET)
        except Exception:
            raise RuntimeError(f"нет листа «{NET}» — сначала merge_network.py")
        build_table(wb, log)
        build_export(wb, log)
        patch_lines(wb, log)
        app.Calculation = -4105
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
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(1)
    for p in sys.argv[1:]:
        run(p)
