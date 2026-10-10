# -*- coding: utf-8 -*-
"""Сечение и потери в «Исходных данных» v3.12 (перенесено из домашней 1.08, 10.10.2026).

    py patch_sections.py <книга.xlsx> [...]    — книга должна быть закрыта; повторный запуск ничего не портит

1. Линии до 50 В (в F «Напряжение» 12 / 24 / 48): потери считаются от тока на этом напряжении
   (Р·1000 / U) и в процентах от него; «Уровень напряжения» = F. Ток T и автомат не меняются
   (автомат стоит до блока питания, на 230 В).
2. Сечение V = наибольшее из: по току (как было), минимальное по коду (Q), по потерям —
   наименьшее сечение, при котором потери ≤ «ΔU линии доп.» (4,5 %), для линий до 50 В —
   ≤ «ΔU доп. до 50 В» (5 %). Значения — «Справочная» BQ32:BQ33.
   Своё число в Q «Сечение (ручное)» (не формула) — главнее, как раньше.
3. Q «мин. сечение по коду» дотягивается на пустые строки ниже 78-й (в шаблоне формула была до 78-й).
Свой скрытый экземпляр Excel; закрывается только он.
"""
import os
import subprocess
import sys

SRC = "Исходные данные"
REF = "Справочная"
PWD = "sx"
LV = "AND(ISNUMBER(RC6),RC6>0,RC6<=50)"

AA = ("=RC25*(RC23*RC13/RC22*RC16+RC24*RC13*RC26/1000)*IF(" + LV + ",N(RC5)*1000/RC6,RC20)")
AB = "=ROUNDUP(RC27/IF(" + LV + ",RC6,(Uн_лин/SQRT(3))*1000)*100,2)"
AC = "=IF(" + LV + ",RC6,IFS(RC8=3,Uн_лин*1000,RC8=1,ROUND((Uн_лин/SQRT(3))*1000,0)))"
LOSS = ("IFERROR(LET(lv_,{lv},i_c,IF(lv_,N(RC5)*1000/RC6,N(RC20)),u_f,IF(lv_,RC6,Uн_лин/SQRT(3)*1000),"
        "lim_,IF(lv_,dU_НН_доп,dU_лин_доп),s_k,INDEX(Кабель_сопротивление,0,1),x_k,INDEX(Кабель_сопротивление,0,2),"
        "r_k,INDEX(Кабель_сопротивление,0,3),d_u,RC25*(r_k*N(RC13)/s_k*RC16+x_k*N(RC13)*RC26/1000)*i_c/u_f*100,"
        "MIN(FILTER(s_k,d_u<=lim_,MAX(s_k)))),0)").format(lv=LV)
MARK = "dU_лин_доп"


def new_v(f):
    """IF(RC17=0,<по току>,RC17) -> своё Q главнее, иначе MAX(по току, Q по коду, по потерям)."""
    key = "IF(RC[-5]=0,"
    i = f.find(key)
    if i < 0 or MARK in f:
        return None
    j, depth = i + len(key), 0
    k = j
    while k < len(f):
        ch = f[k]
        if ch == "(":
            depth += 1
        elif ch == ")":
            depth -= 1
        elif ch == "," and depth == 0:
            break
        k += 1
    ci = f[j:k]
    tail = ",RC[-5])"
    if not f.startswith(tail, k):
        return None
    rep = ("IF(AND(NOT(ISFORMULA(RC17)),ISNUMBER(RC17),N(RC17)>0),RC17,LET(c_i," + ci +
           ",IF(c_i=\"Не исп\",c_i,MAX(c_i,N(RC17)," + LOSS + "))))")
    return f[:i] + rep + f[k + len(tail):]


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
    opts = dict(DrawingObjects=True, Contents=True, Scenarios=True,
                AllowFormattingCells=p.AllowFormattingCells, AllowFormattingColumns=p.AllowFormattingColumns,
                AllowFormattingRows=p.AllowFormattingRows, AllowSorting=p.AllowSorting,
                AllowFiltering=p.AllowFiltering, AllowUsingPivotTables=p.AllowUsingPivotTables)
    ws.Unprotect(PWD)
    return opts


CAPTIONS = {
    "BP32": "Допустимая потеря напряжения в линии ΔUдоп, %",
    "BR32": "ΔUдоп — наибольшая потеря напряжения от щита до потребителя; сечение растёт, пока ΔU ≤ ΔUдоп "
            "(ГОСТ Р 50571.5.52, прил. G)",
    "BP33": "Допустимая потеря напряжения в линии до 50 В ΔUдоп.нн, %",
    "BR33": "ΔUдоп.нн — то же для линий с U ≤ 50 В (12 / 24 / 48 В, столбец F «Напряжение»)",
}


def settings(wb):
    """Допустимые потери на «Справочной» BP32:BR33 + имена dU_лин_доп, dU_НН_доп; подписи — всегда по ГОСТ."""
    ws = wb.Worksheets(REF)
    names = [n.Name for n in wb.Names]
    added = MARK not in names
    if added and any(ws.Range(a).Formula not in ("", None) for a in ("BP32", "BQ32", "BP33", "BQ33")):
        raise RuntimeError("«Справочная» BP32:BQ33 заняты — некуда записать допустимые потери")
    opts = unprotect(ws)
    for a, t in CAPTIONS.items():
        ws.Range(a).Value = t
    if added:
        ws.Range("BQ32").Value = 4.5
        ws.Range("BQ33").Value = 5
        ws.Range("BP31:BR31").Copy()
        ws.Range("BP32:BR33").PasteSpecial(-4122)   # xlPasteFormats
        wb.Application.CutCopyMode = False
        wb.Names.Add(Name="dU_лин_доп", RefersTo="='Справочная'!$BQ$32")
        wb.Names.Add(Name="dU_НН_доп", RefersTo="='Справочная'!$BQ$33")
    if opts:
        ws.Protect(PWD, **opts)
    return added


def runs(ws, col, r0=3, r1=1000):
    """Подряд идущие строки с одинаковой формулой R1C1: [(первая, последняя, формула)]."""
    fs = [x[0] for x in ws.Range(f"{col}{r0}:{col}{r1}").Formula2R1C1]
    out, start = [], r0
    for i in range(1, len(fs) + 1):
        if i == len(fs) or fs[i] != fs[i - 1]:
            out.append((start, r0 + i - 1, fs[i - 1]))
            start = r0 + i
    return out


def patch(wb):
    ws = wb.Worksheets(SRC)
    opts = unprotect(ws)
    n = {"Q": 0, "V": 0, "AA": 0, "AB": 0, "AC": 0}
    rules = {
        "V": lambda f: new_v(f),
        "AA": lambda f: AA if "RC6" not in f and "RC[-7]" in f else None,
        "AB": lambda f: AB if "RC6" not in f and "RC[-1]" in f else None,
        "AC": lambda f: AC if "RC6" not in f and "Uн_лин" in f else None,
    }
    for col, rule in rules.items():
        for a, b, f in runs(ws, col):
            nf = rule(f) if isinstance(f, str) and f.startswith("=") else None
            if nf:
                ws.Range(f"{col}{a}:{col}{b}").Formula2R1C1 = nf
                n[col] += b - a + 1
    q21 = ws.Range("Q21").FormulaR1C1
    if isinstance(q21, str) and q21.startswith("="):
        for a, b, f in runs(ws, "Q", 22):
            if f in ("", None):
                ws.Range(f"Q{a}:Q{b}").FormulaR1C1 = q21
                n["Q"] += b - a + 1
    if opts:
        ws.Protect(PWD, **opts)
    return n


def run(paths):
    app, pid = excel()
    try:
        for path in paths:
            wb = app.Workbooks.Open(os.path.abspath(path), 0, False)
            try:
                app.Calculation = -4135          # вручную на время правки
                s = settings(wb)
                n = patch(wb)
                app.Calculation = -4105
                app.CalculateFull()
                wb.Save()
                print(f"{os.path.basename(path)}: настройки {'добавлены' if s else 'уже были'}; формул: {n}")
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
