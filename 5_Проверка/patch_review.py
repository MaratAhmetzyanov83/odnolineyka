# -*- coding: utf-8 -*-
"""Замечания проверки 11.10.2026 (19–22) для книг v3.12 (после patch_validation.py).

    py patch_review.py <книга.xlsx> [...]   — книга закрыта; повторный запуск ничего не портит

19. Двигатели. Коды В, ПВ, ПУ, ЛФ, НС — двигатель («Справочная» T = 2). Ток с КПД и cos φ по таблице двигателей
    (ближайший больший типоразмер, а не точное совпадение мощности); пусковой ток = ток × Iп/Iн.
    Автомат двигателя — характеристика D; статус «Однолинейки»: «пуск … А ≥ k·In», если пуск попадает в зону
    мгновенного расцепителя (B 3, C 5, D 10 · In).
20. Отключающая способность. «Сеть_щита» строка 42: Icu аппаратов щита, кА (по умолчанию 6). Проверка в C38:E38:
    Iкз max на шинах > Icu — замечание (уходит в «Дерево щитов» и ПЗ).
21. Освещение общедомовых помещений: код LM («Освещение общедомовое»), группа «Пост» (Кс = 1, без Ко).
    Квартирное L остаётся «Быт». Значение Кс — сверить с СП 256.
22. Строка ввода: мощность («Исходные данные» AL -> схема, «В Акад») = Рр щита, как и ток (п. 15).
23. Длина линии («Исходные данные» M): опуск × (2·n − 1), n — «Кол-во потребителей» (G). Пустой G давал −1 —
    опуск вычитался (линия короче). Теперь пусто = 1. DLN заполняет G сам — его книги не меняются.
Плюс общие коды: Щ (распределительный щит, мин. 6 мм²), ЛФ (лифт), НС (насос).
Свой скрытый экземпляр Excel; закрывается только он. COM — только позиционные аргументы.
"""
import os
import signal
import sys

PWD = "sx"
SRC, ODN, REF, NET = "Исходные данные", "Однолинейка", "Справочная", "Сеть_щита"
MOTOR_CODES = ("В", "ПВ", "ПУ", "ЛФ", "НС")
NEW_CODES = [
    dict(P="Щ", Q="Распределительный щит", R=5, S=1, T=1, U="мс", V=5, W=400, X=3, Y=6,
         Z="Распределительные щиты жилой части", AV="Прочее (прямоуг. с подписью)", AW="нет", AZ="Быт"),
    dict(P="ЛФ", Q="Лифт", R=1, S=1, T=2, U="мс", V=5, W=400, X=3, Y=4, Z="Лифты",
         AV="Привод (двигатель)", AW="нет", AZ="Сил"),
    dict(P="НС", Q="Насос", R=1, S=1, T=2, U="мс", V=5, W=400, X=3, Z="Насосы (водоснабжение, дренаж)",
         AV="Привод (двигатель)", AW="нет", AZ="Сил"),
    dict(P="LM", Q="Освещение общедомовое", R=1, S=1, T=1, U="мс", V=3, W=230, X=1, Y=1.5, Z="Рабочее освещение",
         AV="Светильник", AW="нет", AX=10, AZ="Пост"),
]
# типоразмер двигателя: первый в таблице с мощностью ≥ E (RC5)
MOT = "INDEX(Справочная_двигатели,MATCH(TRUE,INDEX(INDEX(Справочная_двигатели,0,1)>=RC5,0),0),{c})"
IS_MOTOR = "IFERROR(VLOOKUP(RC34,Потребители_обозначение,5,FALSE)=2,FALSE)"
BOARD = "INDEX(Разбивка_по_щитам!R1C6:R1000C6,MATCH(RC4,Разбивка_по_щитам!R1C8:R1000C8,0))"
BOARD_P = (f"IFERROR(CHOOSE(MATCH({BOARD},ПЩ_щиты,0),"
           "Нагрузка_щитов!R10C32,Нагрузка_щитов!R19C32,Нагрузка_щитов!R28C32)*1,\"\")")


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


def runs(ws, col, r0=3, r1=1000):
    fs = [x[0] for x in ws.Range(f"{col}{r0}:{col}{r1}").Formula2R1C1]
    out, start = [], r0
    for i in range(1, len(fs) + 1):
        if i == len(fs) or fs[i] != fs[i - 1]:
            out.append((start, r0 + i - 1, fs[i - 1]))
            start = r0 + i
    return out


def apply(ws, col, rule, r0=3):
    n = 0
    for a, b, f in runs(ws, col, r0):
        g = rule(f) if isinstance(f, str) and f.startswith("=") else None
        if g and g != f:
            ws.Range(f"{col}{a}:{col}{b}").Formula2R1C1 = g
            n += b - a + 1
    return n


# ---------------------------------------------------------------- коды «Справочной»
def patch_codes(wb, log):
    ws = wb.Worksheets(REF)
    codes = {ws.Range(f"P{r}").Value: r for r in range(4, 52) if ws.Range(f"P{r}").Value}
    n = 0
    for c in MOTOR_CODES:
        r = codes.get(c)
        if r and ws.Range(f"T{r}").Value != 2:
            ws.Range(f"T{r}").Value = 2
            n += 1
    free = [r for r in range(4, 52) if not ws.Range(f"P{r}").Value]
    added = []
    for d in NEW_CODES:
        if d["P"] in codes:
            continue
        r = free.pop(0)
        for col, v in d.items():
            ws.Range(f"{col}{r}").Value = v
        added.append(d["P"])
    if n or added:
        log.append(f"«{REF}»: двигатель (T = 2) — {n} кодов" + (f"; новые коды: {', '.join(added)}" if added else ""))


# ---------------------------------------------------------------- 19 двигатели
def rule_p(f):
    old3 = "VLOOKUP(RC[-11],Справочная_двигатели,8,FALSE)"
    old1 = "VLOOKUP(RC[-11],Справочная_двигатели,10,TRUE)"
    if old3 not in f:
        return None
    m8, m10 = MOT.format(c=8), MOT.format(c=10)
    return f.replace(old3, m8).replace(old1, f"IF(N({m10})>0,{m10},{m8})")


def rule_t(f):
    key = "IFS(RC[-12]=1,RC[-15]/((Uн_лин/SQRT(3))*RC[-4]),RC[-12]=3,RC[-15]/(SQRT(3)*Uн_лин*RC[-4]))"
    if key not in f or "Справочная_двигатели" in f:
        return None
    eta = f"IFERROR(IF({IS_MOTOR},{MOT.format(c=6)},1),1)"
    return f.replace(key, f"{key}/{eta}")


def rule_ar(f):
    if "Справочная_двигатели,3,FALSE" not in f:
        return None
    r7, r9 = MOT.format(c=7), MOT.format(c=9)
    return (f"=IFERROR(IF({IS_MOTOR},N(RC20)*IF(RC8=1,IF(N({r9})>0,{r9},{r7}),{r7}),0),0)")


def rule_h(f):
    key = '&" С"&'
    if key not in f or '" D"' in f:
        return None
    motor = "IFERROR(VLOOKUP(RC38,Потребители_обозначение,5,FALSE)=2,FALSE)"
    return f.replace(key, f'&IF({motor}," D"," С")&', 1)


def rule_au(f):
    old = 'SUBSTITUTE(SUBSTITUTE(UPPER(RC[-39]),"С","C"),"А","A")'
    if old not in f or '" D"' in f:
        return None
    new = f'SUBSTITUTE(SUBSTITUTE({old}," D"," C")," B"," C")'
    return f.replace(old, new)


START = ('IFERROR(LET(ip_,N(INDEX(\'Исходные данные\'!R1C46:R1000C46,MATCH(RC[-13],\'Исходные данные\'!R1C4:R1000C4,0))),'
         'in_,N(RC[16]),k_,IF(ISNUMBER(SEARCH(" D",RC[-23])),10,IF(ISNUMBER(SEARCH(" B",RC[-23])),3,5)),'
         'IF(AND(ip_>0,in_>0,ip_>=k_*in_),"пуск "&FIXED(ip_,0,TRUE)&" А ≥ "&k_&"·In — нужна характеристика D или больше In","")),"")')


def rule_ae(f):
    if "пуск " in f or not f.endswith(")))"):
        return None
    return f[:-2] + "," + START + "))"


def patch_motors(wb, log):
    s = wb.Worksheets(SRC)
    o = wb.Worksheets(ODN)
    so, oo = unprotect(s), unprotect(o)
    n = {"P": apply(s, "P", rule_p), "T": apply(s, "T", rule_t), "AR": apply(s, "AR", rule_ar),
         "H": apply(o, "H", rule_h), "AU": apply(o, "AU", rule_au), "AE": apply(o, "AE", rule_ae)}
    if so:
        s.Protect(PWD, *so)
    if oo:
        o.Protect(PWD, *oo)
    if any(n.values()):
        log.append("19: двигатели — " + ", ".join(f"{k} {v}" for k, v in n.items() if v))


# ---------------------------------------------------------------- 20 Icu
def patch_icu(wb, log):
    ws = wb.Worksheets(NET)
    if ws.Range("B42").Value:
        return
    ws.Range("B42").Value = "Отключающая способность аппаратов щита Icu"
    ws.Range("C42:E42").Value = (6, 6, 6)
    ws.Range("C42:E42").Interior.Color = 255 + 242 * 256 + 204 * 65536     # светло-жёлтый — ввод
    ws.Range("F42").Value = "кА"
    ws.Range("G42").Value = "Iкз max на шинах ≤ Icu (модульные 4,5 / 6 / 10 кА; ВА — по каталогу). Проверка — строка 38"
    for c in "CDE":
        f = ws.Range(f"{c}38").Formula
        head = f'=IF({c}$28="","",'
        if not f.startswith(head) or "Icu" in f:
            continue
        inner = f[len(head):-1]
        ws.Range(f"{c}38").Formula2 = (
            f'{head}LET(m_,{inner},i_,IF(AND(N({c}36)>0,N({c}42)>0,N({c}36)>N({c}42)),'
            f'"Iкз max "&FIXED({c}36,1,TRUE)&" кА > Icu "&{c}42&" кА — аппараты щита с большей отключающей способностью",""),'
            f'IF(i_="",m_,IF(OR(m_="",m_="ok"),i_,m_&"; "&i_))))')
    log.append(f"20: «{NET}» строка 42 — Icu аппаратов щита (6 кА), проверка в C38:E38")


# ---------------------------------------------------------------- 22 мощность ввода
def rule_al(f):
    if "Нагрузка_щитов!R10C32" in f:
        return None
    return f'=IF(AND(RC34="Ввод",ISNUMBER({BOARD_P})),{BOARD_P},' + f[1:] + ")"


def patch_input_power(wb, log):
    s = wb.Worksheets(SRC)
    so = unprotect(s)
    n = apply(s, "AL", rule_al)
    if so:
        s.Protect(PWD, *so)
    if n:
        log.append(f"22: «{SRC}» AL — мощность ввода = Рр щита, {n} строк")


# ---------------------------------------------------------------- 23 длина при пустом G
def rule_m(f):
    if "(RC[-6]*2-1)" not in f:
        return None
    return f.replace("(RC[-6]*2-1)", "(MAX(1,N(RC[-6]))*2-1)")


def patch_length(wb, log):
    s_ = wb.Worksheets(SRC)
    so = unprotect(s_)
    n = apply(s_, "M", rule_m)
    if so:
        s_.Protect(PWD, *so)
    if n:
        log.append(f"23: «{SRC}» M — пустое «кол-во потребителей» = 1 (опуск не вычитается), {n} строк")


def run(path):
    path = os.path.abspath(path)
    app, pid = excel()
    log = []
    try:
        wb = app.Workbooks.Open(path, 0, False)
        app.Calculation = -4135
        patch_codes(wb, log)
        patch_motors(wb, log)
        patch_icu(wb, log)
        patch_input_power(wb, log)
        patch_length(wb, log)
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
    for s_ in log or ["изменений нет"]:
        print("  " + s_)


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(1)
    for p in sys.argv[1:]:
        run(p)
