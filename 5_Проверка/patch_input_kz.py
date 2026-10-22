# -*- coding: utf-8 -*-
"""Iкз в строке ввода щита — без двойного учёта кабеля ввода, In и характеристика — фидера выше.

    py patch_input_kz.py <книга.xlsx> [...]   — книга закрыта; повторный запуск ничего не портит

«Исходные данные», строка ввода (AH = «Ввод»). Кабель ввода кончается на шинах этого щита.
Было: BL = Zп шин щита («Сеть_щита» ВЩ_Z, уже с кабелем ввода) + 2ρL/S кабеля ввода ещё раз;
BJ = Iкз / In строки ввода (ручной In или 1,6·отходящего), BK — характеристика из «Однолинейки» H (у QS — C).
Пример ЩС-ПК: Zп 0,165 + 0,080 = 0,245 Ом, 897 А / 125 = 7,2 < 10 «мало»; верно 1330 А / 100 (QF10 C100) = 13,3.
Стало (только строка ввода):
  BL — Zп на шинах щита: Ввод 1 — «Сеть_щита» C34:E34, Ввод 2/3 — AB6:AB11 (пусто = как Ввод 1);
  BJ — Iкз / In фидера выше: Ввод 1 — C59:E59 (F6 или ручной C24), Ввод 2/3 — Z6:Z11 (пусто = как Ввод 1);
  BK — кратность по характеристике фидера: «Сеть_щита» I6:I8 (автомат фидера) или C25:E25 (B / C / D).
Отходящие линии не меняются: после правки их BI:BM сверяются со значениями до правки, иначе книга не сохраняется.
Свой скрытый экземпляр Excel; закрывается только он.
"""
import os
import signal
import sys

SRC = "Исходные данные"
PWD = "sx"
R1, R2 = 3, 1000
TAG = '$AH{r}="Ввод"'

K = ("MATCH(INDEX(Разбивка_по_щитам!$F$1:$F$1000,MATCH($D{r},Разбивка_по_щитам!$H$1:$H$1000,0)),ПЩ_щиты,0)")
N = "MAX(1,IFERROR(--MID($D{r},6,2),1))"


def bl_new(r):
    k, n = K.format(r=r), N.format(r=r)
    return (f'IFERROR(IF(OR($D{r}="",NOT(ISNUMBER($V{r}))),"",LET(k_,{k},n_,{n},'
            f'z1_,INDEX(Сеть_щита!$C$34:$E$34,1,k_),'
            f'z2_,IF(n_>=2,INDEX(Сеть_щита!$AB$6:$AB$11,(k_-1)*2+n_-1),""),'
            f'IF(ISNUMBER(z2_),z2_,IF(ISNUMBER(z1_),z1_,"")))),"")')


def in_new(r):
    k, n = K.format(r=r), N.format(r=r)
    return (f'LET(k_,{k},n_,{n},'
            f'i2_,IF(n_>=2,N(INDEX(Сеть_щита!$Z$6:$Z$11,(k_-1)*2+n_-1)),0),'
            f'IF(i2_>0,i2_,N(INDEX(Сеть_щита!$C$59:$E$59,1,k_))))')


def bk_new(r):
    k = K.format(r=r)
    f = (f'IFERROR(LET(k_,{k},t_," "&INDEX(Сеть_щита!$I$6:$I$8,k_)&" "&INDEX(Сеть_щита!$C$25:$E$25,1,k_),'
         f'IF(OR(ISNUMBER(SEARCH(" B",t_)),ISNUMBER(SEARCH(" В",t_))),5,IF(ISNUMBER(SEARCH(" D",t_)),20,10))),10)')
    return (f'IF(OR($D{r}="",$CB{r}=1,NOT(ISNUMBER($V{r}))),"",IF(NOT(ISNUMBER(BJ{r})),'
            f'"не проверено: заполните «Сеть_щита»",IF(BJ{r}>={f},"ok","мало: Iкз < "&{f}&"·In фидера")))')


def wrap(col, f, r):
    """=old -> =IF($AHr="Ввод",new,old)"""
    if not (isinstance(f, str) and f.startswith("=")) or TAG.format(r=r) in f:
        return None
    if col == "BL":
        if "ВЩ_Z" not in f:
            return None
        new = bl_new(r)
    elif col == "BJ":
        if f"BI{r}/" not in f:
            return None
        new = f'IFERROR(IF(ISNUMBER(BI{r}),BI{r}/{in_new(r)},""),"")'
    else:
        if "не проверено" not in f:
            return None
        new = bk_new(r)
    return f'=IF({TAG.format(r=r)},{new},{f[1:]})'


def excel():
    import win32com.client
    import win32process
    app = win32com.client.DispatchEx("Excel.Application")
    app.Visible = False
    app.DisplayAlerts = False
    app.AskToUpdateLinks = False
    app.AutomationSecurity = 3
    return app, win32process.GetWindowThreadProcessId(app.Hwnd)[1]


def values(ws):
    return [list(x) for x in ws.Range(f"BI{R1}:BM{R2}").Value]


def patch(app, wb, log):
    ws = wb.Worksheets(SRC)
    ah = [x[0] for x in ws.Range(f"AH{R1}:AH{R2}").Value]
    app.CalculateFull()
    before = values(ws)
    n = {}
    for col in ("BL", "BJ", "BK"):
        rng = ws.Range(f"{col}{R1}:{col}{R2}")
        old = [x[0] for x in rng.Formula2]
        out, k = [], 0
        for i, f in enumerate(old):
            g = wrap(col, f, R1 + i)
            k += bool(g)
            out.append((g or f,))
        if k:
            rng.Formula2 = tuple(out)
        n[col] = k
    app.CalculateFull()
    after = values(ws)
    bad, inputs = [], []
    for i, (a, b) in enumerate(zip(before, after)):
        r = R1 + i
        if ah[i] == "Ввод":
            if a != b:
                inputs.append(f"{ws.Range(f'D{r}').Text}: Zп {a[3]} → {b[3]}, Iкз {a[0]} → {b[0]}, "
                              f"кратн. {a[1]} → {b[1]}, {a[2]} → {b[2]}")
        elif a != b:
            bad.append(r)
    log.append("«%s» строка ввода: BL %d, BJ %d, BK %d" % (SRC, n["BL"], n["BJ"], n["BK"]))
    log.extend("  " + s for s in inputs)
    if bad:
        log.append(f"ОШИБКА: изменились отходящие линии, строки {bad[:10]} — книга не сохранена")
        return False
    log.append("отходящие линии BI:BM — без изменений")
    return True


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
        ok = patch(app, wb, log)
        if prot:
            ws.Protect(PWD, True, True, True)
        if ok:
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
