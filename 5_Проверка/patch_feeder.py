# -*- coding: utf-8 -*-
"""Данные фидера для схемы нижестоящего щита: номер автомата (QF), автомат, кабель, длина, щит-источник.

    py patch_feeder.py <книга.xlsx> [...]   — книга закрыта; повторный запуск ничего не портит
    (после merge_network.py и patch_inputs.py)

«Сеть_щита», раздел Б (кого питает щит), строки 94–123 — справа, BK:BO «ДЛЯ СХЕМЫ НИЖЕСТОЯЩЕГО»:
  BK QF фидера (из «Однолинейки» этого щита, столбец G строки линии-фидера), BL автомат (H первой строки группы QF),
  BM кабель («Исходные данные» AE), BN длина, м (M), BO щит-источник (A «Однолинейки»).
«Сеть_щита», раздел А: H5:L8 — то же по Вводу 1 (щит 1–3), AL5:AP11 — по Вводам 2 и 3 (строки как U6:U11).
  Заполняет кнопка «Связать с вышестоящим» (связью на BK:BO вышестоящего); вручную — можно вписать самому.
SXDRAW рисует по ним рамку питающего щита: «ВРУ — QF3 3P С80А — кабель … L=…» перед вводом.
Селективность ввода («Однолинейка», строки с кодом «Ввод»):
  AQ — номинал ввода = большее из «по току щита» и ближайшего ряда ≥ 1,6 × наибольшего отходящего (AO других строк щита);
  «Исходные данные» U ввода — тот же номинал (кабель ввода подбирается под автомат);
  G — у выключателя нагрузки обозначение QS (номер общий с QF);
  H — если селективный автомат невозможен (In фидера < 1,6 · In ввода): выключатель нагрузки «QS 3P …А» ≥ In фидера;
  AE — статус (для своего текста автомата в E): In фидера < 1,6 · In ввода -> «неселективно — QS на ввод или фидер больше».
  Свой текст в E (например «QS 3P 250А») — выключатель нагрузки: номинал не разбирается, проверки нет.
Свой скрытый экземпляр Excel; закрывается только он. COM — только позиционные аргументы.
"""
import os
import signal
import sys

NET = "Сеть_щита"
MARK = "ДЛЯ СХЕМЫ НИЖЕСТОЯЩЕГО"
HDR = ["QF фидера", "Автомат фидера", "Кабель фидера", "L фидера, м", "Щит-источник"]


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


def patch(wb, log):
    ws = wb.Worksheets(NET)
    if ws.Range("BK92").Value == MARK:
        return
    if wb.Application.WorksheetFunction.CountA(ws.Range("BK92:BO123")) > 0:
        raise RuntimeError(f"«{NET}» BK92:BO123 занято")
    # раздел Б: что отдать нижестоящему для схемы
    ws.Range("BK92").Value = MARK
    ws.Range("BK92").Font.Bold = True
    ws.Range("BK93:BO93").Value = HDR
    O, S = "Однолинейка!", "'Исходные данные'!"
    for r in range(94, 124):
        b = f"$B{r}"
        ws.Range(f"BK{r}").Formula = f'=IF({b}="","",IFERROR(INDEX({O}$G$1:$G$1000,MATCH({b},{O}$R$1:$R$1000,0))&"",""))'
        ws.Range(f"BL{r}").Formula = f'=IF(BK{r}="","",IFERROR(INDEX({O}$H$1:$H$1000,MATCH(BK{r},{O}$G$1:$G$1000,0))&"",""))'
        ws.Range(f"BM{r}").Formula = f'=IF({b}="","",IFERROR(INDEX({S}$AE$1:$AE$1000,MATCH({b},{S}$D$1:$D$1000,0))&"",""))'
        ws.Range(f"BN{r}").Formula = f'=IF({b}="","",IFERROR(INDEX({S}$M$1:$M$1000,MATCH({b},{S}$D$1:$D$1000,0))*1,""))'
        ws.Range(f"BO{r}").Formula = f'=IF({b}="","",IFERROR(INDEX({O}$A$1:$A$1000,MATCH({b},{O}$R$1:$R$1000,0))&"",""))'
    for rng in ("BK93:BO93",):
        ws.Range(rng).Font.Bold = True
        ws.Range(rng).WrapText = True
        ws.Range(rng).Interior.Color = rgb((221, 235, 247))
    ws.Range("BK93:BO123").Borders.LineStyle = 1
    # раздел А: что пришло от фидера (Ввод 1 — H:L, Вводы 2 и 3 — AL:AP)
    ws.Range("H5:L5").Value = HDR
    ws.Range("AL5:AP5").Value = HDR
    for rng in ("H5:L5", "AL5:AP5"):
        ws.Range(rng).Font.Bold = True
        ws.Range(rng).WrapText = True
        ws.Range(rng).Interior.Color = rgb((221, 235, 247))
    ws.Range("H6:L8").Interior.Color = rgb((255, 242, 204))
    ws.Range("AL6:AP11").Interior.Color = rgb((255, 242, 204))
    ws.Range("H5:L8").Borders.LineStyle = 1
    ws.Range("AL5:AP11").Borders.LineStyle = 1
    for c, w in zip(("H", "I", "J", "K", "L"), (9, 14, 22, 8, 12)):
        if ws.Columns(c).ColumnWidth < w:
            ws.Columns(c).ColumnWidth = w
    for c, w in zip(("AL", "AM", "AN", "AO", "AP"), (9, 14, 22, 8, 12)):
        ws.Columns(c).ColumnWidth = w
    log.append(f"«{NET}»: BK92:BO123 — данные фидеров для нижестоящих; H5:L8 и AL5:AP11 — данные своего фидера")


SER = "{6,10,16,20,25,32,40,50,63,80,100,125,160,200,250,315,400,500,630,800,1000,1250,1600}"


def nom(x):
    return f"IFERROR(INDEX({SER},MATCH(TRUE,INDEX({SER}>={x},0),0)),{x})"


def runs(ws, col, r0=3, r1=1000):
    fs = [x[0] for x in ws.Range(f"{col}{r0}:{col}{r1}").Formula2R1C1]
    out, start = [], r0
    for i in range(1, len(fs) + 1):
        if i == len(fs) or fs[i] != fs[i - 1]:
            out.append((start, r0 + i - 1, fs[i - 1]))
            start = r0 + i
    return out


def rule_aq(f):
    if "MAXIFS(R3C41:R1000C41" in f or not f.startswith("=") or len(f) < 5:
        return None
    sel = nom('1.6*MAXIFS(R3C41:R1000C41,R3C1:R1000C1,RC1,R3C38:R1000C38,"<>Ввод")')
    return f'=IF(AND(RC38="Ввод",RC5<>""),MAX(N({f[1:]}),{sel}),{f[1:]})'


def rule_u(f):
    """«Исходные данные» U (номинал для кабеля) ввода — тот же селективный номинал, что AQ в «Однолинейке»:
    кабель ввода подбирается под автомат. Без цикла: AO «Однолинейки» не зависит от U."""
    if "Однолинейка!R3C41" in f or not f.startswith("="):
        return None
    sel = nom('1.6*MAXIFS(Однолинейка!R3C41:R1000C41,Однолинейка!R3C1:R1000C1,RC78,Однолинейка!R3C38:R1000C38,"<>Ввод")')
    return f'=IF(RC34="Ввод",MAX(N({f[1:]}),{sel}),{f[1:]})'


# In фидера вышестоящего для ввода N: N = 1 — «Сеть_щита» C59:E59 (связь F6 или вручную C24), N ≥ 2 — Z6:Z11
UP_IN = ('IFERROR(IF(IFERROR(--MID(RC18,6,2),1)<=1,INDEX(Сеть_щита!R59C3:R59C5,1,MATCH(RC1,ПЩ_щиты,0)),'
         'INDEX(Сеть_щита!R6C26:R11C26,(MATCH(RC1,ПЩ_щиты,0)-1)*2+IFERROR(--MID(RC18,6,2),1)-1))*1,0)')
SEL_MSG = (f'IFERROR(IF(AND(RC38="Ввод",N(RC47)>0,{UP_IN}>0,{UP_IN}<1.6*N(RC47)),'
           f'"ввод In "&RC47&" А, фидер "&{UP_IN}&" А — неселективно (нужно In фидера ≥ 1,6·In ввода): '
           f'на ввод QS или фидер больше",""),"")')


def rule_h(f):
    """ввод «авто»: если селективный автомат невозможен (In фидера < 1,6·In ввода) — выключатель нагрузки QS ≥ In фидера"""
    key = 'IF(AND(RC[-3]<>"авто",RC[-3]<>"+"),RC[-3],'
    if key not in f or '"QS "' in f:
        return None
    i = f.index(key) + len(key)
    auto = f[i:-2]                                  # текст автомата «авто» (до двух закрывающих скобок)
    qs = f'IF(AND(RC38="Ввод",{UP_IN}>0,{UP_IN}<1.6*N(RC43)),"QS "&RC46&" "&{nom(UP_IN)}&"А",{auto})'
    return f[:i] + qs + "))"


def rule_g(f):
    """обозначение: у выключателя нагрузки (текст автомата «QS …») — QS вместо QF, номер общий (QS1, QS2, QF3…)"""
    if not f.startswith("=") or "QS" in f:
        return None
    return f'=IF(LEFT(RC8,3)="QS ",SUBSTITUTE({f[1:]},"QF","QS"),{f[1:]})'


def rule_ae(f):
    if "неселективно" in f or not f.endswith(")))"):
        return None
    return f[:-2] + "," + SEL_MSG + "))"


def patch_input_sel(wb, log):
    ws = wb.Worksheets("Однолинейка")
    prot = ws.ProtectContents
    if prot:
        p = ws.Protection
        opts = (True, True, True, False, p.AllowFormattingCells, p.AllowFormattingColumns, p.AllowFormattingRows,
                p.AllowInsertingColumns, p.AllowInsertingRows, p.AllowInsertingHyperlinks, p.AllowDeletingColumns,
                p.AllowDeletingRows, p.AllowSorting, p.AllowFiltering, p.AllowUsingPivotTables)
        ws.Unprotect("sx")
    n = {}
    for col, rule in (("AQ", rule_aq), ("H", rule_h), ("G", rule_g), ("AE", rule_ae)):
        n[col] = 0
        for a, b, f in runs(ws, col):
            g = rule(f) if isinstance(f, str) else None
            if g:
                ws.Range(f"{col}{a}:{col}{b}").Formula2R1C1 = g
                n[col] += b - a + 1
    if prot:
        ws.Protect("sx", *opts)
    src = wb.Worksheets("Исходные данные")
    n["U (Исходные)"] = 0
    for a_, b_, f in runs(src, "U"):
        g = rule_u(f) if isinstance(f, str) else None
        if g:
            src.Range(f"U{a_}:U{b_}").Formula2R1C1 = g
            n["U (Исходные)"] += b_ - a_ + 1
    if any(n.values()):
        log.append("«Однолинейка»: селективность ввода — " + ", ".join(f"{k} {v}" for k, v in n.items() if v))


def run(path):
    path = os.path.abspath(path)
    app, pid = excel()
    log = []
    try:
        wb = app.Workbooks.Open(path, 0, False)
        app.Calculation = -4135
        patch(wb, log)
        patch_input_sel(wb, log)
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
