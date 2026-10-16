# -*- coding: utf-8 -*-
"""Узел КМ (АПС): контактор отключает выбранные линии щита по сигналу пожарной сигнализации.

    py patch_km.py <книга.xlsx> [...]      — книга закрыта; повторный запуск ничего не портит

Как пользоваться в книге («Однолинейка»):
  * Узел — своя строка-место: E «авто» (автомат цепи управления), I «КМ АПС». Следующие два места — пустые
    (там SXDRAW рисует контакт реле АПС и катушку КМ). Узлы щита получают имена KM-1, KM-2… (столбец CE).
  * Линии за контактором: в столбце D («Шина») — имя узла, например «KM-1». Они встают на шину
    «L1,L2,L3 (KM-1)» правее узла.
Подбор контактора (CH): первый по порядку в «Справочная» CY:DD, у которого
  AC-3 (AC-7b) >= CF — Iр двигателей на шине КМ (УГО «Привод», BK = MOTOR, и ККБ — компрессоры) и
  AC-1 >= CG — 1,1 × Iр всей шины КМ. Iр — наиболее нагруженная фаза, сумма токов линий, без Кс (в запас).
  Так верно и при смешанной нагрузке (ГОСТ IEC 60947-4-1: ток своей цепи в своей категории применения).
  In фидера в подбор не входит: он касается защиты КМ от КЗ (координация по данным изготовителя).
  Свою модель можно вписать в CH вручную — CI проверит её по тем же условиям.
  CI — проверка узла. «Каталог» N считает подобранные КМ (CH) — они идут в спецификацию.
Данные: Chint NCH8 (инструкция NCH8, табл. 1: AC-1/AC-7a 20/25/40/63 А, AC-7b 9/9/15/20 А),
        Chint NXC-25…100 (каталог NXC: Ith 40…125 А, AC-3 400 В = номер модели).
Свой скрытый экземпляр Excel; закрывается только он.
"""
import os
import signal
import sys

PWD = "sx"
ODN = "Однолинейка"
SPR = "Справочная"
CAT = "Каталог"
TITLE = "КОНТАКТОРЫ для узла КМ (АПС)"

# модель (как «Каталог» C), AC-1, AC-3 (AC-7b), полюса, производитель, наименование, артикул, ширина DIN
KM = [
    ("КМ NCH8-20/40", 20, 9, "4P", "Chint", "Контактор модульный NCH8-20/40 20А 4НО AC220/230В", "", 2),
    ("КМ NCH8-25/40", 25, 9, "4P", "Chint", "Контактор модульный NCH8-25/40 25А 4НО AC220/230В", "", 2),
    ("КМ NCH8-40/40", 40, 15, "4P", "Chint", "Контактор модульный NCH8-40/40 40А 4НО AC220/230В", "256099", 3),
    ("КМ NCH8-63/40", 63, 20, "4P", "Chint", "Контактор модульный NCH8-63/40 63А 4НО AC220/230В", "256101", 3),
    ("КМ NXC-25", 40, 25, "3P", "Chint", "Контактор NXC-25 25А (AC-3) катушка AC230В", "", None),
    ("КМ NXC-32", 50, 32, "3P", "Chint", "Контактор NXC-32 32А (AC-3) катушка AC230В", "", None),
    ("КМ NXC-38", 50, 38, "3P", "Chint", "Контактор NXC-38 38А (AC-3) катушка AC230В", "", None),
    ("КМ NXC-40", 60, 40, "3P", "Chint", "Контактор NXC-40 40А (AC-3) катушка AC230В", "", None),
    ("КМ NXC-50", 80, 50, "3P", "Chint", "Контактор NXC-50 50А (AC-3) катушка AC230В", "", None),
    ("КМ NXC-65", 80, 65, "3P", "Chint", "Контактор NXC-65 65А (AC-3) катушка AC230В", "", None),
    ("КМ NXC-75", 90, 75, "3P", "Chint", "Контактор NXC-75 75А (AC-3) катушка AC230В", "", None),
    ("КМ NXC-85", 100, 85, "3P", "Chint", "Контактор NXC-85 85А (AC-3) катушка AC230В", "", None),
    ("КМ NXC-100", 125, 100, "3P", "Chint", "Контактор NXC-100 100А (AC-3) катушка AC230В", "", None),
    ("КМ NXC-120", 200, 120, "3P", "Chint", "Контактор NXC-120 120А (AC-3) катушка AC/DC230В", "", None),
    ("КМ NXC-160", 200, 160, "3P", "Chint", "Контактор NXC-160 160А (AC-3) катушка AC/DC230В", "", None),
    ("КМ NXC-185", 275, 185, "3P", "Chint", "Контактор NXC-185 185А (AC-3) катушка AC/DC230В", "", None),
]
ROWS = 30                                        # CY4:DD33 — место под свои строки
NAMES = {"КМ_n": "$CY$4:$CY$33", "КМ_модель": "$CZ$4:$CZ$33", "КМ_AC1": "$DA$4:$DA$33", "КМ_AC3": "$DB$4:$DB$33"}
NOTES = [
    "Узел КМ: строка-место с E «авто» и I «КМ АПС»; два следующих места пустые (реле АПС и катушка).",
    "Линии за КМ: в «Однолинейке» D впишите имя узла (KM-1, KM-2…, столбец CE).",
    "Подбор (CH): первый сверху, у кого AC-3 >= Iр двигателей шины КМ и AC-1 >= 1,1 × Iр всей шины КМ.",
    "Порядок строк = предпочтение. Свои модели — в пустые строки; модель как в «Каталоге» C (для спецификации).",
    "NCH8 — инструкция Chint (AC-3 = AC-7b). NXC — каталог Chint (AC-1 = Ith). Артикулы NCH8-20/25 и NXC — уточнить.",
]
NODE = '$BH{r}="KMA"'
COLS = {   # столбец: (заголовок, формула для строки 3)
    "CE": ("Узел КМ\n(имя шины;\nв D линий)",
           '=IF($BH3="KMA","KM-"&COUNTIFS($A$3:$A3,$A3,$BH$3:$BH3,"KMA"),"")'),
    "CF": ("Iр двигателей\nшины КМ, А\n(для AC-3)",
           '=IF($CE3="","",ROUND(MAX(SUMIFS($U$3:$U$1000,$A$3:$A$1000,$A3,$D$3:$D$1000,$CE3,$BK$3:$BK$1000,"MOTOR",$V$3:$V$1000,"*L1*")+SUMIFS($U$3:$U$1000,$A$3:$A$1000,$A3,$D$3:$D$1000,$CE3,$AL$3:$AL$1000,"ККБ*",$BK$3:$BK$1000,"<>MOTOR",$V$3:$V$1000,"*L1*"),SUMIFS($U$3:$U$1000,$A$3:$A$1000,$A3,$D$3:$D$1000,$CE3,$BK$3:$BK$1000,"MOTOR",$V$3:$V$1000,"*L2*")+SUMIFS($U$3:$U$1000,$A$3:$A$1000,$A3,$D$3:$D$1000,$CE3,$AL$3:$AL$1000,"ККБ*",$BK$3:$BK$1000,"<>MOTOR",$V$3:$V$1000,"*L2*"),SUMIFS($U$3:$U$1000,$A$3:$A$1000,$A3,$D$3:$D$1000,$CE3,$BK$3:$BK$1000,"MOTOR",$V$3:$V$1000,"*L3*")+SUMIFS($U$3:$U$1000,$A$3:$A$1000,$A3,$D$3:$D$1000,$CE3,$AL$3:$AL$1000,"ККБ*",$BK$3:$BK$1000,"<>MOTOR",$V$3:$V$1000,"*L3*")),2))'),
    "CG": ("Iр шины КМ\n× 1,1, А\n(для AC-1)",
           '=IF($CE3="","",ROUND(1.1*MAX(SUMIFS($U$3:$U$1000,$A$3:$A$1000,$A3,$D$3:$D$1000,$CE3,$V$3:$V$1000,"*L1*"),SUMIFS($U$3:$U$1000,$A$3:$A$1000,$A3,$D$3:$D$1000,$CE3,$V$3:$V$1000,"*L2*"),SUMIFS($U$3:$U$1000,$A$3:$A$1000,$A3,$D$3:$D$1000,$CE3,$V$3:$V$1000,"*L3*")),2))'),
    "CH": ("Контактор\n(подобран;\nможно вписать)",
           '=IF($CE3="","",IF(MINIFS(КМ_n,КМ_AC3,">="&N($CF3),КМ_AC1,">="&N($CG3),КМ_модель,"?*")=0,"нет в таблице",'
           'INDEX(КМ_модель,MINIFS(КМ_n,КМ_AC3,">="&N($CF3),КМ_AC1,">="&N($CG3),КМ_модель,"?*"))))'),
    "CI": ("Узел КМ:\nпроверка",
           '=IF($CE3="","",IF($CH3="нет в таблице","НЕТ КМ на AC-3 "&$CF3&" А / AC-1 "&$CG3&" А — «Справочная» CY",'
           'IF(COUNTIFS($A$3:$A$1000,$A3,$D$3:$D$1000,$CE3)=0,"нет линий: в D линий впишите "&$CE3,'
           'IF(COUNTIFS($A$3:$A$1000,$A3,$D$3:$D$1000,$CE3,$B$3:$B$1000,"<"&$B3)>0,"линия "&$CE3&" левее узла",'
           'IF(COUNTIFS($A$3:$A$1000,$A3,$B$3:$B$1000,">"&$B3,$B$3:$B$1000,"<="&($B3+2),$E$3:$E$1000,"?*")'
           '+COUNTIFS($A$3:$A$1000,$A3,$B$3:$B$1000,">"&$B3,$B$3:$B$1000,"<="&($B3+2),$R$3:$R$1000,"?*")>0,'
           '"места "&($B3+1)&"–"&($B3+2)&" должны быть пустые",'
           'IF(COUNTIFS($A$3:$A$1000,$A3,$B$3:$B$1000,$B3+3,$C$3:$C$1000,$C3)=0,"узел не помещается на лист (нужны места до "&($B3+3)&")",'
           'IF(ISNA(MATCH($CH3,КМ_модель,0)),"модели нет в «Справочная» CZ",'
           'IF(INDEX(КМ_AC3,MATCH($CH3,КМ_модель,0))<N($CF3),"КМ мал по AC-3",'
           'IF(INDEX(КМ_AC1,MATCH($CH3,КМ_модель,0))<N($CG3),"КМ мал по AC-1","ok")))))))))'),
}
# заголовки первой версии (AC-1 по In фидера) — заменяются
OLD_HDR = {"CF": ("Iр шины КМ, А\n(макс. фаза,\nбез Кс)",), "CG": ("In автомата\nперед КМ, А\n(ввод / фидер)",)}
# BH (аппарат 1): «… АПС …» -> код узла KMA (до проверки «КМ» -> KM)
OLD_BH = 'IF(OR(LEFT(UPPER(RC[-51]),2)="КМ",LEFT(UPPER(RC[-51]),2)="KM"),"KM","X")'
NEW_BH = 'IF(ISNUMBER(SEARCH("АПС",RC[-51])),"KMA",' + OLD_BH + ')'
I2_OLD = "БП…)"
I2_NEW = "БП…;\nКМ АПС — узел)"
# «Каталог» N: подобранные КМ (CH = столбец 86) — в спецификацию
OLD_N = "COUNTIF(Однолинейка!R3C11:R1000C11,RC[-11])"
NEW_N = OLD_N + "+COUNTIF(Однолинейка!R3C86:R1000C86,RC[-11])"


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


def runs(ws, col, r0, r1):
    fs = [x[0] for x in ws.Range(f"{col}{r0}:{col}{r1}").Formula2R1C1]
    out, start = [], r0
    for i in range(1, len(fs) + 1):
        if i == len(fs) or fs[i] != fs[i - 1]:
            out.append((start, r0 + i - 1, fs[i - 1]))
            start = r0 + i
    return out


def fmt(ws, src, dst):
    ws.Range(src).Copy(ws.Range(dst))
    ws.Range(dst).ClearContents()


def spr(wb, log):
    ws = wb.Worksheets(SPR)
    if str(ws.Range("CY2").Value or "").strip() == TITLE:
        have = [str(x[0] or "").strip() for x in ws.Range(f"CZ4:CZ{3 + ROWS}").Value]
        new = [k for k in KM if k[0] not in have]
        if new:
            opts = unprotect(ws)
            free = [4 + i for i, v in enumerate(have) if not v]
            for (m, a1, a3, p, mf, *_), r in zip(new, free):
                ws.Range(f"CZ{r}:DD{r}").Value = ((m, a1, a3, p, mf),)
            if opts:
                ws.Protect(PWD, *opts)
            log.append(f"«{SPR}» CZ — новые контакторы: {', '.join(k[0] for k in new[:len(free)])}")
        return
    if ws.Range("CY1:DD40").Find("*") is not None:
        raise SystemExit(f"«{SPR}» CY:DD заняты — книга не изменена")
    opts = unprotect(ws)
    fmt(ws, "CC1", "CY1")
    ws.Range("CY1").Value = "Контакторы узла КМ (отключение по АПС). Жёлтые — можно менять и дополнять."
    fmt(ws, "CC2", "CY2")
    ws.Range("CY2").Value = TITLE
    fmt(ws, "BP3", "CY3:DD3")
    ws.Range("CY3:DD3").Value = (("№", "Модель (как «Каталог» C)", "AC-1, А", "AC-3 (AC-7b), А", "Полюса",
                                  "Производитель"),)
    fmt(ws, "BP4", f"CY4:CY{3 + ROWS}")
    fmt(ws, "BQ4", f"CZ4:DD{3 + ROWS}")
    ws.Range(f"CY4:CY{3 + ROWS}").Formula = '=IF(CZ4="","",ROW()-3)'
    ws.Range(f"CZ4:DD{3 + len(KM)}").Value = tuple((m, a1, a3, p, mf) for m, a1, a3, p, mf, *_ in KM)
    fmt(ws, "BR14", f"CY{6 + ROWS}:CY{5 + ROWS + len(NOTES)}")
    for i, s in enumerate(NOTES):
        ws.Range(f"CY{6 + ROWS + i}").Value = s
    for c, w in (("CY", 4), ("CZ", 17), ("DA", 8), ("DB", 9), ("DC", 7), ("DD", 10)):
        ws.Columns(c).ColumnWidth = w
    for n in NAMES:
        try:
            wb.Names(n).Delete()
        except Exception:
            pass
    for n, ref in NAMES.items():
        wb.Names.Add(n, f"='{SPR}'!{ref}")
    if opts:
        ws.Protect(PWD, *opts)
    log.append(f"«{SPR}» CY:DD — таблица контакторов ({len(KM)} шт.), имена КМ_*")


def odn(wb, log):
    ws = wb.Worksheets(ODN)
    opts = unprotect(ws)
    last = max(ws.Cells(ws.Rows.Count, c).End(-4162).Row for c in ("A", "H", "AF", "BH"))
    for col, (hdr, f) in COLS.items():
        h = ws.Range(f"{col}2")
        if str(h.Value or "").strip() not in ("", hdr) + OLD_HDR.get(col, ()):
            raise SystemExit(f"«{ODN}» {col}2 занят: «{h.Value}» — книга не изменена")
        if h.Value != hdr:
            ws.Range("CB2").Copy(h)
            h.Value = hdr
            ws.Range("CB3").Copy(ws.Range(f"{col}3:{col}{last}"))
            ws.Range(f"{col}3:{col}{last}").ClearContents()
            ws.Columns(col).ColumnWidth = 13 if col in ("CH", "CI") else 10
            log.append(f"«{ODN}» {col}2 — «{hdr.replace(chr(10), ' ')}»")
        if ws.Range(f"{col}3").Formula2 != f or ws.Range(f"{col}{last}").Formula2 == "":
            # CH — ручные модели (текст без «=») не затираем
            if col == "CH":
                vals = [x[0] for x in ws.Range(f"CH3:CH{last}").Formula2]
                keep = {i: v for i, v in enumerate(vals) if v not in ("", None) and not str(v).startswith("=")}
                ws.Range(f"CH3:CH{last}").Formula2 = f
                for i, v in keep.items():
                    ws.Range(f"CH{3 + i}").Value = v
            else:
                ws.Range(f"{col}3:{col}{last}").Formula2 = f
            log.append(f"«{ODN}» {col}3:{col}{last} — формула")
    if not str(ws.Range("CE1").Value or ""):
        ws.Range("CE1").Value = "УЗЕЛ КМ (АПС)"
    n = 0
    for a, b, f in runs(ws, "BH", 3, last):
        if isinstance(f, str) and OLD_BH in f and '"KMA"' not in f:
            ws.Range(f"BH{a}:BH{b}").Formula2R1C1 = f.replace(OLD_BH, NEW_BH, 1)
            n += b - a + 1
    if n:
        log.append(f"«{ODN}» BH — «КМ АПС» в аппарате 1 -> узел KMA: {n} строк")
    elif not any('"KMA"' in str(f) for _, _, f in runs(ws, "BH", 3, last)):
        log.append(f"«{ODN}» BH — формула не узнана, узел КМ не подключён")
    h = str(ws.Range("I2").Value or "")
    if h.endswith(I2_OLD) and "АПС" not in h:
        ws.Range("I2").Value = h[: -len(I2_OLD)] + I2_NEW
    if opts:
        ws.Protect(PWD, *opts)


def cat(wb, log):
    ws = wb.Worksheets(CAT)
    opts = unprotect(ws)
    vals = ws.Range("A4:C500").Value
    have = {str(r[2] or "").strip() for r in vals}
    free = [4 + i for i, r in enumerate(vals) if not any(r)]
    added = []
    for m, a1, a3, p, mf, name, art, w in KM:
        if m in have or not free:
            continue
        r = free.pop(0)
        ws.Range(f"A{r}:M{r}").Value = (("Аппарат", mf, m, None, None, None, name, art or None, p, a1, None, w,
                                          "шт."),)
        if art:
            ws.Range(f"H{r}").NumberFormat = "@"
            ws.Range(f"H{r}").Value = art
        added.append(m)
    if added:
        log.append(f"«{CAT}» — новые позиции: {', '.join(added)}")
    n = 0
    for a, b, f in runs(ws, "N", 4, 500):
        if isinstance(f, str) and OLD_N in f and "R3C86" not in f:
            ws.Range(f"N{a}:N{b}").Formula2R1C1 = f.replace(OLD_N, NEW_N, 1)
            n += b - a + 1
    if n:
        log.append(f"«{CAT}» N — подобранные КМ (Однолинейка CH) в спецификации: {n} строк")
    if opts:
        ws.Protect(PWD, *opts)


def run(path):
    path = os.path.abspath(path)
    app, pid = excel()
    log = []
    try:
        wb = app.Workbooks.Open(path, 0, False)
        if wb.ReadOnly:                              # открыта у другого — не трогаем
            wb.Close(False)
            raise SystemExit(f"{os.path.basename(path)}: открыта у другого пользователя — пропущена")
        app.Calculation = -4135
        spr(wb, log)
        odn(wb, log)
        cat(wb, log)
        app.Calculation = -4105
        app.CalculateFull()
        if log:
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
