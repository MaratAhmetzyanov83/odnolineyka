# -*- coding: utf-8 -*-
"""Лист «Ведомость_работ» — ведомость объёмов работ (ВОР) из «Спецификации» и «Кабельного_журнала».

    py patch_vor.py <книга.xlsx> [...]      — книга закрыта; повторный запуск ничего не портит
    py patch_vor.py --force <книга.xlsx>    — пересоздать лист (свои слова в G:H — сохраняются)

Графы: № п/п, наименование работ, ед. изм., кол-во, примечание. Лист считается сам (формулы):
  1. Монтаж оборудования — позиции «Спецификации» (поз. с точкой, кол-во > 0), кроме слов «Не включать»
     и кабелей (их длина — из кабельного журнала).
  2. Прокладка кабелей — «Кабельный_журнал» C:D: марка (сечение) и длина с запасом (D3), округление вверх до 1 м;
     марка с длиной 0 (кабель ввода без L) не выводится; ПуГВ — «прокладка провода».
  3. Прокладка труб, лотков, коробов — позиции «Спецификации» со словами «Прокладка».
Слова — жёлтые ячейки G4:H15 (часть наименования, без учёта регистра). Пустой раздел не выводится.
Скрытые столбцы J:U — служебные. Свой скрытый экземпляр Excel; закрывается только он.
"""
import os
import signal
import sys

SHEET = "Ведомость_работ"
MARK = "ВЕДОМОСТЬ ОБЪЁМОВ РАБОТ"
SPEC, KJ = "Спецификация", "Кабельный_журнал"
LAST = 400                                       # строк ВОР
SROWS = (3, 300)                                 # строки «Спецификации»
KROWS = (5, 1000)                                # строки «Кабельного_журнала»
NOT_IN = ["Кабель", "Провод", "Наконечник", "Расходн", "Перемычка", "Изолятор", "Фиксатор"]
LAYING = ["Труба", "Лоток", "Короб", "Гофр", "Металлорукав"]
SECTIONS = ["Монтаж оборудования (по спецификации)", "Прокладка кабелей (по кабельному журналу)",
            "Прокладка труб, лотков, коробов (по спецификации)"]


def excel():
    import win32com.client
    import win32process
    app = win32com.client.DispatchEx("Excel.Application")
    app.Visible = False
    app.DisplayAlerts = False
    app.AskToUpdateLinks = False
    app.AutomationSecurity = 3
    return app, win32process.GetWindowThreadProcessId(app.Hwnd)[1]


def build(wb, ws, keep):
    s0, s1 = SROWS
    k0, k1 = KROWS
    sp, kj = f"'{SPEC}'", f"'{KJ}'"
    ws.Cells.Clear()
    # шапка
    ws.Range("A1").Formula2 = (f'="{MARK}"&IFERROR(" — "&INDEX(Однолинейка!$A$3:$A$1000,'
                               'MATCH("?*",Однолинейка!$A$3:$A$1000,0)),"")')
    ws.Range("A2:E2").Value = (("№ п/п", "Наименование работ", "Ед. изм.", "Кол-во", "Примечание"),)
    # настройки (жёлтые)
    ws.Range("G2").Value = "Слова в наименовании (часть слова, без учёта регистра). Жёлтые — можно менять."
    ws.Range("G3:H3").Value = (("Не включать", "Прокладка (раздел 3)"),)
    g = keep.get("G") or NOT_IN
    h = keep.get("H") or LAYING
    ws.Range("G4:H15").Value = tuple((g[i] if i < len(g) else None, h[i] if i < len(h) else None) for i in range(12))
    ws.Range("G4:H15").Interior.Color = 0x99FFFF
    # служебное: J,K — номер раздела и позиции в нём строки ВОР; L — строка источника;
    # P,Q — класс строки «Спецификации» (1 монтаж / 3 прокладка) и номер в классе 1, R — номер в классе 3;
    # S — номер строки кабельного журнала; U1:U6 — числа строк разделов
    ws.Range("J2:U2").Value = (("раздел", "№ в разд.", "строка ист.", "заголовок", "", "", "класс спец.", "№ кл.1",
                                "№ кл.3", "№ КЖ", "", "итоги"),)
    ws.Range(f"P{s0}:P{s1}").Formula2 = (
        f'=IF(OR({sp}!$B{s0}="",ISERROR(SEARCH(".",{sp}!$B{s0})),IFERROR(--{sp}!$H{s0},0)<=0),0,'
        f'IF(SUMPRODUCT(($G$4:$G$15<>"")*ISNUMBER(SEARCH($G$4:$G$15,{sp}!$C{s0})))>0,0,'
        f'IF(SUMPRODUCT(($H$4:$H$15<>"")*ISNUMBER(SEARCH($H$4:$H$15,{sp}!$C{s0})))>0,3,1)))')
    ws.Range(f"Q{s0}:Q{s1}").Formula2 = f'=IF($P{s0}=1,COUNTIF($P${s0}:$P{s0},1),"")'
    ws.Range(f"R{s0}:R{s1}").Formula2 = f'=IF($P{s0}=3,COUNTIF($P${s0}:$P{s0},3),"")'
    ws.Range(f"S{k0}:S{k1}").Formula2 = (
        f'=IF(AND(ROW()<IFERROR(MATCH("Артикул ПТО",{kj}!$C:$C,0),{k1 + 1}),ISTEXT({kj}!$C{k0}),'
        f'IFERROR(N({kj}!$D{k0}),0)>0),COUNTIF($S${k0 - 1}:$S{k0 - 1},">0")+1,"")')
    ws.Range("T1:T6").Value = (("n1",), ("n2",), ("n3",), ("b1",), ("b2",), ("b3",))
    ws.Range("U1").Formula2 = f"=MAX($Q${s0}:$Q${s1})"
    ws.Range("U2").Formula2 = f"=MAX($S${k0}:$S${k1})"
    ws.Range("U3").Formula2 = f"=MAX($R${s0}:$R${s1})"
    for i in range(3):
        ws.Range(f"U{4 + i}").Formula2 = f"=IF($U${1 + i}>0,$U${1 + i}+1,0)"
    r0 = 3
    n = f"(ROW()-{r0 - 1})"
    ws.Range(f"J{r0}:J{LAST}").Formula2 = (
        f'=IF({n}<=$U$4,1,IF({n}<=$U$4+$U$5,2,IF({n}<=$U$4+$U$5+$U$6,3,"")))')
    ws.Range(f"K{r0}:K{LAST}").Formula2 = (
        f'=IF($J{r0}="","",{n}-1-IF($J{r0}>1,$U$4,0)-IF($J{r0}>2,$U$5,0))')
    ws.Range(f"L{r0}:L{LAST}").Formula2 = (
        f'=IF(N($K{r0})=0,"",IFERROR(CHOOSE($J{r0},MATCH($K{r0},$Q:$Q,0),MATCH($K{r0},$S:$S,0),'
        f'MATCH($K{r0},$R:$R,0)),""))')
    ws.Range(f"M{r0}:M{LAST}").Formula2 = f'=IF(AND($J{r0}<>"",N($K{r0})=0),1,0)'   # заголовок раздела
    # графы ВОР (раздел 1 и 3 — «Спецификация», раздел 2 — «Кабельный_журнал»)
    secs = ",".join(f'"{s}"' for s in SECTIONS)
    ws.Range(f"A{r0}:A{LAST}").Formula2 = (
        f'=IF($J{r0}="","",IF($K{r0}=0,$J{r0}&"",$J{r0}&"."&$K{r0}))')
    ws.Range(f"B{r0}:B{LAST}").Formula2 = (
        f'=IF($J{r0}="","",IF($K{r0}=0,CHOOSE($J{r0},{secs}),IF($L{r0}="","?",'
        f'CHOOSE($J{r0},"Монтаж: "&INDEX({sp}!$C:$C,$L{r0}),IF(ISNUMBER(SEARCH("ПуГВ",INDEX({kj}!$C:$C,$L{r0}))),"Прокладка провода ","Прокладка кабеля ")&INDEX({kj}!$C:$C,$L{r0}),'
        f'"Прокладка: "&INDEX({sp}!$C:$C,$L{r0})))))')
    ws.Range(f"C{r0}:C{LAST}").Formula2 = (
        f'=IF(OR($J{r0}="",N($K{r0})=0,$L{r0}=""),"",IF($J{r0}=2,"м",INDEX({sp}!$G:$G,$L{r0})&""))')
    ws.Range(f"D{r0}:D{LAST}").Formula2 = (
        f'=IF(OR($J{r0}="",N($K{r0})=0,$L{r0}=""),"",IF($J{r0}=2,ROUNDUP(N(INDEX({kj}!$D:$D,$L{r0})),0),'
        f'IFERROR(--INDEX({sp}!$H:$H,$L{r0}),INDEX({sp}!$H:$H,$L{r0}))))')
    ws.Range(f"E{r0}:E{LAST}").Formula2 = (
        f'=IF(OR($J{r0}="",N($K{r0})=0,$L{r0}=""),"",IF($J{r0}=2,"кабельный журнал, запас "'
        f'&ROUND(N({kj}!$D$3)*100,0)&" %","спецификация, поз. "&INDEX({sp}!$B:$B,$L{r0})))')
    # оформление
    for c, w in (("A", 7), ("B", 70), ("C", 9), ("D", 10), ("E", 32), ("F", 2), ("G", 18), ("H", 20)):
        ws.Columns(c).ColumnWidth = w
    ws.Range("A1").Font.Bold = True
    ws.Range("A1").Font.Size = 12
    hdr = ws.Range("A2:E2")
    hdr.Font.Bold = True
    hdr.HorizontalAlignment = -4108
    hdr.Interior.Color = 0xEFEFEF
    ws.Range("G2:H3").Font.Bold = True
    ws.Range("G2").WrapText = False
    body = ws.Range(f"A2:E{LAST}")
    body.VerticalAlignment = -4160
    ws.Range(f"B3:B{LAST}").WrapText = True
    ws.Range(f"E3:E{LAST}").WrapText = True
    ws.Range(f"A3:A{LAST}").HorizontalAlignment = -4108
    ws.Range(f"C3:D{LAST}").HorizontalAlignment = -4108
    # рамки и жирные заголовки разделов — условным форматом (число строк меняется)
    fc = ws.Range(f"A2:E{LAST}").FormatConditions
    fc.Delete()
    c1 = fc.Add(2, 3, '=$J2<>""')
    for b in (-4131, -4152, -4160, -4107):       # xlLeft, xlRight, xlTop, xlBottom
        c1.Borders(b).LineStyle = 1
    c2 = ws.Range(f"A2:E{LAST}").FormatConditions.Add(2, 3, "=$M2=1")   # без функций: язык Excel любой
    c2.Font.Bold = True
    c2.Interior.Color = 0xF7F7F7
    for b in (1, 2, 3, 4):
        ws.Range("A2:E2").Borders(b + 6).LineStyle = 1
    ws.Range("J:U").EntireColumn.Hidden = True
    ps = ws.PageSetup
    try:
        ps.PrintArea = f"$A$1:$E${LAST}"
        ps.PrintTitleRows = "$2:$2"
        ps.Orientation = 1
        ps.Zoom = False
        ps.FitToPagesWide = 1
        ps.FitToPagesTall = False
    except Exception:
        pass


def run(path, force):
    path = os.path.abspath(path)
    app, pid = excel()
    log = []
    try:
        wb = app.Workbooks.Open(path, 0, False)
        if wb.ReadOnly:                              # открыта у другого — не трогаем
            wb.Close(False)
            raise SystemExit(f"{os.path.basename(path)}: открыта у другого пользователя — пропущена")
        names = [s.Name for s in wb.Worksheets]
        for need in (SPEC, KJ):
            if need not in names:
                raise SystemExit(f"{os.path.basename(path)}: нет листа «{need}» — книга не изменена")
        app.Calculation = -4135
        keep = {}
        if SHEET in names:
            ws = wb.Worksheets(SHEET)
            if not force and str(ws.Range("A1").Formula or "").find(MARK) >= 0:
                ws = None
            else:
                for col in ("G", "H"):
                    keep[col] = [str(x[0]).strip() for x in ws.Range(f"{col}4:{col}15").Value if x[0] not in (None, "")]
        else:
            ws = wb.Worksheets.Add(None, wb.Worksheets(SPEC))
            ws.Name = SHEET
        sh = wb.Worksheets(SHEET)
        if sh.Index != wb.Worksheets(SPEC).Index + 1:      # место — сразу после «Спецификации»
            sh.Move(None, wb.Worksheets(SPEC))
            log.append(f"лист «{SHEET}» — после «{SPEC}»")
        if sh.Tab.Color != 0x317DED:                     # оранжевый ярлык, как «Спецификация» (tidy_book)
            sh.Tab.Color = 0x317DED
            log.append(f"лист «{SHEET}» — ярлык оранжевый")
        if ws is not None:
            build(wb, ws, keep)
            log.append(f"лист «{SHEET}» — {'пересоздан' if keep else 'создан'} (ВОР: оборудование, кабели, трубы)")
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
    args = sys.argv[1:]
    force = "--force" in args
    args = [a for a in args if a != "--force"]
    if not args:
        print(__doc__)
        sys.exit(1)
    for p in args:
        run(p, force)
