# -*- coding: utf-8 -*-
"""Порядок в книге v3.12: «Пульт», порядок и цвет листов, оформленная «Справочная», ТКЗ-отчёт, уровни столбцов.

    py tidy_book.py <книга.xlsx> [...]   — книга закрыта; повторный запуск ничего не портит

Формулы, столбцы и строки не меняются, ничего не переносится. Что делается:
  * «Пульт» — первый лист: порядок работы со ссылками, ошибки, итоги (формулами из «Справочной»).
  * «Справочная» — только оформление: у каждого блока заголовок, подпись «что это / где используется»,
    рамка; пустые столбцы между блоками узкие; закреплены строки 1–3.
    Устаревший блок I:J («Ток — кабель», не используется с v3.10) очищается.
  * «ТКЗ» — отчёт для печати без своего ввода: до шин — «Питание_щита», линии — «Исходные данные» BI–BM.
    Старый ручной расчёт по участкам (упрощённый, ввод вторично) удаляется; на лист никто не ссылается.
  * «Краткое руководство» -> «История_версий».
  * Формульный лист «Бирки» удаляется (на него никто не ссылается): бирки и наклейки создаёт кнопка
    надстройки (SX_Tags / SX_Labels) значениями. «Спецификация» остаётся — её заполняет макрос.
  * Листы не скрываются (так решил пользователь). Порядок: Пульт, Исходные данные, Разбивка_по_щитам,
    Однолинейка, Нагрузка_щитов, Кабельный_журнал, Спецификация, В Акад — затем справочники и редкие расчёты
    (серые ярлыки). Скрыты только служебные «Спецификация_авто», «Диапазоны», как в v3.12.
  * «Исходные данные» и «Однолинейка»: уровни столбцов 1 / 2 / 3 (кнопки над заголовками):
    1 — ввод, 2 — + результаты, 3 — всё. Книга сохраняется на уровне 2.
Свой скрытый экземпляр Excel; закрывается только он.
COM: только позиционные аргументы — именованные через позднее связывание Excel теряет
(Move(After=…) уносит лист в новую книгу).
"""
import os
import signal
import sys

PWD = "sx"
REF, PULT = "Справочная", "Пульт"
GUIDE_OLD, GUIDE_NEW = "Краткое руководство", "История_версий"
SRC, ODN = "Исходные данные", "Однолинейка"

# блоки «Справочной»: столбцы, заголовок, подпись (что это / где используется), последняя строка
BLOCKS = [
    ("B:D", "СОПРОТИВЛЕНИЕ ЖИЛ", "ρ и x0 по сечению. Для потерь напряжения («Исходные данные» W, X).", 19),
    ("F:G", "НОМИНАЛЫ АВТОМАТОВ", "Ряд: расчётный ток -> номинал автомата («Исходные данные» U).", 23),
    ("L:N", "МАРКИ КАБЕЛЕЙ", "Короткое обозначение -> марка кабеля. Вводится в «Исходных данных» N.", 51),
    ("P:Z", "КОДЫ ПОТРЕБИТЕЛЕЙ", "Главная таблица. Код из номера линии -> наименование, кабель, U, фазы, категория. "
                                "Новый код — новая строка здесь и строка в «УГО по коду».", 51),
    ("AA:AB", "ТРУБЫ", "Обозначения труб для кабельного журнала.", 7),
    ("AD:AG", "ЭТАЖИ", "Обозначение этажа -> подъём, опуск. Для длины линии.", 19),
    ("AJ:AO", "ПТО (АРТИКУЛЫ КАБЕЛЯ)", "Кабель -> артикул. Для кабельного журнала.", 200),
    ("AQ:AS", "СПОСОБ УПРАВЛЕНИЯ", "Ключ в выноске (AY «Исходных данных») -> цвет номера линии.", 10),
    ("AU:AZ", "ПО КОДУ: УГО, 30 мА, МИН. АВТОМАТ, ГРУППА Кс",
     "Для «Однолинейки»: значок, 30 мА, минимальный автомат, клеммы, группа Кс.", 75),
    ("BA:BC", "ЗНАЧКИ УГО", "Название значка -> блок AutoCAD (SXDRAW). Свой значок — новая строка.", 32),
    ("BD:BD", "ПОЖАРНАЯ", "По коду: да / нет. Для категорий и режимов.", 52),
    ("BE:BF", "НАСТРОЙКИ И ИТОГИ", "Жёлтые — можно менять: мест на лист A3, мин. автомат перед БП. Ниже — итоги.", 26),
    ("BH:BN", "СВОБОДНЫЕ КАНАЛЫ МОДУЛЕЙ", "Считается само по «Однолинейке».", 41),
    ("BP:BR", "Кс (СП 256) И НАСТРОЙКИ СЕТИ", "Кс быт по мощности, Ко сил., Uн, К прокл., допустимые ΔU.", 33),
    ("BT:BV", "КАТЕГОРИИ ДЛЯ ПЗ", "Список категорий нагрузки для листа «Расчёт_нагрузок».", 33),
    ("BX:CA", "ДОПУСТИМЫЙ ТОК (ПУЭ 1.3)", "Iдоп по сечению: одно-, двух-, трёхжильные. Для подбора сечения.", 20),
    ("CC:CC", "СЛАБОТОЧНЫЕ КОДЫ", "Коды линий без тока и нагрузки: DALI, KNX…", 40),
]
OBSOLETE = "I2:J19"
OBSOLETE_NAMES = ["Кабель_номиналы"]
# цвета заголовков блоков по очереди (тёмный — заголовок, светлый — шапка)
PALETTE = [((47, 85, 151), (221, 235, 247)), ((84, 130, 53), (226, 239, 218)), ((191, 143, 0), (255, 242, 204))]

# порядок листов: рабочие, затем справочники и редкие расчёты (все видны — так решил пользователь),
# служебные «Спецификация_авто», «Диапазоны» скрыты, как в v3.12
MAIN = [PULT, SRC, "Разбивка_по_щитам", ODN, "Нагрузка_щитов",
        "Кабельный_журнал", "Спецификация", "В Акад"]
OTHER = [  # (лист, что это) — список на «Пульте»
    (REF, "коды потребителей, УГО, кабели, нормы, настройки"),
    ("Каталог", "аппараты и модули для спецификации"),
    ("Справочная_потребители", "двигатели, светодиоды, диаметры кабелей"),
    ("Сеть_щита", "питание щита (сверху), нижестоящие щиты, токи КЗ по линиям"),
    ("Питание_щита", "питающая линия, ΔU и ток КЗ до шин"),          # книги до merge_network.py
    ("Нижестоящие_щиты", "нагрузка дочерних щитов"),
    ("Режимы", "категории надёжности, режимы питания"),
    ("ИБП", "мощность ИБП и АКБ"),
    ("Расчёт_нагрузок", "таблица нагрузок для ПЗ"),
    ("ТКЗ", "отчёт по токам КЗ для печати"),
    (GUIDE_NEW, "история версий таблицы"),
]
SERVICE = ["Спецификация_авто", "Диапазоны"]
# цвет ярлыка (RGB): ввод — зелёный, результат — оранжевый, пульт — синий, остальные — серый
GREEN, ORANGE, BLUE, GREY = (112, 173, 71), (237, 125, 49), (68, 114, 196), (166, 166, 166)
TAB = {PULT: BLUE, SRC: GREEN, "Разбивка_по_щитам": GREEN, ODN: GREEN,
       "Нагрузка_щитов": ORANGE, "Кабельный_журнал": ORANGE, "Спецификация": ORANGE,
       "Бирки": ORANGE, "Наклейки": ORANGE, "В Акад": ORANGE}

# уровни столбцов: 1 — результат (видно на «2»), 2 — служебное (видно на «3»); не указано — ввод (всегда)
LEVELS = {
    SRC: {1: "P T U V AB AD AT AU AV AW AX AY BC BE BF BG BI BJ BK BM BP BQ BR BS BU BV CB",
          2: "A W X Y Z AA AC AE AF AG AH AI AJ AK AL AM AO AP AQ AR AS BL BT BW BX BY BZ CA"},
    ODN: {1: "C G H M O Q T U V AA",
          2: "W X Z AB AC AD AF:BZ"},
}

STEPS = [
    ("1", SRC, "Линии с планов. AutoCAD: DLN — заполнить, DLN_UPD — обновить, DLN_CHK — проверка."),
    ("2", "Разбивка_по_щитам", "Щит и номер линии для каждой строки."),
    ("3", ODN, "Автоматы, аппараты, модули, клеммы. Белые столбцы — ввод. Столбец AE — статус."),
    ("4", "Нагрузка_щитов", "Проверить нагрузку и фазы щита."),
    ("5", None, "AutoCAD: SXDRAW — построить схему, SXUPDATE — обновить. CPDRAW — компоновка."),
    ("6", "Кабельный_журнал", "Кабельный журнал: кнопка «КЖ в DWG» на ленте."),
    ("7", "Спецификация", "Лента «Однолинейка»: «Обновить спецификацию». Бирки и наклейки — тоже кнопками."),
]


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


def sheet(wb, name):
    try:
        return wb.Worksheets(name)
    except Exception:
        return None


def unprotect_all(wb):
    saved = {}
    for ws in wb.Worksheets:
        if ws.ProtectContents:
            p = ws.Protection
            # порядок аргументов Worksheet.Protect
            saved[ws.Name] = (True, True, True, False,
                              p.AllowFormattingCells, p.AllowFormattingColumns, p.AllowFormattingRows,
                              p.AllowInsertingColumns, p.AllowInsertingRows, p.AllowInsertingHyperlinks,
                              p.AllowDeletingColumns, p.AllowDeletingRows, p.AllowSorting,
                              p.AllowFiltering, p.AllowUsingPivotTables)
            ws.Unprotect(PWD)
    return saved


def protect_all(wb, saved):
    for name, opts in saved.items():
        wb.Worksheets(name).Protect(PWD, *opts)


def format_reference(wb, app, log):
    ws = wb.Worksheets(REF)
    if ws.Range("B1").Value == BLOCKS[0][2]:
        log.append("«Справочная» уже оформлена — пропуск")
        return
    hdr = ws.Range("L2")
    if hdr.MergeCells:                      # общий заголовок L2:AG2 «Заполнить!!!»
        hdr.MergeArea.UnMerge()
    ws.Range(OBSOLETE).Clear()
    for n in OBSOLETE_NAMES:
        try:
            wb.Names(n).Delete()
        except Exception:
            pass
    ws.Range("BH3").Value = None            # старый заголовок блока каналов — теперь в строке 2
    ws.Range("BE2").Value = None            # «…(перенесено с листа «Справочник_схемы»)»
    # I:J очищен — подпись в «Как подбирается автомат» ведёт на допустимый ток
    ws.Range("BE18:BE30").Replace("по сечению (I:J)", "по допустимому току сечения (BX:CA, с К прокл.)", 2)
    for i, (cols, title, note, last) in enumerate(BLOCKS):
        dark, light = PALETTE[i % len(PALETTE)]
        c1, c2 = cols.split(":")
        r1 = ws.Range(f"{c1}1:{c2}1")
        r2 = ws.Range(f"{c1}2:{c2}2")
        r3 = ws.Range(f"{c1}3:{c2}3")
        whole = ws.Range(f"{c1}1:{c2}{last}")
        r2.UnMerge()
        r2.ClearContents()
        r1.Merge()
        r1.Value = note
        r1.WrapText = True
        r1.VerticalAlignment = -4160          # верх
        r1.HorizontalAlignment = -4131        # влево
        r1.Font.Size = 9
        r1.Font.Italic = True
        r1.Font.Color = rgb((89, 89, 89))
        r2.Merge()
        r2.Value = title
        r2.HorizontalAlignment = -4108
        r2.Font.Bold = True
        r2.Font.Color = rgb((255, 255, 255))
        r2.Interior.Color = rgb(dark)
        r3.Interior.Color = rgb(light)
        r3.Font.Bold = True
        r3.WrapText = True
        r3.VerticalAlignment = -4108
        r3.HorizontalAlignment = -4108
        for edge in (7, 8, 9, 10):          # рамка блока: лево, верх, низ, право
            b = whole.Borders(edge)
            b.LineStyle = 1
            b.Weight = -4138                  # средняя
            b.Color = rgb(dark)
        if ws.Columns(c1).ColumnWidth < 12 and c1 == c2:
            ws.Columns(c1).ColumnWidth = 16
    # пустые столбцы между блоками — узкие
    used = set()
    for cols, *_ in BLOCKS:
        c1, c2 = cols.split(":")
        used.update(range(ws.Range(c1 + "1").Column, ws.Range(c2 + "1").Column + 1))
    last_col = ws.Range("CC1").Column
    for c in range(1, last_col + 2):
        if c not in used:
            ws.Columns(c).ColumnWidth = 2
    ws.Rows(1).RowHeight = 60
    ws.Rows(2).RowHeight = 20
    ws.Activate()
    win = app.ActiveWindow
    win.FreezePanes = False
    win.ScrollRow, win.ScrollColumn = 1, 1
    ws.Range("A4").Select()
    win.FreezePanes = True
    win.ScrollColumn = ws.Range("P1").Column - 1    # открывается на кодах потребителей
    ws.Range("P4").Select()
    log.append(f"«Справочная» оформлена: {len(BLOCKS)} блоков, I:J (устар.) очищен")


TKZ = "ТКЗ"
TKZ_TITLE = "ТОКИ КОРОТКОГО ЗАМЫКАНИЯ — ОТЧЁТ"
TKZ_LINES = 998                              # строк линий: как в «Исходных данных» (3…1000)


def rebuild_tkz(wb, log):
    """ТКЗ — отчёт для печати без своего ввода: до шин — из «Питание_щита», линии — из «Исходных данных»."""
    ws = sheet(wb, TKZ)
    if ws is None or ws.Range("A1").Value == TKZ_TITLE:
        return
    ws.Cells.Clear()
    for i in range(ws.Shapes.Count, 0, -1):
        ws.Shapes(i).Delete()
    ws.Cells.Font.Size = 10
    ws.Range("A1").Value = TKZ_TITLE
    ws.Range("A1").Font.Bold = True
    ws.Range("A1").Font.Size = 13
    ws.Range("A2").Value = ("Только для печати, ввода здесь нет. Питающая линия и трансформатор — лист «Питание_щита», "
                            "линии — «Исходные данные» (BI–BM). ГОСТ 28249, ПУЭ 1.7.79.")
    ws.Range("A2").Font.Italic = True
    # 1. до шин щита: по щиту в строке (столбцы C / D / E «Питание_щита», раздел 1 — строки 6 / 7 / 8)
    ws.Range("A4").Value = "1. ДО ШИН ЩИТА"
    hdr1 = ["Щит", "Питается от", "Питающий кабель", "L, м", "In аппарата, А", "ΔU до шин, %",
            "Zп ф–0 на шинах, Ом", "Iкз min (1ф), А", "Iкз max (3ф), кА", "Отключение КЗ на шинах",
            "Термическая стойкость"]
    for j, h in enumerate(hdr1):
        ws.Cells(5, 1 + j).Value = h
    P = "Питание_щита!"
    for k, c in enumerate("CDE"):
        r = 6 + k
        a = f"$A{r}"
        f = [
            f'=IF({P}{c}$28="","",{P}{c}$28)',
            f'=IF({a}="","",IF(N({P}$D${6 + k})>0,"вышестоящий щит (связь)",IF({P}{c}12="","—",{P}{c}12&"")))',
            f'=IF(OR({a}="",N({P}{c}19)=0),"",TRIM({P}{c}18&" "&IF(N({P}{c}21)>1,{P}{c}21&"×","")'
            f'&IF(N({P}{c}20)=0,5,{P}{c}20)&"×"&{P}{c}19))',
            f'=IF(OR({a}="",N({P}{c}22)=0),"",{P}{c}22)',
            f'=IF(OR({a}="",N({P}{c}59)=0),"",{P}{c}59)',
            f'=IF({a}="","",{P}{c}33)',
            f'=IF(OR({a}="",NOT(ISNUMBER({P}{c}34))),"",ROUND({P}{c}34,4))',
            f'=IF(OR({a}="",NOT(ISNUMBER({P}{c}35))),"",ROUND({P}{c}35,0))',
            f'=IF(OR({a}="",NOT(ISNUMBER({P}{c}36))),"",ROUND({P}{c}36,2))',
            f'=IF({a}="","",{P}{c}39&"")',
            f'=IF({a}="","",{P}{c}40&"")',
        ]
        for j, x in enumerate(f):
            ws.Cells(r, 1 + j).Formula = x
    # 2. линии: N — номер строки «Исходных данных» со следующей линией (поиск от предыдущей, без летучих функций)
    ws.Range("A10").Value = "2. ЛИНИИ ЩИТА (КЗ в конце линии)"
    hdr2 = ["№", "Щит", "Линия", "Наименование", "Кабель", "L, м", "In, А", "Zп в конце, Ом",
            "Iкз min, А", "Iкз min / In", "Отключение ≤ 0,4 с", "Iкз max, кА"]
    for j, h in enumerate(hdr2):
        ws.Cells(11, 1 + j).Value = h
    ws.Range("N11").Value = "строка ИД"
    S = "'Исходные данные'!"
    r0, r1 = 12, 12 + TKZ_LINES - 1
    ws.Range(f"N{r0}").Formula = f'=IFERROR(MATCH(TRUE,INDEX(LEN({S}$D$3:$D$1000)>0,0),0)+2,"")'
    ws.Range(f"N{r0 + 1}:N{r1}").Formula = (
        f'=IF(N{r0}="","",IFERROR(N{r0}+MATCH(TRUE,INDEX(LEN(INDEX({S}$D$1:$D$1000,N{r0}+1):{S}$D$1000)>0,0),0),""))')

    def pick(col, rnd=None):
        v = f"INDEX({S}${col}$1:${col}$1000,$N{r0})"
        if rnd is not None:
            v = f"IF(ISNUMBER({v}),ROUND({v},{rnd}),{v}&\"\")"
        return f'=IF($N{r0}="","",{v})'
    cols = [f'=IF($N{r0}="","",ROW()-{r0 - 1})', pick("BZ"), pick("D"), pick("AJ"), pick("AE"), pick("M"),
            pick("U"), pick("BL", 4), pick("BI", 0), pick("BJ", 2), pick("BK"), pick("BM", 2)]
    for j, x in enumerate(cols):
        ws.Range(ws.Cells(r0, 1 + j), ws.Cells(r1, 1 + j)).Formula = x
    ws.Columns("N").Hidden = True
    # оформление
    for rng in (ws.Range("A5:K5"), ws.Range("A11:L11")):
        rng.Font.Bold = True
        rng.WrapText = True
        rng.VerticalAlignment = -4108
        rng.HorizontalAlignment = -4108
        rng.Interior.Color = rgb((221, 235, 247))
    for rng in (ws.Range("A5:K8"), ws.Range(f"A11:L{r1}")):
        rng.Borders.LineStyle = 1
        rng.Borders.Color = rgb((191, 191, 191))
    for addr in ("A4", "A10"):
        ws.Range(addr).Font.Bold = True
    for col, w in zip("ABCDEFGHIJKL", (6, 9, 13, 38, 24, 7, 8, 11, 10, 10, 16, 10)):
        ws.Columns(col).ColumnWidth = w
    ws.Rows(5).RowHeight = 30
    ws.Rows(11).RowHeight = 30
    ps = ws.PageSetup
    ps.PrintArea = f"$A$1:$L${r1}"
    ps.PrintTitleRows = "$11:$11"
    ps.Orientation = 2
    ps.Zoom = False
    ps.FitToPagesWide = 1
    ps.FitToPagesTall = False
    log.append("ТКЗ: отчёт без своего ввода (до шин — «Питание_щита», линии — «Исходные данные»)")


def drop_tags(wb, log):
    """Формульный лист «Бирки» (v3.12) — удалить; лист из значений (создан кнопкой) не трогать."""
    ws = sheet(wb, "Бирки")
    if ws is None or not ws.Range("A3").HasFormula:
        return
    ws.Visible = -1
    ws.Delete()
    log.append("формульный лист «Бирки» удалён — бирки создаёт кнопка надстройки")


def build_pult(wb, log):
    if sheet(wb, PULT):
        return
    ws = wb.Worksheets.Add(wb.Worksheets(1))
    ws.Name = PULT
    ws.Columns("A").ColumnWidth = 2
    ws.Columns("B").ColumnWidth = 24
    ws.Columns("C").ColumnWidth = 12
    ws.Range("B1").Value = "ОДНОЛИНЕЙКА — ПУЛЬТ"
    ws.Range("B1").Font.Bold = True
    ws.Range("B1").Font.Size = 14
    ws.Range("B3").Value = "ПОРЯДОК РАБОТЫ"
    ws.Range("B3").Font.Bold = True
    r = 4
    for num, target, text in STEPS:
        label = f"{num}. {target or 'AutoCAD'}"
        if target:
            ws.Hyperlinks.Add(ws.Cells(r, 2), "", f"'{target}'!A1", target, label)
        else:
            ws.Cells(r, 2).Value = label
        ws.Cells(r, 3).Value = text
        r += 1
    ws.Cells(r, 2).Value = "Ошибки"
    ws.Cells(r, 3).Formula = ('="Однолинейка — ошибок: "&COUNTIF(Однолинейка!$BU$3:$BU$1000,2)'
                              '&", предупреждений: "&COUNTIF(Однолинейка!$BU$3:$BU$1000,1)')
    ws.Range(ws.Cells(r, 2), ws.Cells(r, 3)).Font.Bold = True
    r += 2
    ws.Cells(r, 2).Value = "ИТОГИ ПО ОДНОЛИНЕЙКЕ"
    ws.Cells(r, 2).Font.Bold = True
    for src_row in range(8, 17):            # «Справочная» BE8:BF16 — подпись и число
        r += 1
        ws.Cells(r, 2).Formula = f'=IF(Справочная!$BE${src_row}="","",TRIM(Справочная!$BE${src_row}))'
        ws.Cells(r, 3).Formula = f'=IF(Справочная!$BE${src_row}="","",Справочная!$BF${src_row})'
        ws.Cells(r, 3).HorizontalAlignment = -4152
    r += 1
    ws.Cells(r, 2).Value = "Настройки (мест на лист A3, мин. автомат перед БП) — «Справочная», блок «Настройки и итоги»."
    ws.Cells(r, 2).Font.Italic = True
    r += 2
    ws.Cells(r, 2).Value = "СПРАВОЧНИКИ И РЕДКИЕ РАСЧЁТЫ"
    ws.Cells(r, 2).Font.Bold = True
    for name, text in OTHER:
        r += 1
        ws.Hyperlinks.Add(ws.Cells(r, 2), "", f"'{name}'!A1", name, name)
        ws.Cells(r, 3).Value = text
    ws.Range(f"B3:C{r}").WrapText = False
    log.append("лист «Пульт» собран")


def arrange_sheets(wb, log):
    g = sheet(wb, GUIDE_OLD)
    if g is not None:
        g.Name = GUIDE_NEW
    order = MAIN + [n for n, _ in OTHER] + SERVICE
    present = [n for n in order if sheet(wb, n) is not None]
    # скрытый лист Move молча не двигает — сначала показать все;
    # Move только с позиционным Before: ставим с конца списка каждый лист первым
    for n in present:
        wb.Worksheets(n).Visible = -1
    count = wb.Worksheets.Count
    for n in reversed(present):
        wb.Worksheets(n).Move(wb.Worksheets(1))
        if wb.Worksheets.Count != count:
            raise RuntimeError(f"лист «{n}» ушёл из книги при переносе — книга не сохранена")
    for n in present:
        ws = wb.Worksheets(n)
        ws.Tab.Color = rgb(TAB.get(n, GREY))
        ws.Visible = 0 if n in SERVICE else -1
    others = [ws.Name for ws in wb.Worksheets if ws.Name not in order]
    if others:
        log.append("листы вне списка (не тронуты): " + ", ".join(others))
    log.append(f"листов: {wb.Worksheets.Count}, видимых: {sum(1 for ws in wb.Worksheets if ws.Visible == -1)}")


def set_levels(wb, log):
    for name, levels in LEVELS.items():
        ws = sheet(wb, name)
        if ws is None:
            continue
        try:
            ws.Outline.ShowLevels(0, 8)
        except Exception:
            pass                               # на листе ещё нет групп
        ws.Columns.ClearOutline()
        ws.Columns.Hidden = False
        for lvl, spec in sorted(levels.items()):
            for c in spec.split():
                ws.Columns(c).OutlineLevel = lvl + 1
        ws.Outline.SummaryColumn = 0          # кнопка «+» слева от группы
        ws.Outline.ShowLevels(0, 2)
        log.append(f"«{name}»: уровни столбцов 1/2/3, открыт уровень 2")


def freeze_src(wb, app):
    ws = wb.Worksheets(SRC)
    ws.Activate()
    win = app.ActiveWindow
    if not win.FreezePanes:
        win.ScrollRow, win.ScrollColumn = 1, 1
        ws.Range("E3").Select()
        win.FreezePanes = True
    ws.Range("D3").Select()


def run(path):
    path = os.path.abspath(path)
    app, pid = excel()
    log = []
    try:
        wb = app.Workbooks.Open(path, 0, False)
        app.Calculation = -4135
        app.ScreenUpdating = False
        names0 = [ws.Name for ws in wb.Worksheets]
        saved = unprotect_all(wb)
        # показать «Справочную» на время оформления (окно нужно для закрепления строк)
        wb.Worksheets(REF).Visible = -1
        format_reference(wb, app, log)
        rebuild_tkz(wb, log)
        drop_tags(wb, log)
        build_pult(wb, log)
        arrange_sheets(wb, log)
        set_levels(wb, log)
        freeze_src(wb, app)
        protect_all(wb, saved)
        lost = [n for n in names0 if n not in (GUIDE_OLD, "Бирки") and sheet(wb, n) is None]
        if lost or app.Workbooks.Count != 1:
            raise RuntimeError(f"пропали листы {lost} — книга не сохранена")
        wb.Worksheets(PULT).Activate()
        wb.Worksheets(PULT).Range("A1").Select()
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
    for s in log:
        print("  " + s)


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(1)
    for p in sys.argv[1:]:
        run(p)
