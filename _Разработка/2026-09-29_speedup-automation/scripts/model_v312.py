# -*- coding: utf-8 -*-
"""model_v312.py - внедрение модели «категории надёжности, режимы, ИБП» (v3.12) в книгу «Однолинейка».

    py model_v312.py <книга.xlsx> [--out <другая.xlsx>] [--keep-visible]

Что делает (идемпотентно: повторный запуск перезаписывает формулы и подписи, но НЕ трогает
введённые пользователем значения - категории, схему щита, таблицу «Шины щита», настройки ИБП):

  «Исходные данные» BN:CA      категория (ввод), пожарная (ввод), итоги, шина, тип шины, источник,
                               обеспечена, проверка, служебные столбцы; список секции BB расширен до 1/2/3
  «Справочная»  BD3:BD51       пожарная по коду (пресет), BE16:BF16 - число необеспеченных линий
  «Нагрузка_щитов» AR:BA       расчёт по линиям (категория, пожарная, источник, слагаемые Рр)
  «Питание_щита» раздел 6      схема щита, «Шины щита»; J44:S47 - итоги для вышестоящего щита
  «Режимы», «ИБП»              новые листы (после «Нагрузка_щитов»)
  «Краткое руководство»        строка v3.12 в ведомости изменений, правило 21, правка правила 5
  скрывает «Спецификация_авто» и «Диапазоны» и переносит их в конец книги

Правки только через Excel COM (openpyxl ломает эти книги): отдельный невидимый экземпляр,
DisplayAlerts=False, UpdateLinks=0, Excel закрывается всегда (при зависании убивается свой PID).
Не вставляет строки/столбцы, не меняет листы «Однолинейка» (пароль sx), «Разбивка_по_щитам» и др.
"""
import argparse
import os
import shutil
import subprocess
import sys
import time

CAT = ["I особая", "I", "II", "III"]
CAT_ARR = '{"III","II","I","I особая"}'
SOURCES = ["Ввод 1", "Ввод 2", "Ввод 3", "ИБП", "ДГУ"]
BUS_TYPES = ["Ввод 1", "Ввод 2", "Ввод 3", "ИБП", "ДГУ", "Одиночная"]
NG = "Нагрузка_щитов"
SD = "'Исходные данные'"
LAST = 1000                       # последняя строка расчётных диапазонов книги (как в «Исходные_данные»)

# цвета Excel (BGR)
C_INPUT = 13431551                # светло-жёлтый - ввод
C_HEAD = 14277081                 # серый заголовок
C_SECT = 16247773                 # голубой (заголовки «Справочной»)
C_TAB = 3243501                   # оранжевый ярлык - результаты
C_GREY = 8421504
C_BAD_FILL = 13551615             # RGB(255,199,206)
C_BAD_FONT = 393372               # RGB(156,0,6)
C_OK_FILL = 13561798              # RGB(198,239,206)
C_OK_FONT = 24832                 # RGB(0,97,0)


DEC = ","                        # десятичный разделитель Excel (COM отдаёт NumberFormat в локальной записи)


def setnf(rng, code):
    rng.NumberFormat = code.replace(".", DEC)


def log(*a):
    print(*a, flush=True)


# ------------------------------------------------------------------------------------------------
# Excel
# ------------------------------------------------------------------------------------------------
def _excel_pids():
    try:
        out = subprocess.run(["tasklist", "/FI", "IMAGENAME eq EXCEL.EXE", "/FO", "CSV", "/NH"],
                             capture_output=True).stdout.decode("latin-1")
    except Exception:
        return set()
    s = set()
    for ln in out.splitlines():
        p = [x.strip('"') for x in ln.split('","')]
        if len(p) > 1 and p[0].upper() == "EXCEL.EXE":
            try:
                s.add(int(p[1]))
            except ValueError:
                pass
    return s


class ExcelSession:
    def __enter__(self):
        import pythoncom
        import win32com.client
        import win32process
        pythoncom.CoInitialize()
        self.pythoncom = pythoncom
        self.pid = None
        self.xl = win32com.client.DispatchEx("Excel.Application")     # свой экземпляр
        try:
            self.pid = win32process.GetWindowThreadProcessId(self.xl.Hwnd)[1]
        except Exception:
            pass
        xl = self.xl
        xl.Visible = False
        xl.DisplayAlerts = False
        xl.AskToUpdateLinks = False
        xl.EnableEvents = False
        try:
            xl.AutomationSecurity = 3          # без макросов
        except Exception:
            pass
        return xl

    def __exit__(self, *exc):
        try:
            for wb in list(self.xl.Workbooks):
                try:
                    wb.Close(False)
                except Exception:
                    pass
        except Exception:
            pass
        try:
            self.xl.Quit()
        except Exception:
            pass
        self.xl = None
        try:
            self.pythoncom.CoUninitialize()
        except Exception:
            pass
        if self.pid:
            for _ in range(12):
                if self.pid not in _excel_pids():
                    break
                time.sleep(0.5)
            if self.pid in _excel_pids():
                subprocess.run(["taskkill", "/F", "/PID", str(self.pid)], capture_output=True)
        return False


def sheet(wb, name):
    for ws in wb.Worksheets:
        if ws.Name == name:
            return ws
    return None


def is_empty(cell):
    v = cell.Value
    return v is None or (isinstance(v, str) and v == "")


def put_default(ws, addr, value):
    """Записать значение, только если ячейка пуста (ввод пользователя не затирается)."""
    c = ws.Range(addr)
    if is_empty(c) and not c.HasFormula:
        c.Value = value


def style_head(rng, fill=C_HEAD, bold=True, wrap=True):
    rng.Font.Bold = bold
    rng.Interior.Color = fill
    rng.WrapText = wrap
    rng.HorizontalAlignment = -4108
    rng.VerticalAlignment = -4108
    for i in (7, 8, 9, 10, 11, 12):
        try:
            rng.Borders(i).LineStyle = 1
        except Exception:
            pass


def box(rng):
    for i in (7, 8, 9, 10, 11, 12):
        try:
            rng.Borders(i).LineStyle = 1
        except Exception:
            pass


def add_list_validation(rng, items, sep, msg=None):
    v = rng.Validation
    v.Delete()
    v.Add(3, 1, 1, sep.join(items))
    v.IgnoreBlank = True
    v.InCellDropdown = True
    if msg:
        v.ErrorTitle = "Значение"
        v.ErrorMessage = msg
        v.ShowError = True


# ------------------------------------------------------------------------------------------------
# формулы
# ------------------------------------------------------------------------------------------------
def KC(x):
    """Кс быт по таблице (интерполяция) - как «Нагрузка_щитов» AF7."""
    m = "MATCH(%s,Кс_Р,1)" % x
    return ("IF({x}>=INDEX(Кс_Р,ROWS(Кс_Р)),INDEX(Кс_К,ROWS(Кс_К)),INDEX(Кс_К,{m})+({x}-INDEX(Кс_Р,{m}))"
            "*(INDEX(Кс_К,{m}+1)-INDEX(Кс_К,{m}))/(INDEX(Кс_Р,{m}+1)-INDEX(Кс_Р,{m})))").format(x=x, m=m)


def ngr(col):
    return "%s!$%s$3:$%s$%d" % (NG, col, col, LAST)


def SIFS(sumcol, sel, src=None, fire=None, cat=None):
    a = [ngr(sumcol), ngr("C"), sel]
    if src is not None:
        a += [ngr("AU"), '"%s"' % src]
    if fire is not None:
        a += [ngr("AT"), '"%s"' % fire]
    if cat is not None:
        a += [ngr("AS"), '"%s"' % cat]
    return "SUMIFS(%s)" % ",".join(a)


RAWCOLS = ["AV", "AW", "AX", "AY", "AZ"]         # Руст, Руст быт, Рр сил, Руст пост, ΣP·cos - «Нагрузка_щитов»


# ------------------------------------------------------------------------------------------------
# 1. «Справочная»: пожарная по коду
# ------------------------------------------------------------------------------------------------
FIRE_FORMULA = ('=IF(P@r="","",IF(OR(ISNUMBER(SEARCH("противопожар",Z@r&"")),ISNUMBER(SEARCH("дымоудал",Q@r&"")),'
                'ISNUMBER(SEARCH("противодым",Q@r&"")),ISNUMBER(SEARCH("подпор",Q@r&"")),'
                'ISNUMBER(SEARCH("огнезадерж",Q@r&"")),AND(ISNUMBER(SEARCH("пожарн",Q@r&"")),'
                'ISNUMBER(SEARCH("насос",Q@r&""))),UPPER(LEFT(P@r,3))="СПЗ",UPPER(LEFT(P@r,3))="ПНС"),"да","нет"))')


def step_reference(wb, sep):
    ws = sheet(wb, "Справочная")
    ws.Range("BD2").Value = "ПОЖАР (v3.12)"
    ws.Range("BD3").Value = "Пожарная (да/нет)"
    for a in ("BD2", "BD3"):
        c = ws.Range(a)
        c.Font.Bold = True
        c.Interior.Color = C_SECT
        c.WrapText = True
        c.HorizontalAlignment = -4108
    ws.Columns("BD").ColumnWidth = 14
    for r in range(4, 52):
        c = ws.Range("BD%d" % r)
        if is_empty(c) and not c.HasFormula:                 # пресет; поверх можно записать да/нет
            c.Formula = FIRE_FORMULA.replace("@r", str(r))
    rng = ws.Range("BD4:BD51")
    rng.Interior.Color = C_INPUT
    rng.Locked = False
    rng.HorizontalAlignment = -4108
    add_list_validation(rng, ["да", "нет"], sep)
    ws.Range("BD52").Value = ("Пресет «пожарная» по коду P: категория ПЗ «Системы противопожарной защиты» или дымоудаление, "
                              "подпор, огнезадерживающий, пожарный насос, СПЗ, ПНС. Формулу можно заменить на да / нет.")
    ws.Range("BD52").Font.Size = 8
    ws.Range("BD52").Font.Color = C_GREY
    # сводка: число необеспеченных линий
    ws.Range("BE16").Value = "Линий с необеспеченной категорией (v3.12)"
    ws.Range("BF16").Formula = '=COUNTIF(%s!$BV$3:$BV$%d,"НЕ ОБЕСПЕЧЕНА")' % (SD, LAST)
    for a, b in (("BE15", "BE16"), ("BF15", "BF16")):
        ws.Range(a).Copy(ws.Range(b))
    ws.Range("BE16").Value = "Линий с необеспеченной категорией (v3.12)"
    ws.Range("BF16").Formula = '=COUNTIF(%s!$BV$3:$BV$%d,"НЕ ОБЕСПЕЧЕНА")' % (SD, LAST)


# ------------------------------------------------------------------------------------------------
# 2. «Питание_щита»: раздел 6 (схема, «Шины щита») - до формул «Исходных данных» (имена!)
# ------------------------------------------------------------------------------------------------
R_SCHEME = 64                      # заголовок «Схема щита»
R_VVOD, R_AVR, R_IBP, R_PROV, R_PROV2 = 65, 66, 67, 68, 69
R_BUS_H = 72                       # заголовок таблицы «Шины щита»
R_BUS1, R_BUS2 = 73, 82


def step_power_sections(wb, sep):
    ws = sheet(wb, "Питание_щита")
    def cell(a, v, bold=False, size=10, color=None, wrap=True, align=None):
        c = ws.Range(a)
        c.Value = v
        c.Font.Name = "Arial"
        c.Font.Size = size
        c.Font.Bold = bold
        if color is not None:
            c.Font.Color = color
        c.WrapText = wrap
        if align is not None:
            c.HorizontalAlignment = align
        return c

    ws.Range("B3").Value = ("Категории надёжности, схема щита (вводы, АВР, ИБП) и таблица «Шины щита» — раздел 6 (строка 62); "
                            "результаты — листы «Режимы» и «ИБП».")
    ws.Range("B3").Font.Name = "Arial"
    ws.Range("B3").Font.Size = 9
    ws.Range("B3").Font.Color = C_GREY
    ws.Range("B3").WrapText = False

    cell("B62", "6. КАТЕГОРИИ НАДЁЖНОСТИ: СХЕМА ЩИТА И ШИНЫ (для листов «Режимы» и «ИБП»)", bold=True, wrap=False)
    cell("B63", "Жёлтое — ввод. Требуемая категория линии — «Исходные данные» BN, шина линии — «Однолинейка» D (пусто = основная), "
                "секция — «Исходные данные» BB. Шина «ИБП» / «ДГУ» — по имени или по таблице ниже.", size=9, color=C_GREY, wrap=False)
    # схема щита
    cell("B%d" % R_SCHEME, "Схема щита")
    for i, col in enumerate("CDE"):
        cell("%s%d" % (col, R_SCHEME), "=$B$%d" % (6 + i))
    cell("F%d" % R_SCHEME, "ед.")
    cell("G%d" % R_SCHEME, "Пояснение")
    style_head(ws.Range("B%d:G%d" % (R_SCHEME, R_SCHEME)))
    rows = [
        (R_VVOD, "Число вводов (1 / 2 / 3)", 1, "шт.", "пусто = 1. 2 и более вводов без АВР обеспечивают II категорию, с АВР — I"),
        (R_AVR, "АВР (да / нет)", "нет", "", "АВР на вводах / секционном выключателе; пусто = нет"),
        (R_IBP, "ИБП / ЩГП питается от", "Ввод 1", "", "секция, к которой подключён вход ИБП / ЩГП; вход ИБП учитывается на этом вводе («Режимы»)"),
    ]
    for r, lab, dflt, unit, note in rows:
        cell("B%d" % r, lab)
        for col in "CDE":
            put_default(ws, "%s%d" % (col, r), dflt)
        cell("F%d" % r, unit, align=-4108)
        cell("G%d" % r, note)
        rng = ws.Range("C%d:E%d" % (r, r))
        rng.Interior.Color = C_INPUT
        rng.HorizontalAlignment = -4108
        rng.Font.Name = "Arial"
        rng.Font.Size = 10
        box(ws.Range("B%d:G%d" % (r, r)))
    add_list_validation(ws.Range("C%d:E%d" % (R_VVOD, R_VVOD)), ["1", "2", "3"], sep)
    add_list_validation(ws.Range("C%d:E%d" % (R_AVR, R_AVR)), ["да", "нет"], sep)
    add_list_validation(ws.Range("C%d:E%d" % (R_IBP, R_IBP)), ["Ввод 1", "Ввод 2", "Ввод 3"], sep)
    cell("B%d" % R_PROV, "Обеспечиваемая категория: шина «Ввод»")
    cell("B%d" % R_PROV2, "Обеспечиваемая категория: шина «ИБП», «ДГУ»")
    for col in "CDE":
        cell("%s%d" % (col, R_PROV),
             '=IF(%s$%d="","",IF(IFERROR(--%s%d,1)>=2,IF(LOWER(TRIM(%s%d&""))="да","I","II"),"III"))'
             % (col, R_SCHEME, col, R_VVOD, col, R_AVR), bold=True, align=-4108)
        cell("%s%d" % (col, R_PROV2), '=IF(%s$%d="","","I особая")' % (col, R_SCHEME), bold=True, align=-4108)
    cell("G%d" % R_PROV, "один ввод — III; 2+ ввода без АВР — II; 2+ ввода с АВР — I (шина «Одиночная» — всегда III)")
    cell("G%d" % R_PROV2, "нагрузка на ИБП или ДГУ — I особая")
    box(ws.Range("B%d:G%d" % (R_PROV, R_PROV2)))
    # шины щита
    cell("B71", "Шины щита: тип шины вручную (если шины нет в таблице — тип по имени и секции)", bold=True, wrap=False)
    cell("B%d" % R_BUS_H, "Шина (как в «Однолинейка» D; основная — «основная»)")
    cell("C%d" % R_BUS_H, "Щит (пусто = любой)")
    cell("D%d" % R_BUS_H, "Тип шины")
    cell("G%d" % R_BUS_H, "Пояснение")
    style_head(ws.Range("B%d:G%d" % (R_BUS_H, R_BUS_H)))
    rng = ws.Range("B%d:D%d" % (R_BUS1, R_BUS2))
    rng.Interior.Color = C_INPUT
    rng.Font.Name = "Arial"
    rng.Font.Size = 10
    rng.HorizontalAlignment = -4108
    box(ws.Range("B%d:G%d" % (R_BUS1, R_BUS2)))
    add_list_validation(ws.Range("D%d:D%d" % (R_BUS1, R_BUS2)), BUS_TYPES, sep)
    notes = {R_BUS1: "Типы: Ввод 1 / 2 / 3 — источник по секции; ИБП, ДГУ — отдельные источники (I особая); Одиночная — без резерва (III).",
             R_BUS1 + 1: "Без таблицы: имя шины содержит «ИБП» — ИБП, «ДГУ» — ДГУ; иначе по секции («Исходные данные» BB 1/2/3 → Ввод N, пусто → Ввод 1).",
             R_BUS1 + 2: "Пример: шина «ЩГП» → ИБП. Если имя шины есть в таблице, тип берётся отсюда."}
    for r, t in notes.items():
        cell("G%d" % r, t, size=9, color=C_GREY)
    ws.Columns("G").ColumnWidth = 57.22
    # имена
    for nm, ref in (("ШЩ_шина", "=Питание_щита!$B$%d:$B$%d" % (R_BUS1, R_BUS2)),
                    ("ШЩ_щит", "=Питание_щита!$C$%d:$C$%d" % (R_BUS1, R_BUS2)),
                    ("ШЩ_тип", "=Питание_щита!$D$%d:$D$%d" % (R_BUS1, R_BUS2)),
                    ("СХ_вводы", "=Питание_щита!$C$%d:$E$%d" % (R_VVOD, R_VVOD)),
                    ("СХ_АВР", "=Питание_щита!$C$%d:$E$%d" % (R_AVR, R_AVR)),
                    ("СХ_ИБП_от", "=Питание_щита!$C$%d:$E$%d" % (R_IBP, R_IBP)),
                    ("СХ_кат_ввод", "=Питание_щита!$C$%d:$E$%d" % (R_PROV, R_PROV))):
        try:
            wb.Names(nm).RefersTo = ref
        except Exception:
            wb.Names.Add(nm, ref)


# ------------------------------------------------------------------------------------------------
# 3. «Исходные данные» BN:CA
# ------------------------------------------------------------------------------------------------
SD_COLS = [  # (столбец, заголовок, ввод?)
    ("BN", "Категория надёжности (ввод: I особая / I / II / III; пусто = III)", True),
    ("BO", "Пожарная (ввод: да / нет; пусто = по коду, «Справочная» BD)", True),
    ("BP", "Категория (итог)", False),
    ("BQ", "Пожарная (итог)", False),
    ("BR", "Шина (из «Однолинейки» D; пусто = основная)", False),
    ("BS", "Тип шины", False),
    ("BT", "Источник для режимов", False),
    ("BU", "Категория обеспечена", False),
    ("BV", "Категория: проверка", False),
    ("BW", "служебн.: № необеспеч.", False),
    ("BX", "служебн.: строка «Однолинейки»", False),
    ("BY", "служебн.: шина как в D (группа)", False),
    ("BZ", "служебн.: щит", False),
    ("CA", "служебн.: тип по таблице «Шины щита»", False),
]
ONE_R = "Однолинейка!$R$3:$R$%d" % LAST


def sd_formulas():
    """Формулы «Исходных данных» для строки 3 (ссылки относительные)."""
    f = {}
    f["BP"] = ('=IF($D3="","",IFERROR(INDEX({"I особая","I","II","III"},MATCH(TRIM(BN3&""),{"I особая","I","II","III"},0)),"III"))')
    f["BQ"] = ('=IF($D3="","",IF(TRIM(BO3&"")<>"",IF(LOWER(TRIM(BO3&""))="да","да","нет"),'
               'IF(IFERROR(INDEX(Справочная!$BD$4:$BD$51,MATCH($AH3,Справочная!$P$4:$P$51,0)),"")="да","да","нет")))')
    f["BX"] = '=IF($D3="","",IFERROR(MATCH($D3,%s,0),0))' % ONE_R
    f["BY"] = ('=IF(N($BX3)=0,"",IF(INDEX(Однолинейка!$D$3:$D$1000,$BX3)&""<>"",INDEX(Однолинейка!$D$3:$D$1000,$BX3)&"",'
               'IF(INDEX(Однолинейка!$AG$3:$AG$1000,$BX3)&""="","",IFERROR(LOOKUP(2,1/((Однолинейка!$A$3:$A$1000=INDEX(Однолинейка!$A$3:$A$1000,$BX3))'
               '*(Однолинейка!$AJ$3:$AJ$1000=INDEX(Однолинейка!$AG$3:$AG$1000,$BX3))),Однолинейка!$D$3:$D$1000)&"",""))))')
    f["BR"] = '=IF($D3="","",IF($BY3="","основная",$BY3))'
    f["BZ"] = '=IF($D3="","",IFERROR(INDEX(Разбивка_по_щитам!$F$1:$F$1000,MATCH($D3,Разбивка_по_щитам!$H$1:$H$1000,0))&"",""))'
    f["CA"] = ('=IF($D3="","",IFERROR(INDEX(ШЩ_тип,MATCH(1,INDEX(((ШЩ_щит=$BZ3)+(ШЩ_щит=""))*(ШЩ_шина=$BR3)*(ШЩ_тип<>""),0),0))&"",""))')
    f["BS"] = ('=IF($D3="","",IF($CA3<>"",$CA3,IF(ISNUMBER(SEARCH("ИБП",$BR3)),"ИБП",IF(ISNUMBER(SEARCH("ДГУ",$BR3)),"ДГУ",'
               '"Ввод "&IFERROR(MIN(3,MAX(1,INT(--$BB3))),1)))))')
    f["BT"] = ('=IF($D3="","",IF(OR($BS3="ИБП",$BS3="ДГУ"),$BS3,IF(LEFT($BS3,5)="Ввод ",$BS3,"Ввод "&IFERROR(MIN(3,MAX(1,INT(--$BB3))),1))))')
    f["BU"] = ('=IF($D3="","",IF(OR($BS3="ИБП",$BS3="ДГУ"),"I особая",IF($BS3="Одиночная","III",'
               'IF($BZ3="","III",IFERROR(INDEX(СХ_кат_ввод,MATCH($BZ3,ПЩ_щиты,0)),"III")))))')
    f["BV"] = ('=IF($D3="","",IF(IFERROR(MATCH($BP3,%s,0),1)>IFERROR(MATCH($BU3,%s,0),1),"НЕ ОБЕСПЕЧЕНА","ОК"))' % (CAT_ARR, CAT_ARR))
    f["BW"] = '=IF($BV3="НЕ ОБЕСПЕЧЕНА",COUNTIF($BV$3:$BV3,"НЕ ОБЕСПЕЧЕНА"),"")'
    return f


def step_source_data(wb, sep):
    ws = sheet(wb, "Исходные данные")
    for i, (col, head, is_input) in enumerate(SD_COLS):
        ws.Range("%s1" % col).Value = 66 + i
        c = ws.Range("%s2" % col)
        c.Value = head
        c.Font.Name = "Calibri"
        c.Font.Size = 11
        c.WrapText = True
        c.HorizontalAlignment = -4108
        c.VerticalAlignment = -4108
        box(c)
        if is_input:
            c.Interior.Color = C_INPUT
        if col in ("BW", "BX", "BY", "BZ", "CA"):
            c.Font.Color = C_GREY
            c.Font.Size = 9
    widths = {"BN": 16, "BO": 16, "BP": 13, "BQ": 11, "BR": 16, "BS": 11, "BT": 12, "BU": 13, "BV": 17,
              "BW": 10, "BX": 11, "BY": 14, "BZ": 12, "CA": 14}
    for col, w in widths.items():
        ws.Columns(col).ColumnWidth = w
    body = ws.Range("BN3:CA%d" % LAST)
    body.Font.Name = "Calibri"
    body.Font.Size = 11
    body.HorizontalAlignment = -4108
    box(body)
    ws.Range("BW3:CA%d" % LAST).Font.Color = C_GREY
    for col, f in sd_formulas().items():
        ws.Range("%s3:%s%d" % (col, col, LAST)).Formula = f
    inp = ws.Range("BN3:BO%d" % LAST)
    inp.Locked = False
    inp.Interior.Color = C_INPUT
    add_list_validation(ws.Range("BN3:BN%d" % LAST), CAT, sep, "Выберите: I особая, I, II, III (пусто = III)")
    add_list_validation(ws.Range("BO3:BO%d" % LAST), ["да", "нет"], sep, "да / нет (пусто = по коду потребителя)")
    # секция: допускаем 1 / 2 / 3 (было 1 / 2); в «Нагрузка_щитов» AG:AH секция 3 по-прежнему считается как 1
    add_list_validation(ws.Range("BB3:BB%d" % LAST), ["1", "2", "3"], sep, "Секция: 1, 2 или 3 (пусто = 1)")
    ws.Range("BB2").Value = "Секция (1/2/3; пусто = 1)"
    # одно правило условного форматирования на диапазон
    rng = ws.Range("BN3:BV%d" % LAST)
    have = False
    try:
        for i in range(1, rng.FormatConditions.Count + 1):
            fc = rng.FormatConditions(i)
            if fc.Type == 2 and "НЕ ОБЕСПЕЧЕНА" in str(fc.Formula1):
                have = True
    except Exception:
        pass
    if not have:
        fc = rng.FormatConditions.Add(2, None, '=$BV3="НЕ ОБЕСПЕЧЕНА"')
        fc.Interior.Color = C_BAD_FILL
        fc.Font.Color = C_BAD_FONT


# ------------------------------------------------------------------------------------------------
# 4. «Нагрузка_щитов» AR:BA - расчёт по линиям
# ------------------------------------------------------------------------------------------------
NG_COLS = [
    ("AR", "Строка в «Исходных данных» (служебное)"),
    ("AS", "Категория (итог)"),
    ("AT", "Пожарная"),
    ("AU", "Источник для режимов"),
    ("AV", "Руст (без «Нет»), кВт"),
    ("AW", "Руст быт, кВт"),
    ("AX", "Рр сил (P·Кс), кВт"),
    ("AY", "Руст пост, кВт"),
    ("AZ", "P·cos φ"),
    ("BA", "Категория: проверка"),
]


def ng_formulas():
    f = {}
    f["AR"] = '=IF($A3="","",IFERROR(MATCH($D3,%s!$D$1:$D$%d,0),0))' % (SD, LAST)
    for col, src in (("AS", "BP"), ("AT", "BQ"), ("AU", "BT"), ("BA", "BV")):
        f[col] = '=IF(N($AR3)=0,"",INDEX(%s!$%s$1:$%s$%d,$AR3)&"")' % (SD, src, src, LAST)
    f["AV"] = '=IF(OR($A3="",$A3="Нет"),0,IFERROR(N($H3),0))'
    f["AW"] = '=IF($A3="Быт",IFERROR(N($H3),0),IF($A3="Щит",N($AN3),0))'
    f["AX"] = '=IF($A3="Сил",N($AK3),IF($A3="Щит",N($AP3),0))'
    f["AY"] = '=IF($A3="Пост",IFERROR(N($H3),0),IF($A3="Щит",N($AQ3),0))'
    f["AZ"] = '=IF(OR($A3="",$A3="Нет"),0,IFERROR(N($L3),0))'
    return f


def step_loads(wb, sep):
    ws = sheet(wb, NG)
    for col, head in NG_COLS:
        c = ws.Range("%s2" % col)
        c.Value = head
        c.Font.Name = "Calibri"
        c.Font.Size = 11
        c.WrapText = True
        c.HorizontalAlignment = -4108
        box(c)
        ws.Columns(col).ColumnWidth = 12
    ws.Range("AR2").Font.Color = C_GREY
    ws.Range("AR2").Font.Size = 9
    ws.Columns("AR").ColumnWidth = 11
    ws.Columns("AS").ColumnWidth = 12
    ws.Columns("BA").ColumnWidth = 16
    ws.Range("AR1").Value = "расчёт по линиям для листов «Режимы», «ИБП» (v3.12)"
    ws.Range("AR1").Font.Size = 9
    ws.Range("AR1").Font.Color = C_GREY
    for col, f in ng_formulas().items():
        ws.Range("%s3:%s%d" % (col, col, LAST)).Formula = f
    body = ws.Range("AR3:BA%d" % LAST)
    body.Font.Name = "Calibri"
    body.Font.Size = 11
    body.HorizontalAlignment = -4108
    setnf(ws.Range("AV3:AZ%d" % LAST), "0.00")


# ------------------------------------------------------------------------------------------------
# 5. Лист «Режимы»
# ------------------------------------------------------------------------------------------------
MODE_HDR = ["Ввод 1", "Ввод 2", "Ввод 3", "ИБП", "ДГУ", "Итого"]
MCOL = ["C", "D", "E", "F", "G", "H"]
FMT_P = '0.00;-0.00;"–"'
FMT_K = '0.00;-0.00;"–"'

# раскладка листа (строки)
R = {}
R["sel"] = 3
R["eff"] = 4
R["scheme"] = 5
R["feed"] = 6
R["k"] = 7
R["bad"] = 8
R["sect1"] = 10
R["hdr"] = 11
R["cnt"] = 12
BLOCK = 14                       # строк в блоке режима (заголовок + 8 + 5 слагаемых)
R["m1"] = 13
R["m2"] = R["m1"] + BLOCK + 1    # 28
R["m3"] = R["m2"] + BLOCK + 1    # 43
R["m4"] = R["m3"] + BLOCK + 1    # 58
R["m5"] = R["m4"] + 7            # 65
R["m6"] = R["m5"] + BLOCK + 1    # 80
R["m7"] = R["m6"] + BLOCK + 1    # 95
R["cat"] = R["m7"] + 6           # 101
R["badsec"] = R["cat"] + 9       # 110
R["chk"] = R["badsec"] + 24      # 134
R["svc"] = R["chk"] + 10         # 144


def block_rows(t):
    return dict(t=t, rust=t + 1, kc=t + 2, cos=t + 3, tg=t + 4, P=t + 5, Q=t + 6, S=t + 7, I=t + 8,
                hb=t + 9, hk=t + 10, hs=t + 11, hp=t + 12, hc=t + 13)


def ensure_new_sheets(wb):
    """Листы должны существовать до записи формул (иначе Excel примет «ИБП!» за внешний файл)."""
    if sheet(wb, "Режимы") is None:
        ws = wb.Worksheets.Add(None, wb.Worksheets(NG))
        ws.Name = "Режимы"
    if sheet(wb, "ИБП") is None:
        ws = wb.Worksheets.Add(None, wb.Worksheets("Режимы"))
        ws.Name = "ИБП"


def step_modes(wb, sep):
    ws = sheet(wb, "Режимы")
    ws.Cells.Font.Name = "Arial"
    ws.Cells.Font.Size = 10
    ws.Tab.Color = C_TAB
    ws.Columns("A").ColumnWidth = 1.5
    ws.Columns("B").ColumnWidth = 50
    for col in "CDEFGHIJKL":
        ws.Columns(col).ColumnWidth = 13

    def put(a, v, bold=False, size=10, color=None, fmt=None, align=None, wrap=False, formula=False):
        c = ws.Range(a)
        if formula or (isinstance(v, str) and v.startswith("=")):
            c.Formula = v
        else:
            c.Value = v
        c.Font.Bold = bold
        c.Font.Size = size
        if color is not None:
            c.Font.Color = color
        if fmt:
            setnf(c, fmt)
        if align is not None:
            c.HorizontalAlignment = align
        c.WrapText = wrap
        return c

    put("B1", "РЕЖИМЫ ПИТАНИЯ, КАТЕГОРИИ НАДЁЖНОСТИ", bold=True, size=12)
    put("B2", "Расчёт по линиям «Исходных данных» (категория BN, шина «Однолинейка» D, секция BB, пожарная BO) и схеме щита «Питание_щита» раздел 6. "
              "Пожарные нагрузки в максимум не входят. Серые строки — слагаемые расчёта.", size=9, color=C_GREY)
    put("B3", "Щит (выбрать; пусто = первый)")
    c = ws.Range("C3")
    c.Interior.Color = C_INPUT
    c.HorizontalAlignment = -4108
    v = c.Validation
    v.Delete()
    v.Add(3, 1, 1, "=Питание_щита!$B$6:$B$8")
    v.IgnoreBlank = True
    box(c)
    put("B4", "Расчёт для щита", bold=True)
    put("C4", '=IF(C3<>"",C3&"",IFERROR(Питание_щита!$B$6&"",""))', bold=True, align=-4108)
    put("B5", "Число вводов / АВР (ввод: «Питание_щита» C65:E66)")
    put("C5", '=IFERROR(INDEX(СХ_вводы,MATCH($C$4,ПЩ_щиты,0)),"")', align=-4108)
    put("D5", '=IFERROR(INDEX(СХ_АВР,MATCH($C$4,ПЩ_щиты,0))&"","")', align=-4108)
    put("B6", "ИБП / ЩГП питается от (ввод: «Питание_щита» C67:E67)")
    put("C6", '=IF(IFERROR(INDEX(СХ_ИБП_от,MATCH($C$4,ПЩ_щиты,0))&"","")="","Ввод 1",INDEX(СХ_ИБП_от,MATCH($C$4,ПЩ_щиты,0))&"")', align=-4108)
    put("B7", "Номер щита в книге (1–3)", color=C_GREY)
    put("C7", '=IFERROR(MATCH($C$4,Питание_щита!$B$6:$B$8,0),0)', color=C_GREY, align=-4108)
    put("B8", "Проверка категорий (вся книга): линий не обеспечено", bold=True)
    put("C8", '=COUNTIF(%s!$BV$3:$BV$%d,"НЕ ОБЕСПЕЧЕНА")' % (SD, LAST), bold=True, align=-4108)
    put("D8", '=IF(C8=0,"Все категории обеспечены","Не обеспечено: "&C8&" линий")', bold=True)
    rng = ws.Range("B8:D8")
    rng.FormatConditions.Delete()
    fc = rng.FormatConditions.Add(2, None, "=$C$8>0")
    fc.Interior.Color = C_BAD_FILL
    fc.Font.Color = C_BAD_FONT
    fc2 = rng.FormatConditions.Add(2, None, "=$C$8=0")
    fc2.Interior.Color = C_OK_FILL
    fc2.Font.Color = C_OK_FONT

    sel = "$C$4"
    feed = "$C$6"
    # ---- служебный блок: строки
    S = R["svc"]
    svc = {}
    svc["title"] = S
    svc["hdr"] = S + 1
    svc["nf"] = {s: S + 2 + i for i, s in enumerate(SOURCES)}           # без пожарных
    svc["f"] = {s: S + 7 + i for i, s in enumerate(SOURCES)}            # пожарные
    svc["ibp"] = S + 12
    svc["all"] = S + 13
    svc["v1"] = S + 14
    svc["v2"] = S + 15
    svc["cat_title"] = S + 17
    svc["cat_hdr"] = S + 18
    per = 11
    svc["catk"] = {}
    for k in range(3):
        base = S + 19 + k * per
        svc["catk"][k] = dict(head=base, nf={j: base + 1 + i for i, j in enumerate(CAT)},
                              f={j: base + 5 + i for i, j in enumerate(CAT)}, nf_all=base + 9, f_all=base + 10)

    # ---- ИБП: ссылки на лист «ИБП» (см. step_ibp)
    IBP_PIN = "ИБП!$C$46"       # P вх на ЩГП, кВт
    IBP_COS = "ИБП!$C$39"       # cos φ входа

    # ---- служебные строки: по выбранному щиту
    put("B%d" % svc["title"], "СЛУЖЕБНОЕ: суммы по источникам выбранного щита (не править)", bold=True)
    heads = ["Руст, кВт", "Руст быт", "Рр сил (P·Кс)", "Руст пост", "Σ P·cos φ", "Кс быт", "Рр, кВт", "cos φ", "Qр, квар", "Sр, кВА"]
    for i, h in enumerate(heads):
        put("%s%d" % ("CDEFGHIJKL"[i], svc["hdr"]), h, bold=True, wrap=True, align=-4108)
    style_head(ws.Range("B%d:L%d" % (svc["hdr"], svc["hdr"])))
    put("B%d" % svc["hdr"], "Источник / без пожарных (нет) или пожарные (да)", bold=True, wrap=True)

    def derived(r):
        """H..L по сырым C..G строки r."""
        ws.Range("H%d" % r).Formula = "=IF($D%d=0,0," % r + KC("$D%d" % r) + ")"
        ws.Range("I%d" % r).Formula = "=$D%d*H%d+Ко_сил*$E%d+$F%d" % (r, r, r, r)
        ws.Range("J%d" % r).Formula = "=IF($C%d>0,MIN(1,$G%d/$C%d),0)" % (r, r, r)
        ws.Range("K%d" % r).Formula = "=IF(J%d>0,I%d*TAN(ACOS(J%d)),0)" % (r, r, r)
        ws.Range("L%d" % r).Formula = "=SQRT(I%d^2+K%d^2)" % (r, r)

    for s in SOURCES:
        for fire, key in (("нет", "nf"), ("да", "f")):
            r = svc[key][s]
            put("B%d" % r, "%s, %s" % (s, "без пожарных" if fire == "нет" else "пожарные"))
            for i, rc in enumerate(RAWCOLS):
                ws.Range("%s%d" % ("CDEFG"[i], r)).Formula = '=IF($C$4="",0,%s)' % SIFS(rc, sel, src=s, fire=fire)
            derived(r)
    r = svc["ibp"]
    put("B%d" % r, "Вход ИБП на ЩГП (лист «ИБП»): постоянная нагрузка ввода")
    ws.Range("C%d" % r).Formula = "=N(%s)" % IBP_PIN
    ws.Range("D%d" % r).Value = 0
    ws.Range("E%d" % r).Value = 0
    ws.Range("F%d" % r).Formula = "=C%d" % r
    ws.Range("G%d" % r).Formula = "=C%d*%s" % (r, "IF(N(%s)>0,%s,0.95)" % (IBP_COS, IBP_COS))
    derived(r)
    # объединения: все линии, Ввод 1, Ввод 2 (без входа ИБП) - для сверки
    for key, lab, srcs in (("all", "Все линии щита (без замены ИБП его входом)", SOURCES),
                           ("v1", "Ввод 1: все линии (с пожарными)", ["Ввод 1"]),
                           ("v2", "Ввод 2: все линии (с пожарными)", ["Ввод 2"])):
        r = svc[key]
        put("B%d" % r, lab)
        for i in range(5):
            col = "CDEFG"[i]
            parts = ["%s%d" % (col, svc["nf"][s]) for s in srcs] + ["%s%d" % (col, svc["f"][s]) for s in srcs]
            ws.Range("%s%d" % (col, r)).Formula = "=" + "+".join(parts)
        derived(r)
    # ---- категории по трём щитам
    put("B%d" % svc["cat_title"], "СЛУЖЕБНОЕ: категории по щитам 1–3 (для (б) и «Питание_щита» J44:S47)", bold=True)
    for i, h in enumerate(heads):
        put("%s%d" % ("CDEFGHIJKL"[i], svc["cat_hdr"]), h, bold=True, wrap=True, align=-4108)
    style_head(ws.Range("B%d:L%d" % (svc["cat_hdr"], svc["cat_hdr"])))
    put("B%d" % svc["cat_hdr"], "Щит / категория / пожарные", bold=True, wrap=True)
    for k in range(3):
        d = svc["catk"][k]
        put("B%d" % d["head"], "Щит %d:" % (k + 1), bold=True)
        ws.Range("C%d" % d["head"]).Formula = '=IFERROR(Питание_щита!$B$%d&"","")' % (6 + k)
        ws.Range("C%d" % d["head"]).Font.Bold = True
        hs = "$C$%d" % d["head"]
        for j in CAT:
            for fire, key in (("нет", "nf"), ("да", "f")):
                r = d[key][j]
                put("B%d" % r, "  %s%s" % (j, "" if fire == "нет" else " (пожарные)"))
                for i, rc in enumerate(RAWCOLS):
                    ws.Range("%s%d" % ("CDEFG"[i], r)).Formula = '=IF(%s="",0,%s)' % (hs, SIFS(rc, hs, fire=fire, cat=j))
                derived(r)
        for key, lab in (("nf_all", "  Все категории, без пожарных"), ("f_all", "  Все категории, пожарные")):
            r = d[key]
            put("B%d" % r, lab, bold=True)
            src = "nf" if key == "nf_all" else "f"
            for i in range(5):
                col = "CDEFG"[i]
                ws.Range("%s%d" % (col, r)).Formula = "=" + "+".join("%s%d" % (col, d[src][j]) for j in CAT)
            derived(r)
    setnf(ws.Range("C%d:L%d" % (svc["nf"]["Ввод 1"], svc["catk"][2]["f_all"])), "0.00")
    ws.Range("B%d:L%d" % (svc["title"], svc["catk"][2]["f_all"])).Font.Color = C_GREY
    ws.Range("B%d:L%d" % (svc["hdr"], svc["hdr"])).Font.Color = 0
    ws.Range("B%d:L%d" % (svc["cat_hdr"], svc["cat_hdr"])).Font.Color = 0

    # ---- главная таблица
    put("B%d" % R["sect1"], "1. РЕЖИМЫ ПО ИСТОЧНИКАМ (пожарные нагрузки в максимум не входят)", bold=True, size=11)
    put("B%d" % R["hdr"], "Режим / показатель")
    for i, h in enumerate(MODE_HDR):
        put("%s%d" % (MCOL[i], R["hdr"]), h)
    put("I%d" % R["hdr"], "Пояснение")
    style_head(ws.Range("B%d:I%d" % (R["hdr"], R["hdr"])))
    put("B%d" % R["cnt"], "Линий на источнике (без пожарных), шт.")
    for i, s in enumerate(SOURCES):
        ws.Range("%s%d" % (MCOL[i], R["cnt"])).Formula = '=IF($C$4="",0,COUNTIFS(%s,$C$4,%s,%s$%d,%s,"нет"))' % (
            ngr("C"), ngr("AU"), MCOL[i], R["hdr"], ngr("AT"))
    ws.Range("H%d" % R["cnt"]).Formula = "=SUM(C%d:G%d)" % (R["cnt"], R["cnt"])
    ws.Range("C%d:H%d" % (R["cnt"], R["cnt"])).HorizontalAlignment = -4108

    def vec(kind, srcname, extra=None):
        """5 сырых слагаемых (rust, быт, сил, пост, pcos) для столбца: список из 5 формул (без '=')."""
        return None

    comp_src = ["C", "D", "E", "F", "G"]              # столбцы служебных строк

    def comp(ci, srcs, with_in=None, fire=False):
        """Сумма ci-й компоненты по источникам srcs; with_in - условие (формула) добавления входа ИБП."""
        key = "f" if fire else "nf"
        col = comp_src[ci]
        parts = ["$%s$%d" % (col, svc[key][s]) for s in srcs]
        e = "+".join(parts) if parts else "0"
        if with_in:
            e += "+IF(%s,$%s$%d,0)" % (with_in, col, svc["ibp"])
        return e

    def fill_block(t, title, colspec, expl, fire=False, total=None):
        """colspec: словарь столбец -> (список источников, условие входа ИБП или None); total: список столбцов итога"""
        rr = block_rows(t)
        put("B%d" % t, title, bold=True)
        ws.Range("B%d:I%d" % (t, t)).Interior.Color = C_HEAD
        labels = {"rust": "Руст, кВт", "kc": "Кс (Pр / Руст)", "cos": "cos φ", "tg": "tg φ", "P": "Pр, кВт", "Q": "Qр, квар",
                  "S": "Sр, кВА", "I": "Iр, А", "hb": "  слагаемое: Руст быт, кВт", "hk": "  слагаемое: Кс быт (таблица СП 256)",
                  "hs": "  слагаемое: Рр сил = Σ P·Кс, кВт", "hp": "  слагаемое: Руст пост, кВт", "hc": "  слагаемое: Σ P·cos φ"}
        for key, lab in labels.items():
            put("B%d" % rr[key], lab, bold=key in ("P", "S", "I"), color=C_GREY if key.startswith("h") else None)
        ws.Range("B%d:I%d" % (rr["hb"], rr["hc"])).Font.Color = C_GREY
        ws.Range("B%d:I%d" % (rr["hb"], rr["hc"])).Font.Size = 9
        for col in MCOL:
            if col == "H":
                cols_sum = total
                for key in ("rust", "hb", "hs", "hp", "hc"):
                    ws.Range("H%d" % rr[key]).Formula = "=" + "+".join("%s%d" % (c, rr[key]) for c in cols_sum)
            else:
                srcs, cond = colspec[col]
                for key, ci in (("rust", 0), ("hb", 1), ("hs", 2), ("hp", 3), ("hc", 4)):
                    ws.Range("%s%d" % (col, rr[key])).Formula = "=" + comp(ci, srcs, cond, fire)
            ws.Range("%s%d" % (col, rr["hk"])).Formula = "=IF(%s%d=0,0," % (col, rr["hb"]) + KC("%s%d" % (col, rr["hb"])) + ")"
            ws.Range("%s%d" % (col, rr["P"])).Formula = "=%s%d*%s%d+Ко_сил*%s%d+%s%d" % (col, rr["hb"], col, rr["hk"], col, rr["hs"], col, rr["hp"])
            ws.Range("%s%d" % (col, rr["kc"])).Formula = "=IF(%s%d>0,%s%d/%s%d,0)" % (col, rr["rust"], col, rr["P"], col, rr["rust"])
            ws.Range("%s%d" % (col, rr["cos"])).Formula = "=IF(%s%d>0,MIN(1,%s%d/%s%d),0)" % (col, rr["rust"], col, rr["hc"], col, rr["rust"])
            ws.Range("%s%d" % (col, rr["tg"])).Formula = "=IF(%s%d>0,TAN(ACOS(%s%d)),0)" % (col, rr["cos"], col, rr["cos"])
            ws.Range("%s%d" % (col, rr["Q"])).Formula = "=%s%d*%s%d" % (col, rr["P"], col, rr["tg"])
            ws.Range("%s%d" % (col, rr["S"])).Formula = "=SQRT(%s%d^2+%s%d^2)" % (col, rr["P"], col, rr["Q"])
            ws.Range("%s%d" % (col, rr["I"])).Formula = "=%s%d/(SQRT(3)*Uн_лин)" % (col, rr["S"])
        setnf(ws.Range("C%d:H%d" % (rr["rust"], rr["hc"])), FMT_P)
        setnf(ws.Range("C%d:H%d" % (rr["kc"], rr["tg"])), FMT_K)
        setnf(ws.Range("C%d:H%d" % (rr["hk"], rr["hk"])), '0.000;-0.000;"–"')
        ws.Range("C%d:H%d" % (rr["rust"], rr["hc"])).HorizontalAlignment = -4108
        ws.Range("B%d:H%d" % (rr["P"], rr["P"])).Font.Bold = True
        ws.Range("B%d:H%d" % (rr["S"], rr["I"])).Font.Bold = True
        for i, key in enumerate(("rust", "P", "S", "I")):
            pass
        box(ws.Range("B%d:H%d" % (t, rr["hc"])))
        for i, (key, text) in enumerate(expl.items()):
            put("I%d" % rr[key], text, size=9, color=C_GREY)
        return rr

    cond_feed = lambda n: '%s="%s"' % (feed, n)
    m1 = fill_block(
        R["m1"], "1) НОРМАЛЬНЫЙ РЕЖИМ (СВ отключён, каждая шина на своём источнике)",
        {"C": (["Ввод 1"], cond_feed("Ввод 1")), "D": (["Ввод 2"], cond_feed("Ввод 2")), "E": (["Ввод 3"], cond_feed("Ввод 3")),
         "F": (["ИБП"], None), "G": (["ДГУ"], None)},
        {"rust": "Ввод N включает вход ИБП/ЩГП (лист «ИБП»), если ЩГП питается от него",
         "P": "столбец «ИБП» — нагрузка на выходе ИБП, справочно: в «Итого» не входит (вход ИБП уже во вводе)",
         "I": "симметричный Iр; по перекосу фаз — «Нагрузка_щитов» V7"},
        total=["C", "D", "E", "G"])
    m2 = fill_block(
        R["m2"], "2) АВАРИЙНЫЙ: потерян Ввод 1 (СВ включён, всё на Вводе 2)",
        {"C": ([], None), "D": (["Ввод 1", "Ввод 2"], 'OR(%s="Ввод 1",%s="Ввод 2")' % (feed, feed)),
         "E": (["Ввод 3"], cond_feed("Ввод 3")), "F": (["ИБП"], None), "G": (["ДГУ"], None)},
        {"P": "Кс быт по сумме Руст быт двух секций; вход ИБП переходит на оставшийся ввод"},
        total=["C", "D", "E", "G"])
    m3 = fill_block(
        R["m3"], "3) АВАРИЙНЫЙ: потерян Ввод 2 (СВ включён, всё на Вводе 1)",
        {"C": (["Ввод 1", "Ввод 2"], 'OR(%s="Ввод 1",%s="Ввод 2")' % (feed, feed)), "D": ([], None),
         "E": (["Ввод 3"], cond_feed("Ввод 3")), "F": (["ИБП"], None), "G": (["ДГУ"], None)},
        {}, total=["C", "D", "E", "G"])
    # 4) максимум
    t = R["m4"]
    put("B%d" % t, "4) МАКСИМУМ АВАРИЙНОГО РЕЖИМА (больший из вариантов 2 и 3 по Sр)", bold=True)
    ws.Range("B%d:I%d" % (t, t)).Interior.Color = C_HEAD
    for i, (key, lab) in enumerate((("P", "Pр, кВт"), ("Q", "Qр, квар"), ("S", "Sр, кВА"), ("I", "Iр, А"), ("sc", "Максимум при (какой ввод потерян)"))):
        put("B%d" % (t + 1 + i), lab, bold=key in ("P", "S", "I"))
    for col in MCOL:
        pick = "%s%d>=%s%d" % (col, m2["S"], col, m3["S"])
        ws.Range("%s%d" % (col, t + 1)).Formula = "=IF(%s,%s%d,%s%d)" % (pick, col, m2["P"], col, m3["P"])
        ws.Range("%s%d" % (col, t + 2)).Formula = "=IF(%s,%s%d,%s%d)" % (pick, col, m2["Q"], col, m3["Q"])
        ws.Range("%s%d" % (col, t + 3)).Formula = "=MAX(%s%d,%s%d)" % (col, m2["S"], col, m3["S"])
        ws.Range("%s%d" % (col, t + 4)).Formula = "=MAX(%s%d,%s%d)" % (col, m2["I"], col, m3["I"])
        ws.Range("%s%d" % (col, t + 5)).Formula = '=IF(%s%d<=0,"",IF(%s%d=%s%d,"оба варианта",IF(%s,"без Ввода 1","без Ввода 2")))' % (
            col, t + 3, col, m2["S"], col, m3["S"], pick)
    setnf(ws.Range("C%d:H%d" % (t + 1, t + 4)), FMT_P)
    ws.Range("C%d:H%d" % (t + 1, t + 5)).HorizontalAlignment = -4108
    ws.Range("B%d:H%d" % (t + 3, t + 4)).Font.Bold = True
    box(ws.Range("B%d:H%d" % (t, t + 5)))
    put("I%d" % (t + 1), "по каждому источнику берётся вариант с большим Sр; «Итого» одинаково в обоих вариантах", size=9, color=C_GREY)
    put("I%d" % (t + 2), "это режим для проверки вводов, аппаратов ГРЩ и кабелей (потеря одного ввода)", size=9, color=C_GREY)
    m5 = fill_block(
        R["m5"], "5) ПОТЕРЯНЫ ВСЕ ВВОДЫ (работают только шины ИБП и ДГУ)",
        {"C": ([], None), "D": ([], None), "E": ([], None), "F": (["ИБП"], None), "G": (["ДГУ"], None)},
        {"P": "ИБП — от батареи в течение автономии, затем от ДГУ; пожарные линии см. блок 6"},
        total=["F", "G"])
    m6 = fill_block(
        R["m6"], "6) ПОЖАР: пожарные нагрузки (не участвуют в расчёте максимума, п. 6.9 СП 31-110-2003)",
        {"C": (["Ввод 1"], None), "D": (["Ввод 2"], None), "E": (["Ввод 3"], None), "F": (["ИБП"], None), "G": (["ДГУ"], None)},
        {"P": "линии с признаком «пожарная» («Исходные данные» BQ): СПЗ, дымоудаление, подпор, пожарные насосы и т. п."},
        fire=True, total=["C", "D", "E", "F", "G"])
    t = R["m7"]
    put("B%d" % t, "7) НОРМАЛЬНЫЙ + ПОЖАРНЫЕ (для проверки вводов при пожаре; арифметическая сумма блоков 1 и 6)", bold=True)
    ws.Range("B%d:I%d" % (t, t)).Interior.Color = C_HEAD
    for i, (key, lab) in enumerate((("P", "Pр, кВт"), ("Q", "Qр, квар"), ("S", "Sр, кВА"), ("I", "Iр, А"))):
        put("B%d" % (t + 1 + i), lab, bold=key in ("S", "I", "P"))
    for col in MCOL:
        ws.Range("%s%d" % (col, t + 1)).Formula = "=%s%d+%s%d" % (col, m1["P"], col, m6["P"])
        ws.Range("%s%d" % (col, t + 2)).Formula = "=%s%d+%s%d" % (col, m1["Q"], col, m6["Q"])
        ws.Range("%s%d" % (col, t + 3)).Formula = "=SQRT(%s%d^2+%s%d^2)" % (col, t + 1, col, t + 2)
        ws.Range("%s%d" % (col, t + 4)).Formula = "=%s%d/(SQRT(3)*Uн_лин)" % (col, t + 3)
    setnf(ws.Range("C%d:H%d" % (t + 1, t + 4)), FMT_P)
    ws.Range("C%d:H%d" % (t + 1, t + 4)).HorizontalAlignment = -4108
    box(ws.Range("B%d:H%d" % (t, t + 4)))
    put("I%d" % (t + 1), "столбец «ИБП»: выход ИБП + пожарные линии на его шине (справочно)", size=9, color=C_GREY)

    # ---- (б) по категориям
    t = R["cat"]
    put("B%d" % t, "2. ПО КАТЕГОРИЯМ НАДЁЖНОСТИ (для ТУ и ПЗ; без привязки к секциям, ИБП — нагрузка на выходе)", bold=True, size=11)
    hd = ["Категория", "Линий, шт.", "Pуст, кВт", "Pр, кВт", "Qр, квар", "Sр, кВА", "Пожарные: Pуст, кВт", "Пожарные: Pр, кВт", "Не обеспечено, линий"]
    for i, h in enumerate(hd):
        put("%s%d" % ("BCDEFGHIJ"[i], t + 1), h)
    style_head(ws.Range("B%d:J%d" % (t + 1, t + 1)))
    ws.Rows(t + 1).RowHeight = 40

    def choose(col, key):
        return "CHOOSE($C$7," + ",".join("$%s$%d" % (col, svc["catk"][k][key] if not isinstance(key, tuple) else svc["catk"][k][key[0]][key[1]]) for k in range(3)) + ")"

    for i, j in enumerate(CAT):
        r = t + 2 + i
        put("B%d" % r, j, bold=True)
        ws.Range("C%d" % r).Formula = '=IF($C$4="",0,COUNTIFS(%s,$C$4,%s,$B%d,%s,"нет"))' % (ngr("C"), ngr("AS"), r, ngr("AT"))
        ws.Range("D%d" % r).Formula = "=IFERROR(%s,0)" % choose("C", ("nf", j))
        ws.Range("E%d" % r).Formula = "=IFERROR(%s,0)" % choose("I", ("nf", j))
        ws.Range("F%d" % r).Formula = "=IFERROR(%s,0)" % choose("K", ("nf", j))
        ws.Range("G%d" % r).Formula = "=IFERROR(%s,0)" % choose("L", ("nf", j))
        ws.Range("H%d" % r).Formula = "=IFERROR(%s,0)" % choose("C", ("f", j))
        ws.Range("I%d" % r).Formula = "=IFERROR(%s,0)" % choose("I", ("f", j))
        ws.Range("J%d" % r).Formula = '=IF($C$4="",0,COUNTIFS(%s,$C$4,%s,$B%d,%s,"НЕ ОБЕСПЕЧЕНА"))' % (ngr("C"), ngr("AS"), r, ngr("BA"))
    r = t + 6
    put("B%d" % r, "Итого (все категории)", bold=True)
    ws.Range("C%d" % r).Formula = "=SUM(C%d:C%d)" % (t + 2, t + 5)
    ws.Range("D%d" % r).Formula = "=IFERROR(%s,0)" % choose("C", "nf_all")
    ws.Range("E%d" % r).Formula = "=IFERROR(%s,0)" % choose("I", "nf_all")
    ws.Range("F%d" % r).Formula = "=IFERROR(%s,0)" % choose("K", "nf_all")
    ws.Range("G%d" % r).Formula = "=IFERROR(%s,0)" % choose("L", "nf_all")
    ws.Range("H%d" % r).Formula = "=IFERROR(%s,0)" % choose("C", "f_all")
    ws.Range("I%d" % r).Formula = "=IFERROR(%s,0)" % choose("I", "f_all")
    ws.Range("J%d" % r).Formula = "=SUM(J%d:J%d)" % (t + 2, t + 5)
    setnf(ws.Range("C%d:J%d" % (t + 2, t + 6)), FMT_P)
    setnf(ws.Range("C%d:C%d" % (t + 2, t + 6)), "0")
    setnf(ws.Range("J%d:J%d" % (t + 2, t + 6)), "0")
    ws.Range("C%d:J%d" % (t + 2, t + 6)).HorizontalAlignment = -4108
    ws.Range("B%d:J%d" % (t + 6, t + 6)).Font.Bold = True
    box(ws.Range("B%d:J%d" % (t + 1, t + 6)))
    put("B%d" % (t + 7), "Кс быт берётся по сумме Руст быт своего набора линий, поэтому сумма строк по Pр может отличаться от «Итого». "
                         "Пожарные линии в Pуст, Pр без пожарных не входят.", size=9, color=C_GREY)

    # ---- (в) необеспеченные линии
    t = R["badsec"]
    put("B%d" % t, "3. ЛИНИИ С НЕОБЕСПЕЧЕННОЙ КАТЕГОРИЕЙ (вся книга, первые 20)", bold=True, size=11)
    ws.Range("B%d" % (t + 1)).Formula = '=IF($C$8=0,"Все категории обеспечены","Не обеспечено: "&$C$8&" линий"&IF($C$8>20," (показаны первые 20)",""))'
    ws.Range("B%d" % (t + 1)).Font.Bold = True
    hd = ["Линия — наименование", "Щит", "Требуется", "Обеспечена", "Шина", "Тип шины"]
    for i, h in enumerate(hd):
        put("%s%d" % ("BCDEFG"[i], t + 2), h)
    style_head(ws.Range("B%d:G%d" % (t + 2, t + 2)))
    put("K%d" % (t + 2), "служебн.", size=8, color=C_GREY)
    for n in range(1, 21):
        r = t + 2 + n
        ws.Range("K%d" % r).Formula = '=IFERROR(MATCH(%d,%s!$BW$3:$BW$%d,0),"")' % (n, SD, LAST)
        ws.Range("K%d" % r).Font.Color = C_GREY
        ws.Range("K%d" % r).Font.Size = 8
        ws.Range("B%d" % r).Formula = '=IF($K%d="","",INDEX(%s!$D$3:$D$%d,$K%d)&" — "&INDEX(%s!$AJ$3:$AJ$%d,$K%d))' % (r, SD, LAST, r, SD, LAST, r)
        for col, src in (("C", "BZ"), ("D", "BP"), ("E", "BU"), ("F", "BR"), ("G", "BS")):
            ws.Range("%s%d" % (col, r)).Formula = '=IF($K%d="","",INDEX(%s!$%s$3:$%s$%d,$K%d))' % (r, SD, src, src, LAST, r)
    ws.Range("C%d:G%d" % (t + 3, t + 22)).HorizontalAlignment = -4108
    box(ws.Range("B%d:G%d" % (t + 2, t + 22)))

    # ---- (г) сверка
    t = R["chk"]
    put("B%d" % t, "4. СВЕРКА С «Нагрузка_щитов» (совпадает, если нет ИБП и пожарных линий)", bold=True, size=11)
    for i, h in enumerate(["Показатель", "«Нагрузка_щитов»", "«Режимы»", "Разница"]):
        put("%s%d" % ("BCDE"[i], t + 1), h)
    style_head(ws.Range("B%d:E%d" % (t + 1, t + 1)))
    ws.Rows(t + 1).RowHeight = 28

    def ngc(rows):
        return "IFERROR(CHOOSE($C$7,%s),0)" % ",".join("%s!$%s" % (NG, x) for x in rows)
    checks = [
        ("Руст весь щит, кВт", ngc(["AF$5", "AF$14", "AF$23"]), "$C$%d" % svc["all"]),
        ("Рр весь щит, кВт (аварийный)", ngc(["AF$10", "AF$19", "AF$28"]), "$I$%d" % svc["all"]),
        ("Руст секция 1 = Ввод 1, кВт", ngc(["AG$5", "AG$14", "AG$23"]), "$C$%d" % svc["v1"]),
        ("Рр секция 1 = Ввод 1, кВт", ngc(["AG$10", "AG$19", "AG$28"]), "$I$%d" % svc["v1"]),
        ("Руст секция 2 = Ввод 2, кВт", ngc(["AH$5", "AH$14", "AH$23"]), "$C$%d" % svc["v2"]),
        ("Рр секция 2 = Ввод 2, кВт", ngc(["AH$10", "AH$19", "AH$28"]), "$I$%d" % svc["v2"]),
    ]
    for i, (lab, a, b) in enumerate(checks):
        r = t + 2 + i
        put("B%d" % r, lab)
        ws.Range("C%d" % r).Formula = "=N(%s)" % a
        ws.Range("D%d" % r).Formula = "=%s" % b
        ws.Range("E%d" % r).Formula = "=D%d-C%d" % (r, r)
    r = t + 8
    put("B%d" % r, "Iр весь щит, А: «Нагрузка_щитов» V7 (по фазе) и симметричный")
    ws.Range("C%d" % r).Formula = "=N(%s)" % ngc(["V$7", "V$16", "V$25"])
    ws.Range("D%d" % r).Formula = "=$L$%d/(SQRT(3)*Uн_лин)" % svc["all"]
    ws.Range("E%d" % r).Formula = "=D%d-C%d" % (r, r)
    setnf(ws.Range("C%d:E%d" % (t + 2, t + 8)), "0.00")
    ws.Range("C%d:E%d" % (t + 2, t + 8)).HorizontalAlignment = -4108
    box(ws.Range("B%d:E%d" % (t + 1, t + 8)))
    put("F%d" % (t + 8), "разница — перекос фаз («Нагрузка_щитов» считает по наиболее нагруженной фазе)", size=9, color=C_GREY)
    put("F%d" % (t + 2), "разница ≠ 0: есть ИБП (учтён вход) или другое разбиение на секции (BB), чем на шинах", size=9, color=C_GREY)

    # для листа «ИБП» и «Питание_щита»
    return dict(m1=m1, m2=m2, m3=m3, m4=R["m4"], m5=m5, m6=m6, m7=R["m7"], svc=svc)


# ------------------------------------------------------------------------------------------------
# 6. Лист «ИБП»
# ------------------------------------------------------------------------------------------------
def step_ibp(wb, sep, info):
    ws = sheet(wb, "ИБП")
    ws.Cells.Font.Name = "Arial"
    ws.Cells.Font.Size = 10
    ws.Tab.Color = C_TAB
    ws.Columns("A").ColumnWidth = 1.5
    ws.Columns("B").ColumnWidth = 58
    ws.Columns("C").ColumnWidth = 13
    ws.Columns("D").ColumnWidth = 9
    ws.Columns("E").ColumnWidth = 96
    for col in "FGHIJKLMN":
        ws.Columns(col).ColumnWidth = 9

    m1 = info["m1"]

    def line(r, label, formula=None, unit="", note="", inp=None, fmt=None, bold=False):
        ws.Range("B%d" % r).Value = label
        c = ws.Range("C%d" % r)
        if formula is not None:
            c.Formula = formula
        if inp is not None:
            put_default(ws, "C%d" % r, inp)
            c.Interior.Color = C_INPUT
            box(c)
        if fmt:
            setnf(c, fmt)
        c.HorizontalAlignment = -4108
        c.Font.Bold = bold
        ws.Range("B%d" % r).Font.Bold = bold
        ws.Range("D%d" % r).Value = unit
        ws.Range("D%d" % r).HorizontalAlignment = -4108
        ws.Range("E%d" % r).Value = note
        ws.Range("E%d" % r).Font.Size = 9
        ws.Range("E%d" % r).Font.Color = C_GREY

    def head(r, text):
        ws.Range("B%d" % r).Value = text
        ws.Range("B%d" % r).Font.Bold = True
        ws.Range("B%d:E%d" % (r, r)).Interior.Color = C_HEAD

    ws.Range("B1").Value = "ИБП: мощность, аккумуляторная батарея, вход на ЩГП"
    ws.Range("B1").Font.Bold = True
    ws.Range("B1").Font.Size = 12
    ws.Range("B2").Value = ("Жёлтое — ввод. Лист считает ИБП шины типа «ИБП» выбранного на листе «Режимы» щита. Нагрузка шины берётся из «Режимов» (Pр без пожарных). "
                            "Вход ИБП уходит в «Режимы» как нагрузка ввода, от которого питается ЩГП (вместо мощности выхода).")
    ws.Range("B2").Font.Size = 9
    ws.Range("B2").Font.Color = C_GREY

    line(4, "Щит (выбирается на листе «Режимы» C3)", "=Режимы!$C$4", note="")
    line(5, "Шина типа ИБП в этом щите",
         '=IF(IF(Режимы!$C$4="",0,COUNTIFS(%s,Режимы!$C$4,%s,"ИБП"))>0,"есть","нет")' % (ngr("C"), ngr("AU")),
         note="тип шины: имя шины содержит «ИБП» («Однолинейка» D) или таблица «Шины щита» («Питание_щита» B73:D82)", bold=True)
    line(6, "Состояние", '=IF($C$5="есть","лист активен","нет шины ИБП: лист не используется, вход ИБП = 0")')
    ws.Range("C6").HorizontalAlignment = -4131

    head(8, "1. НАГРУЗКА ШИНЫ ИБП")
    line(9, "P расчётная шины ИБП (без пожарных)", '=IF($C$5="есть",Режимы!$F$%d,0)' % m1["P"], "кВт",
         "Pр из «Режимов» блок 1, столбец «ИБП»: Кс быт по таблице, Ко·Σ(P·Кс) для силовых, постоянные", fmt="0.00")
    line(10, "Добавка вручную (например, пожарные линии на ИБП)", None, "кВт",
         "пожарные линии на шине ИБП в расчётную мощность не входят: если ИБП их питает, впишите сюда их мощность", inp=0, fmt="0.00")
    line(11, "P нагрузки ИБП", '=C9+N(C10)', "кВт", "", fmt="0.00", bold=True)
    line(12, "cos φ нагрузки", '=IF(Режимы!$F$%d>0,Режимы!$F$%d,0.95)' % (m1["cos"], m1["cos"]), "", "средневзвешенный по линиям шины (0,95, если данных нет)", fmt="0.00")
    line(13, "S нагрузки ИБП", '=IF(C12>0,C11/C12,0)', "кВА", "S = P / cos φ", fmt="0.00", bold=True)

    head(15, "2. ПОДБОР ИБП")
    line(16, "Допустимая загрузка ИБП", None, "", "по умолчанию 0,75: запас на рост нагрузки и пусковые токи", inp=0.75, fmt="0.00")
    line(17, "Требуемая S ИБП (по кВА)", '=IF(N(C16)>0,C13/C16,0)', "кВА", "S нагрузки / допустимая загрузка", fmt="0.00")
    line(18, "cos φ на выходе ИБП (паспорт)", None, "", "по умолчанию 0,9; активная мощность ИБП = S ном × cos φ", inp=0.9, fmt="0.00")
    line(19, "Требуемая S ИБП (по кВт)", '=IF(AND(N(C16)>0,N(C18)>0),C11/(C18*C16),0)', "кВА", "P нагрузки / (cos φ вых × допустимая загрузка)", fmt="0.00")
    series = [6, 10, 15, 20, 30, 40, 60, 80, 100, 120, 160, 200]
    for i, v in enumerate(series):
        c = ws.Range("C%d" % (50 + i))
        put_default(ws, "C%d" % (50 + i), v)
        c.Interior.Color = C_INPUT
        c.HorizontalAlignment = -4108
        ws.Range("B%d" % (50 + i)).Value = "  типоразмер %d" % (i + 1)
        ws.Range("D%d" % (50 + i)).Value = "кВА"
        ws.Range("D%d" % (50 + i)).HorizontalAlignment = -4108
        box(ws.Range("B%d:D%d" % (50 + i, 50 + i)))
    ser = "$C$50:$C$61"
    near = '=IF(N(%s)<=0,0,IFERROR(SMALL(%s,COUNTIF(%s,"<"&%s)+1),"нет в ряду"))'
    line(20, "Ряд типоразмеров ИБП (раздел 5, C50:C61)", '=COUNT($C$50:$C$61)', "шт.", "ряд 6…200 кВА можно править в разделе 5 внизу листа (подбирается наименьший типоразмер, не меньше требуемого)", fmt="0")
    line(21, "Ближайший типоразмер по кВА (не меньше требуемого)", near % ("C17", ser, ser, "C17"), "кВА", "наименьший из ряда, не меньше требуемой S", fmt="0")
    line(22, "Ближайший типоразмер по кВт", near % ("C19", ser, ser, "C19"), "кВА", "проверка по активной мощности при cos φ вых", fmt="0")
    line(23, "Принятый типоразмер ИБП", '=IF(AND(ISNUMBER(C21),ISNUMBER(C22)),MAX(C21,C22),IF(N(C17)<=0,0,"нет в ряду"))', "кВА",
         "больший из двух подборов", fmt="0", bold=True)
    line(24, "Загрузка по кВА", '=IF(AND(ISNUMBER(C23),N(C23)>0),C13/C23,"")', "", "", fmt="0%")
    line(25, "Загрузка по кВт", '=IF(AND(ISNUMBER(C23),N(C23)>0,N(C18)>0),C11/(C23*C18),"")', "", "P / (S ном × cos φ вых)", fmt="0%")
    line(26, "Проверка", '=IF(NOT(ISNUMBER(C23)),IF(N(C17)>0,"нет в ряду",""),IF(N(C23)=0,"",IF(AND(C24<=C16+0.000001,C25<=C16+0.000001),"ОК","перегрузка")))', "",
         "загрузка по кВА и по кВт не выше допустимой", bold=True)

    head(28, "3. АККУМУЛЯТОРНАЯ БАТАРЕЯ")
    line(29, "Время автономии", None, "мин", "по заданию на проектирование (требуемое время работы от батареи)", inp=15, fmt="0")
    line(30, "Напряжение батареи U АКБ", None, "В", "по умолчанию 240 В (20 блоков 12 В)", inp=240, fmt="0")
    line(31, "КПД инвертора", None, "", "по умолчанию 0,95", inp=0.95, fmt="0.00")
    line(32, "Коэффициент разряда / старения k", None, "", "по умолчанию 0,8: глубина разряда и запас на износ батареи", inp=0.8, fmt="0.00")
    line(33, "Энергия, отдаваемая нагрузке за время автономии", '=IF(N(C31)>0,C11*C29/60/C31,0)', "кВт·ч", "E = P·t / (60·КПД)", fmt="0.00")
    line(34, "Требуемая ёмкость АКБ", '=IF(AND(N(C30)>0,N(C31)>0,N(C32)>0),C11*1000*(C29/60)/(C30*C31*C32),0)', "А·ч",
         "Ач = P[кВт]·1000 · t[мин]/60 / (U[В] · КПД · k)", fmt="0.0", bold=True)
    line(35, "Энергоёмкость батареи (справочно)", '=C34*C30/1000', "кВт·ч", "Ач · U / 1000", fmt="0.00")

    head(37, "4. ВХОД ИБП → НАГРУЗКА НА ЩГП")
    line(38, "Заряд АКБ", None, "", "по умолчанию 10 % от мощности на входе (добавка на заряд батареи)", inp=0.1, fmt="0%")
    line(39, "cos φ входа ИБП", None, "", "по умолчанию 0,95 (выпрямитель без корректора); с активным корректором до 0,99", inp=0.95, fmt="0.00")
    line(40, "P входа ИБП = P / КПД × (1 + заряд)", '=IF(N(C31)>0,C11/C31*(1+N(C38)),0)', "кВт", "мощность из сети при нормальной работе и заряде батареи", fmt="0.00", bold=True)
    line(41, "S входа ИБП", '=IF(N(C39)>0,C40/C39,0)', "кВА", "", fmt="0.00")
    line(42, "I входа ИБП", '=C41/(SQRT(3)*Uн_лин)', "А", "трёхфазный вход, симметрично", fmt="0.0")
    line(43, "ИБП / ЩГП питается от", '=Режимы!$C$6', "", "выбирается на «Питание_щита» C67:E67")
    line(44, "Учтено на листе «Режимы» в столбце", '=IF($C$5="есть",$C$43,"")', "", "вход ИБП добавляется к постоянной нагрузке этого ввода")
    line(45, "P выхода ИБП (в «Режимах» только справочно, в «Итого» не входит)", '=C11', "кВт", "двойного счёта нет: в «Итого» идёт вход ИБП, а выход показан отдельно", fmt="0.00")
    line(46, "Нагрузка на ЩГП, передаётся в «Режимы»", '=IF($C$5="есть",C40,0)', "кВт", "равна P входа ИБП; в «Режимах» — постоянная нагрузка (Кс = 1) выбранного ввода", fmt="0.00", bold=True)
    head(48, "5. РЯД ТИПОРАЗМЕРОВ ИБП (можно править; порядок любой)")
    ws.Range("B49").Value = "Стандартный ряд мощностей ИБП, кВА"
    ws.Range("B49").Font.Size = 9
    ws.Range("B49").Font.Color = C_GREY
    # условное форматирование: одно правило на диапазон
    rng = ws.Range("C26")
    if rng.FormatConditions.Count == 0:
        fc = rng.FormatConditions.Add(2, None, '=$C$26="перегрузка"')
        fc.Interior.Color = C_BAD_FILL
        fc.Font.Color = C_BAD_FONT
        fc2 = rng.FormatConditions.Add(2, None, '=$C$26="ОК"')
        fc2.Interior.Color = C_OK_FILL
        fc2.Font.Color = C_OK_FONT
    box(ws.Range("B4:E6"))
    ws.Range("B4:E6").Borders(12).LineStyle = 1
    for r in (9, 10, 11, 12, 13, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 29, 30, 31, 32, 33, 34, 35, 38, 39, 40, 41, 42, 43, 44, 45, 46):
        box(ws.Range("B%d:D%d" % (r, r)))
    ws.Range("B2").WrapText = False
    return ws


# ------------------------------------------------------------------------------------------------
# 7. «Питание_щита» J44:S47 (для вышестоящего)
# ------------------------------------------------------------------------------------------------
def step_upstream(wb, info):
    ws = sheet(wb, "Питание_щита")
    svc = info["svc"]
    heads = ["Руст I особая, кВт", "Руст I, кВт", "Руст II, кВт", "Руст III, кВт",
             "Рр I особая, кВт", "Рр I, кВт", "Рр II, кВт", "Рр III, кВт", "Руст пожарных, кВт", "Рр пожарных, кВт"]
    cols = "JKLMNOPQRS"
    ws.Range("J43").Value = "v3.12: категории надёжности и пожарные — для сборки в ГРЩ (Копировать J:S → «Нижестоящие_щиты» → Вставить связь)"
    ws.Range("J43").Font.Name = "Arial"
    ws.Range("J43").Font.Size = 9
    ws.Range("J43").Font.Bold = True
    for i, h in enumerate(heads):
        c = ws.Range("%s44" % cols[i])
        c.Value = h
        c.Font.Name = "Arial"
        c.Font.Size = 10
        ws.Columns(cols[i]).ColumnWidth = 13
    style_head(ws.Range("J44:S44"))
    for k in range(3):
        r = 45 + k
        d = svc["catk"][k]
        for i, j in enumerate(CAT):
            ws.Range("%s%d" % (cols[i], r)).Formula = '=IF($B%d="","",IFERROR(Режимы!$C$%d*1,""))' % (r, d["nf"][j])
            ws.Range("%s%d" % (cols[4 + i], r)).Formula = '=IF($B%d="","",IFERROR(Режимы!$I$%d*1,""))' % (r, d["nf"][j])
        ws.Range("R%d" % r).Formula = '=IF($B%d="","",IFERROR(Режимы!$C$%d*1,""))' % (r, d["f_all"])
        ws.Range("S%d" % r).Formula = '=IF($B%d="","",IFERROR(Режимы!$I$%d*1,""))' % (r, d["f_all"])
    rng = ws.Range("J45:S47")
    rng.Font.Name = "Arial"
    rng.Font.Size = 10
    rng.Font.Bold = True
    rng.HorizontalAlignment = -4108
    setnf(rng, "0.00")
    box(ws.Range("J44:S47"))


# ------------------------------------------------------------------------------------------------
# 8. «Краткое руководство», порядок и скрытие листов
# ------------------------------------------------------------------------------------------------
def step_guide(wb):
    ws = sheet(wb, "Краткое руководство")
    last_e = 0
    for r in range(2, 80):
        v = ws.Range("E%d" % r).Value
        if v is not None and str(v).strip():
            last_e = r
            if str(v).strip().startswith("v3.12"):
                last_e = -r
                break
    if last_e < 0:
        r = -last_e
    else:
        r = last_e + 1
        ws.Range("E%d" % last_e).Copy(ws.Range("E%d" % r))
        ws.Range("F%d" % last_e).Copy(ws.Range("F%d" % r))
    ws.Range("E%d" % r).Value = "v3.12 (категории, режимы, ИБП)"
    ws.Range("F%d" % r).Value = ("Категории надёжности линий, шины и схема щита, проверка «ОК / НЕ ОБЕСПЕЧЕНА»; пожарные нагрузки. "
                                 "Новые листы «Режимы» (нормальный, аварийный, потеря вводов, пожар, по категориям) и «ИБП» (мощность, АКБ, вход на ЩГП). "
                                 "Служебные «Спецификация_авто» и «Диапазоны» скрыты. См. 4_Инструкции/Категории_режимы_ИБП.md.")
    ws.Rows(r).AutoFit()
    # правило 21
    last_c = 0
    found = False
    for rr in range(2, 80):
        v = ws.Range("C%d" % rr).Value
        if v is not None and str(v).strip():
            last_c = rr
            if str(v).startswith("КАТЕГОРИИ"):
                found = True
                last_c = rr
                break
    if found:
        rr = last_c
    else:
        rr = last_c + 1
        ws.Range("B%d" % last_c).Copy(ws.Range("B%d" % rr))
        ws.Range("C%d" % last_c).Copy(ws.Range("C%d" % rr))
        ws.Range("B%d" % rr).Value = int(float(str(ws.Range("B%d" % last_c).Value).replace(",", "."))) + 1
    ws.Rows(rr).RowHeight = 43
    ws.Range("C%d" % rr).Value = ("КАТЕГОРИИ, РЕЖИМЫ, ИБП (v3.12): «Исходные данные» BN — требуемая категория линии, BO — пожарная (да/нет); "
                                  "«Питание_щита» раздел 6 — вводы, АВР, ИБП от, «Шины щита»; результаты — листы «Режимы» и «ИБП».")
    # правило 5 (скрытые листы)
    for r0 in range(2, 40):
        v = ws.Range("C%d" % r0).Value
        if v is not None and str(v).startswith("Листы: синие"):
            ws.Range("C%d" % r0).Value = (
                "Листы: синие — справочники (Справочная, Каталог), зелёные — ввод (Исходные данные > Разбивка по щитам > Однолинейка), "
                "оранжевые — результаты, серые — служебные. «В Акад» — для старого способа (Data Link в AutoCAD): НЕ удалять, НЕ скрывать и НЕ переименовывать. "
                "С v3.12 скрыты «Диапазоны» и «Спецификация_авто» (стоят в конце книги; показать: правая кнопка на ярлыке листа → Показать): не удалять, не переименовывать.")
            break


def step_sheets_order(wb, keep_visible):
    # «Режимы» и «ИБП» уже после «Нагрузка_щитов»; служебные - в конец и скрыть
    for nm in ("Спецификация_авто", "Диапазоны"):
        ws = sheet(wb, nm)
        if ws is None:
            continue
        ws.Visible = -1
        ws.Move(None, wb.Sheets(wb.Sheets.Count))   # позиционно: Move() без аргументов создаёт НОВУЮ книгу!
    for nm in ("Спецификация_авто", "Диапазоны"):
        ws = sheet(wb, nm)
        if ws is not None and not keep_visible:
            ws.Visible = 0            # xlSheetHidden


# ------------------------------------------------------------------------------------------------
# проверка ошибок
# ------------------------------------------------------------------------------------------------
def count_errors(ws, addr):
    try:
        rng = ws.Range(addr)
        e = rng.SpecialCells(-4123, 16)      # xlCellTypeFormulas, xlErrors
        return e.Count, e.Address
    except Exception:
        return 0, ""


def run(path, out=None, keep_visible=False):
    path = os.path.abspath(path)
    out = os.path.abspath(out) if out else path
    with ExcelSession() as xl:
        wb = xl.Workbooks.Open(path, 0, False)               # UpdateLinks=0
        need = ["Краткое руководство", "Справочная", "Исходные данные", "Разбивка_по_щитам", "Однолинейка", "Питание_щита",
                NG, "Диапазоны", "Спецификация_авто"]
        miss = [n for n in need if sheet(wb, n) is None]
        names = set(n.Name for n in wb.Names)
        miss += [n for n in ("Кс_Р", "Кс_К", "Ко_сил", "Uн_лин", "ПЩ_щиты") if n not in names]
        if miss:
            wb.Close(False)
            raise SystemExit("Книга не похожа на однолинейку v3.11: нет %s" % ", ".join(miss))
        global DEC
        DEC = xl.International[2]
        sep = xl.International[4]
        calc0 = xl.Calculation
        try:
            xl.ScreenUpdating = False
            xl.Calculation = -4135                              # ручной пересчёт на время правок
            log("справочная...")
            step_reference(wb, sep)
            log("питание щита: раздел 6...")
            step_power_sections(wb, sep)
            log("исходные данные BN:CA...")
            step_source_data(wb, sep)
            log("нагрузка щитов AR:BA...")
            step_loads(wb, sep)
            ensure_new_sheets(wb)
            log("лист Режимы...")
            info = step_modes(wb, sep)
            log("лист ИБП...")
            step_ibp(wb, sep, info)
            log("питание щита J44:S47...")
            step_upstream(wb, info)
            step_guide(wb)
            step_sheets_order(wb, keep_visible)
            xl.Calculation = -4105                              # автоматический
            xl.CalculateFull()
            bad = 0
            for sh, addr in (("Исходные данные", "BN3:CA%d" % LAST), (NG, "AR3:BA%d" % LAST), ("Режимы", "B1:L200"),
                             ("ИБП", "B1:N60"), ("Питание_щита", "B60:S90"), ("Питание_щита", "J44:S47"),
                             ("Справочная", "BD3:BF20")):
                n, a = count_errors(sheet(wb, sh), addr)
                if n:
                    bad += n
                    log("ОШИБКИ в %s!%s: %d (%s)" % (sh, addr, n, a[:120]))
            log("ошибок формул в новых областях: %d" % bad)
            # порядок: активный лист - первый видимый
            try:
                wb.Worksheets(1).Activate()
            except Exception:
                pass
            try:
                xl.ScreenUpdating = True
            except Exception:
                pass
            if out == path:
                wb.Save()
            else:
                wb.SaveAs(out, 51)
            log("сохранено:", out)
        finally:
            try:
                xl.Calculation = calc0 if calc0 in (-4105, -4135, 2) else -4105
            except Exception:
                pass
            wb.Close(False)
    return out


if __name__ == "__main__":
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("book")
    ap.add_argument("--out")
    ap.add_argument("--keep-visible", action="store_true", help="не скрывать служебные листы")
    a = ap.parse_args()
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass
    run(a.book, a.out, a.keep_visible)
