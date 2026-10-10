# -*- coding: utf-8 -*-
"""«Разбивка_по_щитам» как лист «Однолинейка» v3: столбцы схемы — между G «Начало линии» и H «Номер линии»
(слева направо = сверху вниз по схеме: шина → автомат → аппараты → модуль и канал → клемма → линия).

Столбцы переставляются вырезанием (Excel сам правит ссылки A1 во всех формулах, условном форматировании,
проверке данных). Формулы 1.05, которые берут данные «Разбивки» по НОМЕРУ столбца
(VLOOKUP(…;Разбивка_нумерация_5;k;…) в «Нагрузке щитов» и «Кабельном журнале»), исправляются здесь.
"""
import re

RZ = "Разбивка_по_щитам"
ORDER = ["Место на схеме", "Лист схемы", "Шина", "Группа", "Автомат", "Дифзащита, мА", "QF", "Автомат (итог)",
         "Аппараты", "Модуль", "Модуль (итог)", "Модель", "Канал", "Канал (итог)", "Клемма", "Клемма (итог)",
         "Рисовать"]
LINE_HEAD = "Номер линии"
# имя диапазона -> номер столбца H в нём (с какого номера индексы сдвигаются)
SHIFT_FROM = {"Разбивка_нумерация_5": 4, "Разбивка_нумерация_4": 5, "Разбивка_нумерация_2": 7,
              "Разбивка_нумерация_1": 8}
WIDTHS = {"Место на схеме": 6, "Лист схемы": 5, "Шина": 7, "Группа": 6, "Автомат": 11, "Дифзащита, мА": 9,
          "QF": 7, "Автомат (итог)": 15, "Аппараты": 18, "Модуль": 11, "Модуль (итог)": 7, "Модель": 11,
          "Канал": 6, "Канал (итог)": 6, "Клемма": 7, "Клемма (итог)": 7, "Рисовать": 7}


def headers(ws):
    return [("" if v is None else str(v)) for v in ws.Range("A2:EZ2").Value[0]]


def col_of(ws, name):
    h = headers(ws)
    return h.index(name) + 1      # точное совпадение (с учётом регистра: «номер линии» в N — другой)


def patch(wb, log=print):
    rz = wb.Worksheets(RZ)
    h = headers(rz)
    if h.index(ORDER[0]) < h.index(LINE_HEAD):
        log("порядок «как в Однолинейке» уже есть — пропуск")
        return
    rz.Activate()
    win = wb.Application.ActiveWindow
    win.FreezePanes = False
    n = 0
    for name in ORDER:
        c = col_of(rz, name)
        target = col_of(rz, LINE_HEAD)
        rz.Columns(c).Cut()
        rz.Columns(target).Insert(-4161)      # xlToRight: вставить вырезанное перед «Номером линии»
        n += 1
    wb.Application.CutCopyMode = False
    # индексы VLOOKUP по «Разбивке» (столбцы от H и правее сдвинулись на n)
    pat = re.compile(r"(Разбивка_нумерация_\d)\s*,\s*(\d+)")
    fixed = 0
    for ws in wb.Worksheets:
        if ws.Name == RZ:
            continue
        ur = ws.UsedRange
        vals = ur.Formula
        if not isinstance(vals, tuple):
            continue
        r0, c0 = ur.Row, ur.Column
        for i, row in enumerate(vals):
            for j, f in enumerate(row):
                if isinstance(f, str) and "Разбивка_нумерация_" in f:
                    def sub(m):
                        k = int(m.group(2))
                        frm = SHIFT_FROM.get(m.group(1))
                        return f"{m.group(1)},{k + n if frm and k >= frm else k}"
                    g = pat.sub(sub, f)
                    if g != f:
                        ws.Cells(r0 + i, c0 + j).Formula = g
                        fixed += 1
    # ширины, закрепление: щит и начало линии (A:E свёрнуты)
    for name, w in WIDTHS.items():
        rz.Columns(col_of(rz, name)).ColumnWidth = w
    try:
        win.ScrollColumn = 1
        win.ScrollRow = 1
        rz.Cells(3, col_of(rz, "Начало линии") + 1).Select()
        win.FreezePanes = True
    except Exception:
        pass
    log(f"«Разбивка»: {n} столбцов схемы перед «Номером линии», индексов VLOOKUP исправлено: {fixed}")
