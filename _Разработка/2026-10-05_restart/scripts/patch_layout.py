# -*- coding: utf-8 -*-
"""Оформление: блоки 1.07–1.08 со «Справочной» — на отдельный лист «Настройки» (вырезанием: имена и формулы
Excel поправит сам); в «Справочной» — подписи над таблицами 1.05. Запускается последним (build_all)."""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

TITLE_FILL = 0xF2E1C5     # светло-голубой (BGR)
HEAD_FILL = 0xE7E6E6      # серый
INPUT_FILL = 0xFFFFFF

# (откуда на «Справочной», куда на «Настройках» — левый верхний угол, заголовок блока, пояснение)
MOVES = [   # (откуда на «Справочной» — без старой строки заголовка, куда на «Настройках»)
    ("AL22:AM32", "A6"),
    ("AL3:AO20", "A22"),
    ("AQ3:AW150", "F4"),
    ("AX3:AZ40", "N4"),
    ("BB3:BF60", "R4"),
    ("BH3:BI20", "X4"),
]
OLD_TITLES = ("AL2", "AQ2", "AX2", "BB2", "BH2")
BLOCKS = [   # (ячейка заголовка, заголовок, пояснение строкой ниже, шапка, вся таблица)
    ("A3", "ОБЩИЕ НАСТРОЙКИ", "Меняйте значения в столбце B.", "A5:B5", "A5:B16"),
    ("A20", "ДОПУСТИМЫЙ ДЛИТЕЛЬНЫЙ ТОК, А (ПУЭ табл. 1.3.6, медь, в воздухе)",
     "По нему подбирается сечение. «—» — таких кабелей в таблице ПУЭ нет.", "A22:D22", "A22:D37"),
    ("F2", "КОДЫ ПОТРЕБИТЕЛЕЙ: что подставлять по коду (код = буквы номера линии до первой точки)",
     "Пусто — нет ограничения. Новый код — в первую пустую строку. Наименование, опуск, этаж — как раньше: «Справочная» P:T.",
     "F4:L4", "F4:L60"),
    ("N2", "УГО → БЛОК AutoCAD (SXDRAW)", "Код программы не менять.", "N4:P4", "N4:P11"),
    ("R2", "МОДУЛИ: каталог для схемы, спецификации и наклеек",
     "Каналов — сколько линий на модуль. Ширина — в модулях DIN (для наклеек).", "R4:V4", "R4:V20"),
    ("X2", "МОДЕЛЬ МОДУЛЯ ПО УМОЛЧАНИЮ", "Управление: из выноски (Relay, DIM, DALI) или по коду (шторы, клапаны).",
     "X4:Y4", "X4:Y11"),
]
WIDTHS = {"A": 52, "B": 12, "C": 14, "D": 22, "E": 3, "F": 10, "G": 20, "H": 13, "I": 13, "J": 13, "K": 13, "L": 13,
          "M": 3, "N": 20, "O": 24, "P": 18, "Q": 3, "R": 16, "S": 12, "T": 10, "U": 16, "V": 12, "W": 3, "X": 14, "Y": 16}
SP_LABELS = [("B1", "СЕЧЕНИЕ: x и ρ (для потерь)"), ("F1", "ТОК → АВТОМАТ"),
             ("I1", "АВТОМАТ → СЕЧЕНИЕ (1.05; с 1.06 не используется — см. «Настройки»)"),
             ("L1", "МАРКИ КАБЕЛЕЙ"), ("P1", "КОДЫ ПОТРЕБИТЕЛЕЙ: наименование, опуск, этаж/шлейф, пуск"),
             ("V1", "СПОСОБЫ ПРОКЛАДКИ"), ("Y1", "ЭТАЖИ"), ("AE1", "ПТО: артикулы кабелей")]


def layout(wb, log=print):
    app = wb.Application
    if any(ws.Name == "Настройки" for ws in wb.Worksheets):
        log("«Настройки» уже есть — пропуск")
        return
    sp = wb.Worksheets("Справочная")
    st = wb.Worksheets.Add(After=sp)
    st.Name = "Настройки"
    st.Tab.Color = sp.Tab.Color if sp.Tab.Color else 0xC0A060
    for src, dst in MOVES:
        sp.Range(src).Cut(st.Range(dst))
    for c in OLD_TITLES:
        sp.Range(c).ClearContents()
    st.Range("A1").Value = "НАСТРОЙКИ ТАБЛИЦЫ 1.08 — что таблица подставляет сама. Меняйте значения; столбцы и строки не вставляйте."
    st.Range("A1").Font.Bold = True
    st.Range("A1").Font.Size = 13
    for cell, title, hint, head, table in BLOCKS:
        c = st.Range(cell)
        c.Value = title
        c.Font.Bold = True
        c.Font.Color = 0x7F3F00
        hc = st.Cells(c.Row + 1, c.Column)
        hc.Value = hint
        hc.Font.Italic = True
        hc.Font.Color = 0x595959
        h = st.Range(head)
        h.Interior.Color = HEAD_FILL
        h.Font.Bold = True
        h.WrapText = True
        h.VerticalAlignment = -4108
        t = st.Range(table)
        t.Borders.LineStyle = 1
        t.Borders.Color = 0xBFBFBF
    # общие настройки: подпись + пояснение строкой ниже заголовка уже есть; значения — голубым
    st.Range("A5:B5").Value = (("Параметр", "Значение"),)
    st.Range("B6:B16").Interior.Color = 0xF7EBDD
    st.Range("B6:B16").HorizontalAlignment = -4108
    for col, w in WIDTHS.items():
        st.Columns(col).ColumnWidth = w
    st.Rows(4).RowHeight = 15
    st.Range("A22:D22").RowHeight = 30
    st.Range("F4:L4").RowHeight = 45
    # «Справочная»: подписи таблиц 1.05
    for cell, text in SP_LABELS:
        r = sp.Range(cell)
        r.Value = text
        r.Font.Bold = True
        r.Font.Color = 0x7F3F00
    sp.Range("I3:J20").Font.Color = 0xA6A6A6
    sp.Range("AL1").Value = "Настройки автоподбора и схемы (1.07–1.08) — на листе «Настройки»."
    sp.Range("AL1").Font.Italic = True
    log("Лист «Настройки»: блоки перенесены и оформлены; «Справочная»: подписи таблиц")


if __name__ == "__main__":
    import patch_106
    src, dst = sys.argv[1], sys.argv[2]
    import shutil
    shutil.copyfile(src, dst)
    app = patch_106.excel()
    try:
        wb = app.Workbooks.Open(os.path.abspath(dst), UpdateLinks=0)
        layout(wb)
        wb.Save()
        wb.Close(False)
    finally:
        patch_106.quit_excel(app)
