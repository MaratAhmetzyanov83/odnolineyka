# Длины кабелей ИБП-АКБ (шкаф АКБ рядом с ИБП) и пересчёт формул количества в спецификации (ячейки были текстовые).
import os
import win32com.client

p = os.path.join(os.path.dirname(os.path.abspath(__file__)), "xl", "Однолинейка_ЩГП.xlsx")
LEN = {"ИБП-АКБ.1": 3, "ИБП-АКБ.2": 3, "ИБП-АКБ.4": 3}
xl = win32com.client.DispatchEx("Excel.Application")
xl.Visible = False
xl.DisplayAlerts = False
xl.AskToUpdateLinks = False
xl.EnableEvents = False
try:
    wb = xl.Workbooks.Open(p, UpdateLinks=0, ReadOnly=False)
    ws = wb.Worksheets("Исходные данные")
    for i in range(3, 1001):
        n = ws.Cells(i, 4).Value
        if n in LEN:
            ws.Cells(i, 9).Value = LEN[n]
    sc = wb.Worksheets("Спецификация")
    print("защита листа:", sc.ProtectContents)
    # кол-во числом (формат ячеек текстовый, NumberFormat не меняется): 3 м × 1,2 запас КЖ → 4 м
    QTY = {"2.8": (12, "ИБП-АКБ.1: 3 жилы × 4 м (L=3 м + 20 %)"),
           "2.9": (4, "ИБП-АКБ.2: L=3 м + 20 %"),
           "2.10": (4, "ИБП-АКБ.4: L=3 м + 20 %; ИБП-EPO — добавить по плану")}
    for i in range(45, 70):
        k = str(sc.Cells(i, 2).Value or "")
        if k in QTY:
            sc.Cells(i, 8).Value = QTY[k][0]
            sc.Cells(i, 10).Value = QTY[k][1]
    xl.CalculateFull()
    for i in range(3, 1001):
        n = ws.Cells(i, 4).Value
        if n and str(n).startswith("ИБП"):
            print("исх", i, n, "| I", ws.Cells(i, 9).Text, "| M", ws.Cells(i, 13).Text, "| K", ws.Cells(i, 11).Text)
    kj = wb.Worksheets("Кабельный_журнал")
    for i in range(5, 20):
        if str(kj.Cells(i, 16).Text).startswith("ИБП"):
            print("КЖ", i, kj.Cells(i, 16).Text, "| L", kj.Cells(i, 23).Text)
    for i in range(45, 70):
        if str(sc.Cells(i, 2).Value or "") in ("2.8", "2.9", "2.10"):
            print("спец", sc.Cells(i, 2).Value, "=", sc.Cells(i, 8).Text)
    wb.Save()
    wb.Close(False)
finally:
    xl.Quit()
