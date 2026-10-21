# КЖ книги ЩГП: строк данных не хватает — вставить строки перед «Итого» и протянуть формулы последней строки.
import os
import win32com.client

p = os.path.join(os.path.dirname(os.path.abspath(__file__)), "xl", "Однолинейка_ЩГП.xlsx")
ADD = 6
xl = win32com.client.DispatchEx("Excel.Application")
xl.Visible = False
xl.DisplayAlerts = False
xl.AskToUpdateLinks = False
xl.EnableEvents = False
try:
    wb = xl.Workbooks.Open(p, UpdateLinks=0, ReadOnly=False)
    kj = wb.Worksheets("Кабельный_журнал")
    tot = next(i for i in range(5, 200) if str(kj.Cells(i, 16).Formula).startswith("Итого по кабельному"))
    last = tot - 1
    print("последняя строка данных", last, "| итог", tot)
    kj.Rows("%d:%d" % (tot, tot + ADD - 1)).Insert()
    kj.Rows(last).Copy(kj.Rows("%d:%d" % (last + 1, last + ADD)))
    xl.CutCopyMode = False
    xl.CalculateFull()
    for i in range(5, last + ADD + 1):
        print(i, kj.Cells(i, 16).Text, "|", kj.Cells(i, 18).Text[:40], "|", kj.Cells(i, 19).Text, "| L", kj.Cells(i, 23).Text)
    i = last + ADD + 1
    while i < last + ADD + 12:
        t = [kj.Cells(i, c).Text for c in (16, 17, 21, 23)]
        if any(t):
            print("сводка", i, t)
        i += 1
    wb.Save()
    wb.Close(False)
finally:
    xl.Quit()
