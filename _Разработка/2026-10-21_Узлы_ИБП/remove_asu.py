# Убрать кабели ИБП-АСУ.1/.2 из книги ЩГП (связь с АСУ — не раздел ЭОМ).
import os
import win32com.client

p = os.path.join(os.path.dirname(os.path.abspath(__file__)), "xl", "Однолинейка_ЩГП.xlsx")
NAMES = ("ИБП-АСУ.1", "ИБП-АСУ.2")
xl = win32com.client.DispatchEx("Excel.Application")
xl.Visible = False
xl.DisplayAlerts = False
xl.AskToUpdateLinks = False
xl.EnableEvents = False
try:
    wb = xl.Workbooks.Open(p, UpdateLinks=0, ReadOnly=False)
    # «Исходные данные»: вернуть строки к шаблону (копия ближайшей пустой строки-шаблона ниже)
    ws = wb.Worksheets("Исходные данные")
    rows = [i for i in range(3, 1001) if ws.Cells(i, 4).Value in NAMES]
    tpl = next(i for i in range(max(rows) + 1, 1001) if ws.Cells(i, 4).Formula == "" and ws.Cells(i, 3).Formula == "")
    for i in rows:
        ws.Rows(tpl).Copy(ws.Rows(i))
        print("исходные: строка", i, "← шаблон", tpl)
    xl.CutCopyMode = False
    # «Разбивка»
    rz = wb.Worksheets("Разбивка_по_щитам")
    for i in range(3, 1001):
        if rz.Cells(i, 8).Value in NAMES:
            rz.Cells(i, 8).ClearContents()
            print("разбивка: H%d очищена" % i)
    # спецификация
    sc = wb.Worksheets("Спецификация")
    for i in range(80, 40, -1):
        if any(n in str(sc.Cells(i, 3).Value or "") for n in NAMES):
            print("спец: удалена", sc.Cells(i, 2).Value, sc.Cells(i, 3).Value)
            sc.Rows(i).Delete()
    # коды марок
    sp = wb.Worksheets("Справочная")
    for i in range(4, 52):
        if sp.Cells(i, 12).Value in ("кипэ", "utp5e"):
            print("код удалён", sp.Cells(i, 12).Value)
            sp.Range(sp.Cells(i, 12), sp.Cells(i, 14)).ClearContents()
    xl.CalculateFull()
    # КЖ: лишние пустые строки (#Н/Д) перед «Итого» — оставить одну, как было
    kj = wb.Worksheets("Кабельный_журнал")
    tot = next(i for i in range(5, 200) if str(kj.Cells(i, 16).Formula).startswith("Итого по кабельному"))
    empty = [i for i in range(5, tot) if kj.Cells(i, 16).Text in ("#Н/Д", "#N/A")]
    if len(empty) > 1:
        kj.Rows("%d:%d" % (empty[1], empty[-1])).Delete()
        print("КЖ: удалены пустые строки", empty[1:])
    xl.CalculateFull()
    tot = next(i for i in range(5, 200) if str(kj.Cells(i, 16).Formula).startswith("Итого по кабельному"))
    for i in range(5, tot):
        print("КЖ", i, kj.Cells(i, 16).Text, "|", kj.Cells(i, 19).Text)
    for i in range(45, 70):
        if str(sc.Cells(i, 2).Value or "").startswith("2."):
            print("спец", sc.Cells(i, 2).Value, sc.Cells(i, 3).Value, "=", sc.Cells(i, 8).Text)
    wb.Save()
    wb.Close(False)
finally:
    xl.Quit()
