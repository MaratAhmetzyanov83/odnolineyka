import os, sys
import win32com.client

p = os.path.join(os.path.dirname(os.path.abspath(__file__)), "xl", "Однолинейка_ЩГП.xlsx")
xl = win32com.client.DispatchEx("Excel.Application")
xl.Visible = False
xl.DisplayAlerts = False
xl.AskToUpdateLinks = False
xl.EnableEvents = False
try:
    wb = xl.Workbooks.Open(p, UpdateLinks=0, ReadOnly=False)
    ws = wb.Worksheets("Спецификация")
    c, h = ws.Range("C10"), ws.Range("H10")
    print("old C10:", c.Formula)
    print("old H10:", h.Formula)
    if c.HasFormula or h.HasFormula:
        print("FORMULA - stop")
        wb.Close(False)
        sys.exit(1)
    c.Value = "Выключатель автоматический NXB-63H 3P C63А 10 кА («На ИБП» QF2)"
    h.Value = 1
    wb.Save()
    print("new C10:", ws.Range("C10").Value, "| H10:", ws.Range("H10").Value)
    wb.Close(False)
finally:
    xl.Quit()
