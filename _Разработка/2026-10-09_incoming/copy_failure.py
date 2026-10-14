from pathlib import Path
import win32com.client as w,pythoncom,sys,time
from win32com.client import VARIANT
sys.stdout.reconfigure(encoding='utf-8',line_buffering=True)
p=Path(__file__).parents[1]/'2026-10-09_singleline-repair';p.mkdir(exist_ok=True)
a=w.GetActiveObject('AutoCAD.Application');d=a.ActiveDocument
assert d.Name=='Однолинейка ЩС-ВК.02.dwg'
target=p/'Однолинейка_проверка.dwg'
ss=d.SelectionSets.Add('SX_REPAIR_SNAPSHOT')
try:
    ss.Select(5,pythoncom.Missing,pythoncom.Missing,VARIANT(pythoncom.VT_ARRAY|pythoncom.VT_I2,[410]),VARIANT(pythoncom.VT_ARRAY|pythoncom.VT_VARIANT,['Model']))
    d.WBlock(str(target),ss)
finally: ss.Delete()
print('Unsaved model copied to',target)
a.Documents.Open(str(target));print('Test drawing opened')
