from pathlib import Path
import json,sys
import win32com.client as w
sys.stdout.reconfigure(encoding='utf-8',line_buffering=True)
s=json.loads((Path.home()/'AppData/Roaming/CableJournal/settings.json').read_text(encoding='utf-8'))
x=w.DispatchEx('Excel.Application');x.DisplayAlerts=False;x.EnableEvents=False;x.AutomationSecurity=3
b=None
try:
 b=x.Workbooks.Open(str(Path(s['source'])),UpdateLinks=0,ReadOnly=True,Password='',WriteResPassword='')
 x.CalculateFull();sh=b.Worksheets.Item('Кабельный_журнал')
 for i in range(22,39):
  print(i,sh.Range(f'P{i}:AB{i}').Value2)
  if 24<=i<=32: print('FORMULAS',sh.Range(f'P{i}:AB{i}').Formula)
 print('MERGE',[(a,sh.Range(a).MergeArea.Address) for a in ['P25','P26','S26','P27','Q26','R26']])
finally:
 if b:b.Close(False)
 x.Quit()
