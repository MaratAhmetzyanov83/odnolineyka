from pathlib import Path
import json,sys
import win32com.client as w
sys.stdout.reconfigure(encoding='utf-8',line_buffering=True)
sys.stderr.reconfigure(encoding='utf-8')
s=json.loads((Path.home()/'AppData/Roaming/CableJournal/settings.json').read_text(encoding='utf-8'))
x=w.DispatchEx('Excel.Application');x.DisplayAlerts=False;x.EnableEvents=False;x.AutomationSecurity=3
b=None
try:
 b=x.Workbooks.Open(str(Path(s['source'])),UpdateLinks=0,ReadOnly=True,Password='',WriteResPassword='')
 x.CalculateFull()
 sh=b.Worksheets.Item('Кабельный_журнал')
 for r in range(49,55):
  print('ROW',r,'values',sh.Range(f'P{r}:AB{r}').Value2,'formulas',sh.Range(f'P{r}:AB{r}').Formula)
 print('SRC rows containing line')
 line=sh.Range('P51').Value2
 for name in ['Исходные данные','Разбивка_по_щитам','Однолинейка']:
  t=b.Worksheets.Item(name);vals=t.Range('A1:CA1000').Value2
  for i,row in enumerate(vals,1):
   if line in row:print(name,i,[(j+1,v) for j,v in enumerate(row) if v not in (None,'',0)])
finally:
 if b:b.Close(False)
 x.Quit()
