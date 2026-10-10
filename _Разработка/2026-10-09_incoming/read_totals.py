import json,pathlib,sys,win32com.client as w
sys.stdout.reconfigure(encoding='utf-8'); p=pathlib.Path(__file__).parent
x=w.GetActiveObject('Excel.Application');b=x.Workbooks.Item(1);s=b.Worksheets.Item('Нагрузка_щитов')
r={'path':b.FullName,'sheet':s.Name,'range':'R1:AG32','values':s.Range('R1:AG32').Value,'formulas':s.Range('R1:AG32').Formula}
(p/'summary_source.json').write_text(json.dumps(r,ensure_ascii=False,default=str),encoding='utf-8')
for i,row in enumerate(r['values']): print(i+1,row)
print('FORMULAS',json.dumps(r['formulas'][3:10],ensure_ascii=False))
a=w.GetActiveObject('AutoCAD.Application'); d=a.ActiveDocument
print('DOC',d.Name)
orig=next(q for q in a.Documents if q.Name.endswith('.001.dwg'))
t=next(e for e in orig.ModelSpace if e.ObjectName=='AcDbTable' and abs(e.InsertionPoint[0]-62.5644)<0.1 and abs(e.InsertionPoint[1]-1340.62)<0.1)
print('TABLE',t.ObjectName,t.InsertionPoint,t.Rows,t.Columns,t.Width,t.Height,t.StyleName)
print('TEXT',[[t.GetText(i,j) for j in range(t.Columns)] for i in range(t.Rows)])
print('WIDTHS',[t.GetColumnWidth(j) for j in range(t.Columns)],'HEIGHTS',[t.GetRowHeight(i) for i in range(t.Rows)])
