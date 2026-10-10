import pathlib,json,sys,win32com.client as w
sys.stdout.reconfigure(encoding='utf-8')
p=pathlib.Path(__file__).parent
a=w.GetActiveObject('AutoCAD.Application'); d=next(q for q in a.Documents if q.Name=='Ввод_проверка.dwg')
print('Drawing',d.Name,'WindowState',a.WindowState,'size',a.Width,a.Height)
r={'drawing':d.Name,'tables':[],'outlets':[],'definition':[]}
for e in d.ModelSpace:
    if e.ObjectName=='AcDbTable':
        r['tables'].append({'handle':e.Handle,'layer':e.Layer,'point':e.InsertionPoint,'width':e.Width,'height':e.Height,'rows':e.Rows,'cols':e.Columns,'texts':[[e.GetText(i,j) for j in range(e.Columns)] for i in range(e.Rows)]})
    if e.ObjectName=='AcDbBlockReference' and getattr(e,'EffectiveName',e.Name) in ['Вывод','SX_U_OUT']:
        col=e.TrueColor
        r['outlets'].append({'handle':e.Handle,'point':e.InsertionPoint,'layer':e.Layer,'color':e.Color,'rgb':[col.Red,col.Green,col.Blue],'name':e.Name})
for e in d.Blocks.Item('Вывод'):
    r['definition'].append({'type':e.ObjectName,'color':e.Color,'layer':e.Layer})
(p/'result_details.json').write_text(json.dumps(r,ensure_ascii=False,indent=2),encoding='utf-8')
print(json.dumps(r,ensure_ascii=False,indent=2))
