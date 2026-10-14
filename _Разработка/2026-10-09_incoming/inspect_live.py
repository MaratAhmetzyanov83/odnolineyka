import json, pathlib, sys
import win32com.client as w
sys.stdout.reconfigure(encoding='utf-8')
out=pathlib.Path(__file__).parent
a=w.GetActiveObject('AutoCAD.Application'); d=a.ActiveDocument
entities=[]
for e in d.ModelSpace:
    r={'handle':e.Handle,'type':e.ObjectName,'layer':e.Layer}
    for k in ('InsertionPoint','StartPoint','EndPoint','Coordinates','TextString','Name','EffectiveName','Rotation','XScaleFactor','YScaleFactor','Linetype','LinetypeScale','Color'):
        try: r[k]=getattr(e,k)
        except Exception: pass
    if r['type']=='AcDbBlockReference':
        try: r['atts']=[{'tag':x.TagString,'text':x.TextString,'point':x.InsertionPoint} for x in e.GetAttributes()]
        except Exception: pass
        try: r['dyn']=[{'name':x.PropertyName,'value':x.Value} for x in e.GetDynamicBlockProperties()]
        except Exception: pass
    entities.append(r)
(out/'entities.json').write_text(json.dumps({'path':d.FullName,'entities':entities},ensure_ascii=False,indent=2),encoding='utf-8')
x=w.GetActiveObject('Excel.Application')
books=[]
for b in x.Workbooks:
    r={'path':b.FullName,'active':b.ActiveSheet.Name,'sheets':[s.Name for s in b.Worksheets]}
    for s in b.Worksheets:
        if s.Name=='Однолинейка':
            r['values']=s.Range('A1:BY35').Value
            r['formulas']=s.Range('A1:BY10').Formula
    books.append(r)
(out/'books.json').write_text(json.dumps(books,ensure_ascii=False,indent=2,default=str),encoding='utf-8')
print(json.dumps([{'path':b['path'],'sheets':b['sheets'],'firstrows':b.get('values',[])[:6]} for b in books],ensure_ascii=False))
