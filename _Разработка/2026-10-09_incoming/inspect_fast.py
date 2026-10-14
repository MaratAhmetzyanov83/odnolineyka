import json,pathlib,sys,win32com.client as w
sys.stdout.reconfigure(encoding='utf-8')
p=pathlib.Path(__file__).parent
x=w.GetActiveObject('Excel.Application')
for b in x.Workbooks:
    for s in b.Worksheets:
        if s.Name=='Однолинейка':
            vals=s.Range('A1:BY1000').Value
            (p/'all_rows.json').write_text(json.dumps({'path':b.FullName,'values':vals},ensure_ascii=False,default=str),encoding='utf-8')
            for i,r in enumerate(vals):
                if i<2 or any(r[j] not in (None,'',0) for j in [3,4,8,9,10,11,15,17,24]):
                    print(i+1,{j+1:r[j] for j in range(77) if r[j] not in (None,'',0)})
import time
def retry(f):
    for k in range(20):
        try: return f()
        except Exception:
            time.sleep(0.1)
    return f()
a=w.GetActiveObject('AutoCAD.Application');d=a.ActiveDocument;es=[]
for e in d.ModelSpace:
    r={'h':e.Handle,'t':e.ObjectName,'l':e.Layer}
    if r['t']=='AcDbLine': r.update(a=e.StartPoint,b=e.EndPoint)
    elif r['t']=='AcDbBlockReference':
        r.update(p=e.InsertionPoint,n=e.Name,en=e.EffectiveName,rot=e.Rotation,s=e.XScaleFactor)
        if e.HasAttributes: r['atts']=retry(lambda:[(at.TagString,at.TextString,at.InsertionPoint) for at in e.GetAttributes()])
        if e.IsDynamicBlock: r['dyn']=[(pr.PropertyName,pr.Value) for pr in e.GetDynamicBlockProperties()]
    elif r['t'] in ('AcDbText','AcDbMText'): r.update(p=e.InsertionPoint,txt=e.TextString)
    es.append(r)
(p/'fast_entities.json').write_text(json.dumps(es,ensure_ascii=False,default=str),encoding='utf-8')
print('Entities',len(es))
