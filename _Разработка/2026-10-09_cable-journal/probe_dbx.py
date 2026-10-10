from pathlib import Path
import sys, win32com.client as w, pythoncom
sys.stdout.reconfigure(encoding='utf-8',line_buffering=True)
p=Path(__file__).parent
a=w.GetActiveObject('AutoCAD.Application')
db=a.GetInterfaceObject('ObjectDBX.AxDbDocument.25'); db.Open(str((p/'Образец_КЖ_ЩР-1.dwg').resolve()))
for e in db.ModelSpace:
    if e.ObjectName=='AcDbTable':
        print('TABLE',e.Handle)
        for name in ['Rows','Columns','BreaksEnabled','BreakSpacing','TableBreakHeight','TableBreakFlowDirection','RepeatTopLabels','RepeatBottomLabels','HorzCellMargin','VertCellMargin']:
            try: print(name,getattr(e,name))
            except Exception as x:print(name,str(x))
        if e.Rows>3:
            tb=e
            print('STYLE',e.GetCellTextHeight(0,2),e.GetCellTextStyle(0,2),e.GetCellAlignment(0,2))
clones=db.CopyObjects(w.VARIANT(pythoncom.VT_ARRAY|pythoncom.VT_DISPATCH,(tb,)),db.ModelSpace)
print('RAW',repr(clones))
def flat(v):
    return [w.Dispatch(x) for x in v[0]]
clones=flat(clones)
print('CLONE',clones[0].Handle)
clones[0].Delete()
first=min((l for l in db.Layouts if not l.ModelType),key=lambda l:l.TabOrder)
layout=db.Layouts.Add('TEST'); layout.CopyFrom(first)
objs=[e for e in first.Block if e.ObjectName!='AcDbViewport' or e.Layer=='Defpoints']
clones=flat(db.CopyObjects(w.VARIANT(pythoncom.VT_ARRAY|pythoncom.VT_DISPATCH,tuple(objs)),layout.Block))
print('LAYOUT COPIED',len(clones))
for e in clones:
    if e.ObjectName=='AcDbViewport': e.Target=w.VARIANT(pythoncom.VT_ARRAY|pythoncom.VT_R8,(420.,0.,0.)); print('TARGET',e.Target)
for l in list(db.Layouts):
    if not l.ModelType and l.Name not in (first.Name,'TEST'):
        try:l.Delete()
        except Exception as x:
            print('REMOVE VIA DICTIONARY',l.Name)
            db.Dictionaries.Item('ACAD_LAYOUT').Remove(l.Name)
print('DELETED')
print('LAYOUTS',[(l.Name,l.TabOrder) for l in db.Layouts])
db.SaveAs(str((p/'DBX_проверка.dwg').resolve()))
print('SAVED')
