from pathlib import Path
import json, sys, time
import win32com.client as w

sys.stdout.reconfigure(encoding='utf-8', line_buffering=True)
ROOT = Path(__file__).resolve().parents[2]
OUT = Path(__file__).parent
SOURCE = Path(r'\\Server\DropBox\Работа server\4. Техдепартамент\4.1. Проекты\098 Ташкент-КТ\5 ПИР\5.4_Том 3 ХБК\02_ELD-D1-02-ЭОМ\КЖ\КЖ_ЩР-1.dwg')
import shutil
LOCAL = OUT / 'Образец_КЖ_ЩР-1.dwg'
if not LOCAL.exists(): shutil.copy2(SOURCE, LOCAL)
a = w.GetActiveObject('AutoCAD.Application')
db = a.GetInterfaceObject('ObjectDBX.AxDbDocument.25')
db.Open(str(LOCAL))
def entity(e):
    d = {'type': e.ObjectName, 'handle': e.Handle, 'layer': e.Layer}
    try:
        if e.ObjectName == 'AcDbTable':
            d.update(point=e.InsertionPoint, rows=e.Rows, columns=e.Columns,
                     widths=[e.GetColumnWidth(i) for i in range(e.Columns)],
                     heights=[e.GetRowHeight(i) for i in range(e.Rows)],
                     text=[[e.GetText(i,j) for j in range(e.Columns)] for i in range(e.Rows)],
                     direction=e.Direction, style=e.StyleName)
        elif e.ObjectName == 'AcDbViewport':
            d.update(center=e.Center, height=e.Height, width=e.Width,
                     scale=e.CustomScale, target=e.Target,
                     direction=e.Direction, locked=e.DisplayLocked)
        elif e.ObjectName == 'AcDbBlockReference':
            d.update(name=e.EffectiveName, point=e.InsertionPoint)
            if e.HasAttributes: d['attrs']=[(v.TagString,v.TextString) for v in e.GetAttributes()]
        elif e.ObjectName in ('AcDbText','AcDbMText'):
            d.update(text=e.TextString, point=e.InsertionPoint)
    except Exception as exc: d['error']=str(exc)
    return d
info = {'source': str(SOURCE), 'model': [entity(e) for e in db.ModelSpace],
        'layouts': [{'name': l.Name, 'order': l.TabOrder,
                     'entities': [entity(e) for e in l.Block]} for l in db.Layouts if not l.ModelType]}
(OUT/'template.json').write_text(json.dumps(info,ensure_ascii=False,indent=2),encoding='utf-8')
from collections import Counter
print('MODEL',Counter(e['type'] for e in info['model']))
for e in info['model']:
    if e['type']=='AcDbTable': print('TABLE',e['handle'],e['rows'],e['columns'],e['point'],e['widths'],e['heights'][:8],e['text'][:6])
for l in info['layouts']:
    print('LAYOUT',l['name'],Counter(e['type'] for e in l['entities']))
    for e in l['entities']:
        if e['type'] in ('AcDbViewport','AcDbBlockReference'): print(e)
