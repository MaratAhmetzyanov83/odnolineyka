import pathlib,json,time,sys,win32com.client as w,pythoncom,shutil
from win32com.client import VARIANT
sys.stdout.reconfigure(encoding='utf-8')
p=pathlib.Path(__file__).parent;root=p.parents[1]
path=root/'База блоков компоновки (1).dwg'
a=w.GetActiveObject('AutoCAD.Application')
db=a.GetInterfaceObject('ObjectDBX.AxDbDocument.25');db.Open(str(path))
items=[];objs=[]
for b in db.Blocks:
    if not b.Name.startswith('*') and not b.IsXRef and not b.IsLayout:
        items.append({'name':b.Name,'count':b.Count})
        objs.append(b)
(p/'actual_library_inventory.json').write_text(json.dumps(items,ensure_ascii=False,indent=2),encoding='utf-8')
print(json.dumps(items,ensure_ascii=False),flush=True)
shutil.copy2(path,root/'2_AutoCAD/Блоки/SX_Компоновка_пользовательские.dwg')
a.Documents.Add();time.sleep(.5);dst=a.ActiveDocument
db.CopyObjects(VARIANT(pythoncom.VT_ARRAY|pythoncom.VT_DISPATCH,objs),dst.Blocks)
dst.SaveAs(str(p/'Компоновка_проверка_из_базы.dwg'));time.sleep(.5);dst=a.ActiveDocument
print('Preview',dst.Name,flush=True)
names=[i['name'] for i in items]
code='''(vl-load-com)
(setq ppdoc (vla-get-activedocument (vlax-get-acad-object)) ppms (vla-get-modelspace ppdoc) ppf (open "%s" "w") ppi 0)
(foreach name '%s
  (setq x (* (rem ppi 5) 1000.0) y (* (fix (/ ppi 5)) -1000.0))
  (setq o (vla-insertblock ppms (vlax-3d-point (list x y 0.0)) name 1.0 1.0 1.0 0.0))
  (vla-getboundingbox o 'ppmn 'ppmx)
  (prin1 (list name (vlax-safearray->list ppmn) (vlax-safearray->list ppmx) (if (= (vla-get-isdynamicblock o) :vlax-true) (mapcar '(lambda (p) (list (vla-get-propertyname p) (vlax-variant-value (vla-get-value p)) (if (vlax-property-available-p p 'AllowedValues) (vlax-variant-value (vla-get-allowedvalues p))))) (vlax-invoke o 'GetDynamicBlockProperties)))) ppf) (write-line "" ppf)
  (vla-addtext ppms name (vlax-3d-point (list x (- y 130.0) 0.0)) 14.0)
  (setq ppi (1+ ppi)))
(close ppf) (vla-regen ppdoc 1) (princ)
'''%(str(p/'actual_library_metrics.txt').replace('\\','/'),'('+' '.join('"'+n.replace('"','\\"')+'"' for n in names)+')')
(p/'actual_library_atlas.lsp').write_text(code,encoding='utf-8')
assert dst.Name=='Компоновка_проверка_из_базы.dwg'
dst.SendCommand('(load "'+str(p/'actual_library_atlas.lsp').replace('\\','/')+'")\n')
print('Actual library atlas requested')
