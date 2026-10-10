import pathlib,time,sys,win32com.client as w,pythoncom
from win32com.client import VARIANT
sys.stdout.reconfigure(encoding='utf-8')
p=pathlib.Path(__file__).parent;root=p.parents[1]
a=w.GetActiveObject('AutoCAD.Application')
src=next(q for q in a.Documents if q.Name.startswith('Однолинейка '))
names=['S201','S203','DS201','DS203','ABB S201','Дин','Шина N','VPR-2.5-3L-GY','VPR-6-GY','VPR-6-BU','D-VPR-6','БП DIN','ESB40','F204','3+1']
if '--resume' in sys.argv:
    dst=next(q for q in a.Documents if q.Name=='Компоновка_проверка.dwg');dst.Activate();time.sleep(.5);dst=a.ActiveDocument
else:
    objs=[src.Blocks.Item(n) for n in names]
    a.Documents.Add();time.sleep(.5);dst=a.ActiveDocument
    arr=VARIANT(pythoncom.VT_ARRAY|pythoncom.VT_DISPATCH,objs)
    src.CopyObjects(arr,dst.Blocks)
    library=root/'2_AutoCAD/Блоки/SX_Компоновка_пользовательские.dwg'
    dst.SaveAs(str(library));time.sleep(.5);dst=a.ActiveDocument
    dst.SaveAs(str(p/'Компоновка_проверка.dwg'));time.sleep(.5);dst=a.ActiveDocument
print('Library copied from',src.FullName,flush=True)
print('Preview',dst.FullName,flush=True)
code='''(vl-load-com)
(setq ppdoc (vla-get-activedocument (vlax-get-acad-object)) ppms (vla-get-modelspace ppdoc) ppf (open "%s" "w") ppi 0)
(foreach name '%s
  (setq x (* (rem ppi 5) 300.0) y (* (fix (/ ppi 5)) -180.0))
  (setq o (vla-insertblock ppms (vlax-3d-point (list x y 0.0)) name 1.0 1.0 1.0 0.0))
  (prin1 (list name (if (= (vla-get-isdynamicblock o) :vlax-true) (mapcar '(lambda (p) (list (vla-get-propertyname p) (vlax-variant-value (vla-get-value p)) (if (vlax-property-available-p p 'AllowedValues) (vlax-variant-value (vla-get-allowedvalues p))))) (vlax-invoke o 'GetDynamicBlockProperties)))) ppf) (write-line "" ppf)
  (vla-addtext ppms name (vlax-3d-point (list x (- y 90.0) 0.0)) 7.0)
  (setq ppi (1+ ppi)))
(close ppf) (vla-regen ppdoc 1) (princ)
'''%(str(p/'block_dynamics.txt').replace('\\','/'),'('+' '.join('"'+n+'"' for n in names)+')')
(p/'block_atlas.lsp').write_text(code,encoding='utf-8')
assert dst.Name=='Компоновка_проверка.dwg'
dst.SendCommand('(load "'+str(p/'block_atlas.lsp').replace('\\','/')+'")\n')
print('Atlas requested')
