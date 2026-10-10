import pathlib,json,time,sys,shutil,win32com.client as w,pythoncom
from win32com.client import VARIANT
sys.stdout.reconfigure(encoding='utf-8',line_buffering=True)
p=pathlib.Path(__file__).parent;root=p.parents[1]
a=w.GetActiveObject('AutoCAD.Application')
src=a.GetInterfaceObject('ObjectDBX.AxDbDocument.25');src.Open(str(root/'База блоков компоновки (1).dwg'))
print('Source opened')
shutil.copy2(p/'Компоновка_проверка.dwg',p/'cp_test_base.dwg')
dest=a.GetInterfaceObject('ObjectDBX.AxDbDocument.25');dest.Open(str(p/'cp_test_base.dwg'))
print('Destination is separate DBX test database')
names=['DIN_TH35_L500','EKF_CM_2x7','DKC_CQEN_1600x600x600_FRAME','AV_2P_36','AV_4P_72','DA_4P_135','SX_DYN_MCB']
objs=[src.Blocks.Item(n) for n in names]
src.CopyObjects(VARIANT(pythoncom.VT_ARRAY|pythoncom.VT_DISPATCH,objs),dest.Blocks)
print('Selected definitions copied')
target=p/'Компоновка_проверка_из_базы.dwg'
dest.SaveAs(str(target));print('Test database saved')
a.Documents.Open(str(target));time.sleep(.5)
d=a.ActiveDocument;assert d.Name==target.name
print('Opened',d.Name)
names=['S201','S203','DS201','VPR-2.5-3L-GY']+names
code='''(vl-load-com)
(if (= (getvar "DWGNAME") "Компоновка_проверка_из_базы.dwg")
(progn
(setq ppdoc (vla-get-activedocument (vlax-get-acad-object)) ppms (vla-get-modelspace ppdoc) ppf (open "%s" "w") ppi 0)
(foreach name '%s
  (setq x (* (rem ppi 5) 2000.0) y (* (fix (/ ppi 5)) -2000.0))
  (setq o (vla-insertblock ppms (vlax-3d-point (list x y 0.0)) name 1.0 1.0 1.0 0.0))
  (vla-getboundingbox o 'ppmn 'ppmx)
  (prin1 (list name (vlax-safearray->list ppmn) (vlax-safearray->list ppmx) (if (= (vla-get-isdynamicblock o) :vlax-true) (mapcar '(lambda (p) (list (vla-get-propertyname p) (vlax-variant-value (vla-get-value p)) (if (vlax-property-available-p p 'AllowedValues) (vlax-safearray->list (vlax-variant-value (vla-get-allowedvalues p)))))) (vlax-invoke o 'GetDynamicBlockProperties)))) ppf) (write-line "" ppf)
  (vla-addtext ppms name (vlax-3d-point (list x (- y 150.0) 0.0)) 14.0)
  (setq ppi (1+ ppi)))
(close ppf) (vla-regen ppdoc 1))) (princ)
'''%(str(p/'actual_library_metrics.txt').replace('\\','/'),'('+' '.join('"'+n+'"' for n in names)+')')
(p/'selected_atlas.lsp').write_text(code,encoding='utf-8')
d.SendCommand('(load "'+str(p/'selected_atlas.lsp').replace('\\','/')+'")\n')
print('Selected atlas requested')
