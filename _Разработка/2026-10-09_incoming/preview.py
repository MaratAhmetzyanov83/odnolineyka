import pathlib,json,time,sys,win32com.client as w
from win32com.client import VARIANT
import pythoncom
p=pathlib.Path(__file__).resolve().parent
a=w.GetActiveObject('AutoCAD.Application');d=a.ActiveDocument
fn=str(p/'Ввод_проверка.dwg')
if '--resume' not in sys.argv:
    ss=d.SelectionSets.Add('SX_IN_PREVIEW')
    try:
        ss.Select(5,pythoncom.Missing,pythoncom.Missing,VARIANT(pythoncom.VT_ARRAY|pythoncom.VT_I2,[410]),VARIANT(pythoncom.VT_ARRAY|pythoncom.VT_VARIANT,['Model']))
        d.WBlock(fn,ss)
    finally: ss.Delete()
    print('Snapshot',fn,flush=True)
    a.Documents.Open(fn)
nd=next(q for q in a.Documents if q.Name=='Ввод_проверка.dwg')
nd.Activate()
time.sleep(0.5)
nd=a.ActiveDocument
assert nd.Name=='Ввод_проверка.dwg'
def lsp(v):
    if v is None: return 'nil'
    if isinstance(v,str): return '"'+v.replace('\\','\\\\').replace('"','\\"')+'"'
    if isinstance(v,(int,float)): return str(v)
    return '('+' '.join(lsp(x) for x in v)+')'
vals=json.loads((p/'all_rows.json').read_text(encoding='utf-8'))['values'][2:]
recs=[]
for idx,row in enumerate(vals):
    if row[0]=='ЩР-1' and isinstance(row[1],(int,float)):
        i=int(row[1])-1;recs.append([i,i*20+10,[None if v=='' else v for v in row],idx+3])
script='''(vl-load-com)
(load "%s")
(setq sx:doc (vla-get-activedocument (vlax-get-acad-object)) sx:ms (vla-get-modelspace sx:doc)
      sx:P0 '(55.0 1345.99 0.0) sx:R '(1.0 0.0) sx:D '(0.0 1.0) sx:ang 0.0 sx:tagpanel (sx:ru '(1065 1056 45 49))
      sx:part nil sx:blkcache nil sx:bbcache nil sx:pecache nil sx:custom nil sx:models nil)
(setq sx:preview-recs '%s)
(setq sx:preview-draw (vl-remove-if-not '(lambda (dd) (sx:is1 dd sx:cDraw)) sx:preview-recs))
(sx:open-wb %s)
(setq sx:summary (sx:read-summary sx:tagpanel))
(setq sx:preview-empty-summary (sx:read-summary "missing-panel"))
(sx:xl-release)
(sx:layers) (sx:summary-capture) (sx:adopt-inputs sx:preview-draw)
(sx:clear)
(sx:layers) (sx:roles) (setq sx:style (sx:ensure-style)) (sx:make-blocks T)
(sx:draw-table sx:preview-recs)
(sx:draw-sheets sx:preview-draw sx:tagpanel)
(sx:draw-groups sx:preview-draw)
(sx:draw-modules sx:preview-draw)
(sx:draw-rows sx:preview-draw)
(sx:draw-summary)
(vla-regen sx:doc 1)
(setq sxf (open "%s" "w"))
(prin1 (list sx:ver sx:errs (sx:input-p (nth 5 sx:preview-recs)) (sx:u (sx:sheet-input sx:preview-draw 0)) sx:summary sx:preview-empty-summary) sxf)
(close sxf)
(princ)
'''%(str(p.parents[1]/'2_AutoCAD/Однолинейка/SX_Schema.lsp').replace('\\','/'),lsp(recs),lsp(w.GetActiveObject('Excel.Application').Workbooks.Item(1).FullName),str(p/'preview_result.txt').replace('\\','/'))
(p/'preview.lsp').write_text(script,encoding='utf-8')
nd.SendCommand('(load "'+str(p/'preview.lsp').replace('\\','/')+'")\n')
print('Preview requested',flush=True)
