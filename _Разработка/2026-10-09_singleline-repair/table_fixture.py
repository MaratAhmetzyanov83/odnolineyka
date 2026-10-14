from pathlib import Path
import json,openpyxl,win32com.client as w,time,sys
sys.stdout.reconfigure(encoding='utf-8',line_buffering=True)
p=Path(__file__).parent;root=p.parents[1]
path=json.loads((root/'_Разработка/2026-10-09_incoming/summary_source.json').read_text(encoding='utf-8'))['path'].replace('.01.xlsx','.02.xlsx')
xl=w.GetActiveObject('Excel.Application')
wb=next(b for b in xl.Workbooks if b.FullName.lower()==path.lower())
rows=wb.Worksheets('Однолинейка').Range('A3:BY1000').Value2
panel=next(r[0] for r in rows if r[0])
data=[]
for idx,row in enumerate(rows):
    vals=[None if isinstance(v,int) and v < -2000000000 else v for v in row]
    if vals[0]==panel and isinstance(vals[1],(int,float)):
        i=int(vals[1])-1;data.append([i,20*i+10,vals,idx+3])
print('Fixture panel',panel,'rows',len(data))
def lsp(v):
    if v is None or v=='':return 'nil'
    if isinstance(v,str):return '"'+v.replace('\\','\\\\').replace('"','\\"')+'"'
    if isinstance(v,(int,float)):return str(v)
    return '('+' '.join(lsp(x) for x in v)+')'
code='''(vl-load-com)
(if (= (getvar "DWGNAME") "Однолинейка_проверка.dwg") (progn
(load %s)
(setq rpf (open %s "w") rpstage "init")
(defun *error* (msg) (prin1 (list rpstage msg) rpf) (close rpf) (princ))
(setq sx:doc (vla-get-activedocument (vlax-get-acad-object)) sx:ms (vla-get-modelspace sx:doc)
      sx:P0 '(0.0 0.0 0.0) sx:R '(1.0 0.0) sx:D '(0.0 1.0) sx:ang 0.0 sx:tagpanel "TEST"
      sx:part nil sx:bbcache nil sx:blkcache nil sx:cccache nil sx:bbctx "")
(setq rpdata '%s)
(sx:clear) (sx:layers) (sx:roles) (setq sx:style (sx:ensure-style)) (sx:make-blocks T)
(defun sx:ttry (key fn label / mark)
  (setq rpstage (strcat "begin " key) mark (entlast))
  (sx:try fn label)
  (setq rpstage (strcat "tag " key))
  (sx:tag-new mark key))
(setq rpstage "draw table") (sx:draw-table rpdata)
(setq rpcount 0)
(vlax-for rpob sx:ms
  (if (and (= (vla-get-objectname rpob) "AcDbBlockReference") (= (vla-get-effectivename rpob) "SX_TCOL")) (setq rpcount (1+ rpcount))))
(prin1 (list 'columns rpcount 'expected (length rpdata) 'errors sx:errs) rpf)
(close rpf))) (princ)
'''%(lsp((root/'2_AutoCAD/Однолинейка/SX_Schema.lsp').as_posix()),lsp((p/'table_fixture.txt').as_posix()),lsp(data))
(p/'table_fixture.lsp').write_text(code,encoding='utf-8')
for attempt in range(30):
    try:
        a=w.GetActiveObject('AutoCAD.Application');ds=a.Documents
        d=next(ds.Item(i) for i in range(ds.Count) if ds.Item(i).Name=='Однолинейка_проверка.dwg')
        d.Activate();time.sleep(.2);d=a.ActiveDocument;assert d.Name=='Однолинейка_проверка.dwg'
        d.SendCommand('(load "'+(p/'table_fixture.lsp').as_posix()+'")\n');print('Table fixture requested');break
    except Exception as exc:last=exc;time.sleep(.4)
else:raise last
