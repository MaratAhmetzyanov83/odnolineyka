from pathlib import Path
import json,win32com.client as w,time,sys
sys.stdout.reconfigure(encoding='utf-8',line_buffering=True)
p=Path(__file__).parent;root=p.parents[1]
book=json.loads((root/'_Разработка/2026-10-09_incoming/summary_source.json').read_text(encoding='utf-8'))['path'].replace('.01.xlsx','.02.xlsx')
panel="ЩC-ВК.02"  # Read from the user's open workbook by table_fixture.py.
def quote(s):return '"'+str(s).replace('\\','\\\\').replace('"','\\"')+'"'
code='''(vl-load-com)
(if (= (getvar "DWGNAME") "Однолинейка_проверка.dwg") (progn
(load %s)
(setq rpf (open %s "w") rpstage "init")
(vlax-ldata-put "SX_Schema" "xls" %s)
(vlax-ldata-put "SX_Schema" "panel" %s)
(vlax-ldata-put "SX_Schema" "sheet" "Однолинейка")
(vlax-ldata-put "SX_Schema" "pt" '(0.0 0.0 0.0))
(vlax-ldata-put "SX_Schema" "lib" "C:/missing.dwg")
(setq rpresult (vl-catch-all-apply 'sx:main (list T)))
(prin1 (list rpstage rptag (if (vl-catch-all-error-p rpresult) (vl-catch-all-error-message rpresult)) sx:errs sx:logl) rpf) (write-line "" rpf)
(setq rpcount 0 rptagged 0)
(vlax-for rpob sx:ms
  (if (and (= (vla-get-objectname rpob) "AcDbBlockReference") (= (vla-get-effectivename rpob) "SX_TCOL"))
    (progn
      (setq rpcount (1+ rpcount))
      (if (sx:xd-get (vlax-vla-object->ename rpob)) (setq rptagged (1+ rptagged))))))
(prin1 (list 'version sx:ver 'columns rpcount 'tagged rptagged 'invalid-ename-protected (vl-catch-all-error-p (sx:xd-set nil "TEST"))) rpf)
(vla-save sx:doc)
(close rpf))) (princ)
'''%(quote((root/'2_AutoCAD/Однолинейка/SX_Schema.lsp').as_posix()),quote((p/'reproduce.txt').as_posix()),quote(book),quote(panel))
(p/'reproduce.lsp').write_text(code,encoding='utf-8')
for attempt in range(30):
    try:
        a=w.GetActiveObject('AutoCAD.Application');ds=a.Documents
        d=next(ds.Item(i) for i in range(ds.Count) if ds.Item(i).Name=='Однолинейка_проверка.dwg')
        d.Activate();time.sleep(.3);d=a.ActiveDocument;assert d.Name=='Однолинейка_проверка.dwg'
        d.SendCommand('(load "'+(p/'reproduce.lsp').as_posix()+'")\n');print('Full reproduction requested');break
    except Exception as exc:last=exc;time.sleep(.4)
else:raise last
