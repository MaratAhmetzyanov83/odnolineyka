from pathlib import Path
import win32com.client as w,time
p=Path(__file__).parent;root=p.parents[1]
code='''(vl-load-com)
(if (= (getvar "DWGNAME") "Однолинейка_проверка.dwg") (progn
(load "%s")
(setq rpf (open "%s" "w") rpstage "init")
(defun *error* (msg) (prin1 (list rpstage msg) rpf) (close rpf) (princ))
(setq sx:doc (vla-get-activedocument (vlax-get-acad-object)) sx:ms (vla-get-modelspace sx:doc)
      sx:P0 '(0.0 0.0 0.0) sx:R '(1.0 0.0) sx:D '(0.0 1.0) sx:ang 0.0 sx:tagpanel "TEST")
(sx:layers) (sx:roles) (setq sx:style (sx:ensure-style)) (sx:make-blocks T)
(setq rpstage "head block") (prin1 (list sx:bHEAD (sx:blk-p sx:bHEAD)) rpf) (write-line "" rpf)
(setq rpmark (entlast))
(setq rpstage "head insert") (setq rpob (sx:insv sx:bHEAD sx:vHeadDx 0.0 1.0 0.0 nil sx:TB))
(write-line "head inserted" rpf)
(setq rpstage "tagging") (sx:tag-new rpmark "H")
(write-line "head tagged" rpf)
(setq rpstage "column insert") (setq rpob (sx:ins "SX_TCOL" 10.0 0.0 '(("NAME1" . "TEST") ("CUR" . "10")) sx:TB))
(write-line "column inserted" rpf)
(close rpf))) (princ)
'''%((root/'2_AutoCAD/Однолинейка/SX_Schema.lsp').as_posix(),(p/'debug_table.txt').as_posix())
(p/'debug_table.lsp').write_text(code,encoding='utf-8')
for attempt in range(25):
    try:
        a=w.GetActiveObject('AutoCAD.Application');docs=a.Documents
        d=next(docs.Item(i) for i in range(docs.Count) if docs.Item(i).Name=='Однолинейка_проверка.dwg')
        d.Activate();time.sleep(.3);d=a.ActiveDocument
        assert d.Name=='Однолинейка_проверка.dwg'
        d.SendCommand('(load "'+(p/'debug_table.lsp').as_posix()+'")\n');print('Table diagnostic requested');break
    except Exception as exc: last=exc;time.sleep(.4)
else: raise last
