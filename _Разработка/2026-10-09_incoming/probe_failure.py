from pathlib import Path
import time,win32com.client as w
p=Path(__file__).parent
code='''(vl-load-com)
(if (= (getvar "DWGNAME") "Однолинейка ЩС-ВК.02.dwg")
(progn
(setq sxdpf (open "%s" "w"))
(foreach key '(sx:ver sx:stage sx:errs sx:logl sx:P0 sx:lib sx:path sx:style sx:part sx:pi sx:bHEAD sx:summary)
  (prin1 (list key (eval key)) sxdpf) (write-line "" sxdpf))
(prin1 (list 'source (findfile "SX_Schema.lsp") 'dcl (findfile "SX_Schema.dcl") 'prefix (getvar "DWGPREFIX") 'lispsys (getvar "LISPSYS")) sxdpf)
(write-line "" sxdpf) (close sxdpf))) (princ)
'''%(p/'failure_probe.txt').as_posix()
(p/'probe_failure.lsp').write_text(code,encoding='utf-8')
for attempt in range(25):
    try:
        a=w.GetActiveObject('AutoCAD.Application')
        d=a.ActiveDocument
        if d.Name!='Однолинейка ЩС-ВК.02.dwg': raise RuntimeError('Target is not active')
        d.SendCommand('(load "'+(p/'probe_failure.lsp').as_posix()+'")\n')
        print('Read-only diagnostic requested');break
    except Exception as exc:
        last=exc;time.sleep(.4)
else: raise last
