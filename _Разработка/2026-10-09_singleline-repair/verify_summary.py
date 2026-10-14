from pathlib import Path
import time, sys, win32com.client as w
sys.stdout.reconfigure(encoding='utf-8', line_buffering=True)
p = Path(__file__).parent
out = p / 'verify_summary.txt'
def quote(s):
    return '"' + str(s).replace('\\', '\\\\').replace('"', '\\"') + '"'
code = '''(vl-load-com)
(if (= (getvar "DWGNAME") "Однолинейка_проверка.dwg") (progn
(setq vsf (open %s "w") vsc 0 vst 0 vsq nil)
(vlax-for vsob (vla-get-modelspace (vla-get-activedocument (vlax-get-acad-object)))
  (if (= (vla-get-objectname vsob) "AcDbTable") (setq vst (1+ vst)))
  (if (= (vla-get-objectname vsob) "AcDbBlockReference")
    (cond
      ((= (vla-get-effectivename vsob) "SX_SUMMARY")
        (setq vsc (1+ vsc))
        (prin1 (list 'point (vlax-get vsob 'InsertionPoint)
                     'tag (sx:xd-get (vlax-vla-object->ename vsob))
                     'attrs (mapcar '(lambda (a) (cons (vla-get-tagstring a) (vla-get-textstring a))) (vlax-invoke vsob 'GetAttributes))) vsf)
        (write-line "" vsf)
        (vla-getboundingbox vsob 'vsmn 'vsmx)
        (prin1 (list 'bbox (vlax-safearray->list vsmn) (vlax-safearray->list vsmx)) vsf) (write-line "" vsf))
      ((= (vla-get-effectivename vsob) "QF1")
        (foreach vsp (vlax-invoke vsob 'GetDynamicBlockProperties)
          (if (= (vla-get-propertyname vsp) sx:pVis)
            (setq vsq (cons (vlax-variant-value (vla-get-value vsp)) vsq))))))))
(prin1 (list 'summary-count vsc 'native-tables vst 'qf-visibility vsq 'errors sx:errs) vsf) (write-line "" vsf)
;; Exercise the same capture/delete/insert sequence used by SXUPDATE.
(setq vsoriginal sx:summary vspoint sx:sumpt)
(vlax-for vsob sx:ms
  (if (and (= (vla-get-objectname vsob) "AcDbBlockReference") (= (vla-get-effectivename vsob) "SX_SUMMARY"))
    (vla-put-insertionpoint vsob (vlax-3d-point '(200.0 -50.0 0.0)))))
(sx:summary-capture)
(sx:del-keys sx:tagpanel '("U"))
(sx:draw-summary)
(setq vsmoved nil)
(vlax-for vsob sx:ms
  (if (and (= (vla-get-objectname vsob) "AcDbBlockReference") (= (vla-get-effectivename vsob) "SX_SUMMARY"))
    (setq vsmoved (equal (vlax-get vsob 'InsertionPoint) '(200.0 -50.0 0.0) 0.0001))))
(sx:del-keys sx:tagpanel '("U"))
(setq sx:summary nil)
(sx:draw-summary)
(setq vsempty T)
(vlax-for vsob sx:ms
  (if (and (= (vla-get-objectname vsob) "AcDbBlockReference") (= (vla-get-effectivename vsob) "SX_SUMMARY"))
    (setq vsempty nil)))
(setq sx:summary vsoriginal sx:sumpt vspoint)
(sx:draw-summary)
(prin1 (list 'update-preserves-moved-point vsmoved 'no-empty-summary vsempty) vsf)
(vla-save sx:doc)
(close vsf))) (princ)
''' % quote(out.as_posix())
script = p / 'verify_summary.lsp'
script.write_text(code, encoding='utf-8')
for attempt in range(30):
    try:
        app = w.GetActiveObject('AutoCAD.Application')
        docs = app.Documents
        doc = next(docs.Item(i) for i in range(docs.Count) if docs.Item(i).Name == 'Однолинейка_проверка.dwg')
        doc.Activate()
        time.sleep(.3)
        doc = app.ActiveDocument
        assert doc.Name == 'Однолинейка_проверка.dwg'
        doc.SendCommand('(load ' + quote(script.as_posix()) + ')\n')
        print('Verification requested')
        break
    except Exception as exc:
        last = exc
        time.sleep(.4)
else:
    raise last
