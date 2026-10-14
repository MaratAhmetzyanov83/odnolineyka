import json,pathlib,time,win32com.client as w
p=pathlib.Path(__file__).parent
def lsp(v):
    if v is None or v=='': return 'nil'
    if isinstance(v,str):return '"'+v.replace('\\','\\\\').replace('"','\\"')+'"'
    if isinstance(v,(int,float)):return str(v)
    return '('+' '.join(lsp(x) for x in v)+')'
vals=json.loads((p/'summary_source.json').read_text(encoding='utf-8'))['values']
# A named but empty block still contains a default cos and phase errors.
vals[12][0]='empty-panel'
for rr in range(13,18):
    vals[rr][4]=None
vals[16][4]=.95
for rr in range(13,16): vals[rr][7]=None
code='''(setq zzchecks nil)
(defun zzcheck (name ok) (setq zzchecks (cons (list name (if ok "PASS" "FAIL")) zzchecks)))
(setq zzvals '%s)
(zzcheck "source totals" (equal (sx:summary-values zzvals sx:tagpanel) sx:summary 0.000001))
(zzcheck "empty table with default cos" (null (sx:summary-values zzvals "empty-panel")))
(setq zzbad (subst (subst "bad-number" (nth 4 (nth 4 zzvals)) (nth 4 zzvals)) (nth 4 zzvals) zzvals))
(zzcheck "invalid filled table skipped" (null (sx:summary-values zzbad sx:tagpanel)))
(zzcheck "two decimals with DIMZIN=8" (progn (setq zzz (getvar "DIMZIN")) (setvar "DIMZIN" 8) (setq zzt (sx:summary-num 0.2)) (setvar "DIMZIN" zzz) (= zzt "0,20")))
(setq zzsum 0 zzout 0 zzwhite 0 zzsumob nil)
(vlax-for ob sx:ms
  (setq e (vlax-vla-object->ename ob) xd (sx:xd-get e))
  (if (equal xd (list sx:tagpanel "U")) (setq zzsum (1+ zzsum) zzsumob ob))
  (if (and (= (vla-get-objectname ob) "AcDbBlockReference") (= (vla-get-layer ob) sx:BL)
           (= (vla-get-effectivename ob) sx:bOut))
    (progn (setq zzout (1+ zzout))
      (if (/= (cdr (assoc 420 (entget e))) sx:red-rgb) (setq zzwhite (1+ zzwhite))))))
(zzcheck "exactly one totals table" (= zzsum 1))
(zzcheck "all fallback outlets red" (and (> zzout 0) (= zzwhite 0)))
(zzcheck "table location and dimensions" (and zzsumob (equal (vla-get-width zzsumob) 120.0 0.000001) (equal (vla-get-height zzsumob) 40.0 0.000001) (equal (sx:pt 7.56437443444702 -5.3734190440593) (vlax-safearray->list (vlax-variant-value (vla-get-insertionpoint zzsumob))) 0.000001)))
(zzcheck "summary values and formatting" (and (= (vla-gettext zzsumob 0 1) "5,73") (= (vla-gettext zzsumob 1 1) "3,95") (= (vla-gettext zzsumob 2 1) "6,23") (= (vla-gettext zzsumob 0 4) "1,37") (= (vla-gettext zzsumob 1 4) "1,26") (= (vla-gettext zzsumob 2 4) "1,33") (= (vla-gettext zzsumob 3 1) "0,95") (= (vla-gettext zzsumob 4 1) "0,69")))
;; Selecting totals updates only totals; no line is expanded into selection.
(setq sx:part T zzkeys (sx:sel-units '("U") sx:preview-draw))
(zzcheck "summary-only selection" (and (equal zzkeys '("U")) (null sx:onlyP) (null sx:onlyG) (null sx:onlyS)))
(sx:summary-capture) (sx:del-keys sx:tagpanel zzkeys) (sx:draw-summary)
(setq zzsaved sx:summary sx:summary nil)
(sx:summary-capture) (sx:del-keys sx:tagpanel '("U")) (sx:draw-summary)
(setq zzsum 0)
(vlax-for ob sx:ms (if (equal (sx:xd-get (vlax-vla-object->ename ob)) (list sx:tagpanel "U")) (setq zzsum (1+ zzsum))))
(zzcheck "remove stale table when no totals" (= zzsum 0))
(setq sx:summary zzsaved) (sx:draw-summary) (setq sx:part nil)
(setq f (open "%s" "w")) (prin1 (reverse zzchecks) f) (close f)
(vla-regen sx:doc 1) (vla-save sx:doc) (princ)
'''%(lsp(vals),str(p/'totals_verification.txt').replace('\\','/'))
(p/'verify_totals.lsp').write_text(code,encoding='utf-8')
a=w.GetActiveObject('AutoCAD.Application')
d=next(q for q in a.Documents if q.Name=='Ввод_проверка.dwg');d.Activate();time.sleep(.5)
d=a.ActiveDocument;assert d.Name=='Ввод_проверка.dwg'
d.SendCommand('(load "'+str(p/'verify_totals.lsp').replace('\\','/')+'")\n')
print('Totals verification requested')
