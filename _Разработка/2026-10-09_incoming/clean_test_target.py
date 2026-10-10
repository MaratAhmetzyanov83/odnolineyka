import pathlib,re,win32com.client as w
p=pathlib.Path(__file__).parent
data=(p/'target_check.txt').read_text(encoding='cp1251')
handles=re.findall(r'^\("([0-9A-F]+)" "SX_',data,re.M)
assert len(handles)==426
assert all(int(h,16)>=int('211306',16) for h in handles)
a=w.GetActiveObject('AutoCAD.Application')
d=next(q for q in a.Documents if q.Name=='03_Чертежи ЭОМ_ХБК.dwg')
d.Activate()
d=a.ActiveDocument
assert d.Name=='03_Чертежи ЭОМ_ХБК.dwg'
code='''(if (= (getvar "DWGNAME") %s)
  (progn (setq zzn 0)
    (foreach h '%s
      (if (and (setq e (handent h)) (wcmatch (cdr (assoc 8 (entget e))) "SX_*"))
        (progn (entdel e) (setq zzn (1+ zzn)))))
    (setq f (open "%s" "w")) (prin1 zzn f) (close f)
    (vla-regen (vla-get-activedocument (vlax-get-acad-object)) 1)))
(princ)
'''%('"'+d.Name+'"','('+' '.join('"'+h+'"' for h in handles)+')',str(p/'cleanup_result.txt').replace('\\','/'))
(p/'cleanup.lsp').write_text(code,encoding='utf-8')
d.SendCommand('(load "'+str(p/'cleanup.lsp').replace('\\','/')+'")\n')
print('Cleanup requested for',len(handles),'exact handles')
