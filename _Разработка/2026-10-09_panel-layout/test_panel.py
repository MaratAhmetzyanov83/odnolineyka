import pathlib,json,time,sys,win32com.client as w
sys.stdout.reconfigure(encoding='utf-8',line_buffering=True)
p=pathlib.Path(__file__).parent;root=p.parents[1]
def lsp(v):
    if v is None:return 'nil'
    if isinstance(v,str):return '"'+v.replace('\\','\\\\').replace('"','\\"')+'"'
    if isinstance(v,(int,float)):return str(v)
    return '('+' '.join(lsp(x) for x in v)+')'
path=json.loads((root/'_Разработка/2026-10-09_incoming/summary_source.json').read_text(encoding='utf-8'))['path']
code='''(vl-load-com)
(if (= (getvar "DWGNAME") "Компоновка_проверка_из_базы.dwg")
(progn
(load %s)
(defun *error* (msg) (setq f (open "%s" "w")) (prin1 (list 'error msg cp:stage) f) (close f) (princ))
(setq cp:doc (vla-get-activedocument (vlax-get-acad-object)) cp:measure-cache nil cp:cfg (cp:defaults))
(cp:put 'path %s) (cp:put 'library %s) (cp:put 'panel "ЩР-1")
(cp:put 'gap 60.0) (cp:put 'nrows 2) (cp:put 'slots 24) (cp:put 'reserve 20.0)
(setq cp:stage "read book") (cp:read-book (cp:g 'path))
(setq cp:stage "library") (cp:import-library)
(setq cp:stage "collect") (cp:collect (cp:g 'panel))
(setq cp:cfg (vl-remove-if '(lambda (p) (= (car p) 'sizes)) cp:cfg) cp:sd nil)
(foreach d cp:devices (setq cp:sd (cons (cons (nth 1 d) (list (cp:default-block d) 1.0 0.0)) cp:sd)))
(setq cp:sd (cons '("@XT" "VPR-2.5-3L-GY" 1.0 0.0) cp:sd))
(cp:put 'sizes cp:sd)
(setq cp:stage "sizes") (cp:apply-sizes)
(setq cp:stage "plan") (setq cp:ok (cp:plan))
(if cp:ok
  (progn
    (setq cp:stage "clear test drawing")
    (vlax-for ob (vla-get-modelspace cp:doc) (vla-delete ob))
    (setq cp:stage "build") (setq cp:result (cp:build '(0.0 0.0 0.0)))
    (vla-regen cp:doc 1) (vla-save cp:doc)))
(setq f (open "%s" "w"))
(prin1 (list (cons 'ok cp:ok) (cons 'problem cp:problem) (cons 'devices (length cp:devices)) (cons 'rows cp:totalrows) (cons 'sizes cp:sd) (cons 'positions cp:positions) (cons 'block (if cp:result (vla-get-name cp:result)))) f) (close f)
(cp:lib-close))) (princ)
'''%(lsp(str(root/'2_AutoCAD/Компоновка/SX_Panel.lsp').replace('\\','/')),str(p/'test_result.txt').replace('\\','/'),lsp(path),lsp(str(root/'База блоков компоновки (1).dwg')),str(p/'test_result.txt').replace('\\','/'))
(p/'test_panel.lsp').write_text(code,encoding='utf-8')
from native_runner import run
run(p/'test_panel.lsp')
