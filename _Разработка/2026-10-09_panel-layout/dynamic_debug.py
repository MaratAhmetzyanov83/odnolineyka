from pathlib import Path
import win32com.client as w,time
p=Path(__file__).parent
code='''(vl-load-com)
(if (= (getvar "DWGNAME") "Компоновка_проверка_из_базы.dwg") (progn
(setq dd (vla-get-activedocument (vlax-get-acad-object)) ff (open "%s" "w") stage "create")
(defun *error* (msg) (prin1 (list stage msg) ff) (close ff) (princ))
(setq tt (vla-add (vla-get-blocks dd) (vlax-3d-point '(0.0 0.0 0.0)) "CP_DYNAMIC_DEBUG"))
(setq stage "insert") (setq oo (vla-insertblock tt (vlax-3d-point '(0.0 0.0 0.0)) "SX_DYN_MCB" 1.0 1.0 1.0 0.0))
(setq stage "isdynamic") (prin1 (vla-get-isdynamicblock oo) ff) (write-line "" ff)
(setq stage "getdynamicprops") (setq props (vlax-invoke oo 'GetDynamicBlockProperties))
(foreach pr props
  (setq stage "propertyname") (prin1 (vla-get-propertyname pr) ff)
  (setq stage "allowed") (prin1 (cp:dyn-values pr) ff) (write-line "" ff))
(setq stage "delete") (vla-delete oo) (vla-delete tt) (close ff))) (princ)
'''%(p/'dynamic_debug.txt').as_posix()
(p/'dynamic_debug.lsp').write_text(code,encoding='utf-8')
from native_runner import run
run(p/'dynamic_debug.lsp')
