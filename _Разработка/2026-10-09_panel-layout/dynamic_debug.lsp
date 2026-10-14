(vl-load-com)
(if (= (getvar "DWGNAME") "Компоновка_проверка_из_базы.dwg") (progn
(setq dd (vla-get-activedocument (vlax-get-acad-object)) ff (open "C:/Users/user/Desktop/02_Справочники_CAD_Revit/odnolineyka/_Разработка/2026-10-09_panel-layout/dynamic_debug.txt" "w") stage "create")
(defun *error* (msg) (prin1 (list stage msg) ff) (close ff) (princ))
(setq tt (vla-add (vla-get-blocks dd) (vlax-3d-point '(0.0 0.0 0.0)) "CP_DYNAMIC_DEBUG"))
(setq stage "insert") (setq oo (vla-insertblock tt (vlax-3d-point '(0.0 0.0 0.0)) "SX_DYN_MCB" 1.0 1.0 1.0 0.0))
(setq stage "isdynamic") (prin1 (vla-get-isdynamicblock oo) ff) (write-line "" ff)
(setq stage "getdynamicprops") (setq props (vlax-invoke oo 'GetDynamicBlockProperties))
(foreach pr props
  (setq stage "propertyname") (prin1 (vla-get-propertyname pr) ff)
  (setq stage "allowed") (prin1 (cp:dyn-values pr) ff) (write-line "" ff))
(setq stage "delete") (vla-delete oo) (vla-delete tt) (close ff))) (princ)
