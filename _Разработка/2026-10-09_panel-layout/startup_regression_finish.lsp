(if (= (getvar "DWGNAME") "CPDRAW_regression.dwg")
(progn
(qa:check "controlled_failure_preserves_layout" (and (= (vla-get-count (vla-get-modelspace cp:doc)) 1) (= qa:oldname (vla-get-name qa:root))))
(qa:check "controlled_failure_closes_library_and_excel" (and (null cp:libsrc) (null cp:wb)))
(load "C:/Users/user/Desktop/02_Справочники_CAD_Revit/odnolineyka/2_AutoCAD/Компоновка/SX_Panel.lsp")
(setq qa:r (vl-catch-all-apply 'cp:fail '("Controlled diagnostic message")))
(qa:check "cpfail_exits_without_undefined_error" (and (vl-catch-all-error-p qa:r) (not (vl-string-search "no function definition" (vl-catch-all-error-message qa:r)))))
(vla-save cp:doc) (close qa:f)
)) (princ)
