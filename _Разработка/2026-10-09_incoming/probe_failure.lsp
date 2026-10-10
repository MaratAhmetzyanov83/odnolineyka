(vl-load-com)
(if (= (getvar "DWGNAME") "Однолинейка ЩС-ВК.02.dwg")
(progn
(setq sxdpf (open "C:/Users/user/Desktop/02_Справочники_CAD_Revit/odnolineyka/_Разработка/2026-10-09_incoming/failure_probe.txt" "w"))
(foreach key '(sx:ver sx:stage sx:errs sx:logl sx:P0 sx:lib sx:path sx:style sx:part sx:pi sx:bHEAD sx:summary)
  (prin1 (list key (eval key)) sxdpf) (write-line "" sxdpf))
(prin1 (list 'source (findfile "SX_Schema.lsp") 'dcl (findfile "SX_Schema.dcl") 'prefix (getvar "DWGPREFIX") 'lispsys (getvar "LISPSYS")) sxdpf)
(write-line "" sxdpf) (close sxdpf))) (princ)
