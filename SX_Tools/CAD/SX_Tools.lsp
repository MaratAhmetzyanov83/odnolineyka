;;; SX_Tools.lsp - loader (called by SXTools.dll). ASCII only: folder names are read from disk.
;;; Loads every *.lsp in CAD\ subfolders, adds them to support path and trusted paths.
(vl-load-com)

(defun sxt:subdirs (base / r)
  (foreach d (vl-directory-files base nil -1)
    (if (not (member d '("." "..")))
      (setq r (cons (strcat base "/" d) r))))
  (reverse r))

(defun sxt:win (p) (vl-string-translate "/" "\\" p))

(defun sxt:setup-paths (root dirs / files cur s)
  (vl-catch-all-apply
    '(lambda ()
       (setq files (vla-get-files (vla-get-preferences (vlax-get-acad-object)))
             cur   (vla-get-supportpath files) s cur)
       (foreach d dirs
         (if (not (vl-string-search (strcase (sxt:win d)) (strcase s)))
           (setq s (strcat s ";" (sxt:win d)))))
       (if (/= s cur) (vla-put-supportpath files s))))
  (vl-catch-all-apply
    '(lambda ()
       (setq cur (getvar "TRUSTEDPATHS") s (strcat (sxt:win root) "\\..."))
       (if (not (vl-string-search (strcase s) (strcase cur)))
         (setvar "TRUSTEDPATHS" (strcat cur ";" s))))))

(defun sxt:load-all (/ root cad dirs n err)
  (setq root (vl-string-right-trim "/" (vl-string-translate "\\" "/" sxt:root))
        cad  (strcat root "/CAD")
        dirs (sxt:subdirs cad))
  (sxt:setup-paths root dirs)
  (setq n 0)
  (foreach d dirs
    (foreach f (vl-directory-files d "*.lsp" 1)
      (setq err (vl-catch-all-apply 'load (list (strcat d "/" f))))
      (if (vl-catch-all-error-p err)
        (princ (strcat "\nSX Tools: load error " f ": " (vl-catch-all-error-message err)))
        (setq n (1+ n)))))
  (princ (strcat "\nSX Tools: LISP files loaded - " (itoa n) ". Ribbon tab: SX."))
  (princ))

(if (and (boundp 'sxt:root) sxt:root
         (or (null (boundp 'c:SXDRAW)) (and (boundp 'sxt:force) sxt:force)))
  (sxt:load-all))
(princ)
