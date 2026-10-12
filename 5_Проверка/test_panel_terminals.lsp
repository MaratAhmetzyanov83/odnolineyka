;;; Проверка одиночных клемм. Для проверки наличия блоков открыть копию базы DWG.
(defun qa:check (label ok)
  (write-line (strcat label "=" (if ok "PASS" "FAIL")) qa:panel-report))
(defun qa:row (cable phase count / row i)
  (setq i 44)
  (repeat 44
    (setq row (cons (cond ((= i 17) cable) ((= i 22) phase) ((= i 44) count)) row) i (1- i))) row)
(defun qa:terminal (label cable phase count expected / result)
  (setq cp:terminalgroups nil cp:failure nil)
  (setq result (vl-catch-all-apply 'cp:terminal-set (list "XT11" (qa:row cable phase count))))
  (qa:check label (and (not (vl-catch-all-error-p result))
                      (equal (mapcar '(lambda (d) (nth 2 d)) (cadr (car cp:terminalgroups))) expected))))
(qa:terminal "singlephase_4_three_single_terminals" "3x4" "L1" 1 '("VPR-4-GY" "VPR-4-BU" "VPR-4-PE-YG"))
(qa:check "singlephase_pin_labels" (equal (mapcar 'car (cadr (car cp:terminalgroups))) '("XT11:L" "XT11:N" "XT11:PE")))
(qa:terminal "singlephase_2_5_uses_single_terminals" "3x2,5" "L2" 1 '("VPR-2.5-GY" "VPR-2.5-BU" "VPR-2.5-PE-YG"))
(qa:terminal "singlephase_1_5_PE_alias" "3x1.5" "L3" 1 '("VPR-1.5-GY" "VPR-1.5-BU" "VPR-1.5-PE"))
(qa:terminal "intermediate_section_rounded_up" "3x3" "L1" 1 '("VPR-4-GY" "VPR-4-BU" "VPR-4-PE-YG"))
(qa:terminal "singlephase_6" "3x6" "L1" 1 '("VPR-6-GY" "VPR-6-BU" "VPR-6-PE-YG"))
(qa:terminal "singlephase_10" "3x10" "L1" 1 '("VPR-10-GY" "VPR-10-BU" "VPR-10-PE-YG"))
(qa:terminal "singlephase_16" "3x16" "L1" 1 '("VPR-16-GY" "VPR-16-BU" "VPR-16-PE-YG"))
(qa:terminal "threephase_unchanged" "5x4" "L1,L2,L3" 3 '("VPR-4-GY" "VPR-4-GY" "VPR-4-GY" "VPR-4-BU" "VPR-4-PE-YG"))
(qa:terminal "threephase_from_phase_list" "5x4" "L1,L2,L3" nil '("VPR-4-GY" "VPR-4-GY" "VPR-4-GY" "VPR-4-BU" "VPR-4-PE-YG"))
(setq cp:terminalgroups nil cp:failure nil)
(qa:check "missing_section_still_rejected" (and (vl-catch-all-error-p (vl-catch-all-apply 'cp:terminal-set (list "XT11" (qa:row "" "L1" 1)))) cp:failure))
(setq cp:terminalgroups nil cp:failure nil)
(qa:check "oversize_section_still_rejected" (and (vl-catch-all-error-p (vl-catch-all-apply 'cp:terminal-set (list "XT11" (qa:row "3x25" "L1" 1)))) cp:failure))
;;; Габариты из каталога; функции измерения заменяются только на время расчёта раскладки.
(defun qa:layout ( / cp:measure cp:terminalgroups cp:cfg cp:devices caps rows tight pos row right)
  (defun cp:measure (name scale angle state)
    (list 0.0 0.0 (if (= name "D-VPR-4") 2.2 6.2) 55.9))
  (setq cp:cfg '((nrows . 1) (slots . 2) (pitch . 18.0) (reserve . 0.0) (gap . 60.0) (devh . 90.0)
                 (sizes ("XT|VPR-2.5-3L-GY" "VPR-2.5-3L-GY" 1.0 0.0)))
        cp:devices '(("QF1" "QF" "AV_1P_18" 1.0 90.0 "QF" 3 "AV_1P_18" 1.0 0.0))
        cp:terminalgroups nil)
  (cp:terminal-set "XT11" (qa:row "3x4" "L1" 1))
  (cp:terminal-set "XT12" (qa:row "3x4" "L2" 1))
  (qa:check "single_terminal_layout_fits" (cp:plan))
  (qa:check "ordinary_endcaps_on_each_row" (and (= (length cp:caps) 2) (vl-every '(lambda (d) (= (nth 7 d) "D-VPR-4")) cp:caps)))
  (setq tight T row -1 right 0.0)
  (foreach pos cp:xtpositions
    (if (= (cadr pos) row) (if (not (equal (nth 2 pos) right 0.00001)) (setq tight nil)))
    (setq row (cadr pos) right (+ (nth 2 pos) (nth 3 (car pos)))))
  (qa:check "terminals_touch_and_groups_stay_together"
    (and tight (= cp:xtrows 2) (= (length cp:terminals) 6)
         (equal (mapcar 'cadr cp:xtpositions) '(0 0 0 0 1 1 1 1))
         (not (vl-some '(lambda (d) (wcmatch (nth 7 d) "*-3L-*")) cp:terminals)))))
(qa:layout)
(qa:check "real_library_has_all_single_terminals_and_endcaps"
 (vl-every '(lambda (name) (tblsearch "BLOCK" name))
  '("VPR-1.5-GY" "VPR-1.5-BU" "VPR-1.5-PE" "D-VPR-1.5"
    "VPR-2.5-GY" "VPR-2.5-BU" "VPR-2.5-PE-YG" "D-VPR-2.5"
    "VPR-4-GY" "VPR-4-BU" "VPR-4-PE-YG" "D-VPR-4"
    "VPR-6-GY" "VPR-6-BU" "VPR-6-PE-YG" "D-VPR-6"
    "VPR-10-GY" "VPR-10-BU" "VPR-10-PE-YG" "D-VPR-10"
    "VPR-16-GY" "VPR-16-BU" "VPR-16-PE-YG" "D-VPR-16")))
(princ)
