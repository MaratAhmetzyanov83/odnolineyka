from pathlib import Path
p=Path(__file__).parent;f=p/'SX_Panel.source.lsp';s=f.read_text(encoding='utf-8')
i=s.index('(defun cp:collect ')
s=s[:i]+'''(defun cp:cross-add ( / counts d r parts ph poles idx need ways banks name prefix k)
  ;; Счёт отходящих подключений по фазам + одно питающее подключение на полюс.
  (setq counts '(0 0 0 0) cp:cross-unknown 0)
  (foreach d cp:devices
    (if (= (nth 5 d) "QF")
      (progn
        (setq r (nth (- (nth 6 d) 3) cp:rows) parts (cp:qparts (cp:at r 8))
              ph (cp:norm (cp:at r 22)) poles (car parts))
        (cond ((member poles '("3P" "4P"))
               (setq counts (list (1+ (car counts)) (1+ (cadr counts)) (1+ (caddr counts)) (nth 3 counts))))
              ((member ph '("L1" "L2" "L3"))
               (setq idx (vl-position ph '("L1" "L2" "L3")))
               (setq counts (mapcar '(lambda (n j) (if (= j idx) (1+ n) n)) counts '(0 1 2 3))))
              (t ;; Без фазы не угадываем: предусматриваем подключение на каждой фазе.
               (setq cp:cross-unknown (1+ cp:cross-unknown)
                     counts (list (1+ (car counts)) (1+ (cadr counts)) (1+ (caddr counts)) (nth 3 counts)))))
        (if (member poles '("2P" "4P")) (setq counts (list (car counts) (cadr counts) (caddr counts) (1+ (nth 3 counts))))))))
  (setq cp:cross-counts counts need (apply 'max counts)
        ways (cond ((<= need 6) 7) ((<= need 10) 11) (t 15))
        banks (max 1 (fix (+ (/ (float need) (1- ways)) 0.999999)))
        name (strcat "EKF_CM_4x" (itoa ways)) cp:cross-banks banks cp:cross-ways ways k 1 prefix nil)
  (repeat banks
    (setq prefix (append prefix (list (list (strcat "XRM" (itoa k)) "CROSS|DISTRIBUTION" name nil (cp:g 'devh) "CROSS" 0))) k (1+ k)))
  (setq cp:devices (append prefix cp:devices)))
''' + s[i:]
s=s.replace('  cp:devices)\n(defun cp:apply-sizes', '  (if cp:devices (cp:cross-add))\n  cp:devices)\n(defun cp:apply-sizes')
# Existing defaults select the exact EKF_CM block model; enforce capacity after user changes it.
s=s.replace('  (setq cp:xtrows 0 cp:xtcap 1)', '''  (foreach d cp:devices
    (if (= (nth 5 d) "CROSS")
      (progn
        (setq crossname (nth 7 d))
        (if (or (null crossname) (not (wcmatch crossname "EKF_CM_4x7,EKF_CM_4x11,EKF_CM_4x15")))
          (setq cp:problem "Для кросс-модуля выберите EKF_CM_4x7, EKF_CM_4x11 или EKF_CM_4x15 из вашей базы.")
          (if (< (* cp:cross-banks (1- (atoi (substr crossname 10)))) (apply 'max cp:cross-counts))
            (setq cp:problem "В выбранном кросс-модуле недостаточно подключений. Одно подключение каждого полюса оставлено под питание."))))))
  (setq cp:xtrows 0 cp:xtcap 1)''')
s=s.replace('(defun cp:plan ( / usable packed maxh bb)', '(defun cp:plan ( / usable packed maxh bb d crossname)')
s=s.replace('  (cp:draw-list (+ w 35.0) h))', '''  (cp:draw-list (+ w 35.0) h)
  (cp:text (/ w 2.0) -22.0 4.0
    (strcat "Кросс-модуль: L1=" (itoa (car cp:cross-counts)) "; L2=" (itoa (cadr cp:cross-counts))
            "; L3=" (itoa (caddr cp:cross-counts)) "; N=" (itoa (nth 3 cp:cross-counts)) "; + питание на полюс")
    (+ w 20.0) "CP_TEXT")
  (if (> cp:cross-unknown 0)
    (cp:text (/ w 2.0) -34.0 3.5 "Фазы части автоматов не заданы: место предусмотрено на каждой фазе." w "CP_TEXT")))''')
s=s.replace(' (cons (cons key size) cp:size-data)))', ' (cons (cons key size) cp:size-data)))')
f.write_text(s,encoding='utf-8')
print('Cross-module added for all QFs')
