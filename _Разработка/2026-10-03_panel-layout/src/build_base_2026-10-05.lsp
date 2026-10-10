;;; Пополнение «База блоков компоновки.dwg»: Mean Well KNX-20E-640, EKF КБР, EKF кросс-модули.
;;; Сетка базы: ячейки 117.2 x 155.8 от X 2517.4, рейка («Дин.рейка», *U7) на 82.69 выше низа ряда.
(setvar "CMDECHO" 0)
(setvar "OSMODE" 0)
(setvar "ATTREQ" 0)
(setq dir "C:\\Users\\marat\\AppData\\Local\\Temp\\claude\\C--Users-marat-Desktop-SX------------\\9b70fd8b-8c16-431e-8972-6668244a1be4\\scratchpad\\")

(defun cell (x0 y0 x1 y1)
  (entmake (list '(0 . "LWPOLYLINE") '(100 . "AcDbEntity") '(8 . "0") '(100 . "AcDbPolyline")
                 '(90 . 4) '(70 . 1)
                 (list 10 x0 y1) (list 10 x1 y1) (list 10 x1 y0) (list 10 x0 y0))))

(defun cap (s x y)
  (entmake (list '(0 . "MTEXT") '(100 . "AcDbEntity") '(8 . "0") '(100 . "AcDbMText")
                 (list 10 x y 0.0) '(40 . 10.0) '(41 . 0.0) '(71 . 2) '(72 . 1) (cons 1 s) '(7 . "Standard"))))

(defun ins (name x y)
  (entmake (list '(0 . "INSERT") '(100 . "AcDbEntity") '(8 . "0") '(100 . "AcDbBlockReference")
                 (cons 2 name) (list 10 x y 0.0) '(41 . 1.0) '(42 . 1.0) '(43 . 1.0) '(50 . 0.0))))

;; вставить файл как блок tmp, расчленить один раз, вернуть новые объекты
(defun ins-explode (tmp file pt / mark e res)
  (command "_.-INSERT" (strcat tmp "=" file) pt 1 1 0)
  (setq mark (entlast))
  (command "_.EXPLODE" mark)
  (setq e (entnext mark) res '())
  (while e (setq res (cons e res) e (entnext e)))
  (command "_.-PURGE" "_B" tmp "_N")
  (reverse res))

;; рейка: копия существующей «Дин.рейка» из ряда VPR (Y 1123.49)
(setq rail nil ss (ssget "X" '((0 . "INSERT"))) i 0)
(while (< i (sslength ss))
  (setq e (ssname ss i) p (cdr (assoc 10 (entget e))))
  (if (and (equal (car p) 2527.543 0.01) (equal (cadr p) 1123.4896 0.01)) (setq rail e))
  (setq i (1+ i)))
(defun rail-at (y) (command "_.COPY" rail "" (list 2527.54323342981 1123.489645820265) (list 2527.54323342981 y)))

(setq GX 2517.4 CW 117.2 RH 155.8)

;;; ---- 1. Mean Well KNX-20E-640 — в ряд KNX, ячейка после ZCL-8HT230
(setq cx (+ 3283.8 (/ CW 2.0)))
(cell 3283.8 1508.2 (+ 3283.8 CW) 1664.0)
(cap "MW KNX-20E-640" cx 1640.8)
(ins-explode "MWTMP" (strcat dir "MeanWell_KNX-20E-640.dwg") (list (- cx 26.25) 1573.489645818215))

;;; ---- 2. EKF КБР — ряд под VPR: 6 исполнений (статические блоки) + динамический EKF_KBR
(setq yb 885.0 yt 1040.8 yr (+ yb 82.69))
(rail-at yr)
(command "_.-INSERT" (strcat "EKF_KBR=" dir "ekf1.dwg") (list -1000.0 -1000.0) 1 1 0)
(setq kbrdyn (entlast))
(setq k 0)
(foreach v '(("EKF_KBR_80" "kbr80-rus" 30.0 65.0) ("EKF_KBR_125" "kbr125-rus" 31.0 77.0)
             ("EKF_KBR_160" "kbr160-rus" 41.0 77.0) ("EKF_KBR_250" "kbr250-rus" 48.0 97.0)
             ("EKF_KBR_400" "kbr400-rus" 48.0 97.0) ("EKF_KBR_500" "plc-kbr500" 49.0 96.0))
  (setq x0 (+ GX (* k CW)) cx (+ x0 (/ CW 2.0)))
  (cell x0 yb (+ x0 CW) yt)
  (cap (cadr v) cx (- yt 6.0))
  (ins (car v) (- cx (/ (caddr v) 2.0)) (- yr (/ (cadddr v) 2.0)))
  (setq k (1+ k)))
;; динамический КБР (видимость в свойствах) — в седьмую ячейку, по центру своего габарита
(setq x0 (+ GX (* 6 CW)) cx (+ x0 (/ CW 2.0)))
(cell x0 yb (+ x0 CW) yt)
(cap "EKF_KBR dyn" cx (- yt 6.0))
(setq ed (entget kbrdyn))
(entmod (subst (list 10 (- cx 15.0) (- yr 32.5) 0.0) (assoc 10 ed) ed))

;;; ---- 3. EKF кросс-модули — 2 ряда (2- и 4-полюсные) по 3 ячейки 156 мм, динамические с выбранной видимостью
(setq CW2 156.0)
(setq r2 (list 729.2 885.0) r4 (list 573.4 729.2))
(rail-at (+ (car r2) 82.69))
(rail-at (+ (car r4) 82.69))
(setq IP (list 6000.0 0.0))
;; смещение в файле образцов -> (подпись ряд колонка ширина высота)
(setq cmmap '(((0.0 130.0) "cm207-rus" 2 0 70.5 50.0) ((95.0 130.0) "cm211-rus" 2 1 106.5 50.0)
              ((226.0 130.0) "cm215-rus" 2 2 136.5 50.0) ((0.0 0.0) "cm407-rus" 4 0 70.5 90.7)
              ((95.0 0.0) "cm411-rus" 4 1 106.5 90.7) ((226.0 0.0) "cm415-rus" 4 2 136.5 90.7)))
(foreach e (ins-explode "CMTMP" (strcat dir "ekf2.dwg") IP)
  (setq ed (entget e))
  (cond
    ((= (cdr (assoc 0 ed)) "TEXT") (entdel e))
    ((= (cdr (assoc 0 ed)) "INSERT")
     (setq p (cdr (assoc 10 ed))
           off (list (- (car p) (car IP)) (- (cadr p) (cadr IP)))
           m (vl-some '(lambda (r) (if (equal (car r) off 0.01) r)) cmmap))
     (if m
       (progn
         (setq row (if (= (nth 2 m) 2) r2 r4)
               x0 (+ GX (* (nth 3 m) CW2)) cx (+ x0 (/ CW2 2.0))
               yr (+ (car row) 82.69))
         (cell x0 (car row) (+ x0 CW2) (cadr row))
         (cap (nth 1 m) cx (- (cadr row) 6.0))
         (entmod (subst (list 10 (- cx (/ (nth 4 m) 2.0)) (- yr (/ (nth 5 m) 2.0)) 0.0) (assoc 10 ed) ed)))
       (entdel e)))
    (t (entdel e))))

(command "_.AUDIT" "_Y")
(command "_.ZOOM" "_E")
(princ "\nBUILD-BASE-DONE\n")
(princ)
