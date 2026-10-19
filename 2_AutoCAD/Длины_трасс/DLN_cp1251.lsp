;;; ================================================================
;;;  DLN v21 — длины кабельных линий из планов в таблицу однолинейки
;;;  Книги: 1.06 / 1.05 / v3.12 (лист «Исходные данные»); v4, v5 (лист «Линии») — архив.
;;;  Команды:  DLN      — заполнить таблицу (первый раз)
;;;            DLN_UPD  — обновить после правок планов: длины, кол-во, текст выноски;
;;;                       новые линии — по порядку номеров; пропавшие — подсветка
;;;                       и «нет на плане»; «Потолок/Пол», мощность, кабель не трогает.
;;;                       Линия, чья трасса жива на другом листе, — не «нет на плане».
;;;            DLN_SVR  — сверка плана и таблицы по ВСЕМ листам (линии щита на разных
;;;                       этажах, трасса до стояка): лист «Сверка» + пометки, таблицу не меняет
;;;            DLN_CHK  — только проверка чертежа;  DLN_XLS — сбросить путь к таблице
;;;  Перед записью DLN спрашивает таблицу: Enter — та же, «Другая» — выбрать книгу.
;;;  Из модели: рамка двумя точками — берутся только трассы и выноски в рамке
;;;  (Enter — видовые экраны всех листов). На листе — рамки видовых экранов.
;;;  ДСУП: точка РЕ с КУП в CK («Однолинейка») — длина от КУП до точки, трасса проверяется
;;;  до конца трассы своей КУП; сама КУП — от щита до КУП (обычная линия).
;;;  Колонка «DWG: листы» — на каких листах видна линия. Стояк (блок на слое *Стояк*) —
;;;  как щит: трасса этажа, дошедшая до стояка, считается дошедшей.
;;;
;;;  Логика:
;;;   1. Берутся окна видовых экранов планов (шириной в модели >= *dl-minw*).
;;;   2. Полилинии на слоях "*Трасс*", у которых начало или конец в окне.
;;;   3. Мультивыноски: номер линии из первой строки текста,
;;;      стрелка должна лежать на полилинии той же системы (допуск *dl-tol*).
;;;   4. Полилиния без выноски, начинающаяся в конце трассы группы,
;;;      считается продолжением этой группы (допуск *dl-jtol*).
;;;   5. Длины по типу линии: Continuous = гофра ПВХ, JIS_02_0.7 = гофра ПНД,
;;;      Lotok = лоток. Участки гофры, идущие вдоль полилинии лотка
;;;      (слой "*Лоток*" или тип линии Lotok не на слое трасс), — тоже лоток.
;;;   6. В лист "Исходные данные" пишутся колонки «DWG: …» (ПВХ, ПНД, лоток, хэндлы,
;;;      примечание) и «Доп. информация из выноски» — ищутся по заголовку; в J «Лоток» —
;;;      формула от «DWG: лоток». Гофра ПВХ / ПНД — с плана раздельно, опуски
;;;      (итог - по плану) — к гофре по «Потолок/Пол». Строки выше 21-й (вводы и т.п.)
;;;      только обновляются на месте. Книга НЕ сохраняется.
;;;   7. Книга v4 / v5 (есть лист "Линии") — архив: пишутся только колонки DWG и
;;;      «Потолок/Пол» по номеру линии.
;;; ================================================================
(vl-load-com)

(setq *dl-tol*   60.0          ; мм, стрелка выноски -> полилиния
      *dl-jtol*  60.0          ; мм, стык концов полилиний
      *dl-rtol*  400.0         ; мм, выноска на чужой трассе -> свободная трасса рядом
      *dl-chk*   "DLN_Проверка"  ; слой пометок (не печатается)
      *dl-mark*  350.0         ; мм, радиус кружка пометки
      *dl-txth*  180.0         ; мм, высота текста пометки
      *dl-gtol*  500.0         ; мм, «почти стык» концов — считается разрывом трассы
      *dl-ptol*  300.0         ; мм, конец трассы у блока щита или стояка
      *dl-panlay* "*Щит*,*Стояк*" ; слои блоков щита и стояка
      *dl-btol*  150.0         ; мм, вершина трассы -> габарит блока потребителя
      *dl-minw*  20000.0       ; мм, мин. ширина окна видового экрана плана
      *dl-scale* 0.001         ; единицы чертежа -> метры
      *dl-pvh*   "CONTINUOUS"
      *dl-pnd*   "JIS_02_0.7"
      *dl-lot*   "*LOTOK*,*ЛОТОК*" ; тип линии лотка (маска)
      *dl-lotlay* "*Лоток*"    ; слои полилиний лотка
      *dl-ltol*  100.0         ; мм, кабель от оси лотка — считается «в лотке»
      *dl-lstep* 100.0         ; мм, шаг проверки кабеля вдоль лотка
      *dl-lmin*  500.0         ; мм, участки вдоль лотка короче — не лоток (пересечения)
      *dl-sheet* "Исходные данные"
      *dl-row0*  21            ; первая строка, куда DLN пишет линии
      *dl-block* 15            ; строк в блоке
      *dl-gap*   2             ; пустых строк между блоками
      *dl-row1*  1000)
(setq *dl-sheet-v3* *dl-sheet* *dl-row0-v3* *dl-row0* *dl-v4* nil)
;; книга v4 (лист «Линии»): строки с 3-й, без сортировки и формул
(defun dl:v4-setup (book / sh)
  (setq sh (vl-catch-all-apply 'vlax-get-property (list (vlax-get-property book 'Worksheets) 'Item "Линии")))
  (if (vl-catch-all-error-p sh)
    (setq *dl-v4* nil *dl-sheet* *dl-sheet-v3* *dl-row0* *dl-row0-v3*)
    (setq *dl-v4* T *dl-sheet* "Линии" *dl-row0* 3))
  *dl-v4*)
;; v4: последняя строка таблицы «Линии» и расширение таблицы до строки row
(defun dl:v4-table (sh / los lo)
  (setq los (vl-catch-all-apply 'vlax-get-property (list sh 'ListObjects)))
  (if (not (vl-catch-all-error-p los))
    (setq lo (vl-catch-all-apply 'vlax-invoke-method (list los 'Item "Линии"))))
  (if (and lo (not (vl-catch-all-error-p lo))) lo))
(defun dl:v4-last (sh / lo rg)
  (if (setq lo (dl:v4-table sh))
    (progn (setq rg (vlax-get-property lo 'Range))
           (+ (vlax-get-property rg 'Row) (vlax-get-property (vlax-get-property rg 'Rows) 'Count) -1))))
(defun dl:v4-resize (sh row / lo rg ncol)
  (if (setq lo (dl:v4-table sh))
    (progn (setq rg (vlax-get-property lo 'Range)
                 ncol (vlax-get-property (vlax-get-property rg 'Columns) 'Count))
           (vl-catch-all-apply 'vlax-invoke-method
             (list lo 'Resize (dl:rng sh (strcat "A" (itoa (vlax-get-property rg 'Row)) ":" (dl:colname ncol) (itoa row))))))))

;;; ---------- мелкие помощники ----------
(defun dl:2d (p) (list (car p) (cadr p)))
(defun dl:3d (p) (list (car p) (cadr p) (if (caddr p) (caddr p) 0.0)))
(defun dl:r2 (x) (/ (fix (+ (* x 100.0) 0.5)) 100.0))
(defun dl:hnd (e) (cdr (assoc 5 (entget e))))

(defun dl:lt (ed / lt)
  (setq lt (cdr (assoc 6 ed)))
  (if (or (null lt) (= (strcase lt) "BYLAYER"))
    (setq lt (cdr (assoc 6 (tblsearch "LAYER" (cdr (assoc 8 ed)))))))
  (if lt (strcase lt) "CONTINUOUS"))

;; "0_UNC_2 Розетки. Трассы" -> "0_UNC_2 Розетки"
(defun dl:sys (lay / i)
  (if (setq i (vl-string-position 46 lay nil T)) (substr lay 1 i) lay))

;; "0_UNC_3.1 Свет. Трассы" -> "3"
(defun dl:major (lay / i s)
  (setq s "" i 7)
  (if (wcmatch (strcase lay) "0_UNC_*")
    (while (and (<= i (strlen lay)) (wcmatch (substr lay i 1) "#"))
      (setq s (strcat s (substr lay i 1)) i (1+ i))))
  (if (= s "") lay s))

;; первая строка MText без кодов форматирования
(defun dl:plain (s) (dl:plain2 s nil))
;; all = T: все строки через пробел; nil: только первая строка
(defun dl:plain2 (s all / i n c nx out stop)
  (setq i 1 n (strlen s) out "")
  (while (and (<= i n) (not stop))
    (setq c (substr s i 1))
    (cond
      ((or (= c "{") (= c "}")) (setq i (1+ i)))
      ((= c "\\")
       (setq nx (substr s (1+ i) 1))
       (cond
         ((or (= nx "P") (= nx "X")) (if all (setq out (strcat out " ") i (+ i 2)) (setq stop T)))
         ((member nx '("\\" "{" "}")) (setq out (strcat out nx) i (+ i 2)))
         ((member (strcase nx) '("L" "O" "K")) (setq i (+ i 2)))
         ((= nx "~") (setq out (strcat out " ") i (+ i 2)))
         ((and (= (strcase nx) "U") (= (substr s (+ i 2) 1) "+"))
          (setq out (strcat out "?") i (+ i 7)))
         (T (setq i (+ i 2))
            (while (and (<= i n) (/= (substr s i 1) ";")) (setq i (1+ i)))
            (setq i (1+ i)))))
      ((= c "\n") (if all (setq out (strcat out " ") i (1+ i)) (setq stop T)))
      (T (setq out (strcat out c) i (1+ i)))))
  (vl-string-trim " \t" out))

;; "К.1.01.1-(К.1.1) ..." -> "К.1.01.1"; не номер линии -> nil
;; номер линии = первое «слово» первой строки выноски:
;;   "R.1.05.3", "R.0.001г.1", "R.1.001E.1", "R.0.00232", "ОЗДС01", "К.1.01.1-(К.1.1)" -> "К.1.01.1"
;; признаки номера: начинается с буквы, есть цифра, нет «=», длина 2..40
(defun dl:grp (s / i c tok)
  (setq s (vl-string-trim " \t" s) i 1 tok "")
  (while (and (<= i (strlen s)) (not (member (setq c (substr s i 1)) '(" " "\t" "("))))
    (setq tok (strcat tok c) i (1+ i)))
  (setq tok (vl-string-right-trim ".,;:-" tok))
  (if (and (> (strlen tok) 1) (<= (strlen tok) 40)
           (/= (dl:script (substr tok 1 1)) "")
           (wcmatch tok "*#*")
           (not (vl-string-search "=" tok)))
    tok))

;;; ---------- окна видовых экранов ----------
;; окно ВЭ в модели: (x0 y0 x1 y1 (замороженные слои)) или nil
(defun dl:vp-win (ed / c tg h pw ph sc cx cy pc clip o mn mx p0 p1 fr)
  (if (and (assoc 12 ed) (assoc 17 ed) (> (cdr (assoc 41 ed)) 0.0))
    (progn
      (setq c (cdr (assoc 12 ed)) tg (cdr (assoc 17 ed)) h (cdr (assoc 45 ed))
            pw (cdr (assoc 40 ed)) ph (cdr (assoc 41 ed)) sc (/ h ph)
            cx (+ (car tg) (car c)) cy (+ (cadr tg) (cadr c))
            pc (cdr (assoc 10 ed))
            p0 (list (- cx (/ (* pw sc) 2.0)) (- cy (/ h 2.0)))
            p1 (list (+ cx (/ (* pw sc) 2.0)) (+ cy (/ h 2.0))))
      ;; подрезанный ВЭ: берём габарит контура подрезки
      (if (and (setq clip (cdr (assoc 340 ed))) (entget clip)
               (not (eq clip (cdr (assoc -1 ed))))
               (not (vl-catch-all-error-p
                      (vl-catch-all-apply 'vla-GetBoundingBox (list (setq o (vlax-ename->vla-object clip)) 'mn 'mx)))))
        (setq mn (vlax-safearray->list mn) mx (vlax-safearray->list mx)
              p0 (list (+ cx (* (- (car mn) (car pc)) sc)) (+ cy (* (- (cadr mn) (cadr pc)) sc)))
              p1 (list (+ cx (* (- (car mx) (car pc)) sc)) (+ cy (* (- (cadr mx) (cadr pc)) sc)))))
      ;; слои, замороженные в этом ВЭ: коды 341 и xdata ACAD (1003)
      (foreach x ed
        (if (and (= (car x) 341) (= (type (cdr x)) 'ENAME) (entget (cdr x)))
          (setq fr (cons (strcase (cdr (assoc 2 (entget (cdr x))))) fr))))
      (foreach x (cdr (assoc -3 (entget (cdr (assoc -1 ed)) '("ACAD"))))
        (foreach y (cdr x)
          (if (and (= (car y) 1003) (not (member (strcase (cdr y)) fr)))
            (setq fr (cons (strcase (cdr y)) fr)))))
      ;; 6-й элемент: (pcx pcy mcx mcy k лист) — для перехода к пометке на листе
      (list (car p0) (cadr p0) (car p1) (cadr p1) fr
            (list (car pc) (cadr pc) cx cy (/ ph h) (cdr (assoc 410 ed)))))))

(defun dl:width (w) (- (caddr w) (car w)))

;; какие ВЭ брать:
;;  - курсор внутри ВЭ на листе -> этот ВЭ
;;  - на листе в пространстве листа -> выбрать рамки ВЭ (Enter = все крупные ВЭ листа)
;;  - в модели -> все крупные ВЭ всех листов
(defun dl:pick-wins (/ ss i w res p0 p1)
  (cond
    ((and (= (getvar "TILEMODE") 0) (> (getvar "CVPORT") 1))
     (setq ss (ssget "_X" (list '(0 . "VIEWPORT") (cons 410 (getvar "CTAB")) (cons 69 (getvar "CVPORT")))))
     (princ "\nDLN: беру активный видовой экран."))
    ((= (getvar "TILEMODE") 0)
     (princ "\nDLN: выберите рамки видовых экранов планов <Enter — все на листе>: ")
     (if (null (setq ss (ssget '((0 . "VIEWPORT")))))
       (setq ss (ssget "_X" (list '(0 . "VIEWPORT") (cons 410 (getvar "CTAB")))))))
    (T
     (if (setq p0 (getpoint "\nDLN: первый угол рамки плана <Enter — видовые экраны всех листов>: "))
       (setq p1 (getcorner p0 "\nDLN: второй угол рамки: ")))))
  (if (and p0 p1)
    ;; рамка в модели: окно (x0 y0 x1 y1 нет-замороженных (cx cy cx cy 1 "Model"))
    (list (list (min (car p0) (car p1)) (min (cadr p0) (cadr p1))
                (max (car p0) (car p1)) (max (cadr p0) (cadr p1)) nil
                (list 0.0 0.0 0.0 0.0 1.0 "Model")))
    (progn
      (if (and (null ss) (= (getvar "TILEMODE") 1))
        (progn (princ "\nDLN: беру крупные видовые экраны всех листов.")
               (setq ss (ssget "_X" '((0 . "VIEWPORT"))))))
      (dl:ss-wins ss))))
(defun dl:ss-wins (ss / i w res)
  (if ss
    (repeat (setq i (sslength ss))
      (if (and (setq w (dl:vp-win (entget (ssname ss (setq i (1- i))))))
               (>= (dl:width w) *dl-minw*))
        (setq res (cons w res)))))
  res)
;; крупные ВЭ всех листов — для сверки щита, чьи линии на разных листах
(defun dl:all-wins () (dl:ss-wins (ssget "_X" '((0 . "VIEWPORT")))))
;; листы, на которых видна точка pt слоя lay (слой не заморожен в ВЭ)
(defun dl:win-tabs (pt lay / r tab)
  (foreach w *dl-vps*
    (if (and (< (car w) (car pt) (caddr w)) (< (cadr w) (cadr pt) (cadddr w))
             (not (member (strcase lay) (nth 4 w)))
             (not (member (setq tab (nth 5 (nth 5 w))) r)))
      (setq r (cons tab r))))
  r)

(defun dl:layer-on (lay / tb)
  (setq tb (tblsearch "LAYER" lay))
  (and tb (= 0 (logand 1 (cdr (assoc 70 tb)))) (> (cdr (assoc 62 tb)) 0)))
(defun dl:inwin (p lay wins)
  (and (dl:layer-on lay)
  (vl-some '(lambda (w) (and (< (car w) (car p) (caddr w))
                             (< (cadr w) (cadr p) (cadddr w))
                             (not (member (strcase lay) (nth 4 w)))))
           wins)))

;;; ---------- сбор полилиний и выносок ----------
;; запись полилинии: (ename sys major lt len start end)
(defun dl:polys (wins / ss i e ed sp ep res lots cab)
  (setq *dl-trays* (dl:trays wins))
  (if (setq ss (ssget "_X" '((0 . "LWPOLYLINE") (8 . "*Трасс*") (410 . "Model"))))
    (repeat (setq i (sslength ss))
      (setq e (ssname ss (setq i (1- i))) ed (entget e)
            sp (vlax-curve-getStartPoint e) ep (vlax-curve-getEndPoint e))
      (if (or (dl:inwin sp (cdr (assoc 8 ed)) wins) (dl:inwin ep (cdr (assoc 8 ed)) wins))
        (setq res (cons (list e (dl:sys (cdr (assoc 8 ed))) (dl:major (cdr (assoc 8 ed)))
                              (dl:lt ed)
                              (vlax-curve-getDistAtParam e (vlax-curve-getEndParam e))
                              sp ep
                              (mapcar 'cdr (vl-remove-if-not '(lambda (x) (= (car x) 10)) ed))) res)))))
  ;; Lotok на слое трасс поверх кабеля — это лоток, а не трасса: убираем в *dl-trays*
  (foreach p res
    (if (wcmatch (nth 3 p) *dl-lot*) (setq lots (cons p lots)) (setq cab (cons (cons (car p) (dl:bb4 (car p))) cab))))
  (foreach p lots
    (if (>= (dl:along (car p) (nth 4 p) cab) (* 0.5 (nth 4 p)))
      (setq res (vl-remove p res) *dl-trays* (cons (cons (car p) (dl:bb4 (car p))) *dl-trays*))))
  res)

;;; ---------- лоток ----------
;; маска слоя для ssget: спецсимволы экранируются
(defun dl:wesc (s / r)
  (setq r "")
  (foreach c (vl-string->list s)
    (if (member c (vl-string->list "#@.*?~[]-,`")) (setq r (strcat r "`")))
    (setq r (strcat r (chr c))))
  r)
(defun dl:bb4 (e / mn mx)
  (if (not (vl-catch-all-error-p
             (vl-catch-all-apply 'vla-GetBoundingBox (list (vlax-ename->vla-object e) 'mn 'mx))))
    (append (dl:2d (vlax-safearray->list mn)) (dl:2d (vlax-safearray->list mx)))))
(defun dl:bb-hit (a b d)
  (and (<= (- (car a) d) (caddr b)) (<= (car b) (+ (caddr a) d))
       (<= (- (cadr a) d) (cadddr b)) (<= (cadr b) (+ (cadddr a) d))))
;; полилинии лотка в окнах: ((ename x0 y0 x1 y1) ...)
;; слой "*Лоток*" или тип линии Lotok (у объекта или слоя), но не слой трасс
(defun dl:trays (wins / tb nm lays flt ss i e ed lay bb res)
  (while (setq tb (tblnext "LAYER" (null tb)))
    (setq nm (cdr (assoc 2 tb)))
    (if (and (not (wcmatch nm "*Трасс*"))
             (or (wcmatch (strcase nm) (strcase *dl-lotlay*))
                 (wcmatch (strcase (cdr (assoc 6 tb))) *dl-lot*)))
      (setq lays (cons (dl:wesc nm) lays))))
  (setq flt (append '((0 . "LWPOLYLINE,POLYLINE,LINE,ARC") (410 . "Model") (-4 . "<OR") (6 . "*[Ll][Oo][Tt][Oo][Kk]*,*[Лл][Оо][Тт][Оо][Кк]*"))
                    (if lays (list (cons 8 (dl:join lays ","))))
                    '((-4 . "OR>"))))
  (if (setq ss (ssget "_X" flt))
    (repeat (setq i (sslength ss))
      (setq e (ssname ss (setq i (1- i))) ed (entget e) lay (cdr (assoc 8 ed)))
      (if (and (not (wcmatch lay "*Трасс*"))
               (dl:layer-on lay)
               (setq bb (dl:bb4 e))
               (vl-some '(lambda (w) (and (dl:bb-hit bb w 0.0) (not (member (strcase lay) (nth 4 w))))) wins))
        (setq res (cons (cons e bb) res)))))
  res)
;; длина участков кривой e (длина L), идущих вдоль кривых cands ((ename x0 y0 x1 y1) ...)
(defun dl:along (e L cands / bb trs n ds k pt run acc)
  (setq acc 0.0 run 0.0)
  (if (and cands (> L 0.0) (setq bb (dl:bb4 e)))
    (progn
      (foreach t0 cands (if (dl:bb-hit bb (cdr t0) *dl-ltol*) (setq trs (cons (car t0) trs))))
      (if trs
        (progn
          (setq n (max 1 (fix (+ 0.999 (/ L *dl-lstep*)))) ds (/ L n) k 0)
          (repeat n
            (setq pt (vlax-curve-getPointAtDist e (* (+ k 0.5) ds)) k (1+ k))
            (if (and pt
                     (vl-some '(lambda (tr / c)
                                 (and (setq c (vlax-curve-getClosestPointTo tr (dl:3d pt)))
                                      (<= (distance (dl:2d c) (dl:2d pt)) *dl-ltol*)))
                              trs))
              (setq run (+ run ds))
              (setq acc (if (>= run *dl-lmin*) (+ acc run) acc) run 0.0)))
          (if (>= run *dl-lmin*) (setq acc (+ acc run)))))))
  acc)
;; длина участков полилинии кабеля p, идущих вдоль лотка (ед. чертежа)
(defun dl:inlot (p) (dl:along (car p) (nth 4 p) *dl-trays*))

;; пары (первая последняя) вершин каждой линии выноски — из данных объекта
(defun dl:ml-pairs-dxf (e / in first last res)
  (foreach x (entget e)
    (cond ((= (car x) 304) (setq in T first nil last nil))
          ((= (car x) 305) (if (and in first) (setq res (cons (list first last) res))) (setq in nil))
          ((and in (= (car x) 10)) (if (null first) (setq first (cdr x))) (setq last (cdr x)))))
  res)
(defun dl:ml-pairs (obj / res k v n)
  (vl-catch-all-apply
    '(lambda ()
       (setq k 0)
       (repeat (vla-get-LeaderCount obj)
         (foreach li (vlax-invoke obj 'GetLeaderLineIndexes k)
           (setq v (vlax-invoke obj 'GetLeaderLineVertices li))
           (if (>= (length v) 3)
             (setq n (length v)
                   res (cons (list (list (car v) (cadr v) (caddr v))
                                   (list (nth (- n 3) v) (nth (- n 2) v) (nth (- n 1) v)))
                             res))))
         (setq k (1+ k)))))
  res)

;; запись выноски: (grp layer pairs handle)
;; всё из выноски, кроме номера линии
(defun dl:extra (txt g / full)
  (setq full (dl:plain2 txt T))
  (while (vl-string-search "  " full) (setq full (vl-string-subst " " "  " full)))
  (if (= (vl-string-search g full) 0) (setq full (substr full (1+ (strlen g)))))
  (vl-string-trim " -;," full))
(defun dl:labels (wins / ss i e o txt g pairs res)
  (if (setq ss (ssget "_X" '((0 . "MULTILEADER") (410 . "Model"))))
    (repeat (setq i (sslength ss))
      (setq e (ssname ss (setq i (1- i))) o (vlax-ename->vla-object e)
            txt (vl-catch-all-apply 'vla-get-TextString (list o)))
      (if (and (= (type txt) 'STR)
               (setq g (dl:grp (dl:plain txt)))
               (or (null *dl-filter*) (wcmatch (strcase g) (strcase *dl-filter*)))
               (setq pairs (cond ((dl:ml-pairs-dxf e)) ((dl:ml-pairs o))))
               (dl:inwin (car (car pairs)) (cdr (assoc 8 (entget e))) wins))
        (setq res (cons (list g (cdr (assoc 8 (entget e))) pairs (dl:hnd e) (dl:extra txt g)) res)))))
  res)

;;; ---------- сопоставление ----------
(defun dl:getg (e) (cdr (assoc e *dl-asg*)))
(defun dl:setg (e gl) (setq *dl-asg* (subst (cons e gl) (assoc e *dl-asg*) *dl-asg*)))

(defun dl:cands (pt / r cp)
  (foreach p *dl-polys*
    (setq cp (vlax-curve-getClosestPointTo (car p) (dl:3d pt)))
    (if (<= (distance (dl:2d pt) (dl:2d cp)) *dl-tol*) (setq r (cons p r))))
  r)

(defun dl:enddist (pt p)
  (min (distance (dl:2d pt) (dl:2d (nth 5 p))) (distance (dl:2d pt) (dl:2d (nth 6 p)))))

(defun dl:hub (pt / n)
  (setq n 0)
  (foreach p *dl-polys*
    (if (<= (distance (dl:2d pt) (dl:2d (nth 5 p))) *dl-jtol*) (setq n (1+ n)))
    (if (<= (distance (dl:2d pt) (dl:2d (nth 6 p))) *dl-jtol*) (setq n (1+ n))))
  (>= n 3))

(defun dl:match (labels / issues hit pt c same best changed own g gl p keep kd bd bestf)
  (setq *dl-asg* (mapcar '(lambda (p) (cons (car p) nil)) *dl-polys*)
        *dl-hits* nil)
  ;; 1. выноски
  (foreach lb labels
    (foreach pair (nth 2 lb)
      (setq hit nil)
      (foreach pt pair (if (and (not hit) (setq c (dl:cands pt))) (setq hit (list pt c))))
      (if (not hit)
        (setq issues (cons (list (car lb) "выноска не касается трассы" (car pair)) issues))
        (progn
          (setq pt (car hit) c (cadr hit)
                same (vl-remove-if-not '(lambda (p) (= (nth 1 p) (dl:sys (nth 1 lb)))) c))
          (if (null same)
            (setq same (vl-remove-if-not '(lambda (p) (= (nth 2 p) (dl:major (nth 1 lb)))) c)))
          (if (null same)
            (setq issues (cons (list (car lb) "выноска на трассе другой системы" pt) issues))
            (progn
              (setq best (car same))
              (foreach p (cdr same) (if (< (dl:enddist pt p) (dl:enddist pt best)) (setq best p)))
              (setq *dl-hits* (cons (list (car lb) (car best) pt) *dl-hits*))
              (if (not (member (car lb) (dl:getg (car best))))
                (dl:setg (car best) (cons (car lb) (dl:getg (car best)))))))))))
  ;; 2. полилиния с 2+ группами: убрать группу, у которой есть своя трасса
  (setq changed T)
  (while changed
    (setq changed nil)
    (foreach a *dl-asg*
      (if (> (length (setq gl (dl:getg (car a)))) 1)
        (foreach g gl
          (setq own (vl-some '(lambda (b) (and (not (eq (car b) (car a))) (equal (cdr b) (list g)))) *dl-asg*))
          (if (and own (> (length (dl:getg (car a))) 1))
            (progn (dl:setg (car a) (vl-remove g (dl:getg (car a)))) (setq changed T)))))))
  ;; 2б. на общей трассе остаётся группа, чья стрелка ближе к её концу;
  ;;     остальные уходят на свободную трассу, конец которой рядом со стрелкой
  (foreach a *dl-asg*
    (if (> (length (setq gl (dl:getg (car a)))) 1)
      (progn
        (setq p (assoc (car a) *dl-polys*) keep nil kd 1e99)
        (foreach g gl
          (foreach hh *dl-hits*
            (if (and (= (car hh) g) (eq (cadr hh) (car a)) (< (dl:enddist (caddr hh) p) kd))
              (setq kd (dl:enddist (caddr hh) p) keep g))))
        (foreach g gl
          (if (/= g keep)
            (progn
              (setq bestf nil bd *dl-rtol*)
              (foreach hh *dl-hits*
                (if (and (= (car hh) g) (eq (cadr hh) (car a)))
                  (foreach p2 *dl-polys*
                    (if (and (null (dl:getg (car p2))) (= (nth 1 p2) (nth 1 p))
                             (<= (dl:enddist (caddr hh) p2) bd))
                      (setq bd (dl:enddist (caddr hh) p2) bestf p2)))))
              (if bestf
                (progn (dl:setg (car a) (vl-remove g (dl:getg (car a))))
                       (dl:setg (car bestf) (list g))))))))))
  ;; 3. цепочки: конец одной = начало другой (но не через щит, где сходятся 3+ трасс)
  (setq changed T)
  (while changed
    (setq changed nil)
    (foreach p *dl-polys*
      (if (null (dl:getg (car p)))
        (foreach p2 *dl-polys*
          (if (and (null (dl:getg (car p)))
                   (not (eq (car p) (car p2)))
                   (= (length (dl:getg (car p2))) 1)
                   (= (nth 1 p) (nth 1 p2))
                   (vl-some '(lambda (a) (vl-some '(lambda (b) (and (<= (distance (dl:2d a) (dl:2d b)) *dl-jtol*)
                                                                    (not (dl:hub a))))
                                                  (list (nth 5 p2) (nth 6 p2))))
                            (list (nth 5 p) (nth 6 p))))
            (progn (dl:setg (car p) (dl:getg (car p2))) (setq changed T)))))))
  issues)

;;; ---------- потребители (блоки) ----------
(defun dl:bbox (e / mn mx)
  (if (not (vl-catch-all-error-p
             (vl-catch-all-apply 'vla-GetBoundingBox (list (vlax-ename->vla-object e) 'mn 'mx))))
    (list (vlax-safearray->list mn) (vlax-safearray->list mx))))
(defun dl:rect-d (p bb / dx dy)
  (setq dx (max (- (car (car bb)) (car p)) 0.0 (- (car p) (car (cadr bb))))
        dy (max (- (cadr (car bb)) (cadr p)) 0.0 (- (cadr p) (cadr (cadr bb)))))
  (sqrt (+ (* dx dx) (* dy dy))))
(defun dl:inc (key al) (if (assoc key al) (subst (cons key (1+ (cdr (assoc key al)))) (assoc key al) al) (cons (cons key 1) al)))
;; имя блока (для динамических — настоящее имя)
(defun dl:bname (e / r)
  (setq r (vl-catch-all-apply 'vla-get-EffectiveName (list (vlax-ename->vla-object e))))
  (if (vl-catch-all-error-p r) (cdr (assoc 2 (entget e))) r))
;; расстояние от точки до отрезка
(defun dl:pt-seg (p a b / dx dy L tt)
  (setq dx (- (car b) (car a)) dy (- (cadr b) (cadr a)) L (+ (* dx dx) (* dy dy))
        tt (if (= L 0.0) 0.0
             (max 0.0 (min 1.0 (/ (+ (* (- (car p) (car a)) dx) (* (- (cadr p) (cadr a)) dy)) L)))))
  (distance (dl:2d p) (list (+ (car a) (* tt dx)) (+ (cadr a) (* tt dy)))))
;; расстояние от отрезка до габарита блока (0 — отрезок проходит через блок)
(defun dl:seg-rect-d (a b bb / x0 y0 x1 y1 dx dy t0 t1 ok p q tt)
  (setq x0 (car (car bb)) y0 (cadr (car bb)) x1 (car (cadr bb)) y1 (cadr (cadr bb))
        dx (- (car b) (car a)) dy (- (cadr b) (cadr a)) t0 0.0 t1 1.0 ok T)
  (foreach pq (list (cons (- dx) (- (car a) x0)) (cons dx (- x1 (car a)))
                    (cons (- dy) (- (cadr a) y0)) (cons dy (- y1 (cadr a))))
    (if ok
      (progn
        (setq p (car pq) q (cdr pq))
        (if (equal p 0.0 1e-12)
          (if (< q 0.0) (setq ok nil))
          (progn
            (setq tt (/ q p))
            (if (< p 0.0)
              (if (> tt t1) (setq ok nil) (setq t0 (max t0 tt)))
              (if (< tt t0) (setq ok nil) (setq t1 (min t1 tt)))))))))
  (if ok
    0.0
    (min (dl:rect-d a bb) (dl:rect-d b bb)
         (dl:pt-seg (list x0 y0) a b) (dl:pt-seg (list x0 y1) a b)
         (dl:pt-seg (list x1 y0) a b) (dl:pt-seg (list x1 y1) a b))))
(defun dl:poly-rect-d (vs bb / d)
  (setq d 1e99)
  (while (cdr vs)
    (setq d (min d (dl:seg-rect-d (car vs) (cadr vs) bb)) vs (cdr vs)))
  d)
(defun dl:lex< (a b)   ; сравнение списков чисел (d dv da) с допуском 1 мм
  (cond ((null a) nil)
        ((< (car a) (- (car b) 1.0)) T)
        ((> (car a) (+ (car b) 1.0)) nil)
        (T (dl:lex< (cdr a) (cdr b)))))
;; блок "…Оборудование" той же системы относится к группе, чья трасса проходит через него
;; или ближе всех (не дальше *dl-btol*). При равенстве: у кого вершина трассы ближе,
;; затем у кого стрелка выноски ближе. Блоки у 5+ групп (щит) и полные ничьи — пропуск.
(defun dl:count-blocks (wins / gp ss i e ed lay bb sys cands d dv da best second cnt ctr seen key)
  ;; группа -> её полилинии
  (foreach p *dl-polys*
    (if (= (length (dl:getg (car p))) 1)
      (setq gp (if (assoc (car (dl:getg (car p))) gp)
                 (subst (append (assoc (car (dl:getg (car p))) gp) (list p))
                        (assoc (car (dl:getg (car p))) gp) gp)
                 (cons (list (car (dl:getg (car p))) p) gp)))))
  (if (setq ss (ssget "_X" '((0 . "INSERT") (8 . "*Оборуд*") (410 . "Model"))))
    (repeat (setq i (sslength ss))
      (setq e (ssname ss (setq i (1- i))) ed (entget e) lay (cdr (assoc 8 ed)) cands nil)
      ;; положение блока — по центру его геометрии (точка вставки бывает далеко от самого блока)
      (if (and (setq bb (dl:bbox e))
               (setq ctr (list (/ (+ (car (car bb)) (car (cadr bb))) 2.0)
                               (/ (+ (cadr (car bb)) (cadr (cadr bb))) 2.0)))
               (dl:inwin ctr lay wins)
               ;; блок, лежащий точно поверх такого же, считаем один раз
               (not (member (setq key (list (strcase (dl:bname e)) (fix (/ (car ctr) 10.0)) (fix (/ (cadr ctr) 10.0)))) seen)))
        (progn
          (setq seen (cons key seen)
                sys (dl:major lay))
          (foreach x gp
            (if (= (nth 2 (cadr x)) sys)
              (progn
                (setq d (apply 'min (mapcar '(lambda (p) (dl:poly-rect-d (nth 7 p) bb)) (cdr x))))
                (if (<= d *dl-btol*)
                  (progn
                    (setq dv 1e99 da 1e99)
                    (foreach p (cdr x) (foreach v (nth 7 p) (setq dv (min dv (dl:rect-d v bb)))))
                    (foreach hh *dl-hits* (if (= (car hh) (car x)) (setq da (min da (dl:rect-d (caddr hh) bb)))))
                    (setq cands (cons (list (list d dv da) (car x)) cands)))))))
          (if (and cands (< (length cands) 5))
            (progn
              (setq best (car cands) second nil)
              (foreach c (cdr cands)
                (if (dl:lex< (car c) (car best)) (setq second best best c)
                  (if (or (null second) (dl:lex< (car c) (car second))) (setq second c))))
              (if (or (null second) (dl:lex< (car best) (car second)))
                (setq cnt (dl:inc (cadr best) cnt)))))))))
  cnt)

;; итог по группе: (grp pvh pnd handles note кол-во_блоков выноска лоток листы)
(defun dl:summary (labels / names res pvh pnd lot hs notes gl L lt lp tabs)
  (foreach lb labels (if (not (member (car lb) names)) (setq names (cons (car lb) names))))
  (foreach g (reverse names)
    (setq pvh 0.0 pnd 0.0 lot 0.0 hs nil notes nil tabs nil)
    ;; листы: где видны стрелки выносок и начала трасс группы
    (foreach lb labels
      (if (= (car lb) g)
        (foreach pr (nth 2 lb) (setq tabs (append tabs (dl:win-tabs (car pr) (nth 1 lb)))))))
    (foreach p *dl-polys*
      (setq gl (dl:getg (car p)) L (* (nth 4 p) *dl-scale*) lt (nth 3 p))
      (cond
        ((equal gl (list g))
         (setq hs (cons (dl:hnd (car p)) hs)
               tabs (append tabs (dl:win-tabs (nth 5 p) (cdr (assoc 8 (entget (car p)))))
                                 (dl:win-tabs (nth 6 p) (cdr (assoc 8 (entget (car p)))))))
         (cond ((wcmatch lt *dl-lot*) (setq lot (+ lot L)))
               ((or (= lt *dl-pvh*) (= lt *dl-pnd*))
                (setq lp (min L (* (dl:inlot p) *dl-scale*)) lot (+ lot lp))
                (if (= lt *dl-pvh*) (setq pvh (+ pvh (- L lp))) (setq pnd (+ pnd (- L lp)))))
               (T (setq notes (cons (strcat "тип линии " lt " " (rtos L 2 1) " м не учтён") notes)))))
        ((and (> (length gl) 1) (member g gl))
         (setq notes (cons (strcat "полилиния " (dl:hnd (car p)) " общая с "
                                   (apply 'strcat (mapcar '(lambda (x) (if (= x g) "" (strcat x " "))) gl))
                                   "- не засчитана") notes)))))
    (if (and (null hs) (null notes)) (setq notes (list "трасса на плане не найдена")))
    (setq res (cons (list g (dl:r2 pvh) (dl:r2 pnd) (reverse hs)
                          (if notes (apply 'strcat (mapcar '(lambda (x) (strcat x "; ")) (reverse notes))) "")
                          (cdr (assoc g *dl-cnt*))
                          (dl:join (dl:uniq (vl-remove "" (mapcar '(lambda (lb) (if (= (car lb) g) (nth 4 lb) "")) labels))) "; ")
                          (dl:r2 lot)
                          (vl-sort (dl:uniq tabs) '<))
                    res)))
  (reverse res))

(defun dl:uniq (l / r) (foreach x l (if (not (member x r)) (setq r (cons x r)))) (reverse r))
(defun dl:join (l sep / r) (foreach x l (setq r (if r (strcat r sep x) x))) (if r r ""))

;;; ---------- Excel ----------
(defun dl:vv (x / r)
  (if (= (type x) 'VARIANT)
    (if (vl-catch-all-error-p (setq r (vl-catch-all-apply 'vlax-variant-value (list x)))) nil r)
    x))
;; выполнить шаг Excel; при ошибке — сказать, на каком шаге, и остановиться
(defun dl:wait (ms / t0) (setq t0 (getvar "MILLISECS")) (while (< (- (getvar "MILLISECS") t0) ms)))
(defun dl:try (step f args / r n)
  (setq n 0)
  (while (and (vl-catch-all-error-p (setq r (vl-catch-all-apply f args))) (< (setq n (1+ n)) 5))
    (dl:wait 300))
  (if (vl-catch-all-error-p r)
    (progn (princ (strcat "\nDLN: ошибка Excel на шаге [" step "]: " (vl-catch-all-error-message r)))
           (exit)))
  r)
(defun dl:rng (sh a) (dl:try (strcat "ячейка " a) 'vlax-get-property (list sh 'Range a)))
(defun dl:put (sh a v)
  (if (or (null v) (= v "") (and (numberp v) (= v 0.0)))
    (dl:try (strcat "очистка " a) 'vlax-invoke-method (list (dl:rng sh a) 'ClearContents))
    (progn
      (if (and (= (type v) 'STR) (member (substr v 1 1) '("=" "+" "-" "@"))) (setq v (strcat "'" v)))
      (dl:try (strcat "запись " a) 'vlax-put-property (list (dl:rng sh a) 'Value2 v)))))
(defun dl:ru (f)
  (foreach x '(("ROUNDUP(" . "ОКРУГЛВВЕРХ(") ("COUNT(" . "СЧЁТ(") ("SUM(" . "СУММ(")
               ("MAX(" . "МАКС(") ("OR(" . "ИЛИ(") ("IF(" . "ЕСЛИ(") ("N(" . "Ч("))
    (setq f (dl:subst-all (car x) (cdr x) f)))
  (vl-string-translate "," ";" f))
(defun dl:subst-all (find repl str / i)
  (setq i 0)
  (while (setq i (vl-string-search find str i))
    (setq str (strcat (substr str 1 i) repl (substr str (+ i (strlen find) 1)))
          i (+ i (strlen repl))))
  str)
;; формула: Formula -> FormulaLocal (рус.) -> Value2 с "="
(defun dl:putf (sh a f / r)
  (setq r (dl:rng sh a))
  (cond
    ((and (/= *dl-fmode* 2) (/= *dl-fmode* 3)
          (not (vl-catch-all-error-p (vl-catch-all-apply 'vlax-put-property (list r 'Formula f)))))
     (setq *dl-fmode* 1))
    ((and (/= *dl-fmode* 3)
          (not (vl-catch-all-error-p (vl-catch-all-apply 'vlax-put-property (list r 'FormulaLocal (dl:ru f))))))
     (setq *dl-fmode* 2))
    ((not (vl-catch-all-error-p (vl-catch-all-apply 'vlax-put-property (list r 'Value2 f))))
     (setq *dl-fmode* 3))
    (T (princ (strcat "\nDLN: Excel не принимает формулу в " a " ни одним способом: " f)) (exit))))
(defun dl:col (sh c / v)
  (setq v (dl:vv (dl:try (strcat "чтение колонки " c) 'vlax-get-property
                         (list (dl:rng sh (strcat c (itoa *dl-row0*) ":" c (itoa *dl-row1*))) 'Value2))))
  (mapcar '(lambda (row) (dl:vv (car row))) (dl:try (strcat "разбор колонки " c) 'vlax-safearray->list (list v))))
;; формулы колонки (как в строке формул): "" — пусто, "=..." — формула
(defun dl:colf (sh c / v)
  (setq v (dl:vv (dl:try (strcat "чтение формул " c) 'vlax-get-property
                         (list (dl:rng sh (strcat c (itoa *dl-row0*) ":" c (itoa *dl-row1*))) 'Formula))))
  (mapcar '(lambda (row) (dl:str (dl:vv (car row)))) (dl:try (strcat "разбор формул " c) 'vlax-safearray->list (list v))))
(defun dl:str (x) (cond ((null x) "") ((= (type x) 'STR) (vl-string-trim " " x)) (T (vl-princ-to-string x))))

(defun dl:book (path / xl book wbs)
  (setq xl (vl-catch-all-apply 'vlax-get-object (list "Excel.Application")))
  (if (or (null xl) (vl-catch-all-error-p xl))
    (setq xl (dl:try "запуск Excel" 'vlax-create-object (list "Excel.Application"))))
  (if (null xl) (progn (princ "\nDLN: не удалось запустить Excel.") (exit)))
  (vl-catch-all-apply 'vlax-put-property (list xl 'Visible :vlax-true))
  (setq wbs (dl:try "список книг" 'vlax-get-property (list xl 'Workbooks)))
  (vlax-for b wbs
    (if (= (strcase (vlax-get-property b 'FullName)) (strcase path)) (setq book b)))
  (if (null book) (setq book (dl:try (strcat "открытие " path) 'vlax-invoke-method (list wbs 'Open path))))
  (if (= (vl-catch-all-apply 'vlax-get-property (list book 'ReadOnly)) :vlax-true)
    (princ "\nDLN: ВНИМАНИЕ — книга открыта только для чтения, сохранить не получится."))
  book)


(defun dl:split (s / i res)
  (while (setq i (vl-string-search " " s))
    (if (> i 0) (setq res (cons (substr s 1 i) res)))
    (setq s (substr s (+ i 2))))
  (if (/= s "") (setq res (cons s res)))
  (reverse res))
(defun dl:alive (h / e) (and (setq e (handent h)) (entget e)))
;; хэндлы строки таблицы по номеру линии
(defun dl:hands-of (n names hands tnames thands / i)
  (cond ((setq i (vl-position n names)) (nth i hands))
        ((and tnames (setq i (vl-position n tnames))) (nth i thands))
        (T "")))
;; трасса строки жива в чертеже и не отдана другой группе — значит, она на другом листе
(defun dl:elsewhere (hs sum / used)
  (foreach s sum (setq used (append (nth 3 s) used)))
  (vl-some '(lambda (h) (and (dl:alive h) (not (member h used)))) (dl:split hs)))
(defun dl:colname (n / s)
  (setq s "")
  (while (> n 0) (setq s (strcat (chr (+ 65 (rem (1- n) 26))) s) n (/ (1- n) 26)))
  s)
;; поиск колонки по заголовку в строке 2: mode 0 — точно, 1 — начинается с, 2 — содержит
(defun dl:hcol (key mode / i r)
  (setq i 1 key (strcase key))
  (foreach h *dl-hdrs*
    (if (and (null r)
             (cond ((= mode 0) (= h key))
                   ((= mode 1) (= (vl-string-search key h) 0))
                   (T (vl-string-search key h))))
      (setq r i))
    (setq i (1+ i)))
  r)
;; колонка обязательная: нет — остановка
(defun dl:hreq (key mode / r)
  (if (setq r (dl:hcol key mode)) (dl:colname r)
    (progn (princ (strcat "\nDLN: в строке 2 листа не найдена колонка «" key "».")) (exit))))
;; колонка DWG: нет — создаём справа от последней
(defun dl:hnew (sh key mode title / r)
  (if (setq r (dl:hcol key mode)) (dl:colname r)
    (progn
      (setq *dl-lastcol* (1+ *dl-lastcol*))
      (dl:put sh (strcat (dl:colname *dl-lastcol*) "2") title)
      (setq *dl-hdrs* (append *dl-hdrs* (list (strcase title))))
      (dl:colname *dl-lastcol*))))

;; цвет подсветки строки «нет на плане» (светло-красный, RGB 255,199,206)
(setq *dl-miss-color* 13551615 *dl-row-top* 3)
(defun dl:paint (sh a on / it)
  (setq it (vl-catch-all-apply 'vlax-get-property (list (dl:rng sh a) 'Interior)))
  (if (not (vl-catch-all-error-p it))
    (if on
      (vl-catch-all-apply 'vlax-put-property (list it 'Color *dl-miss-color*))
      (if (equal (dl:vv (vl-catch-all-apply 'vlax-get-property (list it 'Color))) (float *dl-miss-color*) 0.5)
        (vl-catch-all-apply 'vlax-put-property (list it 'ColorIndex -4142))))))
;; строки выше *dl-row0* (вводы, ЩБП и т.п., вписанные вручную): только найти и обновить на месте
(defun dl:col-top (sh c / v)
  (if (> *dl-row0* *dl-row-top*)
    (progn
      (setq v (dl:vv (dl:try (strcat "чтение колонки " c) 'vlax-get-property
                             (list (dl:rng sh (strcat c (itoa *dl-row-top*) ":" c (itoa (1- *dl-row0*)))) 'Value2))))
      (mapcar '(lambda (row) (dl:vv (car row))) (dl:try (strcat "разбор колонки " c) 'vlax-safearray->list (list v))))))
;; подсветка «нет на плане» / «полилинии удалены» — после раскладки строк, по номеру линии
(defun dl:paint-all (sh cD marked / v i n last)
  (setq v (mapcar '(lambda (row) (dl:str (dl:vv (car row))))
            (vlax-safearray->list (dl:vv (dl:try "чтение номеров" 'vlax-get-property
              (list (dl:rng sh (strcat cD (itoa *dl-row-top*) ":" cD (itoa *dl-row1*))) 'Value2)))))
        i *dl-row-top* last *dl-row-top*)
  (foreach n v (if (/= n "") (setq last i)) (setq i (1+ i)))
  (setq i *dl-row-top*)
  (foreach n v
    (if (<= i (+ last *dl-block*))
      (dl:paint sh (strcat cD (itoa i)) (and (/= n "") (member n marked))))
    (setq i (1+ i))))
;; найти строку линии g: точное совпадение, затем «те же знаки, другие буквы»
(defun dl:find-row (g names row0 / row i)
  (setq i row0)
  (foreach n names (if (and (null row) (= n g)) (setq row i)) (setq i (1+ i)))
  (if (null row)
    (progn
      (setq i row0)
      (foreach n names
        (if (and (null row) (/= n "") (= (dl:norm n) (dl:norm g)))
          (setq row i *dl-renamed* (cons (list g n) *dl-renamed*)))
        (setq i (1+ i)))))
  row)

;; mode: 'fill — DLN (первое заполнение), 'upd — DLN_UPD (обновить после правок планов)
(defun dl:write (path sum mode / book sh cD cG cI cJ cK cL cM cS cPV cPN cLT cHN cNT cEX cSH
                            names tnames hands thands svals tsvals jf lref rs i last g row newn upd n isnew sv nlines ans)
  (setq book (dl:book path)
        *dl-app* (vlax-get-property book 'Application)
        *dl-v4* (dl:v4-setup book)
        *dl-kup* (dl:read-kup book)
        *dl-fmode* nil
        sh (dl:try (strcat "лист " *dl-sheet*) 'vlax-get-property
                   (list (dl:try "листы книги" 'vlax-get-property (list book 'Worksheets)) 'Item *dl-sheet*)))
  (if (= (vl-catch-all-apply 'vlax-get-property (list sh 'ProtectContents)) :vlax-true)
    (progn (princ (strcat "\nDLN: лист «" *dl-sheet* "» защищён — снимите защиту.")) (exit)))
  ;; заголовки строки 2
  (dl:read-hdrs sh)
  (setq cD (dl:hreq "Номер линии" 0)
        names (mapcar 'dl:str (dl:col sh cD))
        nlines 0)
  (foreach n names (if (/= n "") (setq nlines (1+ nlines))))
  ;; DLN по уже заполненной таблице — предложить режим «обновить»
  (if (and (= mode 'fill) (> nlines 0) (not *dl-v4*))
    (progn
      (initget "Обновить Заполнить _Update Fill")
      (setq ans (getkword (strcat "\nDLN: в таблице уже " (itoa nlines)
                                  " линий. Для правок планов — DLN_UPD. [Обновить/Заполнить] <Обновить>: ")))
      (if (or (null ans) (= ans "Update")) (setq mode 'upd))))
  (vl-catch-all-apply 'vlax-put-property (list *dl-app* 'Calculation -4135))
  (vl-catch-all-apply 'vlax-put-property (list *dl-app* 'ScreenUpdating :vlax-false))
  (setq cG (dl:hreq "Кол-во потребителей" 1)
        cI (dl:hreq "Длина линии (по плану)" 1)
        cK (dl:hreq "Гофра ПВХ" 1)
        cL (dl:hreq "Гофра ПНД" 1)
        cM (if *dl-v4* (dl:hreq "Длина, м" 0) (dl:hreq "Длина линии (Итог)" 1))
        cS (dl:hreq "Потолок" 1)
        cPV (dl:hnew sh "DWG: гофра ПВХ" 1 "DWG: гофра ПВХ по плану, м")
        cPN (dl:hnew sh "DWG: гофра ПНД" 1 "DWG: гофра ПНД по плану, м")
        cHN (dl:hnew sh "DWG: хэндлы" 1 "DWG: хэндлы полилиний")
        cNT (dl:hnew sh "DWG: примечание" 1 "DWG: примечание")
        cEX (dl:hnew sh "выноск" 2 "Доп. информация из выноски")
        cLT (dl:hnew sh "DWG: лоток" 1 "DWG: лоток по плану, м")
        cSH (dl:hnew sh "DWG: листы" 1 "DWG: листы плана")
        cJ (if (dl:hcol "Лоток" 1) (dl:colname (dl:hcol "Лоток" 1))))
  (setq tnames (if *dl-v4* nil (mapcar 'dl:str (dl:col-top sh cD)))
        tsvals (if *dl-v4* nil (dl:col-top sh cS))
        *dl-tnames* (append tnames names)
        *dl-renamed* nil
        *dl-codes* (dl:read-codes book)
        hands (mapcar 'dl:str (dl:col sh cHN))
        svals (dl:col sh cS)
        jf (if cJ (dl:colf sh cJ))
        last (1- *dl-row0*) i *dl-row0*)
  (foreach n names (if (/= n "") (setq last i)) (setq i (1+ i)))
  (if (and *dl-v4* (dl:v4-last sh)) (setq last (dl:v4-last sh)))
  (setq newn 0 upd 0)
  (foreach s sum
    (setq g (car s) isnew nil
          row (dl:find-row g names *dl-row0*))
    (if (and (null row) tnames) (setq row (dl:find-row g tnames *dl-row-top*)))
    (if (null row)
      (progn
        (setq last (1+ last) row last rs (itoa row) newn (1+ newn) isnew T)
        (dl:put sh (strcat cD rs) g)
        (dl:put sh (strcat cG rs) 1))
      (setq upd (1+ upd)))
    (setq rs (itoa row))
    (if (nth 5 s) (dl:put sh (strcat cG rs) (nth 5 s)))
    (if (not *dl-v4*)
      (progn
        ;; длина по плану = ПВХ + ПНД + лоток с плана
        (dl:putf sh (strcat cI rs)
          (strcat "=IF(COUNT(" cPV rs "," cPN rs "," cLT rs ")=0,\"\",ROUNDUP(SUM(" cPV rs "," cPN rs "," cLT rs "),0))"))
        ;; J «Лоток»: формула от «DWG: лоток», если лоток найден или в J не ручное число
        (if (and cJ (or (> (nth 7 s) 0.0)
                        (and (>= row *dl-row0*) (member (substr (dl:str (nth (- row *dl-row0*) jf)) 1 1) '("" "=")))))
          (dl:putf sh (strcat cJ rs) (strcat "=IF(N(" cLT rs ")=0,\"\",ROUNDUP(N(" cLT rs "),0))")))
        ;; гофра — с плана раздельно (лоток в неё не входит); опуски (итог - по плану) — к гофре
        ;; своего типа: потолок (1) или пусто — ПВХ, пол (2) — ПНД
        (dl:putf sh (strcat cK rs)
          (strcat "=IF(" cI rs "=\"\",\"\",ROUNDUP(N(" cPV rs "),0)+IF(" cS rs "=2,0,MAX(0,N(" cM rs ")-N(" cI rs "))))"))
        (dl:putf sh (strcat cL rs)
          (strcat "=IF(" cI rs "=\"\",\"\",ROUNDUP(N(" cPN rs "),0)+IF(" cS rs "=2,MAX(0,N(" cM rs ")-N(" cI rs ")),0))"))))
    (dl:put sh (strcat cPV rs) (nth 1 s))
    (dl:put sh (strcat cPN rs) (nth 2 s))
    (dl:put sh (strcat cLT rs) (nth 7 s))
    (dl:put sh (strcat cHN rs) (dl:join (nth 3 s) " "))
    (dl:put sh (strcat cNT rs) (nth 4 s))
    (dl:put sh (strcat cEX rs) (nth 6 s))
    (dl:put sh (strcat cSH rs) (dl:join (nth 8 s) ", "))
    ;; «Потолок/Пол»: при заполнении — если пусто; при обновлении — только у новых линий
    (setq sv (if (>= row *dl-row0*) (nth (- row *dl-row0*) svals) (nth (- row *dl-row-top*) tsvals)))
    (if (and (nth 3 s) (or isnew (and (= mode 'fill) (null sv))))
      (dl:put sh (strcat cS rs) (if (>= (nth 1 s) (nth 2 s)) 1 2))))
  ;; строки таблицы, которых нет на плане (в пределах тех же «префикс.уровень»): примечание + подсветка
  ;; линии щита бывают на разных листах: трасса жива в чертеже, но вне выбранных
  ;; видовых экранов — это другой лист, а не «нет на плане»
  (setq thands (if tnames (mapcar 'dl:str (dl:col-top sh cHN)))
        *dl-missing* (vl-remove-if '(lambda (n) (dl:elsewhere (dl:hands-of n names hands tnames thands) sum))
                                   (dl:missing *dl-tnames* sum))
        *dl-marked* *dl-missing* i *dl-row-top*)
  (foreach n tnames
    (if (member n *dl-missing*) (dl:put sh (strcat cNT (itoa i)) "нет на плане"))
    (setq i (1+ i)))
  (setq i *dl-row0*)
  (foreach n names
    (if (member n *dl-missing*) (dl:put sh (strcat cNT (itoa i)) "нет на плане"))
    (setq i (1+ i)))
  (setq *dl-added* nil)
  (foreach s sum
    (if (not (vl-some '(lambda (n) (= (dl:norm n) (dl:norm (car s)))) *dl-tnames*))
      (setq *dl-added* (cons (car s) *dl-added*))))
  ;; строки, чьи полилинии удалены из чертежа (другие листы не трогаем)
  (setq i *dl-row0*)
  (foreach n names
    (if (and (/= n "") (not (dl:insum n sum))
             (/= (nth (- i *dl-row0*) hands) "")
             (not (vl-some 'dl:alive (dl:split (nth (- i *dl-row0*) hands)))))
      (progn (dl:put sh (strcat cPV (itoa i)) nil) (dl:put sh (strcat cPN (itoa i)) nil)
             (dl:put sh (strcat cLT (itoa i)) nil)
             (dl:put sh (strcat cHN (itoa i)) nil)
             (dl:put sh (strcat cSH (itoa i)) nil)
             (dl:put sh (strcat cNT (itoa i)) "полилинии удалены из чертежа")
             (setq *dl-marked* (cons n *dl-marked*))))
    (setq i (1+ i)))
  (if *dl-v4*
    (if (> newn 0) (dl:v4-resize sh last))
    (progn (dl:sort sh cD) (dl:paint-all sh cD *dl-marked*)))
  (dl:restore-xl)
  (list upd newn mode))
(defun dl:restore-xl ()
  (if *dl-app*
    (progn (vl-catch-all-apply 'vlax-put-property (list *dl-app* 'ScreenUpdating :vlax-true))
           (vl-catch-all-apply 'vlax-put-property (list *dl-app* 'Calculation -4105)))))

;;; ---------- сортировка строк по номеру линии ----------
(setq *dl-order* '("Ввод" "KV" "На ИБП" "От ИБП" "R" "RС" "В" "ТП" "ЭK" "EK" "TE-ТП" "TE-O" "TE" "SH"
                   "L" "DALI" "LD" "LED" "АПС" "РЕЛЕ" "К" "K" "ПУ" "ПВ" "КМН" "УВ" "ПН" "WD" "КУП" "PE"))
(defun dl:pad (n w) (substr (strcat "00000" n) (1+ (- (+ 5 (strlen n)) w))))
;; "R.1.05.10" -> "05|R|0001.0005.0010"; пустая строка -> в конец
;; "R.0.001г.1" -> "05|R|00000.00001г.00001"; числа в частях номера сравниваются как числа,
;; буквы после цифр (г, д, Э, Т, E…) — после них; номер без точек (ОЗДС01) — в конец списка префиксов
(defun dl:splitc (s ch / i r)
  (while (setq i (vl-string-search ch s)) (setq r (cons (substr s 1 i) r) s (substr s (+ i 2))))
  (reverse (cons s r)))
(defun dl:segkey (seg / i d)
  (setq i 0 d "")
  (while (and (< i (strlen seg)) (wcmatch (substr seg (1+ i) 1) "#"))
    (setq d (strcat d (substr seg (1+ i) 1)) i (1+ i)))
  (strcat (dl:pad (if (= d "") "0" d) 5) (substr seg (1+ i))))
(defun dl:sortkey (name / segs pre idx)
  (if (= name "")
    "99|~"
    (progn
      (setq segs (dl:splitc name ".") pre (car segs) idx (vl-position pre *dl-order*))
      (strcat (dl:pad (itoa (if idx idx 98)) 2) "|" pre "|" (dl:join (mapcar 'dl:segkey (cdr segs)) ".")))))

;; раскладка: с *dl-row0* по *dl-block* линий, затем *dl-gap* пустых строк и т.д.
;; строки с номером линии получают место по порядку номеров, пустые строки
;; (пробелы и незанятые) заполняют промежутки; всё переставляется одной сортировкой Excel
(defun dl:slot (k) (+ (* (/ k *dl-block*) (+ *dl-block* *dl-gap*)) (rem k *dl-block*)))

(defun dl:isarray (sh a)
  (= (vl-catch-all-apply 'vlax-get-property (list (dl:rng sh a) 'HasArray)) :vlax-true))
;; записать в строку row значения/формулы (R1C1) колонок c1..c2
(defun dl:write-run (sh row c1 c2 vals / sa)
  (setq sa (vlax-make-safearray vlax-vbVariant (cons 1 1) (cons 1 (length vals))))
  (vlax-safearray-fill sa
    (list (mapcar '(lambda (v)
                     (cond ((null v) (vlax-make-variant ""))
                           ((and (= (type v) 'STR) (member (substr v 1 1) '("+" "-" "@")))
                            (vlax-make-variant (strcat "'" v)))
                           (T (vlax-make-variant v))))
                  vals)))
  (dl:try (strcat "перенос строки " (itoa row)) 'vlax-put-property
          (list (dl:rng sh (strcat (dl:colname c1) (itoa row) ":" (dl:colname c2) (itoa row))) (dl:f2prop) sa)))
;; Excel 365: Formula2R1C1 (без «@» в LET/XLOOKUP); старый Excel — FormulaR1C1
(defun dl:f2prop ()
  (if (null *dl-f2*)
    (setq *dl-f2* (if (vl-catch-all-error-p
                        (vl-catch-all-apply 'vlax-get-property (list (dl:rng *dl-sh-any* "A1") 'Formula2R1C1)))
                    'FormulaR1C1 'Formula2R1C1)))
  *dl-f2*)

;; перестановка строк без сортировки Excel: строки блока читаются (FormulaR1C1)
;; и переписываются на свои места; колонки с формулами массива не трогаются
(defun dl:sort (sh cD / ncols names i data used maxi npos k srcs fills block arr p src row c c1 vals)
  (setq *dl-sh-any* sh *dl-f2* nil
        ncols *dl-lastcol*
        names (mapcar 'dl:str (dl:col sh cD)) i 0 maxi -1)
  (foreach n names
    (if (/= n "") (setq data (cons (cons (dl:sortkey n) i) data) maxi i))
    (setq i (1+ i)))
  (if data
    (progn
      (setq data (vl-sort data '(lambda (a b) (< (car a) (car b)))) k 0)
      ;; (позиция . исходная_строка)
      (setq data (mapcar '(lambda (d) (setq k (1+ k)) (cons (dl:slot (1- k)) (cdr d))) data)
            used (mapcar 'cdr data)
            npos (1+ (max maxi (apply 'max (mapcar 'car data)))))
      ;; пустые строки по порядку заполняют свободные позиции
      (setq i 0 fills nil)
      (repeat npos (if (not (member i used)) (setq fills (cons i fills))) (setq i (1+ i)))
      (setq fills (reverse fills) p 0 srcs nil)
      (repeat npos
        (setq srcs (cons (if (assoc p data) (cdr (assoc p data)) (progn (setq src (car fills) fills (cdr fills)) src)) srcs)
              p (1+ p)))
      (setq srcs (reverse srcs))
      (setq p -1)
      (if (vl-some '(lambda (x) (/= x (setq p (1+ p)))) srcs)
        (progn
          (setq block (mapcar '(lambda (r) (mapcar 'dl:vv r))
                        (vlax-safearray->list
                          (dl:vv (dl:try "чтение строк" 'vlax-get-property
                                   (list (dl:rng sh (strcat "A" (itoa *dl-row0*) ":" (dl:colname ncols)
                                                            (itoa (+ *dl-row0* npos -1)))) (dl:f2prop))))))
                c 1 arr nil)
          (repeat ncols
            (if (dl:isarray sh (strcat (dl:colname c) (itoa *dl-row0*))) (setq arr (cons c arr)))
            (setq c (1+ c)))
          (setq p 0)
          (foreach src srcs
            (if (/= src p)
              (progn
                (setq row (+ *dl-row0* p) c 1 c1 nil vals nil)
                ;; куски подряд идущих колонок без формул массива
                (repeat ncols
                  (if (member c arr)
                    (progn (if c1 (dl:write-run sh row c1 (1- c) (reverse vals))) (setq c1 nil vals nil))
                    (progn (if (null c1) (setq c1 c)) (setq vals (cons (nth (1- c) (nth src block)) vals))))
                  (setq c (1+ c)))
                (if c1 (dl:write-run sh row c1 ncols (reverse vals)))))
            (setq p (1+ p))))))))

;;; ---------- проверка: пометки на чертеже ----------
(defun dl:doc () (vla-get-ActiveDocument (vlax-get-acad-object)))
(defun dl:chk-clear (/ ss i)
  (if (setq ss (ssget "_X" (list (cons 8 *dl-chk*))))
    (repeat (setq i (sslength ss)) (entdel (ssname ss (setq i (1- i))))))
  (setq *dl-marks* nil *dl-cur* nil)
  (vl-catch-all-apply 'vlax-ldata-delete (list "DLN" "marks"))
  (princ))
(defun dl:chk-layer (/ l)
  (if (not (tblsearch "LAYER" *dl-chk*))
    (progn
      (setq l (vla-Add (vla-get-Layers (dl:doc)) *dl-chk*))
      (vla-put-Color l 1)
      (vla-put-Plottable l :vlax-false))))
;; пометка: кружок + текст «[N] причина»
(defun dl:mark (pt txt / ms o)
  (setq ms (vla-get-ModelSpace (dl:doc))
        *dl-nmark* (1+ *dl-nmark*)
        txt (strcat "[" (itoa *dl-nmark*) "] " txt))
  (if (null pt) (progn (princ (strcat "\n  " txt " (без места на плане)")) (setq pt nil))
  (progn
  (setq o (vla-AddCircle ms (vlax-3d-point (dl:3d pt)) *dl-mark*))
  (vla-put-Layer o *dl-chk*) (vla-put-Color o 1)
  (setq o (vla-AddMText ms (vlax-3d-point (list (+ (car pt) *dl-mark*) (+ (cadr pt) *dl-mark*) 0.0))
                        (* *dl-txth* 25.0) txt))
  (vla-put-Height o *dl-txth*) (vla-put-Layer o *dl-chk*) (vla-put-Color o 1)
  (setq *dl-marks* (append *dl-marks* (list (list *dl-nmark* (list (car pt) (cadr pt)) txt))))
  (princ (strcat "\n  " txt)))))

(defun dl:polypt (e) (vlax-curve-getPointAtDist e (/ (vlax-curve-getDistAtParam e (vlax-curve-getEndParam e)) 2.0)))
(defun dl:starts-at-hub (p) (or (dl:hub (nth 5 p)) (dl:hub (nth 6 p))))
(defun dl:first-arrow (g lbls / r)
  (foreach hh *dl-hits* (if (and (null r) (= (car hh) g)) (setq r (caddr hh))))
  (if (null r) (foreach lb lbls (if (and (null r) (= (car lb) g)) (setq r (car (car (nth 2 lb)))))))
  r)

;; --- алфавит в номерах
(setq *dl-cyr* "АВЕКМНОРСТХУавекмнорстху" *dl-lat* "ABEKMHOPCTXYabekmhopctxy")
(defun dl:norm (s) (strcase (vl-string-translate *dl-cyr* *dl-lat* s)))
(defun dl:script (s / i c lat cyr)
  (setq i 0)
  (while (< i (strlen s))
    (setq i (1+ i) c (ascii (substr s i 1)))
    (cond ((or (<= 65 c 90) (<= 97 c 122)) (setq lat T))
          ((or (<= 1024 c 1279) (<= 192 c 255)) (setq cyr T))))
  (cond ((and lat cyr) "M") (lat "L") (cyr "C") (T "")))
;; префикс = буквы до первой точки или цифры: "R.0.001г.1" -> "R", "ОЗДС01" -> "ОЗДС"
(defun dl:prefix (g / i r c)
  (setq i 1 r "")
  (while (and (<= i (strlen g)) (/= (setq c (substr g i 1)) ".") (not (wcmatch c "#")))
    (setq r (strcat r c) i (1+ i)))
  r)
(defun dl:scope (g / i j)   ; "R.1.05.3" -> "R.1"
  (if (and (setq i (vl-string-search "." g)) (setq j (vl-string-search "." g (1+ i))))
    (substr g 1 j) g))
(defun dl:abc-issue (g / pre alt)
  (setq pre (dl:prefix g))
  (cond
    ((and *dl-codes* (member pre *dl-codes*)) nil)
    ((and *dl-codes* (setq alt (vl-some '(lambda (c) (if (= (dl:norm c) (dl:norm pre)) c)) *dl-codes*)))
     (strcat "в номере буквы другого алфавита: «" pre "», в Справочной «" alt "»"))
    ((= (dl:script pre) "M") (strcat "в номере смешаны кириллица и латиница: «" pre "»"))
    (*dl-codes* (strcat "префикса «" pre "» нет в Справочной"))))

;; --- таблица
(defun dl:read-codes (book / sh v r x rng)
  (setq sh (vl-catch-all-apply 'vlax-get-property (list (vlax-get-property book 'Worksheets) 'Item "Коды"))
        rng "A4:A1000")
  (if (vl-catch-all-error-p sh)
    (setq sh (vl-catch-all-apply 'vlax-get-property
               (list (vlax-get-property book 'Worksheets) 'Item "Справочная"))
          rng "P4:P200"))
  (if (not (vl-catch-all-error-p sh))
    (progn
      (setq v (vl-catch-all-apply 'vlax-get-property (list (vlax-get-property sh 'Range rng) 'Value2)))
      (if (not (vl-catch-all-error-p v))
        (foreach row (vlax-safearray->list (dl:vv v))
          (if (/= (setq x (dl:str (dl:vv (car row)))) "") (setq r (cons x r)))))))
  r)
;; ДСУП: точки РЕ от КУП — «Однолинейка», R (линия) и колонка «ДСУП: от КУП» (CK).
;; *dl-kup*: ((линия . КУП) ...) — трасса такой линии идёт от КУП, а не от щита.
(defun dl:read-kup (book / sh hdr i c v1 v2 k)
  (setq *dl-kup* nil
        sh (vl-catch-all-apply 'vlax-get-property (list (vlax-get-property book 'Worksheets) 'Item "Однолинейка")))
  (if (not (vl-catch-all-error-p sh))
    (progn
      (setq hdr (vl-catch-all-apply 'vlax-get-property (list (vlax-get-property sh 'Range "A2:DZ2") 'Value2)) i 1)
      (if (not (vl-catch-all-error-p hdr))
        (foreach h (car (vlax-safearray->list (dl:vv hdr)))
          (if (and (null c) (= (vl-string-search "ДСУП: ОТ КУП" (strcase (dl:str (dl:vv h)))) 0)) (setq c i))
          (setq i (1+ i))))
      (if c
        (progn
          (setq v1 (vl-catch-all-apply 'vlax-get-property (list (vlax-get-property sh 'Range "R3:R1000") 'Value2))
                v2 (vl-catch-all-apply 'vlax-get-property
                     (list (vlax-get-property sh 'Range (strcat (dl:colname c) "3:" (dl:colname c) "1000")) 'Value2)))
          (if (not (or (vl-catch-all-error-p v1) (vl-catch-all-error-p v2)))
            (mapcar '(lambda (a b)
                       (setq k (dl:str (dl:vv (car b))))
                       (if (and (/= k "") (/= (dl:str (dl:vv (car a))) ""))
                         (setq *dl-kup* (cons (cons (dl:str (dl:vv (car a))) k) *dl-kup*))))
                    (vlax-safearray->list (dl:vv v1)) (vlax-safearray->list (dl:vv v2))))))))
  *dl-kup*)
(defun dl:kup-of (g) (cdr (vl-some '(lambda (x) (if (= (dl:norm (car x)) (dl:norm g)) x)) *dl-kup*)))
;; точка pt у конца трассы КУП k
(defun dl:at-kup (pt k)
  (vl-some '(lambda (p) (and (= (length (dl:getg (car p))) 1)
                             (= (dl:norm (car (dl:getg (car p)))) (dl:norm k))
                             (<= (dl:enddist pt p) *dl-ptol*)))
           *dl-polys*))

(defun dl:insum (n sum)
  (or (assoc n sum) (vl-some '(lambda (s) (= (dl:norm (car s)) (dl:norm n))) sum)))
(defun dl:missing (names sum / scopes r)
  (foreach s sum (if (not (member (dl:scope (car s)) scopes)) (setq scopes (cons (dl:scope (car s)) scopes))))
  (foreach n names
    (if (and (/= n "") (not (dl:insum n sum)) (member (dl:scope n) scopes)
             (or (null *dl-filter*) (wcmatch (strcase n) (strcase *dl-filter*))))
      (setq r (cons n r))))
  (reverse r))
;; только прочитать таблицу (для DLN_CHK)
(defun dl:read-table (path / book sh)
  (setq book (vl-catch-all-apply 'dl:book (list path)))
  (if (not (vl-catch-all-error-p book))
    (progn
      (setq *dl-v4* (dl:v4-setup book)
            *dl-codes* (dl:read-codes book)
            *dl-kup* (dl:read-kup book)
            sh (vl-catch-all-apply 'vlax-get-property
                 (list (vlax-get-property book 'Worksheets) 'Item *dl-sheet*)))
      (if (not (vl-catch-all-error-p sh))
        (setq *dl-tnames* (mapcar 'dl:str
                            (mapcar '(lambda (row) (dl:vv (car row)))
                              (vlax-safearray->list
                                (dl:vv (vlax-get-property
                                         (vlax-get-property sh 'Range
                                           (strcat (dl:hcol-letter sh "Номер линии") (itoa *dl-row0*) ":"
                                                   (dl:hcol-letter sh "Номер линии") (itoa *dl-row1*)))
                                         'Value2))))))))))
(defun dl:hcol-letter (sh key / hdr i r)
  (setq hdr (mapcar '(lambda (x) (strcase (dl:str (dl:vv x))))
              (car (vlax-safearray->list (dl:vv (vlax-get-property (vlax-get-property sh 'Range "A2:DZ2") 'Value2)))))
        i 1)
  (foreach h hdr (if (and (null r) (= h (strcase key))) (setq r i)) (setq i (1+ i)))
  (dl:colname (if r r 4)))
;; путь к таблице: в чертеже (ldata), запасной — реестр
(defun dl:xls-path (ask / p)
  (setq p (vlax-ldata-get "DLN" "xls"))
  (if (or (null p) (not (findfile p))) (setq p (getenv "DLN_XLS")))
  (if (and ask p (/= p "") (findfile p))
    (progn
      (initget "Другая _Other")
      (if (= (getkword (strcat "\nDLN: таблица «" (vl-filename-base p) "». [Другая] <Enter — эта>: ")) "Other")
        (setq p (getfiled "Таблица однолинейки (xlsx / xlsm)" p "xlsx;xlsm" 0)))))
  (if (and ask (or (null p) (= p "") (not (findfile p))))
    (setq p (getfiled "Таблица однолинейки (xlsx / xlsm)" (getvar "DWGPREFIX") "xlsx;xlsm" 0)))
  (if (and p (/= p "") (findfile p))
    (progn (vlax-ldata-put "DLN" "xls" p) (setenv "DLN_XLS" p) p)))

;; --- щиты и стояки на плане: трасса с другого этажа идёт до стояка — это тоже «дошла»
(defun dl:panels (wins / ss i e bb res)
  (if (setq ss (ssget "_X" (list '(0 . "INSERT") (cons 8 *dl-panlay*) '(410 . "Model"))))
    (repeat (setq i (sslength ss))
      (setq e (ssname ss (setq i (1- i))))
      (if (setq bb (dl:bbox e)) (setq res (cons bb res)))))
  res)
(defun dl:at-panel (pt)
  (or (dl:hub pt) (vl-some '(lambda (bb) (<= (dl:rect-d pt bb) *dl-ptol*)) *dl-pan*)))

(defun dl:check (labels issues sum / systems gl hubs g own done gp ok best bd d e iss n k)
  (setq *dl-nmark* 0)
  (dl:chk-clear)
  (dl:chk-layer)
  (setq *dl-pan* (dl:panels nil))
  (princ "\nDLN: проверка выносок и трасс:")
  ;; 1. выноска мимо трассы / на трассе другой системы
  (foreach is issues (dl:mark (caddr is) (strcat (car is) ": " (cadr is))))
  ;; 2. на одной трассе выноски разных групп
  (foreach a *dl-asg*
    (if (> (length (setq gl (cdr a))) 1)
      (foreach hh *dl-hits*
        (if (eq (cadr hh) (car a))
          (dl:mark (caddr hh) (strcat (car hh) ": на этой трассе выноски разных групп ("
                                      (dl:join gl ", ") ") — длина не засчитана"))))))
  ;; 3. выноска стоит на трассе, отданной другой группе
  (foreach hh *dl-hits*
    (setq gl (dl:getg (cadr hh)))
    (if (and gl (not (member (car hh) gl)))
      (dl:mark (caddr hh) (strcat (car hh) ": выноска стоит на трассе группы " (dl:join gl ", ")))))
  ;; 4. несколько отдельных трасс от щита с одним номером
  (foreach s sum
    (setq g (car s) hubs 0)
    (foreach p *dl-polys* (if (and (equal (dl:getg (car p)) (list g)) (dl:starts-at-hub p)) (setq hubs (1+ hubs))))
    (if (> hubs 1)
      (dl:mark (dl:first-arrow g labels)
               (strcat g ": " (itoa hubs) " отдельные трассы от щита с одним номером — проверьте, не дубль ли"))))
  ;; 5. группа без длины
  (foreach s sum
    (if (and (null (nth 3 s)) (not (vl-some '(lambda (is) (= (car is) (car s))) issues))
             (not (vl-some '(lambda (a) (and (> (length (cdr a)) 1) (member (car s) (cdr a)))) *dl-asg*))
             (dl:first-arrow (car s) labels))
      (dl:mark (dl:first-arrow (car s) labels) (strcat (car s) ": длина не найдена"))))
  ;; 6. трасса группы не доходит до щита; точка РЕ от КУП (CK) — до конца трассы своей КУП
  (foreach s sum
    (if (nth 3 s)
      (progn
        (setq ok nil k (dl:kup-of (car s)))
        (foreach p *dl-polys*
          (if (and (equal (dl:getg (car p)) (list (car s)))
                   (if k
                     (or (dl:at-kup (nth 5 p) k) (dl:at-kup (nth 6 p) k))
                     (or (dl:at-panel (nth 5 p)) (dl:at-panel (nth 6 p)))))
            (setq ok T)))
        (if (not ok)
          (dl:mark (dl:first-arrow (car s) labels)
                   (strcat (car s) (if k (strcat ": трасса не доходит до трассы КУП " k
                                                 (if (dl:insum k sum) "" " (у КУП нет трассы на плане)"))
                                         ": трасса не доходит до щита или стояка")))))))
  ;; 7. буквы в номере
  (setq done nil)
  (foreach lb labels
    (if (and (not (member (car lb) done)) (setq iss (dl:abc-issue (car lb))))
      (progn (setq done (cons (car lb) done))
             (dl:mark (dl:first-arrow (car lb) labels) (strcat (car lb) ": " iss)))))
  ;; 8. трасса без выноски / разрыв трассы
  (foreach lb labels (if (not (member (dl:sys (nth 1 lb)) systems)) (setq systems (cons (dl:sys (nth 1 lb)) systems))))
  (foreach p *dl-polys*
    (if (and (null (dl:getg (car p))) (member (nth 1 p) systems))
      (progn
        (setq best nil bd 1e99)
        (foreach e (list (nth 5 p) (nth 6 p))
          (if (not (dl:at-panel e))
            (foreach p2 *dl-polys*
              (if (and (= (length (dl:getg (car p2))) 1) (= (nth 1 p2) (nth 1 p)))
                (foreach e2 (list (nth 5 p2) (nth 6 p2))
                  (setq d (distance (dl:2d e) (dl:2d e2)))
                  (if (and (> d *dl-jtol*) (<= d *dl-gtol*) (< d bd) (not (dl:at-panel e2)))
                    (setq bd d best (list e (car (dl:getg (car p2)))))))))))
        (if best
          (dl:mark (car best) (strcat "разрыв трассы " (rtos bd 2 0) " мм: похоже на продолжение " (cadr best)
                                      " (" (rtos (* *dl-scale* (nth 4 p)) 2 1) " м не засчитано)"))
          (dl:mark (dl:polypt (car p))
                   (strcat "трасса без выноски, " (rtos (* *dl-scale* (nth 4 p)) 2 1) " м (" (dl:hnd (car p)) ")"))))))
  ;; 9. сверка с таблицей
  (if *dl-tnames*
    (progn
      (setq n *dl-missing*)
      (if n (princ (strcat "\n  в таблице есть, на плане нет (" (itoa (length n)) "): " (dl:join n ", ")
                           (if *dl-chk-only* "\n  (линии щита на других листах — сверка по всем листам: DLN_SVR)" ""))))
      (foreach rn *dl-renamed*
        (dl:mark (dl:first-arrow (car rn) labels)
                 (strcat (car rn) ": в таблице «" (cadr rn) "» — те же знаки, но другие буквы (кириллица/латиница); записано в эту строку")))
      (setq n nil)
      (foreach s sum
        (if (not (vl-some '(lambda (tn) (= (dl:norm tn) (dl:norm (car s)))) *dl-tnames*)) (setq n (cons (car s) n))))
      (if n (princ (strcat "\n  на плане есть, в таблице не было — добавлены (" (itoa (length n)) "): " (dl:join (reverse n) ", "))))))
  (vl-catch-all-apply 'vlax-ldata-put (list "DLN" "marks" *dl-marks*))
  (vl-catch-all-apply 'vlax-ldata-put (list "DLN" "vps" *dl-vps*))
  (if (= *dl-nmark* 0)
    (princ " замечаний нет.")
    (princ (strcat "\n  пометок на чертеже: " (itoa *dl-nmark*)
                   " — DLN_NEXT / DLN_PREV: переход по пометкам, DLN_CLR: убрать"))))

;;; ---------- навигатор по пометкам ----------
(defun dl:inbox (pt w) (and (< (car w) (car pt) (caddr w)) (< (cadr w) (cadr pt) (cadddr w))))
(defun dl:goto (m / pt w vp tab ps h)
  (setq pt (cadr m) tab (getvar "CTAB"))
  ;; ВЭ на текущем листе, иначе любой ВЭ с этой точкой
  (foreach w *dl-vps* (if (and (null vp) (dl:inbox pt w) (= (strcase (nth 5 (nth 5 w))) (strcase tab))) (setq vp w)))
  (if (and (null vp) (/= (strcase tab) "MODEL"))
    (foreach w *dl-vps* (if (and (null vp) (dl:inbox pt w) (/= (strcase (nth 5 (nth 5 w))) "MODEL")) (setq vp w))))
  (if (and vp (/= (strcase tab) "MODEL"))
    (progn
      (setq w (nth 5 vp))
      (if (/= (strcase (nth 5 w)) (strcase tab)) (setvar "CTAB" (nth 5 w)))
      (vla-put-MSpace (dl:doc) :vlax-false)   ; в пространство листа — масштаб ВЭ не меняется
      (setq ps (list (+ (car w) (* (- (car pt) (caddr w)) (nth 4 w)))
                     (+ (cadr w) (* (- (cadr pt) (cadddr w)) (nth 4 w))) 0.0)
            h (* 14.0 *dl-mark* (nth 4 w)))
      (command "_.ZOOM" "_C" ps h))
    (progn
      (if (/= (getvar "TILEMODE") 1) (setvar "TILEMODE" 1))
      (command "_.ZOOM" "_C" (list (car pt) (cadr pt) 0.0) (* 14.0 *dl-mark*))))
  (princ (strcat "\n" (itoa (car m)) "/" (itoa (length *dl-marks*)) "  " (caddr m))))
(defun dl:nav (step)
  (if (null *dl-marks*) (setq *dl-marks* (vlax-ldata-get "DLN" "marks") *dl-vps* (vlax-ldata-get "DLN" "vps")))
  (if (null *dl-marks*)
    (princ "\nDLN: пометок нет — сначала DLN или DLN_CHK.")
    (progn
      (setq *dl-cur* (if *dl-cur* (+ *dl-cur* step) (if (> step 0) 0 (1- (length *dl-marks*)))))
      (if (>= *dl-cur* (length *dl-marks*)) (setq *dl-cur* 0))
      (if (< *dl-cur* 0) (setq *dl-cur* (1- (length *dl-marks*))))
      (dl:goto (nth *dl-cur* *dl-marks*))))
  (princ))
(defun c:DLN_NEXT () (dl:nav 1))
(defun c:DLN_PREV () (dl:nav -1))

;;; ---------- команды ----------
(defun c:DLN () (dl:run 'fill))
(defun c:DLN_UPD () (dl:run 'upd))
(defun dl:run (mode / *error* path wins labels issues sum res bad ss)
  (defun *error* (m)
    (dl:restore-xl)
    (if (not (wcmatch (strcase m) "*CANCEL*,*QUIT*,*EXIT*")) (princ (strcat "\nDLN: " m)))
    (princ))
  (setq *dl-app* nil)
  (setq path (dl:xls-path T))
  (if (null path)
    (princ "\nDLN: таблица не выбрана.")
    (progn
      (setq wins (dl:pick-wins) *dl-vps* wins)
      (if (null wins) (progn (princ "\nDLN: не найден ни один видовой экран плана.") (exit)))
      (setq *dl-filter* (getstring "\nDLN: какие линии брать, напр. R* или L*,LD*,LED* <все>: "))
      (if (= *dl-filter* "") (setq *dl-filter* nil))
      (princ "\nDLN: читаю чертёж...")
      (setq wins wins
            *dl-polys* (dl:polys wins)
            labels (dl:labels wins)
            issues (dl:match labels)
            *dl-cnt* (dl:count-blocks wins)
            sum (dl:summary labels))
      (foreach w wins (princ (strcat "\n  окно плана: X " (rtos (car w) 2 0) ".." (rtos (caddr w) 2 0)
                                        "  Y " (rtos (cadr w) 2 0) ".." (rtos (cadddr w) 2 0)
                                        "  заморожено слоёв в ВЭ: " (itoa (length (nth 4 w))))))
      (princ (strcat "\n  окон планов: " (itoa (length wins))
                     "  трасс: " (itoa (length *dl-polys*))
                     "  лотков: " (itoa (length *dl-trays*))
                     "  выносок: " (itoa (length labels))
                     "  линий: " (itoa (length sum))))
      (setq res (dl:write path sum mode))
      (princ (strcat "\n  режим: " (if (= (caddr res) 'upd) "обновить" "заполнить")
                     "  обновлено строк: " (itoa (car res)) ", добавлено: " (itoa (cadr res))
                     "  (формулы: " (cond ((= *dl-fmode* 1) "Formula") ((= *dl-fmode* 2) "FormulaLocal") ((= *dl-fmode* 3) "Value2") (T "-")) ")"))
      (dl:check labels issues sum)
      ;; выделить полилинии без группы или с конфликтом
      (setq ss (ssadd))
      (foreach a *dl-asg* (if (/= (length (cdr a)) 1) (ssadd (car a) ss)))
      (if (> (sslength ss) 0)
        (progn (sssetfirst nil ss)
               (princ (strcat "\n  выделено проблемных полилиний: " (itoa (sslength ss))))))
      (princ "\nDLN: готово. Таблица открыта в Excel, сохраните её после проверки.")))
  (princ))

(defun c:DLN_XLS () (vlax-ldata-delete "DLN" "xls") (setenv "DLN_XLS" "") (princ "\nDLN: путь к таблице сброшен.") (princ))
(defun c:DLN_CLR () (dl:chk-clear) (princ "\nDLN: пометки проверки удалены.") (princ))
;; только проверка, без записи в Excel
(defun c:DLN_CHK (/ wins labels issues path sum)
  (setq wins (dl:pick-wins) *dl-vps* wins *dl-tnames* nil *dl-codes* nil *dl-renamed* nil *dl-kup* nil)
  (if (null wins) (progn (princ "\nDLN: не найден ни один видовой экран плана.") (exit)))
  (setq *dl-filter* (getstring "\nDLN: какие линии проверять, напр. R* или L*,LD*,LED* <все>: "))
  (if (= *dl-filter* "") (setq *dl-filter* nil))
  (setq *dl-polys* (dl:polys wins) labels (dl:labels wins) issues (dl:match labels)
        *dl-cnt* nil)
  ;; таблицу только читаем — для сверки номеров и Справочной
  (if (setq path (dl:xls-path nil)) (dl:read-table path))
  (foreach lb labels
    (if (and *dl-tnames* (not (member (car lb) *dl-tnames*)) (not (assoc (car lb) *dl-renamed*)))
      (foreach tn *dl-tnames*
        (if (and (/= tn "") (= (dl:norm tn) (dl:norm (car lb))) (not (assoc (car lb) *dl-renamed*)))
          (setq *dl-renamed* (cons (list (car lb) tn) *dl-renamed*))))))
  (setq sum (dl:summary labels)
        *dl-missing* (if *dl-tnames* (dl:missing *dl-tnames* sum))
        *dl-chk-only* T)
  (dl:check labels issues sum)
  (setq *dl-chk-only* nil)
  (princ))

;;; ---------- DLN_SVR: сверка плана и таблицы щита ----------
;; Линии одного щита бывают на разных листах (этажи, трасса до стояка), поэтому
;; сверка всегда идёт по всем листам чертежа. Таблица только читается; итог —
;; лист «Сверка» в конце книги и пометки на плане. Книга не сохраняется.
(setq *dl-svr-pri* '(("нет на плане" . 1) ("нет в таблице" . 2) ("длина изменилась" . 3)
                     ("другие буквы" . 4) ("трасса без выноски" . 5) ("нет длины на плане" . 6)
                     ("номер повторяется" . 7) ("длина не записана" . 8) ("потребителей другое число" . 9)
                     ("нет в таблице, другой щит?" . 10) ("не проверена" . 11) ("совпадает" . 12)))
(defun dl:read-hdrs (sh / i)
  (setq *dl-hdrs* (mapcar '(lambda (x) (strcase (dl:str (dl:vv x))))
                          (car (dl:try "чтение заголовков" 'vlax-safearray->list
                                 (list (dl:vv (dl:try "заголовки" 'vlax-get-property
                                                (list (dl:rng sh "A2:DZ2") 'Value2)))))))
        *dl-lastcol* 0 i 1)
  (foreach h *dl-hdrs* (if (/= h "") (setq *dl-lastcol* i)) (setq i (1+ i))))
(defun dl:hopt (key mode / r) (if (setq r (dl:hcol key mode)) (dl:colname r)))
;; колонка целиком: строки выше *dl-row0* и основные; индекс k -> строка *dl-row-top* + k
(defun dl:colall (sh c) (if c (append (dl:col-top sh c) (dl:col sh c))))
(defun dl:r1 (x) (if (numberp x) (/ (fix (+ (* x 10.0) 0.5)) 10.0)))
(defun dl:nz (x) (if (numberp x) x 0.0))
(defun dl:svr-sheet (book / shs n last sh)
  (setq shs (vlax-get-property book 'Worksheets)
        sh (vl-catch-all-apply 'vlax-get-property (list shs 'Item "Сверка")))
  (if (vl-catch-all-error-p sh)
    (progn
      (setq n (vlax-get-property shs 'Count) last (vlax-get-property shs 'Item n)
            sh (dl:try "новый лист «Сверка»" 'vlax-invoke-method (list shs 'Add)))
      ;; только аргумент Before: новый перед последним, затем последний перед новым
      (vlax-invoke-method sh 'Move last)
      (vl-catch-all-apply 'vlax-invoke-method (list last 'Move sh))
      (vlax-put-property sh 'Name "Сверка"))
    (dl:try "очистка листа «Сверка»" 'vlax-invoke-method (list (vlax-get-property sh 'Cells) 'Clear)))
  sh)
;; записать строки rows (списки одной длины) начиная с ячейки A<r0>
(defun dl:put-block (sh r0 rows / nc sa i j)
  (setq nc (length (car rows))
        sa (vlax-make-safearray vlax-vbVariant (cons 0 (1- (length rows))) (cons 0 (1- nc))) i 0)
  (foreach r rows
    (setq j 0)
    (foreach v r
      (vlax-safearray-put-element sa i j
        (cond ((null v) "")
              ((and (= (type v) 'STR) (member (substr v 1 1) '("=" "+" "-" "@"))) (strcat "'" v))
              (T v)))
      (setq j (1+ j)))
    (setq i (1+ i)))
  (dl:try "запись листа «Сверка»" 'vlax-put-property
          (list (dl:rng sh (strcat "A" (itoa r0) ":" (dl:colname nc) (itoa (+ r0 (length rows) -1)))) 'Value2 sa)))

(defun c:DLN_SVR (/ *error* path wins labels issues sum book sh cD cG cPV cPN cLT cHN
                    tn tg tpv tpn tlt thn n k s old new st note out tsc psc used tabs pc
                    nprob cnt ok)
  (defun *error* (m)
    (if (not (wcmatch (strcase m) "*CANCEL*,*QUIT*,*EXIT*")) (princ (strcat "\nDLN: " m)))
    (princ))
  (setq path (dl:xls-path T))
  (if (null path) (progn (princ "\nDLN: таблица не выбрана.") (exit)))
  (setq wins (dl:all-wins) *dl-vps* wins)
  (if (null wins) (progn (princ "\nDLN: не найден ни один видовой экран плана.") (exit)))
  (setq *dl-filter* (getstring "\nDLN: какие линии сверять, напр. R* или L*,LD*,LED* <все>: "))
  (if (= *dl-filter* "") (setq *dl-filter* nil))
  (princ (strcat "\nDLN: сверка по всем листам, видовых экранов: " (itoa (length wins)) "..."))
  (setq *dl-polys* (dl:polys wins) labels (dl:labels wins) issues (dl:match labels)
        *dl-cnt* (dl:count-blocks wins) sum (dl:summary labels))
  ;; таблица — только чтение
  (setq book (dl:book path)
        *dl-v4* (dl:v4-setup book)
        sh (dl:try (strcat "лист " *dl-sheet*) 'vlax-get-property
                   (list (dl:try "листы книги" 'vlax-get-property (list book 'Worksheets)) 'Item *dl-sheet*)))
  (dl:read-hdrs sh)
  (setq cD (dl:hreq "Номер линии" 0)
        cG (dl:hopt "Кол-во потребителей" 1)
        cPV (dl:hopt "DWG: гофра ПВХ" 1) cPN (dl:hopt "DWG: гофра ПНД" 1)
        cLT (dl:hopt "DWG: лоток" 1) cHN (dl:hopt "DWG: хэндлы" 1)
        tn (mapcar 'dl:str (dl:colall sh cD))
        tg (dl:colall sh cG) tpv (dl:colall sh cPV) tpn (dl:colall sh cPN) tlt (dl:colall sh cLT)
        thn (mapcar 'dl:str (dl:colall sh cHN)))
  (foreach x sum (if (not (member (dl:scope (car x)) psc)) (setq psc (cons (dl:scope (car x)) psc))))
  (foreach x tn (if (and (/= x "") (not (member (dl:scope x) tsc))) (setq tsc (cons (dl:scope x) tsc))))
  ;; пометки на плане — заново
  (setq *dl-nmark* 0)
  (dl:chk-clear)
  (dl:chk-layer)
  ;; 1. строки таблицы
  (setq k 0)
  (foreach n tn
    (if (and (/= n "") (or (null *dl-filter*) (wcmatch (strcase n) (strcase *dl-filter*))))
      (progn
        (setq s (cond ((assoc n sum))
                      ((vl-some '(lambda (x) (if (= (dl:norm (car x)) (dl:norm n)) x)) sum)))
              old nil new nil tabs "" pc nil note "")
        (if s
          (progn
            (setq used (cons (car s) used)
                  new (+ (nth 1 s) (nth 2 s) (nth 7 s))
                  tabs (dl:join (nth 8 s) ", ")
                  pc (nth 5 s))
            (if (or (numberp (nth k tpv)) (numberp (nth k tpn)) (numberp (nth k tlt)))
              (setq old (+ (dl:nz (nth k tpv)) (dl:nz (nth k tpn)) (dl:nz (nth k tlt)))))
            (cond
              ((/= (car s) n)
               (setq st "другие буквы" note (strcat "на плане «" (car s) "»: кириллица / латиница"))
               (dl:mark (dl:first-arrow (car s) labels) (strcat (car s) ": в таблице «" n "» — другие буквы")))
              ((null (nth 3 s)) (setq st "нет длины на плане" note (nth 4 s)))
              ((null old) (setq st "длина не записана" note "запустите DLN_UPD"))
              ((>= (abs (- new old)) 0.5)
               (setq st "длина изменилась" note "запустите DLN_UPD")
               (dl:mark (dl:first-arrow n labels)
                        (strcat n ": длина изменилась, было " (rtos old 2 1) " м, стало " (rtos new 2 1) " м")))
              ((and (numberp pc) (numberp (nth k tg)) (/= (fix (nth k tg)) pc))
               (setq st "потребителей другое число" note "проверьте блоки на плане"))
              (T (setq st "совпадает"))))
          (cond
            ((dl:elsewhere (nth k thn) sum)
             (setq st "трасса без выноски" note "трасса в чертеже есть, выноски с номером нет"))
            ((member (dl:scope n) psc) (setq st "нет на плане" note "выноски нет ни на одном листе"))
            (T (setq st "не проверена" note "этого префикса и этажа нет на листах этого DWG"))))
        (if (and (= st "совпадает") (> (length (vl-remove-if-not '(lambda (x) (= x n)) tn)) 1))
          (setq st "номер повторяется" note "номер в таблице дважды"))
        (setq out (cons (list n st tabs (dl:r1 old) (dl:r1 new) (nth k tg) pc (+ *dl-row-top* k) note) out))))
    (setq k (1+ k)))
  ;; 2. линии плана, которых нет в таблице
  (foreach s sum
    (if (not (member (car s) used))
      (progn
        (if (member (dl:scope (car s)) tsc)
          (progn (setq st "нет в таблице" note "новая линия этого щита? DLN_UPD добавит")
                 (dl:mark (dl:first-arrow (car s) labels) (strcat (car s) ": на плане есть, в таблице нет")))
          (setq st "нет в таблице, другой щит?" note "префикс и этаж не из этой таблицы"))
        (setq out (cons (list (car s) st (dl:join (nth 8 s) ", ") nil
                              (dl:r1 (+ (nth 1 s) (nth 2 s) (nth 7 s))) nil (nth 5 s) nil note) out)))))
  ;; порядок: сначала расхождения, внутри — по номеру линии
  (setq out (vl-sort out '(lambda (a b / pa pb)
                            (setq pa (cdr (assoc (cadr a) *dl-svr-pri*)) pb (cdr (assoc (cadr b) *dl-svr-pri*)))
                            (if (= pa pb) (< (dl:sortkey (car a)) (dl:sortkey (car b))) (< pa pb)))))
  ;; лист «Сверка»
  (setq sh (dl:svr-sheet book))
  (dl:put-block sh 1 (list (list "Сверка плана и таблицы щита")
                           (list (strcat "DWG: " (getvar "DWGNAME") ";  листов с планами: "
                                         (itoa (length (dl:uniq (mapcar '(lambda (w) (nth 5 (nth 5 w))) wins))))
                                         ";  " (menucmd "M=$(edtime,$(getvar,date),DD.MO.YYYY HH:MM)")
                                         (if *dl-filter* (strcat ";  фильтр: " *dl-filter*) "")))))
  (dl:put-block sh 4 (list (list "Номер линии" "Статус" "Листы DWG" "Длина в таблице, м" "Длина на плане, м"
                                 "Потребителей в таблице" "Потребителей на плане" "Строка таблицы" "Что сделать")))
  (if out (dl:put-block sh 5 out))
  (vl-catch-all-apply 'vlax-put-property (list (vlax-get-property (dl:rng sh "A1") 'Font) 'Bold :vlax-true))
  (vl-catch-all-apply 'vlax-put-property (list (vlax-get-property (dl:rng sh "A4:I4") 'Font) 'Bold :vlax-true))
  (setq k 5 nprob 0 cnt nil)
  (foreach r out
    (setq ok (member (cadr r) '("совпадает" "не проверена" "нет в таблице, другой щит?"))
          cnt (dl:inc (cadr r) cnt))
    (if (not ok) (progn (setq nprob (1+ nprob)) (dl:paint sh (strcat "B" (itoa k)) T)))
    (setq k (1+ k)))
  (vl-catch-all-apply 'vlax-invoke-method (list (vlax-get-property (dl:rng sh "A4:I4") 'EntireColumn) 'AutoFit))
  (vl-catch-all-apply 'vlax-invoke-method (list book 'Activate))
  (vl-catch-all-apply 'vlax-invoke-method (list sh 'Activate))
  (vl-catch-all-apply 'vlax-ldata-put (list "DLN" "marks" *dl-marks*))
  (vl-catch-all-apply 'vlax-ldata-put (list "DLN" "vps" *dl-vps*))
  ;; итог в командной строке
  (princ (strcat "\nDLN: сверка — линий в отчёте: " (itoa (length out)) ", расхождений: " (itoa nprob)))
  (foreach p (vl-sort cnt '(lambda (a b) (< (cdr (assoc (car a) *dl-svr-pri*)) (cdr (assoc (car b) *dl-svr-pri*)))))
    (princ (strcat "\n  " (car p) ": " (itoa (cdr p)))))
  (if (> *dl-nmark* 0)
    (princ (strcat "\n  пометок на плане: " (itoa *dl-nmark*) " — DLN_NEXT / DLN_PREV, убрать: DLN_CLR")))
  (princ "\nDLN: лист «Сверка» открыт в Excel. Книга не сохранена.")
  (princ))

(princ "\nDLN v17 загружен. Команды: DLN (заполнить), DLN_UPD (обновить после правок планов), DLN_SVR (сверка по всем листам), DLN_CHK (только проверка), DLN_NEXT / DLN_PREV (по пометкам), DLN_CLR (убрать пометки), DLN_XLS (сменить таблицу)")
(princ)
