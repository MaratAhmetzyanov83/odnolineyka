;;; ============================================================================
;;;  SX_Schema.lsp  v3.13 —  однолинейная схема из Excel на ВАШИХ блоках
;;;
;;;  Команды:
;;;    SXDRAW    — окно: файл Excel, лист «Однолинейка», щит, точка вставки,
;;;                библиотека блоков (DWG) -> строит схему целиком.
;;;    SXCLEAR   — удалить всё построенное (слои SX_*).
;;;    SXBLOCKS  — пересоздать служебные блоки SX_* (столбец таблицы, канал).
;;;
;;;  Используются блоки из вашего чертежа 02_Однолинейка:
;;;    QF1 (динамический: Видимость1 = QF / QF_L+N / QFD, Расстояние1 — точка на шину),
;;;    Шина (динамический: длина), XT, БП (трансформатор), Драйвер (блок
;;;    питания), MK-5-1, Розетка1, лампа ($RECOVER_230831092053-0),
;;;    Сервопривод-1, EK.1, Привод, _Боковик_схема Юником.
;;;  Если блока нет в текущем чертеже — он копируется из библиотеки
;;;  (ваш DWG, путь задаётся в окне). Если нет и там — рисуется простой
;;;  заменитель SX_*.
;;;  Размещение (уровни, масштабы, повороты) снято с вашего чертежа.
;;;
;;;  Все строки-литералы в коде — латиница (работает при LISPSYS 0 и 1),
;;;  русские имена блоков и надписи хранятся кодами Unicode (sx:ru).
;;; ============================================================================

(vl-load-com)
(setq sx:ver "v3.13")


;;; ---------------------------------------------------------------- настройки
(if (null *sx-sheet*) (setq *sx-sheet* 17))   ; мест (столбцов) на один лист A3
(if (null *sx-step*)  (setq *sx-step* 20.0))  ; шаг столбца, мм
(setq *sx-range* "A3:BY1000")                  ; данные на листе «Однолинейка» (BY = «УГО (рус)», нужен для «кабельного вывода»)
(setq *sx-style* "SX_GOST")                    ; текстовый стиль схемы (создаётся/обновляется при каждом запуске)
(setq *sx-face* "GOST 2.304 type A")           ; шрифт (имя гарнитуры, как в окне «Текстовые стили»)
(setq sx:red-rgb 10033950 sx:red-aci 244)       ; тёмно-красный 153,27,30 (как БП в вашем DWG)
(setq sx:ltN "JIS_02_0.7" sx:ltsN 4.0 sx:ltsB 5.0)  ; границы щита: JIS_02_0.7, масштаб 5         ; нейтраль: тип и масштаб линии как в DWG

;;; ---- ПОДСТРОЙКА ВЫСОТ (мм по вертикали, "-" = вниз) -----------------------------------------
;;; sx:dTop - сдвиг верхней штриховой границы щита и подписи щита от края рамки листа (v = 232).
;;;           -7.83 даёт 224.17, как в вашем чертеже 02_Однолинейка (граница на 7.8 мм ниже кромки рамки).
;;;           Хотите ниже: -10.0, -12.0 и т.д. (шины и автоматы не сдвигаются).
;;; sx:dLow - сдвиг НИЖНЕЙ части схемы вниз (было 0.0): нижняя граница щита, клеммы XT, шина PE,
;;;           засечки на линиях, соединения "в кабель выше". Автоматы, шины, модули и каналы стоят как были.
;;;           Предел: нижняя граница щита не ниже v = 106 (ячейка кабеля таблицы занимает 75..105),
;;;           т.е. sx:dLow не меньше -8.0. Хотите как раньше: 0.0.
(setq sx:dTop -7.83 sx:dLow -7.0)

;;; уровни по высоте, мм от низа таблицы (сняты с вашего чертежа; нижняя часть смещена на sx:dLow)
(setq sx:vTab 105.0 sx:vBottom (+ 114.5 sx:dLow) sx:vPE (+ 125.6 sx:dLow) sx:vJoin (+ 128.5 sx:dLow)
      sx:vXTtop (+ 137.22 sx:dLow) sx:vXTN (+ 134.22 sx:dLow) sx:vXTPE (+ 132.22 sx:dLow) sx:vXTbot (+ 131.22 sx:dLow)
      sx:vBridge (+ 138.3 sx:dLow) sx:vChBase 139.26 sx:vIN 158.1
      sx:vBoxBot 145.91 sx:vBoxTop 155.91 sx:vKNX 161.0 sx:vComb 162.96
      sx:vQFins 217.24 sx:vQFcontact 36.3 sx:vQFlen 54.3
      sx:vBusL 216.05 sx:vBusN 211.05 sx:vBus2L 205.22 sx:vBus2N 200.22
      sx:vIN0 162.96 sx:qfBody 54.31 sx:nDx -8.36
      ;; аппарат (БП/трансформатор) сбоку от автомата, как QF46 в 02_Однолинейка
      sx:vSideB 172.25 sx:vSideRet 168.53 sx:sideLead 5.78 sx:sideX1 5.64 sx:sideX2 17.7
      sx:vTop (+ 232.0 sx:dTop) sx:vHeadDx -55.0
      sx:vTickQ (/ (+ 162.96 sx:vXTtop) 2.0) sx:vTickLine (+ 119.5 sx:dLow)   ; засечки: на спуске от автомата (посередине) / на линии канала
      sx:vOutTop 73.0)   ; кабельный вывод в строке УГО: верх символа (ячейка УГО 62..75)
(setq sx:xtb sx:vXTbot)   ; низ клеммы текущей строки (у двухуровневой клеммы выше на 2 мм)

;;; столбцы листа «Однолинейка» (A = 1)
(setq sx:cPanel 1 sx:cPos 2 sx:cBusName 4 sx:cQFtxt 8 sx:cDev1 9
      sx:cCab 17 sx:cLine 18 sx:cName 19 sx:cPow 20 sx:cI 21 sx:cPh 22
      sx:cGrp 33 sx:cMod 34 sx:cModel 35 sx:cQFown 36 sx:cCh 15 sx:cXTtxt 16
      sx:cJoin 40 sx:cXT 51
      sx:cDraw 56 sx:cBus2 57 sx:cRcd 58 sx:cChT 59 sx:cDevC1 60 sx:cSym 63 sx:cHasLine 64
      sx:cUgoRu 77)   ; BY: название УГО по-русски (пусто = у линии нет кода УГО)
(setq sx:cStat 31 sx:rowBase 3)   ; AE = status of the row; first Excel row of the range A3:BL


;;; ---------------------------------------------------------------- утилиты
(defun sx:v (x)
  (cond ((= (type x) 'VARIANT)
         (if (member (vlax-variant-type x) '(0 1 10)) nil (sx:v (vlax-variant-value x))))
        ((and (= (type x) 'STR) (= x "")) nil)
        (t x)))

(defun sx:s (x)
  (cond ((null x) "")
        ((= (type x) 'STR) x)
        ((= (type x) 'INT) (itoa x))
        ((= (type x) 'REAL) (if (equal x (fix x) 1e-9) (itoa (fix x)) (rtos x 2 2)))
        (t "")))

(defun sx:num (x) (vl-string-translate "." "," (sx:s x)))   ; число с запятой

(defun sx:n (x)
  (cond ((numberp x) x)
        ((and (= (type x) 'STR) (distof x 2)) (distof x 2))
        (t 0)))

(defun sx:split (s / res w ch k)
  (setq res '() w "" k 1)
  (while (<= k (strlen s))
    (setq ch (substr s k 1))
    (if (= ch " ")
      (progn (if (/= w "") (setq res (cons w res))) (setq w ""))
      (setq w (strcat w ch)))
    (setq k (1+ k)))
  (if (/= w "") (setq res (cons w res)))
  (reverse res))

;;; разбить текст на строки не длиннее n символов (по словам), не больше m строк
(defun sx:wrap (s n m / words lines cur)
  (setq words (sx:split s) lines '() cur "")
  (foreach w words
    (cond ((= cur "") (setq cur w))
          ((<= (+ (strlen cur) 1 (strlen w)) n) (setq cur (strcat cur " " w)))
          (t (setq lines (append lines (list cur)) cur w))))
  (if (/= cur "") (setq lines (append lines (list cur))))
  (if (> (length lines) m)
    (setq lines (append (sx:head lines (1- m)) (list (sx:join (sx:tail lines (1- m)))))))
  lines)
(defun sx:tail (lst k) (repeat k (setq lst (cdr lst))) lst)
(defun sx:head (lst k / res) (repeat k (setq res (append res (list (car lst))) lst (cdr lst))) res)
(defun sx:join (lst / res)
  (setq res "")
  (foreach x lst (setq res (if (= res "") x (strcat res " " x))))
  res)

;;; rec = (i u row)
(defun sx:c (rec col) (nth (1- col) (caddr rec)))
(defun sx:u (rec) (cadr rec))
(defun sx:i (rec) (car rec))
(defun sx:xr (rec) (nth 3 rec))   ; Excel row number of the record
(defun sx:is1 (rec col) (= (sx:n (sx:c rec col)) 1))
(defun sx:sheet (i) (/ i *sx-sheet*))
(defun sx:nth0 (k lst) (if (and lst (< k (length lst))) (nth k lst) ""))
(defun sx:has (rec col) (/= (sx:s (sx:c rec col)) ""))

;;; ---------------------------------------------------------------- журнал / отладка
(setq sx:logl nil sx:errs nil sx:stage "start")
(defun sx:str (x) (vl-prin1-to-string x))
(defun sx:log (s) (setq sx:logl (cons s sx:logl)) s)
(defun sx:stg (s) (setq sx:stage s) (sx:log (strcat "== " s)))
;;; выполнить fn; ошибка не прерывает отрисовку, а попадает в журнал
(defun sx:try (fn label / e)
  (setq sx:stage label)
  (setq e (vl-catch-all-apply fn nil))
  (if (vl-catch-all-error-p e)
    (progn
      (setq e (strcat label ": " (vl-catch-all-error-message e)))
      (setq sx:errs (cons e sx:errs))
      (sx:log (strcat "ERROR " e))
      nil)
    e))
(defun sx:logdir ( / f)
  (cond ((setq f (findfile "SX_Schema.lsp")) (vl-filename-directory f))
        ((and sx:path (/= sx:path "")) (vl-filename-directory sx:path))
        (t (getvar "DWGPREFIX"))))
(defun sx:logdump ( / fn f)
  (setq fn (strcat (sx:logdir) "\\SX_log.txt"))
  (if (setq f (open fn "w"))
    (progn
      (foreach s (reverse sx:logl) (write-line s f))
      (close f)
      (princ (strcat "\nLog: " fn)))))
(defun sx:rowinfo (d)
  (strcat "place " (itoa (1+ (sx:i d))) " row=" (sx:str (sx:head (caddr d) 22))
          " helpers=" (sx:str (sx:head (sx:tail (caddr d) 32) 32))))

;;; ---------------------------------------------------------------- русский текст
(defun sx:cp (c)
  (cond ((< c 128) c)
        ((and (getvar "LISPSYS") (>= (getvar "LISPSYS") 1)) c)
        ((and (>= c 1040) (<= c 1103)) (+ (- c 1040) 192))
        ((= c 1025) 168) ((= c 1105) 184) ((= c 8212) 151) ((= c 8230) 133)
        ((= c 171) 171) ((= c 187) 187) ((= c 8470) 185)
        (t 63)))
(defun sx:ru (codes) (apply 'strcat (mapcar '(lambda (c) (chr (sx:cp c))) codes)))

;;; ---------------------------------------------------------------- геометрия
(defun sx:pt (u v)
  (list (+ (car sx:P0) (* u (car sx:R)) (* v (car sx:D)))
        (+ (cadr sx:P0) (* u (cadr sx:R)) (* v (cadr sx:D)))
        (if (caddr sx:P0) (caddr sx:P0) 0.0)))

(defun sx:line (u1 v1 u2 v2 lay)
  (entmake (append (list '(0 . "LINE") (cons 8 lay) (cons 10 (sx:pt u1 v1)) (cons 11 (sx:pt u2 v2)))
                   (cond ((and sx:LN (= lay sx:LN)) (list (cons 48 sx:ltsN)))
                         ((and sx:B (= lay sx:B)) (list (cons 48 sx:ltsB)))
                         (t '())))))

;;; перемычка с фасками между столбцами a < b на уровне v (как в 02_Однолинейка:
;;; вверх 1 мм, фаска 2x2, горизонталь, фаска, вниз)
(defun sx:bridge (a b v lay)
  (cond ((< (- b a) 0.5) nil)
        ((> (- b a) 4.5)
         (sx:line a v a (+ v 1.0) lay) (sx:line a (+ v 1.0) (+ a 2.0) (+ v 3.0) lay)
         (sx:hline (+ a 2.0) (- b 2.0) (+ v 3.0) lay)
         (sx:line (- b 2.0) (+ v 3.0) b (+ v 1.0) lay) (sx:line b (+ v 1.0) b v lay))
        (t (sx:line a v b v lay))))

;;; точка подключения (кружок)
(setq sx:dotR 0.5)
(defun sx:dot (u v lay)
  (entmake (list '(0 . "CIRCLE") (cons 8 lay) (cons 10 (sx:pt u v)) (cons 40 sx:dotR))))

;;; засечки фазности на вертикальной линии: 1 фаза — одна, 3 фазы — три
(defun sx:ticks (u v n lay / k dv)
  (setq k 0)
  (repeat n
    (setq dv (* 1.3 (- k (/ (1- n) 2.0))))
    (sx:line (- u 1.6) (+ v dv -1.6) (+ u 1.6) (+ v dv 1.6) lay)
    (setq k (1+ k))))
;;; число фаз строки: «фаз гр.» (AR) или по тексту фазы (L1,L2,L3)
(defun sx:nph (d / n ph)
  (setq n (sx:n (sx:c d 44)) ph (sx:s (sx:c d sx:cPh)))
  (if (or (>= n 3) (vl-string-search "," ph)) 3 1))

;;; горизонталь с разрывом на границе листов: в конце листа стрелка,
;;; на новом листе линия начинается со стрелки (как у шины)
(setq sx:brkGap 2.5 sx:arrL 2.0 sx:arrH 1.4)
(defun sx:arr (x v dir lay)   ; остриё в x, dir = 1 вправо
  (sx:line (- x (* dir sx:arrL)) (+ v sx:arrH) x v lay)
  (sx:line (- x (* dir sx:arrL)) (- v sx:arrH) x v lay))
(defun sx:hline (u1 u2 v lay / a b w k bnd x)
  (setq a (min u1 u2) b (max u1 u2) w (* *sx-sheet* *sx-step*))
  (setq k (1+ (fix (/ a w))) x a)
  (while (< (setq bnd (* k w)) (- b 0.01))
    (sx:line x v (- bnd sx:brkGap) v lay)
    (sx:arr (- bnd sx:brkGap) v 1 lay)
    (sx:arr (+ bnd sx:brkGap sx:arrL) v 1 lay)
    (setq x (+ bnd sx:brkGap sx:arrL) k (1+ k)))
  (sx:line x v b v lay))

;;; цепочка перемычек по отсортированному списку столбцов
(defun sx:chain (lst v lay / prev)
  (setq prev nil)
  (foreach x (vl-sort lst '<)
    (if prev (sx:bridge prev x v lay))
    (setq prev x)))

;;; окрасить вставку в тёмно-красный (для блоков с цветом ПоБлоку)
(defun sx:red (ob / e ed)
  (if (and ob (setq e (vlax-vla-object->ename ob)) (setq ed (entget e)))
    (entmod (append (vl-remove-if '(lambda (x) (member (car x) '(62 420))) ed)
                    (list (cons 62 sx:red-aci) (cons 420 sx:red-rgb))))))

(defun sx:rect (u1 v1 u2 v2 lay)
  (sx:line u1 v1 u2 v1 lay) (sx:line u2 v1 u2 v2 lay)
  (sx:line u2 v2 u1 v2 lay) (sx:line u1 v2 u1 v1 lay))

(defun sx:text (u v h s lay)
  (setq s (sx:s s))
  (if (/= s "")
    (entmake (list '(0 . "TEXT") (cons 8 lay) (cons 10 (sx:pt u v)) (cons 40 h) (cons 1 s)
                   (cons 50 sx:ang) (cons 7 sx:style)))))


;;; ---------------------------------------------------------------- вставка блоков
;;; вставка с атрибутами по ТЕГАМ: pairs = (("TAG" . "value") ...)
(defun sx:ins (name u v pairs lay / ob tag val)
  (setq ob (vla-insertblock sx:ms (vlax-3d-point (sx:pt u v)) name 1.0 1.0 1.0 sx:ang))
  (vla-put-layer ob lay)
  (if (= (vla-get-hasattributes ob) :vlax-true)
    (foreach a (vlax-invoke ob 'GetAttributes)
      (setq tag (strcase (vla-get-tagstring a)))
      (vla-put-textstring a (sx:gost (if (setq val (assoc tag pairs)) (sx:s (cdr val)) "")))
      (if (setq val (assoc tag sx:fitlen)) (sx:attfit a (cdr val)))
      (if (setq val (assoc tag sx:cells))
        (progn
          (setq sx:bbctx (strcat name "|" tag))
          (sx:attclamp a (cadr (sx:pt u (+ v (cadr val)))) (cadr (sx:pt u (+ v (caddr val)))))
          (setq sx:bbctx "")))))
  ob)

;;; ячейки нижней таблицы по высоте (от низа столбца): надпись обязана быть внутри
(setq sx:cells '(("NAME1" 0.8 29.2) ("NAME2" 0.8 29.2) ("NAME3" 0.8 29.2)
                 ("CUR" 30.4 37.6) ("POW" 38.4 45.6) ("PH" 46.4 53.6) ("LINE" 54.4 61.6)
                 ("CAB1" 75.8 104.2) ("CAB2" 75.8 104.2)))
;;; реальные габариты надписи (bounding box): если длиннее ячейки — сжать, если вылезает — сдвинуть внутрь
(setq sx:bbctx "" sx:bbcache nil sx:cccache nil sx:blkcache nil)
(defun sx:attbox-raw (a / mn mx r)
  (setq r (vl-catch-all-apply 'vla-getboundingbox (list a 'mn 'mx)))
  (if (not (vl-catch-all-error-p r)) (list (vlax-safearray->list mn) (vlax-safearray->list mx))))
;;; опорная точка надписи: точка вставки (выравнивание влево) или точка выравнивания
(defun sx:att-anchor (a kind / p)
  (setq p (vl-catch-all-apply 'vlax-get (list a (if (= kind 0) 'InsertionPoint 'TextAlignmentPoint))))
  (if (and (not (vl-catch-all-error-p p)) (listp p) (= (length p) 3)) p))
;;; bbox с кэшем: ключ = (блок|тег, текст, ширина); в кэше смещения углов bbox от опорной точки.
;;; sx:bbctx задаёт вызывающий (высота, стиль, поворот, выравнивание постоянны для блока|тега);
;;; при пустом контексте кэш не используется. Результат тот же, что у vla-getboundingbox.
(defun sx:attbox (a / key hit anc kind bb)
  (if (= sx:bbctx "")
    (sx:attbox-raw a)
    (progn
      (setq key (list sx:bbctx (vla-get-textstring a) (vla-get-scalefactor a))
            hit (assoc key sx:bbcache))
      (cond
        ((and hit (setq anc (sx:att-anchor a (cadr hit))))
         (list (mapcar '+ anc (nth 2 hit)) (mapcar '+ anc (nth 3 hit))))
        (t
         (setq bb (sx:attbox-raw a))
         (if (and bb (not hit))
           (progn
             (setq kind (vla-get-alignment a))
             (if (not (member kind '(3 5)))
               (progn
                 (setq kind (if (= kind 0) 0 1) anc (sx:att-anchor a kind))
                 (if anc
                   (setq sx:bbcache (cons (list key kind (mapcar '- (car bb) anc) (mapcar '- (cadr bb) anc))
                                          sx:bbcache)))))))
         bb)))))
(defun sx:attclamp (a y0 y1 / bb h dy)
  (if (and (/= (vla-get-textstring a) "") (setq bb (sx:attbox a)))
    (progn
      (setq h (- (cadr (cadr bb)) (cadr (car bb))))
      ;; повёрнутый текст (вдоль столбца): длина идёт по Y — сжимаем по ширине
      (if (and (> h (- y1 y0)) (> (abs (sin (vla-get-rotation a))) 0.7))
        (progn (vla-put-scalefactor a (* (vla-get-scalefactor a) (/ (- y1 y0) h)))
               (setq bb (sx:attbox a))))
      (if bb
        (progn
          (setq dy (cond ((< (cadr (car bb)) y0) (- y0 (cadr (car bb))))
                         ((> (cadr (cadr bb)) y1) (- y1 (cadr (cadr bb))))
                         (t 0.0)))
          (if (/= dy 0.0)
            (vla-move a (vlax-3d-point '(0.0 0.0 0.0)) (vlax-3d-point (list 0.0 dy 0.0)))))))))

;;; длина надписи не больше maxl (иначе сжать по ширине) — текст не выходит за ячейку
(setq sx:fitlen '(("NAME1" . 27.0) ("NAME2" . 27.0) ("NAME3" . 27.0) ("CAB1" . 27.0) ("CAB2" . 27.0)
                  ("LINE" . 18.0) ("CUR" . 18.0) ("POW" . 18.0) ("PH" . 18.0)))
;;; в шрифте ГОСТ тип А нет знака «№» -> заменяем на sx:numsign
(setq sx:numsign nil)   ; "N" — заменять «№», если в шрифте нет этого знака
(defun sx:gost (txt / k ns)
  (setq ns (chr (sx:cp 8470)))
  (while (and sx:numsign (setq k (vl-string-search ns txt)))
    (setq txt (strcat (substr txt 1 k) sx:numsign (substr txt (+ k 2)))))
  txt)
(defun sx:attfit (a maxl / tb w)
  (setq tb (vl-catch-all-apply 'textbox
             (list (list (cons 1 (vla-get-textstring a)) (cons 40 (vla-get-height a))
                         (cons 41 (vla-get-scalefactor a)) (cons 7 (vla-get-stylename a))))))
  (if (and tb (not (vl-catch-all-error-p tb)))
    (progn
      (setq w (- (car (cadr tb)) (car (car tb))))
      (if (> w maxl) (vla-put-scalefactor a (* (vla-get-scalefactor a) (/ maxl w)))))))

;;; вставка с масштабом/поворотом и атрибутами ПО ПОРЯДКУ: vals = ("знач1" "знач2" ...)
(defun sx:insv (name u v sc rot vals lay / ob k atts)
  (setq ob (vla-insertblock sx:ms (vlax-3d-point (sx:pt u v)) name sc sc sc (+ sx:ang (* pi (/ rot 180.0)))))
  (vla-put-layer ob lay)
  (if (and vals (= (vla-get-hasattributes ob) :vlax-true))
    (progn
      (setq k 0 atts (vlax-invoke ob 'GetAttributes))
      (foreach a atts
        (if (< k (length vals)) (vla-put-textstring a (sx:s (nth k vals))))
        (setq k (1+ k)))))
  ob)

;;; динамические свойства
(defun sx:dynprop (ob pname / res)
  (if (= (vla-get-isdynamicblock ob) :vlax-true)
    (foreach p (vlax-invoke ob 'GetDynamicBlockProperties)
      (if (= (strcase (vla-get-propertyname p)) (strcase pname)) (setq res p))))
  res)
(defun sx:dynset (ob pname val / p)
  (if (setq p (sx:dynprop ob pname))
    (not (vl-catch-all-error-p
           (vl-catch-all-apply 'vla-put-value
             (list p (vlax-make-variant val (vlax-variant-type (vla-get-value p)))))))))
(defun sx:dynget (ob pname / p)
  (if (setq p (sx:dynprop ob pname)) (vlax-variant-value (vla-get-value p))))

(defun sx:bbox (ob / mn mx)
  (vla-getboundingbox ob 'mn 'mx)
  (list (vlax-safearray->list mn) (vlax-safearray->list mx)))

;;; подогнать линейное свойство так, чтобы край блока (axis 0 = X max, 1 = Y max) встал в target
(defun sx:fit (ob pname axis target / d0 cur delta ok)
  (if (setq d0 (sx:dynget ob pname))
    (progn
      (setq cur (nth axis (cadr (sx:bbox ob))) delta (- cur target))
      (if (> (abs delta) 0.2)
        (progn
          (sx:dynset ob pname (- d0 delta))
          (setq cur (nth axis (cadr (sx:bbox ob))))
          (if (> (abs (- cur target)) 0.3)
            (progn (sx:dynset ob pname (+ d0 delta))
                   (setq cur (nth axis (cadr (sx:bbox ob))))))))
      (setq ok (<= (abs (- cur target)) 0.3))))
  ok)


;;; верхняя точка подключения автомата QF1: самая верхняя видимая окружность на оси блока
;;; (в мировых Y). Берётся из определения (анонимного) блока — без учёта атрибутов.
(defun sx:dynkey (ob / res v)
  (setq res (list (vla-get-effectivename ob)))
  (if (= (vla-get-isdynamicblock ob) :vlax-true)
    (foreach p (vlax-invoke ob 'GetDynamicBlockProperties)
      (setq v (vlax-variant-value (vla-get-value p)))
      (if (= (type v) 'SAFEARRAY) (setq v (vlax-safearray->list v)))
      (setq res (cons v res))))
  res)
;;; определение блока: (Y верхней видимой окружности на оси, Y начала блока) или nil
(defun sx:circ-def (ob / bd org best c e)
  (setq bd (vl-catch-all-apply 'vla-item (list (vla-get-blocks sx:doc) (vla-get-name ob))))
  (if (not (vl-catch-all-error-p bd))
    (progn
      (setq org (vlax-get bd 'Origin))
      (vlax-for e bd
        (if (and (= (vla-get-objectname e) "AcDbCircle") (= (vla-get-visible e) :vlax-true))
          (progn
            (setq c (vlax-get e 'Center))
            (if (and (< (abs (- (car c) (car org))) 0.1) (or (null best) (> (cadr c) best)))
              (setq best (cadr c))))))
      (list best (cadr org)))))
;;; кэш по (имя эффективного блока + значения всех динамических свойств): геометрия от них зависит
(defun sx:circ-top (ob / key hit def)
  (setq key (vl-catch-all-apply 'sx:dynkey (list ob)))
  (if (vl-catch-all-error-p key) (setq key nil))
  (setq hit (if key (assoc key sx:cccache)))
  (if hit
    (setq def (cdr hit))
    (progn
      (setq def (sx:circ-def ob))
      (if (and def key) (setq sx:cccache (cons (cons key def) sx:cccache)))))
  (if (and def (car def))
    (+ (cadr (vlax-get ob 'InsertionPoint)) (* (- (car def) (cadr def)) (vla-get-yscalefactor ob)))))

;;; поставить точку подключения автомата ровно на шину (параметр Расстояние1)
(defun sx:qf-align (ob target / p0 c0 c1 k)
  (setq p0 (sx:dynget ob sx:pDist) c0 (sx:circ-top ob))
  (if (and (numberp p0) c0)
    (progn
      (sx:dynset ob sx:pDist (+ p0 1.0))
      (setq c1 (sx:circ-top ob))
      (if (and c1 (> (abs (setq k (- c1 c0))) 1e-4))
        (progn
          (sx:dynset ob sx:pDist (+ p0 1.0 (/ (- target c1) k)))
          (setq c1 (sx:circ-top ob))
          (if (and c1 (< (abs (- c1 target)) 0.2)) T c1))
        (progn (sx:dynset ob sx:pDist p0) c0)))))

;;; ---------------------------------------------------------------- слои
(defun sx:ltype (lt / lts)
  (setq lts (vla-get-linetypes sx:doc))
  (if (vl-catch-all-error-p (vl-catch-all-apply 'vla-item (list lts lt)))
    (vl-catch-all-apply 'vla-load (list lts lt "acadiso.lin")))
  (not (vl-catch-all-error-p (vl-catch-all-apply 'vla-item (list lts lt)))))

(defun sx:layer (name col lt / lays l)
  (setq lays (vla-get-layers sx:doc))
  (setq l (vl-catch-all-apply 'vla-item (list lays name)))
  (if (vl-catch-all-error-p l) (setq l (vla-add lays name)))
  (vla-put-color l col)
  (if (and lt (sx:ltype lt)) (vla-put-linetype l lt))
  name)

(defun sx:ensure-style ( / st r)
  (setq st (vl-catch-all-apply 'vla-item (list (vla-get-textstyles sx:doc) *sx-style*)))
  (if (vl-catch-all-error-p st)
    (setq st (vl-catch-all-apply 'vla-add (list (vla-get-textstyles sx:doc) *sx-style*))))
  (if (not (vl-catch-all-error-p st))
    (progn
      ;; шрифт задаётся по имени гарнитуры (TrueType), как в окне стилей AutoCAD
      (setq r (vl-catch-all-apply 'vla-setfont (list st *sx-face* :vlax-false :vlax-false 204 34)))
      (if (vl-catch-all-error-p r) (sx:log (strcat "style: setfont failed " (vl-catch-all-error-message r))))
      (vl-catch-all-apply 'vla-put-width (list st 1.0))
      (vl-catch-all-apply 'vla-put-height (list st 0.0))))
  (if (tblsearch "STYLE" *sx-style*) *sx-style* (getvar "TEXTSTYLE")))

(defun sx:layers ()
  (setq sx:L   (sx:layer "SX_LINES" 7 nil)
        sx:BL  (sx:layer "SX_BLOCKS" 7 nil)
        sx:TB  (sx:layer "SX_TABLE" 7 nil)
        sx:LN   (sx:layer "SX_N" 4 (if (sx:ltype sx:ltN) sx:ltN "HIDDEN2"))
        sx:PE  (sx:layer "SX_PE" 3 "CENTER2")
        sx:KNX (sx:layer "SX_KNX" 1 nil)
        sx:T   (sx:layer "SX_TEXT" 7 nil)
        sx:B   (sx:layer "SX_BOUND" 7 (if (sx:ltype sx:ltN) sx:ltN "DASHED2"))))

;;; ---------------------------------------------------------------- ваши блоки
;;; роль -> (имя блока, масштаб, поворот°, смещение v вставки, верх для кабеля/высота)
(defun sx:roles ()
  (setq sx:bQF     "QF1"
        sx:bBUS    (sx:ru '(1064 1080 1085 1072))
        sx:bHEAD   (sx:ru '(95 1041 1086 1082 1086 1074 1080 1082 95 1089 1093 1077 1084 1072 32 1070 1085 1080 1082 1086 1084))
        sx:bXT     "XT"
        sx:bXT2    (sx:ru '(88 84 32 1073 1077 1079 32 1079 1077 1084 1083 1080))   ; XT без земли: 2 кружка (L, N)
        sx:bOut    (sx:ru '(1042 1099 1074 1086 1076))                          ; кабельный вывод (динамический блок, диск с хвостом)
        sx:bT      (sx:ru '(1041 1055))          ; трансформатор: вставка сверху, высота 13.8
        sx:bPS     (sx:ru '(1044 1088 1072 1081 1074 1077 1088))         ; блок питания: вставка снизу, высота 13.7
        sx:bMK     "MK-5-1"         ; вставка сверху, масштаб 0.6111
        sx:pVis    (sx:ru '(1042 1080 1076 1080 1084 1086 1089 1090 1100 49))        ; имя свойства видимости
        sx:pDist   (sx:ru '(1056 1072 1089 1089 1090 1086 1103 1085 1080 1077 49))       ; имя свойства длины
        sx:vBusVis (sx:ru '(1053 1072 1095 1072 1083 1086 45 1087 1088 1086 1076 1086 1083 1078 1077 1085 1080 1077)))    ; состояние шины
  ;; УГО: (символ блок масштаб поворот v-вставки v-верх)
  (setq sx:ugo
    (list (list "SOCKET" (sx:ru '(1056 1086 1079 1077 1090 1082 1072 49)) 0.042788 180.0 72.41 72.41)
          (list "LAMP"   "$RECOVER_230831092053-0" 1.0 0.0 68.10 72.40)
          (list "SERVO"  (sx:ru '(1057 1077 1088 1074 1086 1087 1088 1080 1074 1086 1076 45 49)) 0.79921 0.0 64.06 72.05)
          (list "VALVE"  "EK.1" 1.0 0.0 64.06 70.56)
          (list "MOTOR"  (sx:ru '(1055 1088 1080 1074 1086 1076)) 2.5 0.0 68.10 71.10)
          (list "HEAT"   "SX_U_HEAT" 1.0 0.0 68.50 75.0)
          (list "BOX"    "SX_U_BOX" 1.0 0.0 68.50 75.0)))
  ;; аппараты: (код блок масштаб вставка-сверху? высота)
  (setq sx:devs
    (list (list "T"  sx:bT  1.0 T   13.8)
          (list "PS" sx:bPS 1.0 nil 13.7)
          (list "MK" sx:bMK 0.6111 T 9.35)
          (list "KM" "SX_KM" 1.0 T 10.0)
          (list "X"  "SX_DEV" 1.0 T 10.0))))


;;; ---------------------------------------------------------------- служебные блоки
(defun bL (x1 y1 x2 y2) (list '(0 . "LINE") '(8 . "0") (list 10 x1 y1 0.0) (list 11 x2 y2 0.0)))
(defun bLc (x1 y1 x2 y2 c) (append (bL x1 y1 x2 y2) (list (cons 62 c))))
(defun bC (x y r) (list '(0 . "CIRCLE") '(8 . "0") (list 10 x y 0.0) (cons 40 r)))
(defun bCc (x y r c) (append (bC x y r) (list (cons 62 c))))
(defun bA (x y r a1 a2) (list '(0 . "ARC") '(8 . "0") (list 10 x y 0.0) (cons 40 r) (cons 50 a1) (cons 51 a2)))
(defun bR (x1 y1 x2 y2) (list (bL x1 y1 x2 y1) (bL x2 y1 x2 y2) (bL x2 y2 x1 y2) (bL x1 y2 x1 y1)))
(defun bT (x y h s just rot)
  (append (list '(0 . "TEXT") '(8 . "0") (list 10 x y 0.0) (cons 40 h) (cons 1 s) (cons 50 rot) (cons 7 sx:style))
          (cond ((= just "M") (list '(72 . 4) (list 11 x y 0.0)))
                ((= just "R") (list '(72 . 2) (list 11 x y 0.0)))
                (t '()))))
(defun bAtt (tag x y h just rot col)
  (append (list '(0 . "ATTDEF") '(8 . "0") (list 10 x y 0.0) (cons 40 h) '(1 . "") (cons 3 tag) (cons 2 tag)
                '(70 . 0) (cons 50 rot) (cons 7 sx:style))
          (if col (list (cons 62 col)) '())
          (cond ((= just "M") (list '(72 . 4) (list 11 x y 0.0)))
                ((= just "R") (list '(72 . 2) (list 11 x y 0.0)))
                (t '()))))

(defun sx:defblock (name ents force / hasatt)
  (if (or force (not (tblsearch "BLOCK" name)))
    (progn
      (setq hasatt (vl-some '(lambda (e) (= (cdr (assoc 0 e)) "ATTDEF")) ents))
      (entmake (list '(0 . "BLOCK") (cons 2 name) (cons 70 (if hasatt 2 0)) '(10 0.0 0.0 0.0)))
      (foreach e ents (entmake e))
      (entmake '((0 . "ENDBLK") (8 . "0"))))))

(defun sx:make-blocks (force / hp)
  (setq hp (/ pi 2.0))
  ;; канал модуля — геометрия канала из вашего блока «24-1»; вставка в низ провода OUT
  (sx:defblock "SX_CH"
    (list (bL 0.0 0.0 0.0 3.86) (bC 0.0 4.86 1.0) (bL 1.1 4.3 -1.1 5.42) (bL 0.0 5.86 0.0 8.96)
          (bL 0.0 14.96 3.69 8.57) (bL 0.0 17.84 0.0 14.96) (bC 0.0 18.84 1.0) (bL 1.1 19.4 -1.1 18.28)
          (bL 0.0 19.84 0.0 23.7) (bT 1.82 17.78 2.5 "IN" "L" 0.0) (bT 1.82 3.8 2.5 "OUT" "L" 0.0)
          (bAtt "CH" -0.6 11.02 2.5 "R" 0.0 nil)) force)
  ;; заменители (если вашего блока нет ни в чертеже, ни в библиотеке)
  (sx:defblock "SX_QF"
    (list (bL -0.8 -37.1 0.8 -35.5) (bL -0.8 -35.5 0.8 -37.1) (bL 0.0 -45.0 -3.8 -36.9) (bL 0.0 -45.0 0.0 -54.3)
          (bT -6.2 -27.4 2.5 "QF" "L" 0.0)
          (bAtt "N" -3.2 -27.4 2.5 "L" 0.0 nil) (bAtt "P" -6.6 -30.4 2.5 "L" 0.0 nil)
          (bAtt "A" -7.6 -33.3 2.5 "L" 0.0 nil) (bAtt "MA" -9.4 -36.1 2.5 "L" 0.0 nil)) force)
  (sx:defblock "SX_T" (list (bL 0.0 0.0 0.0 -2.5) (bC 0.0 -5.0 2.5) (bC 0.0 -8.8 2.5) (bL 0.0 -11.3 0.0 -13.8)
                            (bAtt "N" -6.2 -5.5 2.5 "L" 0.0 nil) (bAtt "W" -10.3 -9.1 2.5 "L" 0.0 nil)
                            (bAtt "V" -8.8 -12.9 2.5 "L" 0.0 nil)) force)
  (sx:defblock "SX_PS" (append (list (bL 0.0 0.0 0.0 2.5) (bL 0.0 11.2 0.0 13.7)) (bR -2.2 2.5 2.2 11.2)
                               (list (bT 0.0 6.85 2.0 "=" "M" 0.0) (bAtt "N" -4.8 8.3 2.5 "L" 0.0 nil))) force)
  (sx:defblock "SX_MKX" (append (list (bL 0.0 0.0 0.0 -3.0) (bL 0.0 -12.3 0.0 -15.3)) (bR -2.0 -12.3 2.0 -3.0)
                                (list (bL -2.0 -12.3 2.0 -3.0) (bAtt "N" 3.0 -8.8 2.5 "L" 0.0 nil))) force)
  (sx:defblock "SX_KM" (list (bL 0.0 0.0 0.0 -1.5) (bL 0.0 -8.5 0.0 -10.0) (bL 0.0 -8.5 -3.0 -2.0)
                             (bA 0.0 -1.5 0.8 pi (* 2 pi)) (bAtt "N" -9.0 -6.0 2.5 "L" 0.0 nil)) force)
  (sx:defblock "SX_DEV" (append (list (bL 0.0 0.0 0.0 -1.0) (bL 0.0 -9.0 0.0 -10.0)) (bR -3.0 -9.0 3.0 -1.0)
                                (list (bAtt "N" -9.0 -6.0 2.2 "L" 0.0 nil))) force)
  (sx:defblock "SX_XT" (list (bCc 0.0 -1.0 1.0 30) (bCc 0.0 -3.0 1.0 5) (bCc 0.0 -5.0 1.0 3)
                             (bAtt "N" 1.7 0.1 2.5 "L" 0.0 nil)) force)
  ;; двухуровневая клемма без земли (2 проводника на одну клемму) и отдельный кружок PE
  (sx:defblock "SX_XT2" (list (bCc 0.0 -1.0 1.0 30) (bCc 0.0 -3.0 1.0 5)
                              (bAtt "N" 1.7 0.1 2.5 "L" 0.0 nil)) force)
  (sx:defblock "SX_XTPE" (list (bCc 0.0 0.0 1.0 3) (bLc 1.003 0.753 -0.753 -1.003 3) (bLc 0.753 1.003 -1.003 -0.753 3)) force)
  ;; кабельный вывод (если блока «Вывод» нет): линия вниз со стрелкой, вставка в верхней точке
  (sx:defblock "SX_U_OUT" (list (bL 0.0 0.0 0.0 -9.0) (bL 0.0 -9.0 -1.2 -6.6) (bL 0.0 -9.0 1.2 -6.6)) force)
  (sx:defblock "SX_U_HEAT" (append (list (bL 0.0 6.5 0.0 2.5)) (bR -3.0 -2.5 3.0 2.5)
                                   (list (bL -2.0 0.0 -1.0 1.2) (bL -1.0 1.2 0.0 -1.2) (bL 0.0 -1.2 1.0 1.2) (bL 1.0 1.2 2.0 0.0))) force)
  (sx:defblock "SX_U_BOX" (append (list (bL 0.0 6.5 0.0 2.5)) (bR -3.0 -2.5 3.0 2.5)
                                  (list (bAtt "CODE" 0.0 0.0 1.6 "M" 0.0 nil))) force)
  (sx:defblock "SX_U_LAMP" (list (bC 0.0 0.0 4.3) (bL -2.9 -2.9 2.9 2.9) (bL -2.9 2.9 2.9 -2.9)) force)
  (sx:defblock "SX_U_SOCKET" (list (bL 0.0 0.0 0.0 -3.8) (bA 0.0 -6.77 3.0 0.0 pi) (bL -3.0 -6.77 3.0 -6.77)) force)
  (sx:defblock "SX_U_M" (list (bC 0.0 0.0 1.2) (bT 0.0 0.0 1.2 "M" "M" 0.0)) force)
  ;; две доп. клеммы фаз (оранжевые, как в блоке XT) над XT — для трёхфазных линий
  (sx:defblock "SX_XT3"
    (apply 'append
      (mapcar '(lambda (dy)
                 (list (list '(0 . "CIRCLE") '(8 . "0") (list 10 0.0 (- dy 1.0) 0.0) '(40 . 1.0) '(62 . 32) '(420 . 13461800))
                       (list '(0 . "LINE") '(8 . "0") (list 10 1.0033 (- dy 0.2467) 0.0) (list 11 -0.7533 (- dy 2.0033) 0.0) '(62 . 32) '(420 . 13461800))
                       (list '(0 . "LINE") '(8 . "0") (list 10 0.7533 (+ dy 0.0033) 0.0) (list 11 -1.0033 (- dy 1.7533) 0.0) '(62 . 32) '(420 . 13461800))))
              '(2.0 4.0))) force)
  ;; столбец нижней таблицы (ширина 20): вставка в центре низа столбца
  (sx:defblock "SX_TCOL"
    (append
      ;; ячейки таблицы до строки УГО (75); марка кабеля — просто текстом над таблицей, без ячейки
      (list (bL 10.0 0.0 10.0 75.0))
      (mapcar '(lambda (y) (bL -10.0 y 10.0 y)) '(0.0 30.0 38.0 46.0 54.0 62.0 75.0))
      (list (bAtt "NAME1" -3.6 15.0 2.5 "M" hp nil) (bAtt "NAME2" 0.0 15.0 2.5 "M" hp nil)
            (bAtt "NAME3" 3.6 15.0 2.5 "M" hp nil)
            (bAtt "CUR" 0.0 34.0 2.5 "M" 0.0 nil) (bAtt "POW" 0.0 42.0 2.5 "M" 0.0 nil)
            (bAtt "PH" 0.0 50.0 2.5 "M" 0.0 nil) (bAtt "LINE" 0.0 58.0 2.5 "M" 0.0 nil)
            (bAtt "CAB1" -2.6 90.0 2.5 "M" hp nil) (bAtt "CAB2" 2.6 90.0 2.5 "M" hp nil))) force)
  (princ))

(defun c:SXBLOCKS ()
  (setq sx:doc (vla-get-activedocument (vlax-get-acad-object)))
  (setq sx:style (if (tblsearch "STYLE" *sx-style*) *sx-style* (getvar "TEXTSTYLE")))
  (sx:make-blocks T)
  (princ (strcat "\n" (sx:ru '(1057 1083 1091 1078 1077 1073 1085 1099 1077 32 1073 1083 1086 1082 1080 32 83 88 95 42 32 1087 1077 1088 1077 1089 1086 1079 1076 1072 1085 1099 46))))
  (princ))

;;; ---------------------------------------------------------------- библиотека блоков
;;; библиотека блоков: DWG открывается один раз (проверка + импорт) и закрывается в sx:lib-close.
;;; Если библиотека открыта в этом же AutoCAD — берём открытый документ
;;; (ObjectDBX не может открыть файл, который открыт в редакторе).
(setq sx:libsrc nil sx:libdbx nil)
(defun sx:lib-close ()
  (if sx:libdbx (vl-catch-all-apply 'vlax-release-object (list sx:libdbx)))
  (setq sx:libsrc nil sx:libdbx nil))

(defun sx:lib-open (path / full dbx r d)
  (if (null sx:libsrc)
    (if (setq full (findfile path))
      (progn
        (vlax-for d (vla-get-documents (vlax-get-acad-object))
          (if (and (null sx:libsrc) (= (strcase (vla-get-fullname d)) (strcase full)) (not (equal d sx:doc)))
            (setq sx:libsrc d)))
        (if sx:libsrc
          (sx:log "lib: taken from the open document")
          (progn
            (setq dbx (vl-catch-all-apply 'vla-getinterfaceobject
                        (list (vlax-get-acad-object) (strcat "ObjectDBX.AxDbDocument." (substr (getvar "ACADVER") 1 2)))))
            (cond
              ((or (null dbx) (vl-catch-all-error-p dbx))
               (sx:log (strcat "lib: ObjectDBX not available " (if dbx (vl-catch-all-error-message dbx) ""))))
              ((vl-catch-all-error-p (setq r (vl-catch-all-apply 'vla-open (list dbx full))))
               (sx:log (strcat "lib: cannot open " full ": " (vl-catch-all-error-message r)))
               (vl-catch-all-apply 'vlax-release-object (list dbx)))
              (t (setq sx:libdbx dbx sx:libsrc dbx))))))))
  sx:libsrc)

;;; какие из names есть в библиотеке (без копирования)
(defun sx:lib-found (path names / src blks b res)
  (if (and names path (/= path "") (findfile path) (setq src (sx:lib-open path)))
    (progn
      (setq blks (vla-get-blocks src))
      (foreach n names
        (setq b (vl-catch-all-apply 'vla-item (list blks n)))
        (if (not (vl-catch-all-error-p b))
          (progn (setq res (cons n res)) (vl-catch-all-apply 'vlax-release-object (list b)))))))
  (reverse res))

;;; скопировать недостающие блоки из DWG-библиотеки
(defun sx:import-blocks (path names / src objs b sa r blks)
  (setq names (vl-remove-if '(lambda (n) (tblsearch "BLOCK" n)) names))
  (cond
    ((null names) nil)
    ((and path (findfile path))
     (if (setq src (sx:lib-open path))
       (progn
         (setq objs '() blks (vla-get-blocks src))
         (foreach n names
           (setq b (vl-catch-all-apply 'vla-item (list blks n)))
           (if (not (vl-catch-all-error-p b)) (setq objs (cons b objs))))
         (sx:log (strcat "lib: found " (itoa (length objs)) " of " (itoa (length names))))
         (if objs
           (progn
             (setq sa (vlax-make-safearray vlax-vbobject (cons 0 (1- (length objs)))))
             (vlax-safearray-fill sa objs)
             (setq r (vl-catch-all-apply 'vla-copyobjects (list src sa (vla-get-blocks sx:doc))))
             (if (vl-catch-all-error-p r) (sx:log (strcat "lib: copy error " (vl-catch-all-error-message r))))
             (foreach b objs (vl-catch-all-apply 'vlax-release-object (list b))))))))
    (t (sx:log (strcat "lib: file not found: " (sx:str path)))))
  (sx:lib-close))

;;; наличие блока в чертеже с кэшем (кэш сбрасывается в sx:main и после импорта)
(defun sx:blk-p (n / hit)
  (setq hit (assoc n sx:blkcache))
  (if (null hit)
    (setq hit (cons n (if (tblsearch "BLOCK" n) T nil)) sx:blkcache (cons hit sx:blkcache)))
  (cdr hit))

(defun sx:blk (name fallback) (if (sx:blk-p name) name fallback))


;;; ---------------------------------------------------------------- Excel
(setq sx:xl nil sx:xl-created nil sx:wb nil sx:wb-opened nil
      sx:sheets nil sx:si 0 sx:panels nil sx:pi 0 sx:path "")

(defun sx:xl-app ( / xl)
  (if (null sx:xl)
    (progn
      (setq xl (vl-catch-all-apply 'vlax-get-object (list "Excel.Application")))
      (if (or (null xl) (vl-catch-all-error-p xl))
        (setq xl (vlax-create-object "Excel.Application") sx:xl-created T))
      (setq sx:xl xl)))
  sx:xl)

;;; освободить COM-объект (не бросает ошибку)
(defun sx:rel (o)
  (if (and o (= (type o) 'VLA-OBJECT)) (vl-catch-all-apply 'vlax-release-object (list o)))
  nil)

(defun sx:cell (sh addr / rg r)
  (setq rg (vl-catch-all-apply 'vlax-get-property (list sh 'Range addr)))
  (if (vl-catch-all-error-p rg) (error (vl-catch-all-error-message rg)))
  (setq r (vl-catch-all-apply 'vlax-get-property (list rg 'Value2)))
  (sx:rel rg)
  (if (vl-catch-all-error-p r) (error (vl-catch-all-error-message r)))
  (sx:v r))

;;; освободить листы и книгу, взятые в sx:open-wb (книги, открытые нами, закрываются в sx:xl-release)
(defun sx:free-book ()
  (foreach p sx:sheets (sx:rel (cdr p)))
  (if (and sx:wb (not (member sx:wb sx:wb-opened))) (sx:rel sx:wb))
  (setq sx:sheets nil sx:wb nil))

(defun sx:open-wb (path / xl wbs wb full w sh shs)
  (setq xl (sx:xl-app) wb nil)
  (sx:free-book)
  (setq wbs (vlax-get-property xl 'Workbooks))
  (vlax-for w wbs
    (if (null wb)
      (progn
        (setq full (vlax-get-property w 'FullName))
        (if (= (strcase full) (strcase path)) (setq wb w) (sx:rel w)))
      (sx:rel w)))
  (if (null wb)
    (progn
      (setq wb (vl-catch-all-apply 'vlax-invoke-method
                 (list wbs 'Open path :vlax-false :vlax-true)))
      (if (vl-catch-all-error-p wb)
        (setq wb nil)
        (progn (setq sx:wb-opened (cons wb sx:wb-opened))
               (vl-catch-all-apply 'vlax-invoke-method (list xl 'Calculate))))))
  (sx:rel wbs)
  (setq sx:wb wb sx:sheets nil sx:si 0 sx:panels nil sx:pi 0)
  (if wb
    (progn
      (setq sx:path path)
      (setq shs (vlax-get-property wb 'Worksheets))
      (vlax-for sh shs
        (setq sx:sheets (append sx:sheets (list (cons (vlax-get-property sh 'Name) sh)))))
      (sx:rel shs)
      (foreach p sx:sheets
        (if (= (sx:s (sx:cell (cdr p) "BC1")) "SX_DATA")
          (setq sx:si (vl-position p sx:sheets))))
      (sx:load-panels)))
  wb)

(defun sx:load-panels ( / sh rg vals p)
  (setq sx:panels nil sx:pi 0)
  (if (and sx:sheets (setq sh (cdr (nth sx:si sx:sheets))))
    (progn
      (setq rg (vl-catch-all-apply 'vlax-get-property (list sh 'Range "A3:A1000")))
      (if (not (vl-catch-all-error-p rg))
        (progn
          (setq vals (vl-catch-all-apply
                       '(lambda ()
                          (vlax-safearray->list (vlax-variant-value (vlax-get-property rg 'Value2))))
                       nil))
          (sx:rel rg)
          (if (not (vl-catch-all-error-p vals))
            (foreach row vals
              (setq p (sx:s (sx:v (car row))))
              (if (and (/= p "") (not (member p sx:panels))) (setq sx:panels (append sx:panels (list p))))))))))
  (if (and (getenv "SX_LastPanel") (member (getenv "SX_LastPanel") sx:panels))
    (setq sx:pi (vl-position (getenv "SX_LastPanel") sx:panels)))
  sx:panels)

(defun sx:read-data ( / sh rg vals)
  (if (setq sh (cdr (nth sx:si sx:sheets)))
    (progn
      (setq rg (vlax-get-property sh 'Range *sx-range*))
      (setq vals (vl-catch-all-apply
                   '(lambda ()
                      (vlax-safearray->list (vlax-variant-value (vlax-get-property rg 'Value2))))
                   nil))
      (sx:rel rg)
      (if (vl-catch-all-error-p vals) (error (vl-catch-all-error-message vals)))
      (mapcar '(lambda (row) (mapcar 'sx:v row)) vals))))

(defun sx:xl-release ()
  (sx:free-book)
  (foreach wb sx:wb-opened
    (vl-catch-all-apply 'vlax-invoke-method (list wb 'Close :vlax-false))
    (sx:rel wb))
  (if (and sx:xl sx:xl-created) (vl-catch-all-apply 'vlax-invoke-method (list sx:xl 'Quit)))
  (if sx:xl (sx:rel sx:xl))
  (setq sx:xl nil sx:xl-created nil sx:wb nil sx:wb-opened nil sx:sheets nil))

;;; ---------------------------------------------------------------- очистка
(defun sx:clear ( / ss k)
  (if (setq ss (ssget "_X" '((8 . "SX_*"))))
    (progn
      (setq k (sslength ss))
      (repeat k (entdel (ssname ss (setq k (1- k))))))))

(defun c:SXCLEAR () (sx:clear) (princ))

;;; ---------------------------------------------------------------- отрисовка
(defun sx:content-p (d)
  (or (sx:has d sx:cQFown) (sx:c d sx:cCh) (sx:is1 d sx:cXT)))

(defun sx:draw-table (recs / nm cab k)
  (if (sx:blk-p sx:bHEAD) (sx:insv sx:bHEAD sx:vHeadDx 0.0 1.0 0.0 nil sx:TB))
  (foreach d recs
   (sx:try (function (lambda ()
    (setq nm (sx:wrap (sx:s (sx:c d sx:cName)) 20 3)
          cab (sx:s (sx:c d sx:cCab))
          k (vl-string-search " L=" cab))
    (if (= (length nm) 1) (setq nm (list "" (car nm))))
    (setq nm (list (sx:nth0 0 nm) (sx:nth0 1 nm) (sx:nth0 2 nm)))
    (sx:ins "SX_TCOL" (sx:u d) 0.0
            (list (cons "NAME1" (nth 0 nm)) (cons "NAME2" (nth 1 nm)) (cons "NAME3" (nth 2 nm))
                  (cons "CUR" (sx:num (sx:c d sx:cI))) (cons "POW" (sx:num (sx:c d sx:cPow)))
                  (cons "PH" (sx:s (sx:c d sx:cPh))) (cons "LINE" (sx:s (sx:c d sx:cLine)))
                  (cons "CAB1" (vl-string-translate "." "," (if k (substr cab 1 k) cab)))
                  (cons "CAB2" (vl-string-translate "." "," (if k (substr cab (+ k 2)) ""))))
            sx:TB)))
    (strcat "table " (sx:rowinfo d)))))

;;; шина: ваш динамический блок «Шина» (L и N) или две линии
(defun sx:bus (us ue vL vN name / ob)
  (if (sx:blk-p sx:bBUS)
    (progn
      (setq ob (sx:insv sx:bBUS (+ us 5.0) vL 1.0 0.0 (list name) sx:L))
      (sx:dynset ob sx:pVis sx:vBusVis)
      (if (not (sx:fit ob sx:pDist 0 (car (sx:pt (- ue 1.0) vL))))
        (progn (sx:line (+ us 5) vL (- ue 1) vL sx:L) (sx:line (+ us 5) vN (- ue 1) vN sx:LN))))
    (progn
      (sx:line (+ us 3) vL (- ue 1) vL sx:L) (sx:line (+ us 3) vN (- ue 1) vN sx:LN)
      (sx:text (+ us 4) (+ vL 1) 2.5 name sx:T) (sx:text (+ us 4) (+ vN 1) 2.5 "N" sx:T))))

(defun sx:draw-sheets (draw panel / sheets s us ue onsheet b2 segs k)
  (setq sheets '())
  (foreach d draw
    (if (and (sx:content-p d) (not (member (sx:sheet (sx:i d)) sheets)))
      (setq sheets (cons (sx:sheet (sx:i d)) sheets))))
  (sx:log (strcat "sheets=" (sx:str sheets)))
  (foreach s sheets
   (sx:try (function (lambda ()
    (setq us (* s *sx-sheet* *sx-step*) ue (* (1+ s) *sx-sheet* *sx-step*))
    (sx:bus us ue sx:vBusL sx:vBusN "L1,L2,L3")
    (sx:text (+ us 2) (+ sx:vTop 1.5) 2.5 panel sx:T)
    (sx:line (+ us 2) sx:vTop (- ue 2) sx:vTop sx:B)
    (sx:line (+ us 2) sx:vBottom (- ue 2) sx:vBottom sx:B)   ; граница щита снизу
    (setq onsheet (vl-remove-if-not '(lambda (d) (= (sx:sheet (sx:i d)) s)) draw))
    (setq b2 (vl-remove-if-not '(lambda (d) (and (sx:has d sx:cQFown) (sx:is1 d sx:cBus2))) onsheet))
    ;; вторые шины — до конца листа; если на листе несколько разных, каждая до начала следующей
    (setq segs '())
    (foreach d (vl-sort b2 '(lambda (a b) (< (sx:u a) (sx:u b))))
      (if (or (null segs) (/= (car (car segs)) (sx:s (sx:c d sx:cBusName))))
        (setq segs (cons (list (sx:s (sx:c d sx:cBusName)) (sx:u d)) segs))))
    (setq segs (reverse segs) k 0)
    (foreach sg segs
      (sx:bus (if (= k 0) us (- (cadr sg) 9.0))
              (if (< (1+ k) (length segs)) (- (cadr (nth (1+ k) segs)) 11.0) ue)
              sx:vBus2L sx:vBus2N (car sg))
      (setq k (1+ k)))
    (if (vl-some '(lambda (d) (sx:is1 d sx:cXT)) onsheet)
      (sx:line (+ us 3) sx:vPE (- ue 1) sx:vPE sx:PE))))
    (strcat "sheet " (itoa (1+ s))))))

;;; «2P С16А 30мА» -> (полюса номинал утечка)
(defun sx:qf-parts (q / toks p a ma)
  (setq toks (sx:split (sx:s (sx:c q sx:cQFtxt))) p "" a "" ma "")
  (foreach tk toks
    (cond ((wcmatch (strcase tk) "#P") (setq p tk))
          ((= a "") (setq a tk))
          (t (setq ma (if (= ma "") tk (strcat ma " " tk))))))
  (list p a ma))

(defun sx:digits (s / r ch k)
  (setq r "" k 1)
  (while (<= k (strlen s))
    (setq ch (substr s k 1))
    (if (wcmatch ch "#") (setq r (strcat r ch)))
    (setq k (1+ k)))
  r)

;;; аппараты группы: список (код текст) по порядку
(defun sx:dev-list (q / k res code)
  (setq k 0 res '())
  (repeat 3
    (setq code (sx:s (sx:c q (+ sx:cDevC1 k))))
    (if (/= code "") (setq res (append res (list (list code (sx:s (sx:c q (+ sx:cDev1 k))))))))
    (setq k (1+ k)))
  res)

(defun sx:dev-spec (code) (cond ((assoc code sx:devs)) (t (assoc "X" sx:devs))))

;;; аппарат(ы) после автомата. side = T: БП/трансформатор сбоку справа (как QF46 в DWG),
;;; иначе — друг под другом под автоматом.
(defun sx:dev-ins (dv spec x top / blk toks w v n ob)
  (setq blk (nth 1 spec) toks (sx:split (cadr dv)))
  (cond
    ((= (car dv) "T")
     (setq sx:tcount (1+ sx:tcount) w "" v "")
     (foreach tk toks
       (cond ((wcmatch (strcase tk) "*W") (setq w (sx:digits tk)))
             ((wcmatch (strcase tk) "*V") (setq v (sx:digits tk)))))
     (setq n (list (itoa sx:tcount) w v)))
    ((= (car dv) "PS") (setq sx:pcount (1+ sx:pcount) n (list (itoa sx:pcount))))
    (t (setq n (list (cadr dv)))))
  (if (not (sx:blk-p blk))
    (setq blk (cdr (assoc blk (list (cons sx:bT "SX_T") (cons sx:bPS "SX_PS") (cons sx:bMK "SX_MKX")
                                    (cons "SX_KM" "SX_KM") (cons "SX_DEV" "SX_DEV"))))))
  (setq ob (sx:insv blk x (if (nth 3 spec) top (- top (nth 4 spec))) (if (= blk "SX_MKX") 0.6111 (nth 2 spec)) 0.0 n sx:BL))
  (if (member (car dv) '("T" "PS")) (sx:red ob))
  ob)

(defun sx:draw-qf (q u side nbyp / devs H top parts vIns busL ob name spec r xs xd vTop)
  (setq devs (sx:dev-list q) H 0.0)
  (foreach dv devs (setq H (+ H (nth 4 (sx:dev-spec (car dv))))))
  (setq busL (if (sx:is1 q sx:cBus2) sx:vBus2L sx:vBusL)
        parts (sx:qf-parts q)
        name (sx:s (sx:c q sx:cQFown)))
  (if (and side devs (> (- (if (sx:is1 q sx:cBus2) sx:vBus2N sx:vBusN) 3.0) (+ sx:vSideRet sx:sideLead H)))
    (setq vIns (+ sx:vSideB sx:qfBody))
    (setq side nil vIns (+ sx:vQFins H)))
  ;; автомат: точка подключения блока — ровно на шину
  (if (sx:blk-p sx:bQF)
    (progn
      (setq ob (sx:insv sx:bQF u vIns 1.0 0.0 (list (sx:digits name) (nth 0 parts) (nth 1 parts) (nth 2 parts)) sx:BL))
      ;; видимость: QFD — с N; 1P без модуля — QF_L+N (N показывает блок); через модуль — QF (N обходом)
      (sx:dynset ob sx:pVis (cond ((sx:is1 q sx:cRcd) "QFD")
                                  ((and (= (strcase (nth 0 parts)) "1P") (not nbyp)) "QF_L+N")
                                  (t "QF")))
      (setq r (sx:qf-align ob (cadr (sx:pt u busL))))
      (cond ((= r T) nil)
            ((sx:fit ob sx:pDist 1 (cadr (sx:pt u busL))) nil)
            ((numberp r) (sx:line u busL u (- r (cadr sx:P0)) sx:L))
            (t (sx:line u busL u (- vIns sx:vQFcontact) sx:L))))
    (progn
      (sx:insv "SX_QF" u vIns 1.0 0.0 (list (sx:digits name) (nth 0 parts) (nth 1 parts) (nth 2 parts)) sx:BL)
      (sx:line u busL u (- vIns sx:vQFcontact) sx:L)))
  (if side
    ;; сбоку: низ автомата -> вправо -> вверх -> аппараты сверху вниз -> возврат к столбцу -> вниз на вход
    (progn
      (setq xs (+ u sx:sideX1) xd (+ u sx:sideX2) vTop (+ sx:vSideRet sx:sideLead H) top vTop)
      (sx:line u sx:vSideB xs sx:vSideB sx:L)
      (sx:line xs sx:vSideB xs vTop sx:L)
      (sx:line xs vTop xd vTop sx:L)
      (foreach dv devs
        (setq spec (sx:dev-spec (car dv)))
        (sx:dev-ins dv spec xd top)
        (setq top (- top (nth 4 spec))))
      (sx:line xd top xd sx:vSideRet sx:L)
      (sx:line xd sx:vSideRet u sx:vSideRet sx:L)
      (sx:line u sx:vSideRet u sx:vIN0 sx:L))
    ;; под автоматом
    (progn
      (setq top (- vIns sx:vQFlen))
      (foreach dv devs
        (setq spec (sx:dev-spec (car dv)))
        (sx:dev-ins dv spec u top)
        (setq top (- top (nth 4 spec)))))))

(defun sx:draw-groups (draw / gid u busN members feeds chF xtF chU xtU xts xN hi side devs lowv)
  (setq sx:tcount 0 sx:pcount 0)
  (foreach q draw
   (sx:try (function (lambda ()
    (if (sx:has q sx:cQFown)
      (progn
        (setq gid (sx:s (sx:c q sx:cQFown)) u (sx:u q)
              busN (if (sx:is1 q sx:cBus2) sx:vBus2N sx:vBusN)
              devs (sx:dev-list q))
        ;; БП/трансформатор сбоку, если соседний справа столбец (на том же листе) без автомата
        (setq lowv (and devs (not (vl-some '(lambda (dv) (not (member (car dv) '("T" "PS")))) devs))))
        (setq side (and lowv
                        (= (sx:sheet (1+ (sx:i q))) (sx:sheet (sx:i q)))
                        (not (vl-some '(lambda (d) (and (= (sx:i d) (1+ (sx:i q))) (sx:has d sx:cQFown))) draw))))
        (setq members (vl-remove-if-not '(lambda (d) (= (sx:s (sx:c d sx:cGrp)) gid)) draw))
        (setq feeds (vl-remove-if-not
                      '(lambda (d)
                         (and (sx:content-p d)
                              (/= (sx:s (sx:c d sx:cChT)) "=")
                              (or (sx:c d sx:cCh)
                                  (and (sx:is1 d sx:cXT) (not (sx:is1 d sx:cJoin))))))
                      members))
        (setq chF (vl-remove-if-not '(lambda (d) (sx:c d sx:cCh)) feeds)
              xtF (vl-remove-if '(lambda (d) (sx:c d sx:cCh)) feeds)
              chU (mapcar 'sx:u chF)
              xtU (mapcar 'sx:u xtF))
        ;; N отдельной линией (обход модуля) — только у групп 230 В через модуль
        (sx:draw-qf q u side (and chU (not lowv)))
        ;; входы каналов: перемычки IN-IN с фасками
        (if chU (sx:chain (cons u chU) sx:vIN0 sx:L))
        ;; клеммы без канала: спуск от автомата и перемычки по верху клемм
        (if xtU
          (progn
            (if (not (member u chU))
              (progn (sx:line u sx:vComb u sx:vXTtop sx:L)
                     (sx:ticks u sx:vTickQ (sx:nph q) sx:L)))
            (sx:chain (cons u xtU) sx:vXTtop sx:L)))
        ;; нейтраль: только у групп через модуль (230 В) — спуск слева от автомата, между блоками,
        ;; и вправо по клеммам N. У групп без модуля N показывает сам блок автомата.
        (setq xts (vl-remove-if-not '(lambda (d) (sx:is1 d sx:cXT)) members))
        (if (and chU xts (not lowv))  ; = nbyp
          (progn
            (setq xN (+ (min u (apply 'min chU)) sx:nDx)
                  hi (apply 'max (mapcar 'sx:u xts)))
            (sx:line xN busN xN sx:vXTN sx:LN)
            (sx:dot xN busN sx:LN)
            (sx:hline xN hi sx:vXTN sx:LN)))))))
    (strcat "group " (sx:rowinfo q)))))

;;; ёмкость модулей (модель -> число каналов) из именованного диапазона «Модули_спр»
(defun sx:read-models ( / nms n rr r vals res)
  (setq nms (vl-catch-all-apply 'vlax-get-property (list sx:wb 'Names)))
  (if (not (vl-catch-all-error-p nms))
    (progn
      (setq n (vl-catch-all-apply 'vlax-get-property
                (list nms 'Item (sx:ru '(1052 1086 1076 1091 1083 1080 95 1089 1087 1088)))))
      (if (not (vl-catch-all-error-p n))
        (progn
          (setq rr (vl-catch-all-apply 'vlax-get-property (list n 'RefersToRange)))
          (if (not (vl-catch-all-error-p rr))
            (progn
              (setq vals (vl-catch-all-apply
                           '(lambda () (vlax-safearray->list
                                         (vlax-variant-value (vlax-get-property rr 'Value2))))
                           nil))
              (sx:rel rr)
              (if (not (vl-catch-all-error-p vals))
                (foreach row vals
                  (setq r (mapcar 'sx:v row))
                  (if (and (car r) (numberp (nth 2 r))) (setq res (cons (cons (sx:s (car r)) (fix (nth 2 r))) res)))))))
          (sx:rel n)))
      (sx:rel nms)))
  res)

;;; повёрнутый текст (rot, град. к оси схемы)
(defun sx:textr (u v h s lay rot)
  (setq s (sx:s s))
  (if (/= s "")
    (entmake (list '(0 . "TEXT") (cons 8 lay) (cons 10 (sx:pt u v)) (cons 40 h) (cons 1 s)
                   (cons 50 (+ sx:ang (* pi (/ rot 180.0)))) (cons 7 sx:style)))))

;;; вертикальная линия разрыва (ЕСКД): выходит за контур на 1.5 мм, в середине — зигзаг
(defun sx:breakv (u v1 v2 lay / m)
  (setq m (/ (+ v1 v2) 2.0))
  (sx:line u (- v1 1.5) u (- m 1.6) lay)
  (sx:line u (- m 1.6) (+ u 1.2) (- m 0.6) lay)
  (sx:line (+ u 1.2) (- m 0.6) (- u 1.2) (+ m 0.6) lay)
  (sx:line (- u 1.2) (+ m 0.6) u (+ m 1.6) lay)
  (sx:line u (+ m 1.6) u (+ v2 1.5) lay))

(defun sx:draw-modules (draw / segs key item ds umin umax chs cap xl xr knx)
  (setq segs '() knx '())
  (foreach d draw
    (if (and (sx:has d sx:cMod) (sx:c d sx:cCh) (/= (sx:s (sx:c d sx:cChT)) "="))
      (progn
        (setq key (strcat (sx:s (sx:c d sx:cMod)) "|" (itoa (sx:sheet (sx:i d)))))
        (if (setq item (assoc key segs))
          (setq segs (subst (cons key (append (cdr item) (list d))) item segs))
          (setq segs (cons (list key d) segs))))))
  (foreach item segs
   (sx:try (function (lambda ()
    (setq ds (cdr item)
          umin (apply 'min (mapcar 'sx:u ds))
          umax (apply 'max (mapcar 'sx:u ds))
          chs (mapcar '(lambda (d) (fix (sx:n (sx:c d sx:cCh)))) ds)
          cap (cdr (assoc (sx:s (sx:c (car ds) sx:cModel)) sx:models))
          xl (- umin 7.1) xr (+ umax 6.9))
    ;; контур модуля; если каналы идут дальше / начинаются не с 1-го — линия разрыва
    (sx:line xl sx:vBoxBot xr sx:vBoxBot sx:L)
    (sx:line xl sx:vBoxTop xr sx:vBoxTop sx:L)
    (if (> (apply 'min chs) 1) (sx:breakv xl sx:vBoxBot sx:vBoxTop sx:L) (sx:line xl sx:vBoxBot xl sx:vBoxTop sx:L))
    (if (or (null cap) (< (apply 'max chs) cap))
      (sx:breakv xr sx:vBoxBot sx:vBoxTop sx:L)
      (sx:line xr sx:vBoxBot xr sx:vBoxTop sx:L))
    ;; шина KNX и подписи
    (sx:line (- umin 4.3) sx:vBoxTop (- umin 4.3) sx:vKNX sx:KNX)
    (setq knx (cons (list (- umin 4.3) xr) knx))
    (sx:textr (- umin 4.7) (+ sx:vBoxTop 0.6) 2.2 "KNX" sx:KNX 90.0)
    (sx:text (- umin 6.2) (+ sx:vBoxBot 0.9) 2.5 (sx:s (sx:c (car ds) sx:cMod)) sx:KNX)
    (sx:text (+ umin 3.0) (- sx:vBoxTop 3.4) 2.2 (sx:s (sx:c (car ds) sx:cModel)) sx:KNX)))
    (strcat "module " (car item))))
  ;; шина KNX — одна линия через все модули щита, между листами — разрыв со стрелками
  (if knx
    (sx:try (function (lambda ()
      (sx:hline (apply 'min (mapcar 'car knx)) (apply 'max (mapcar 'cadr knx)) sx:vKNX sx:KNX)))
      "knx bus")))

;;; у строки нет кода УГО: BK = "BOX", а название УГО (BY) пустое (или 0)
(defun sx:ugo-noby (d sym)
  (and (= sym "BOX") (member (sx:s (sx:c d sx:cUgoRu)) '("" "0")) T))

;;; кабельный вывод: линия от клеммы к верху символа + блок «Вывод» (или простой значок SX_U_OUT)
(defun sx:draw-outlet (d u / blk)
  (setq blk (sx:blk sx:bOut "SX_U_OUT"))
  (sx:line u sx:xtb u sx:vOutTop sx:L)
  (sx:insv blk u sx:vOutTop 1.0 0.0 nil sx:BL))

(defun sx:draw-ugo (d u / sym spec blk ob)
  (setq sym (sx:s (sx:c d sx:cSym)))
  (if (= sym "") (setq sym "BOX"))
  (if (or (sx:ugo-noby d sym)
          (and (not (assoc sym sx:ugo)) (not (sx:blk-p sym))))
    ;; нет кода УГО или код неизвестен -> кабельный вывод (не красная заглушка)
    (sx:draw-outlet d u)
    (progn
      (setq spec (cond ((assoc sym sx:ugo))
                       (t (list sym sym 1.0 0.0 68.10 72.40)))   ; свой блок из таблицы значков
            blk (nth 1 spec))
      (if (not (sx:blk-p blk))
        (setq blk (cond ((= sym "SOCKET") "SX_U_SOCKET") ((= sym "LAMP") "SX_U_LAMP")
                        ((member sym '("MOTOR" "SERVO" "VALVE")) "SX_U_M") (t "SX_U_BOX"))
              spec (list sym blk 1.0 0.0 (nth 4 spec) (nth 5 spec))))
      (sx:line u sx:xtb u (nth 5 spec) sx:L)
      (setq ob (sx:insv blk u (nth 4 spec) (nth 2 spec) (nth 3 spec)
                        (if (= sym "BOX") (list (sx:s (sx:c d sx:cLine))) (list "")) sx:BL))
      (if (/= sym "SOCKET") (sx:red ob))
      ob)))

;;; подпись клеммы: справа от кружков, на уровне верхнего, ниже перемычки
(setq sx:xtLblDx 2.2 sx:xtLblDv -1.0)
(defun sx:xt-label (ob u / bb tx ty)
  (if (and ob (= (vla-get-hasattributes ob) :vlax-true))
    (foreach a (vlax-invoke ob 'GetAttributes)
      (setq sx:bbctx (strcat "xt|" (vla-get-name ob) "|" (vla-get-tagstring a)))
      (if (and (/= (vla-get-textstring a) "") (setq bb (sx:attbox a)))
        (progn
          (setq tx (car (sx:pt (+ u sx:xtLblDx) 0.0)) ty (cadr (sx:pt 0.0 (+ sx:vXTtop sx:xtLblDv))))
          (vla-move a (vlax-3d-point (list (car (car bb)) (/ (+ (cadr (car bb)) (cadr (cadr bb))) 2.0) 0.0))
                      (vlax-3d-point (list tx ty 0.0)))))))
  (setq sx:bbctx "")
  ob)

;;; есть ли в определении блока окружность цвета 3 (кружок PE); кэш по имени
(setq sx:pecache nil)
(defun sx:blk-pe-p (name / hit e ed found)
  (setq hit (assoc name sx:pecache))
  (if (null hit)
    (progn
      (setq found nil e (tblobjname "BLOCK" name))
      (while (and e (not found) (setq e (entnext e)))
        (setq ed (entget e))
        (if (and (= (cdr (assoc 0 ed)) "CIRCLE") (= (cdr (assoc 62 ed)) 3)) (setq found T)))
      (setq hit (cons name found) sx:pecache (cons hit sx:pecache))))
  (cdr hit))

;;; если блок клеммы динамический (со свойством видимости) - выставить состояние явно:
;;; обычная клемма - состояние с PE / тремя уровнями, двухуровневая - с "2". Список состояний пишется в журнал.
(defun sx:xt-vis (ob two / p vals pick)
  (if (and ob (setq p (sx:dynprop ob sx:pVis)))
    (progn
      (setq vals (vl-catch-all-apply
                   '(lambda () (vlax-safearray->list (vlax-variant-value (vla-get-allowedvalues p)))) nil))
      (if (vl-catch-all-error-p vals) (setq vals nil))
      (sx:log (strcat "XT visibility states: " (sx:str vals)))
      (setq pick (vl-some '(lambda (v) (if (wcmatch (strcase v) (if two "*2*" "*PE*,*3*")) v)) vals))
      (if pick (sx:dynset ob sx:pVis pick)))))

(defun sx:draw-rows (draw / prev u bf xt two blkn ob xseen)
  (setq prev nil xseen nil)
  (foreach d draw
   (sx:try (function (lambda ()
    (setq u (sx:u d) bf (sx:s (sx:c d sx:cChT)) sx:xtb sx:vXTbot)
    (if (and (sx:c d sx:cCh) (/= bf "="))
      (progn
        (sx:ins "SX_CH" u sx:vChBase (list (cons "CH" (sx:s (sx:c d sx:cCh)))) sx:BL)
        (cond
          ((sx:is1 d sx:cXT) (sx:line u sx:vChBase u sx:vXTtop sx:L))
          ;; канал без клеммы (напр. датчики протечки): сразу линия к УГО
          ((sx:is1 d sx:cHasLine)
           (sx:line u sx:vChBase u sx:vXTbot sx:L)
           (sx:draw-ugo d u)))))
    (if (sx:is1 d sx:cXT)
      (progn
        ;; двухуровневая клемма (XT без земли) - только когда на клемму приходят два проводника:
        ;; параллельная линия «=» или повтор номера XT; обычная клемма - XT с кружком PE
        (setq xt (sx:s (sx:c d sx:cXTtxt))
              two (or (and (= bf "=") prev) (and (/= xt "") (member (strcase xt) xseen)))
              blkn (if two (sx:blk sx:bXT2 "SX_XT2") (sx:blk sx:bXT "SX_XT")))
        (if (/= xt "") (setq xseen (cons (strcase xt) xseen)))
        (if two (setq sx:xtb (+ sx:vXTbot 2.0)))
        (if (and (= bf "=") prev)
          (sx:bridge (sx:u prev) u sx:vXTtop sx:L))
        (setq ob (sx:insv blkn u sx:vXTtop 1.0 0.0 (list xt) sx:BL))
        (sx:xt-vis ob two)
        (sx:xt-label ob u)
        ;; у обычной клеммы должен быть кружок PE; если в блоке XT его нет - добавить отдельный
        (if (and (not two) (not (sx:blk-pe-p blkn)))
          (sx:insv "SX_XTPE" u sx:vXTPE 1.0 0.0 nil sx:BL))
        ;; трёхфазная линия: ещё две клеммы фаз над XT
        (if (= (sx:nph d) 3) (sx:insv "SX_XT3" u sx:vXTtop 1.0 0.0 nil sx:BL))
        ;; отвод PE (у двухуровневой клеммы земли нет)
        (if (not two)
          (progn
            (sx:line (+ u 0.7) (- sx:vXTPE 0.7) (+ u 2.5) (- sx:vXTPE 2.5) sx:PE)
            (sx:line (+ u 2.5) (- sx:vXTPE 2.5) (+ u 2.5) sx:vPE sx:PE)
            (sx:dot (+ u 2.5) sx:vPE sx:PE)))
        (if (and (sx:c d sx:cCh) (sx:is1 d sx:cHasLine) (not (sx:is1 d sx:cJoin)))
          (sx:ticks u sx:vTickLine (sx:nph d) sx:L))
        (cond
          ((and (sx:is1 d sx:cJoin) prev)
           (sx:hline (sx:u prev) u sx:vJoin sx:L)
           (sx:line u sx:xtb u sx:vJoin sx:L))
          ((sx:is1 d sx:cHasLine)
           (sx:draw-ugo d u)))
        (setq prev d)))))
    (strcat "row " (sx:rowinfo d)))))

;;; ---------------------------------------------------------------- диалог
(defun sx:dcl-file ( / f)
  (cond ((setq f (findfile "SX_Schema.dcl")) f)
        ((and (getenv "SX_DclFile") (findfile (getenv "SX_DclFile"))) (getenv "SX_DclFile"))
        ((setq f (getfiled (sx:ru '(1059 1082 1072 1078 1080 1090 1077 32 1092 1072 1081 1083 32 1076 1080 1072 1083 1086 1075 1072 32 83 88 95 83 99 104 101 109 97 46 100 99 108)) "SX_Schema.dcl" "dcl" 0))
         (setenv "SX_DclFile" f) f)))

(defun sx:fill-list (key lst idx)
  (start_list key)
  (mapcar 'add_list lst)
  (end_list)
  (if lst (set_tile key (itoa idx))))

(defun sx:pt-info ()
  (if sx:inspt
    (strcat (sx:ru '(1058 1086 1095 1082 1072 32 1074 1089 1090 1072 1074 1082 1080 58 32)) (rtos (car sx:inspt) 2 1) "; " (rtos (cadr sx:inspt) 2 1))
    (sx:ru '(1058 1086 1095 1082 1072 32 1074 1089 1090 1072 1074 1082 1080 32 1085 1077 32 1091 1082 1072 1079 1072 1085 1072))))

(defun sx:on-sheet (val)
  (setq sx:si (atoi val))
  (sx:load-panels)
  (sx:fill-list "panel" sx:panels sx:pi)
  (set_tile "status" (if sx:panels "" (sx:ru '(1053 1072 32 1083 1080 1089 1090 1077 32 1085 1077 1090 32 1097 1080 1090 1086 1074 32 1074 32 1089 1090 1086 1083 1073 1094 1077 32 65 32 8212 32 1074 1099 1073 1077 1088 1080 1090 1077 32 1083 1080 1089 1090 32 171 1054 1076 1085 1086 1083 1080 1085 1077 1081 1082 1072 187 46)))))

(defun sx:dlg-save ( / n s)
  (setq n (atoi (get_tile "nsheet")))
  (if (> n 0) (setq *sx-sheet* n))
  (setq s (distof (get_tile "step") 2))
  (if (and s (> s 0)) (setq *sx-step* s))
  (setq sx:do-clear (= (get_tile "clear") "1"))
  (setq sx:lib (get_tile "lib"))
  (if (/= (get_tile "path") sx:path) (setq sx:path-typed (get_tile "path"))))

(defun sx:dialog ( / dcl id res path p)
  (if (null (setq dcl (sx:dcl-file))) (exit))
  (setq id (load_dialog dcl) res 5)
  (if (< id 0) (exit))
  (setq sx:status "")
  (while (> res 1)
    (if (not (new_dialog "sx_main" id)) (exit))
    (set_tile "path" sx:path)
    (sx:fill-list "sheet" (mapcar 'car sx:sheets) sx:si)
    (sx:fill-list "panel" sx:panels sx:pi)
    (set_tile "ptinfo" (sx:pt-info))
    (set_tile "nsheet" (itoa *sx-sheet*))
    (set_tile "step" (rtos *sx-step* 2 1))
    (set_tile "clear" (if sx:do-clear "1" "0"))
    (set_tile "lib" (if sx:lib sx:lib ""))
    (action_tile "lib" "(setq sx:lib $value)")
    (action_tile "libbrowse" "(sx:dlg-save) (done_dialog 4)")
    (set_tile "status" sx:status)
    (action_tile "path" "(setq sx:path-typed $value)")
    (action_tile "browse" "(sx:dlg-save) (done_dialog 3)")
    (action_tile "pick" "(sx:dlg-save) (done_dialog 2)")
    (action_tile "sheet" "(sx:on-sheet $value)")
    (action_tile "panel" "(setq sx:pi (atoi $value))")
    (action_tile "accept" "(sx:dlg-save) (done_dialog 1)")
    (action_tile "cancel" "(done_dialog 0)")
    (setq sx:path-typed nil sx:status "")
    (setq res (start_dialog))
    (if (and sx:path-typed (/= sx:path-typed sx:path) (/= sx:path-typed ""))
      (progn
        (if (not (sx:open-wb sx:path-typed)) (setq sx:status (sx:ru '(1053 1077 32 1091 1076 1072 1083 1086 1089 1100 32 1086 1090 1082 1088 1099 1090 1100 32 1092 1072 1081 1083 32 69 120 99 101 108 46))))
        (if (= res 1) (setq res 5))))
    (cond
      ((= res 3)
       (setq path (getfiled (sx:ru '(1060 1072 1081 1083 32 69 120 99 101 108 32 1089 32 1086 1076 1085 1086 1083 1080 1085 1077 1081 1082 1086 1081)) (if (/= sx:path "") sx:path "") "xlsx;xlsm;xls" 0))
       (if path
         (if (sx:open-wb path)
           (setenv "SX_LastXls" path)
           (setq sx:status (sx:ru '(1053 1077 32 1091 1076 1072 1083 1086 1089 1100 32 1086 1090 1082 1088 1099 1090 1100 32 1092 1072 1081 1083 32 69 120 99 101 108 46))))))
      ((= res 4)
       (if (setq path (getfiled (sx:ru '(1063 1077 1088 1090 1105 1078 32 1089 32 1074 1072 1096 1080 1084 1080 32 1073 1083 1086 1082 1072 1084 1080 32 40 1073 1080 1073 1083 1080 1086 1090 1077 1082 1072 41)) (if sx:lib sx:lib "") "dwg" 0))
         (setq sx:lib path)))
      ((= res 2)
       (if (setq p (getpoint (strcat "\n" (sx:ru '(1059 1082 1072 1078 1080 1090 1077 32 1083 1077 1074 1099 1081 32 1085 1080 1078 1085 1080 1081 32 1091 1075 1086 1083 32 1090 1072 1073 1083 1080 1094 1099 32 1089 1093 1077 1084 1099 58 32)))))
         (progn (setq sx:inspt (trans p 1 0))
                (vlax-ldata-put "SX_Schema" "pt" sx:inspt))))
      ((= res 1)
       (cond
         ((null sx:wb) (setq sx:status (sx:ru '(1042 1099 1073 1077 1088 1080 1090 1077 32 1092 1072 1081 1083 32 69 120 99 101 108 46)) res 5))
         ((null sx:inspt) (setq sx:status (sx:ru '(1059 1082 1072 1078 1080 1090 1077 32 1090 1086 1095 1082 1091 32 1074 1089 1090 1072 1074 1082 1080 32 40 1082 1085 1086 1087 1082 1072 32 171 1059 1082 1072 1079 1072 1090 1100 32 1090 1086 1095 1082 1091 187 41 46)) res 5))
         ((null sx:panels) (setq sx:status (sx:ru '(1053 1072 32 1083 1080 1089 1090 1077 32 1085 1077 1090 32 1097 1080 1090 1086 1074 32 1074 32 1089 1090 1086 1083 1073 1094 1077 32 65 32 8212 32 1074 1099 1073 1077 1088 1080 1090 1077 32 1083 1080 1089 1090 32 171 1054 1076 1085 1086 1083 1080 1085 1077 1081 1082 1072 187 46)) res 5))))))
  (unload_dialog id)
  (= res 1))

;;; ---------------------------------------------------------------- проверка перед рисованием
;;; блоки, которые понадобятся для строк к рисованию (без служебных SX_*)
(defun sx:need-blocks (draw / res spec sym)
  (setq res (list sx:bQF sx:bBUS sx:bHEAD sx:bXT))
  (foreach d draw
    (foreach dv (sx:dev-list d)
      (setq spec (sx:dev-spec (car dv)))
      (if (not (member (nth 1 spec) res)) (setq res (append res (list (nth 1 spec))))))
    (if (sx:is1 d sx:cHasLine)
      (progn
        (setq sym (sx:s (sx:c d sx:cSym)))
        (if (= sym "") (setq sym "BOX"))
        (if (sx:ugo-noby d sym)
          (setq sym sx:bOut)   ; нет кода УГО -> кабельный вывод
          (progn
            (setq spec (assoc sym sx:ugo))
            (if spec (setq sym (nth 1 spec)))))
        (if (not (member sym res)) (setq res (append res (list sym)))))))
  (vl-remove-if '(lambda (n) (wcmatch (strcase n) "SX_*")) res))

(defun sx:join-sep (lst sep / res)
  (setq res "")
  (foreach x lst (setq res (if (= res "") x (strcat res sep x))))
  res)

;;; строки без номера места (заполняет sx:main): (номер строки Excel ...)
(setq sx:nopos nil)

;;; предупреждения о возможных наложениях линий (список строк); ничего не рисует
(defun sx:validate-layout (recs draw / res qfids id pl lo hi key seenk prevx)
  (setq res '() qfids '() seenk '())
  (foreach x sx:nopos
    (setq res (cons (strcat (sx:ru '(1089 1090 1088 1086 1082 1072 32 69 120 99 101 108 32)) (itoa x) (sx:ru '(58 32 1077 1089 1090 1100 32 1076 1072 1085 1085 1099 1077 44 32 1085 1086 32 1085 1077 32 1079 1072 1076 1072 1085 32 1085 1086 1084 1077 1088 32 1084 1077 1089 1090 1072 32 45 32 1089 1090 1088 1086 1082 1072 32 1085 1077 32 1073 1091 1076 1077 1090 32 1085 1072 1088 1080 1089 1086 1074 1072 1085 1072))) res)))
  (foreach d draw
    (if (and (sx:has d sx:cQFown) (not (member (sx:s (sx:c d sx:cQFown)) qfids)))
      (setq qfids (append qfids (list (sx:s (sx:c d sx:cQFown)))))))
  ;; (1) клемма или канал без нарисованного автомата выше
  (foreach d draw
    (if (and (not (sx:has d sx:cQFown)) (or (sx:c d sx:cCh) (sx:is1 d sx:cXT)))
      (if (not (member (sx:s (sx:c d sx:cGrp)) qfids))
        (setq res (cons (strcat (sx:ru '(1089 1090 1088 1086 1082 1072 32 69 120 99 101 108 32)) (itoa (sx:xr d)) (sx:ru '(44 32 1084 1077 1089 1090 1086 32)) (itoa (1+ (sx:i d)))
                                (sx:ru '(58 32 1082 1083 1077 1084 1084 1072 32 1080 1083 1080 32 1082 1072 1085 1072 1083 32 1073 1077 1079 32 1072 1074 1090 1086 1084 1072 1090 1072 32 1074 1099 1096 1077 32 45 32 1083 1080 1085 1080 1103 32 1087 1086 1074 1080 1089 1085 1077 1090 32 1073 1077 1079 32 1087 1080 1090 1072 1085 1080 1103 59 32 1074 1087 1080 1096 1080 1090 1077 32 1072 1074 1090 1086 1084 1072 1090 32 1074 32 69 32 1087 1077 1088 1074 1086 1081 32 1089 1090 1088 1086 1082 1080 32 1075 1088 1091 1087 1087 1099 32 1080 1083 1080 32 1087 1086 1089 1090 1072 1074 1100 1090 1077 32 1074 32 89 32 171 1074 1088 1091 1095 1085 1091 1102 187))) res)))))
  ;; (2) группы перемешаны: между местами группы стоит место из другой группы (перемычки лягут на чужие линии)
  (foreach id qfids
    (setq pl (mapcar 'sx:i (vl-remove-if-not '(lambda (d) (and (sx:content-p d) (= (sx:s (sx:c d sx:cGrp)) id))) draw)))
    (if (> (length pl) 1)
      (progn
        (setq lo (apply 'min pl) hi (apply 'max pl))
        (foreach d draw
          (if (and (sx:content-p d) (> (sx:i d) lo) (< (sx:i d) hi) (/= (sx:s (sx:c d sx:cGrp)) id))
            (progn
              (setq key (strcat id "|" (itoa (sx:i d))))
              (if (not (member key seenk))
                (setq seenk (cons key seenk)
                      res (cons (strcat (sx:ru '(1075 1088 1091 1087 1087 1072 32)) id (sx:ru '(32 40 1084 1077 1089 1090 1072 32)) (itoa (1+ lo)) "-" (itoa (1+ hi)) (sx:ru '(41 58 32 1084 1077 1078 1076 1091 32 1085 1080 1084 1080 32 1084 1077 1089 1090 1086 32)) (itoa (1+ (sx:i d)))
                                        (sx:ru '(32 40 1089 1090 1088 1086 1082 1072 32 69 120 99 101 108 32)) (itoa (sx:xr d)) (sx:ru '(41 32 1080 1079 32 1076 1088 1091 1075 1086 1081 32 1075 1088 1091 1087 1087 1099 32 1080 1083 1080 32 1073 1077 1079 32 1072 1074 1090 1086 1084 1072 1090 1072 32 45 32 1087 1077 1088 1077 1084 1099 1095 1082 1080 32 1083 1103 1075 1091 1090 32 1085 1072 32 1077 1075 1086 32 1083 1080 1085 1080 1080 59 32 1084 1077 1089 1090 1072 32 1086 1076 1085 1086 1081 32 1075 1088 1091 1087 1087 1099 32 1076 1086 1083 1078 1085 1099 32 1080 1076 1090 1080 32 1087 1086 1076 1088 1103 1076))) res)))))))))
  ;; (3) «=» стоит не рядом с предыдущей клеммой
  (setq prevx nil)
  (foreach d draw
    (if (sx:is1 d sx:cXT)
      (progn
        (if (and prevx (= (sx:s (sx:c d sx:cChT)) "=") (/= (- (sx:i d) (sx:i prevx)) 1))
          (setq res (cons (strcat (sx:ru '(1089 1090 1088 1086 1082 1072 32 69 120 99 101 108 32)) (itoa (sx:xr d)) (sx:ru '(44 32 1084 1077 1089 1090 1086 32)) (itoa (1+ (sx:i d)))
                                  (sx:ru '(58 32 171 61 187 32 1089 1090 1086 1080 1090 32 1085 1077 32 1088 1103 1076 1086 1084 32 1089 32 1087 1088 1077 1076 1099 1076 1091 1097 1077 1081 32 1082 1083 1077 1084 1084 1086 1081 32 40 1084 1077 1089 1090 1086 32)) (itoa (1+ (sx:i prevx)))
                                  (sx:ru '(41 32 45 32 1087 1077 1088 1077 1084 1099 1095 1082 1072 32 1087 1077 1088 1077 1089 1077 1095 1105 1090 32 1095 1091 1078 1080 1077 32 1083 1080 1085 1080 1080 59 32 171 61 187 32 1076 1086 1083 1078 1085 1072 32 1080 1076 1090 1080 32 1089 1088 1072 1079 1091 32 1079 1072 32 1089 1074 1086 1077 1081 32 1083 1080 1085 1080 1077 1081))) res)))
        (setq prevx d))))
  (reverse res))

;;; список проблем (строки) или nil; ничего не рисует и не меняет в чертеже
(defun sx:validate (panel recs draw / probs miss0 found reasons st seenpos seenxt item key)
  (setq probs '())
  (cond ((or (null panel) (null recs))
         (setq probs (cons (strcat (sx:ru '(1097 1080 1090 32 1085 1077 32 1085 1072 1081 1076 1077 1085 32 1074 32 1076 1072 1085 1085 1099 1093 58 32)) (sx:s panel)) probs)))
        ((null draw)
         (setq probs (cons (sx:ru '(1085 1077 1090 32 1089 1090 1088 1086 1082 32 1082 32 1088 1080 1089 1086 1074 1072 1085 1080 1102 32 40 100 114 97 119 61 49 41)) probs))))
  (if draw
    (progn
      ;; (а) блоки: нет ни в чертеже, ни в библиотеке -> заглушки SX_*
      (setq miss0 (vl-remove-if '(lambda (n) (tblsearch "BLOCK" n)) (sx:need-blocks draw)))
      (if miss0
        (progn
          (setq found (if (and sx:lib (/= sx:lib "")) (sx:lib-found sx:lib miss0)))
          (foreach n miss0
            (if (not (member n found))
              (setq probs (cons (strcat (sx:ru '(1073 1083 1086 1082 32 1085 1077 32 1085 1072 1081 1076 1077 1085 32 1074 32 1095 1077 1088 1090 1077 1078 1077 32 1080 32 1073 1080 1073 1083 1080 1086 1090 1077 1082 1077 44 32 1073 1091 1076 1077 1090 32 1079 1072 1075 1083 1091 1096 1082 1072 58 32)) n) probs))))))
      ;; (б) строки: статус AE, фаза и мощность у строки с линией
      (foreach d draw
        (setq reasons '() st (sx:s (sx:c d sx:cStat)))
        (if (/= st "") (setq reasons (list (strcat (sx:ru '(1089 1090 1072 1090 1091 1089 58 32)) st))))
        (if (sx:has d sx:cLine)
          (progn
            (if (not (sx:has d sx:cPh)) (setq reasons (append reasons (list (sx:ru '(1085 1077 1090 32 1092 1072 1079 1099))))))
            (if (not (sx:has d sx:cPow)) (setq reasons (append reasons (list (sx:ru '(1085 1077 1090 32 1084 1086 1097 1085 1086 1089 1090 1080))))))))
        (if reasons
          (setq probs (cons (strcat (sx:ru '(1089 1090 1088 1086 1082 1072 32 69 120 99 101 108 32)) (itoa (sx:xr d)) (sx:ru '(44 32 1084 1077 1089 1090 1086 32)) (itoa (1+ (sx:i d))) ": "
                                    (sx:join-sep reasons ", "))
                            probs))))
      ;; дубли номеров мест (по всем строкам: колонка таблицы строится у каждой) и клемм (по строкам к рисованию)
      (setq seenpos '() seenxt '())
      (foreach d recs
        (setq key (sx:i d) item (assoc key seenpos))
        (if item
          (setq seenpos (subst (append item (list (sx:xr d))) item seenpos))
          (setq seenpos (cons (list key (sx:xr d)) seenpos))))
      (foreach d draw
        (if (and (sx:is1 d sx:cXT) (sx:has d sx:cXTtxt))
          (progn
            (setq key (strcase (sx:s (sx:c d sx:cXTtxt))) item (assoc key seenxt))
            (if item
              (setq seenxt (subst (append item (list (sx:xr d))) item seenxt))
              (setq seenxt (cons (list key (sx:xr d)) seenxt))))))
      (foreach item (reverse seenpos)
        (if (> (length item) 2)
          (setq probs (cons (strcat (sx:ru '(1076 1074 1072 32 1084 1077 1089 1090 1072 32 1089 32 1086 1076 1085 1080 1084 32 1085 1086 1084 1077 1088 1086 1084 32)) (itoa (1+ (car item))) (sx:ru '(32 40 1089 1090 1088 1086 1082 1080 32 69 120 99 101 108 32))
                                    (sx:join-sep (mapcar 'itoa (cdr item)) ", ")
                                    (sx:ru '(41 58 32 1083 1080 1085 1080 1080 32 1080 32 1090 1072 1073 1083 1080 1094 1072 32 1083 1103 1075 1091 1090 32 1076 1088 1091 1075 32 1085 1072 32 1076 1088 1091 1075 1072 59 32 1088 1072 1079 1085 1077 1089 1080 1090 1077 32 1087 1086 32 1088 1072 1079 1085 1099 1084 32 1084 1077 1089 1090 1072 1084 32 1080 1083 1080 32 1087 1086 1089 1090 1072 1074 1100 1090 1077 32 1074 32 89 32 171 1074 1088 1091 1095 1085 1091 1102 187 32 1083 1080 1096 1085 1077 1081 32 1089 1090 1088 1086 1082 1077)))
                            probs))))
      (foreach item (reverse seenxt)
        (if (> (length item) 2)
          (setq probs (cons (strcat (sx:ru '(1076 1091 1073 1083 1100 32 1082 1083 1077 1084 1084 1099 32)) (car item) (sx:ru '(32 40 1089 1090 1088 1086 1082 1080 32))
                                    (sx:join-sep (mapcar 'itoa (cdr item)) ", ") ")")
                            probs))))
      ;; наложения линий: клемма без автомата, перемешанные группы, «=» не рядом, строки без номера места
      (foreach p (sx:validate-layout recs draw) (setq probs (cons p probs)))))
  (reverse probs))

;;; вывод в командную строку (первые 30) и в журнал (все)
(defun sx:report (probs / k)
  (princ (strcat "\n" (sx:ru '(1055 1088 1086 1074 1077 1088 1082 1072 32 1087 1077 1088 1077 1076 32 1088 1080 1089 1086 1074 1072 1085 1080 1077 1084 44 32 1085 1072 1081 1076 1077 1085 1086 32 1087 1088 1086 1073 1083 1077 1084 58 32)) (itoa (length probs))))
  (sx:log (strcat "CHECK problems: " (itoa (length probs))))
  (setq k 0)
  (foreach p probs
    (if (< k 30) (princ (strcat "\n  " p)))
    (sx:log (strcat "CHECK " p))
    (setq k (1+ k)))
  (if (> k 30) (princ (strcat "\n  ... +" (itoa (- k 30)) " (see SX_log.txt)"))))

;;; T = рисовать (Enter = Yes)
(defun sx:ask-go ( / r)
  (initget "Yes No")
  (setq r (getkword (strcat "\n" (sx:ru '(1056 1080 1089 1086 1074 1072 1090 1100 32 1074 1089 1105 32 1088 1072 1074 1085 1086 63 32 91 1044 1072 47 1053 1077 1090 93 32 60 1044 1072 62 32 40 89 101 115 47 78 111 41 58 32)))))
  (/= r "No"))

;;; ---------------------------------------------------------------- команда
(if (null sx:do-clear) (setq sx:do-clear T))

(defun c:SXDRAW () (sx:main nil))
(defun c:SXUPDATE () (sx:main T))

;;; параметры построения, запомненные в самом чертеже (для SXUPDATE)
(defun sx:save-params (panel)
  (vlax-ldata-put "SX_Schema" "xls" sx:path)
  (vlax-ldata-put "SX_Schema" "panel" panel)
  (vlax-ldata-put "SX_Schema" "sheet" (car (nth sx:si sx:sheets)))
  (if sx:lib (vlax-ldata-put "SX_Schema" "lib" sx:lib)))

;;; SXUPDATE: взять всё из чертежа без окна; T = готово к построению
(defun sx:load-params ( / xls pan sht p)
  (setq xls (vlax-ldata-get "SX_Schema" "xls")
        pan (vlax-ldata-get "SX_Schema" "panel")
        sht (vlax-ldata-get "SX_Schema" "sheet"))
  (if (vlax-ldata-get "SX_Schema" "lib") (setq sx:lib (vlax-ldata-get "SX_Schema" "lib")))
  (cond
    ((or (null xls) (null pan) (null sx:inspt))
     (princ "\nSXUPDATE: no saved settings in this drawing - run SXDRAW once.") nil)
    ((not (findfile xls))
     (princ (strcat "\nSXUPDATE: file not found: " xls)) nil)
    ((not (sx:open-wb xls))
     (princ (strcat "\nSXUPDATE: cannot open " xls)) nil)
    (T
     (if (and sht (setq p (assoc sht sx:sheets)))
       (progn (setq sx:si (vl-position p sx:sheets)) (sx:load-panels)))
     (if (member pan sx:panels)
       (progn (setq sx:pi (vl-position pan sx:panels))
              (princ (strcat "\nSXUPDATE: " xls " / " pan)) T)
       (progn (princ (strcat "\nSXUPDATE: panel not found in table: " pan)) nil)))))

(defun sx:main (upd / *error* data recs panel pos i row draw oldcmd maxi sym probs xrow ridx)
  (defun *error* (msg)
    (if sx:doc (vl-catch-all-apply 'vla-endundomark (list sx:doc)))
    (if oldcmd (setvar "CMDECHO" oldcmd))
    (vl-catch-all-apply 'sx:xl-release nil)
    (vl-catch-all-apply 'sx:lib-close nil)
    (if (not (member msg '("Function cancelled" "quit / exit abort")))
      (progn
        (sx:log (strcat "FATAL at [" sx:stage "]: " msg))
        (princ (strcat "\nSXDRAW error at [" sx:stage "]: " msg))
        (vl-catch-all-apply 'sx:logdump nil)))
    (princ))
  (setq sx:logl nil sx:errs nil
        sx:bbcache nil sx:cccache nil sx:blkcache nil sx:bbctx "" sx:pecache nil sx:nopos nil)
  (sx:stg "init")
  (sx:log (strcat "SX_Schema " sx:ver))
  (sx:log (strcat "LISPSYS=" (sx:str (getvar "LISPSYS")) " ACADVER=" (getvar "ACADVER")
                  " DWG=" (getvar "DWGPREFIX") (getvar "DWGNAME")))
  (setq sx:doc (vla-get-activedocument (vlax-get-acad-object))
        sx:ms (vla-get-modelspace sx:doc)
        oldcmd (getvar "CMDECHO"))
  (setvar "CMDECHO" 0)
  (setq sx:inspt (vlax-ldata-get "SX_Schema" "pt"))
  (if (null sx:lib) (setq sx:lib (getenv "SX_LibDwg")))
  (if upd
    (progn
      (sx:stg "update params")
      (if (not (sx:load-params)) (progn (sx:xl-release) (setvar "CMDECHO" oldcmd) (exit))))
    (progn
      (sx:stg "open last xlsx")
      (if (and (getenv "SX_LastXls") (findfile (getenv "SX_LastXls")))
        (sx:open-wb (getenv "SX_LastXls")))
      (sx:stg "dialog")
      (if (not (sx:dialog)) (progn (sx:xl-release) (setvar "CMDECHO" oldcmd) (exit)))))
  (sx:log (strcat "path=" (sx:str sx:path) " sheet#=" (sx:str sx:si) " panels=" (sx:str sx:panels)
                  " pi=" (sx:str sx:pi) " pt=" (sx:str sx:inspt) " lib=" (sx:str sx:lib)
                  " nsheet=" (sx:str *sx-sheet*) " step=" (sx:str *sx-step*)))
  (sx:stg "after dialog")
  (setq panel (nth sx:pi sx:panels))
  (if panel (setenv "SX_LastPanel" panel))
  (setenv "SX_LastXls" sx:path)
  (sx:save-params panel)
  ;; система координат: вдоль X, вверх Y
  (setq sx:P0 sx:inspt sx:R '(1.0 0.0) sx:D '(0.0 1.0) sx:ang 0.0)
  (setq sx:style (sx:ensure-style))
  (sx:log (strcat "style=" sx:style))
  (sx:stg "read excel")
  (setq data (sx:read-data))
  (setq sx:models (vl-catch-all-apply 'sx:read-models nil))
  (if (vl-catch-all-error-p sx:models) (setq sx:models nil))
  (sx:log (strcat "models=" (sx:str sx:models)))
  (sx:xl-release)
  (sx:log (strcat "rows read=" (itoa (length data))))
  ;; записи: (i u row), i = № места - 1
  (sx:stg "filter rows")
  (setq recs '() ridx 0)
  (foreach row data
    (setq xrow (+ sx:rowBase ridx) ridx (1+ ridx))
    (setq pos (nth (1- sx:cPos) row))
    (if (and (= (sx:s (nth (1- sx:cPanel) row)) panel) (numberp pos))
      (progn
        (setq i (1- (fix pos)))
        (setq recs (cons (list i (+ (* i *sx-step*) (/ *sx-step* 2.0)) row xrow) recs)))
      ;; строка щита с данными, но без номера места: не попадёт на чертёж
      (if (and (= (sx:s (nth (1- sx:cPanel) row)) panel) (not (numberp pos))
               (or (/= (sx:s (nth (1- sx:cLine) row)) "") (/= (sx:s (nth 4 row)) "")))
        (setq sx:nopos (append sx:nopos (list xrow))))))
  (setq recs (vl-sort recs '(lambda (a b) (< (car a) (car b)))))
  (setq draw (vl-remove-if-not '(lambda (d) (sx:is1 d sx:cDraw)) recs))
  (sx:log (strcat "recs=" (itoa (length recs)) " draw=" (itoa (length draw))))
  (if recs (sx:log (strcat "first " (sx:rowinfo (car recs)))))
  ;; проверка перед рисованием: все проблемы разом, до очистки и построения
  (sx:stg "validate")
  (sx:roles)
  (setq probs (sx:validate panel recs draw))
  (if probs
    (progn
      (sx:report probs)
      (if (not (sx:ask-go))
        (progn
          (sx:log "check: cancelled by user")
          (sx:lib-close)
          (setvar "CMDECHO" oldcmd)
          (sx:logdump)
          (princ (strcat "\n" (sx:ru '(1054 1090 1084 1077 1085 1077 1085 1086 32 1087 1086 1089 1083 1077 32 1087 1088 1086 1074 1077 1088 1082 1080 44 32 1095 1077 1088 1090 1105 1078 32 1085 1077 32 1080 1079 1084 1077 1085 1105 1085 46))))
          (exit))))
    (sx:log "check: OK"))
  ;; рисуем
  (vla-startundomark sx:doc)
  (sx:stg "clear") (if sx:do-clear (sx:clear))
  (sx:stg "layers") (sx:layers)
  (sx:stg "roles") (sx:roles)
  (sx:stg "make-blocks") (sx:make-blocks T)
  (if (and sx:lib (/= sx:lib "")) (setenv "SX_LibDwg" sx:lib))
  (sx:stg "import-blocks")
  (setq sx:custom nil)
  (foreach d draw
    (setq sym (sx:s (sx:c d sx:cSym)))
    (if (and (/= sym "") (not (assoc sym sx:ugo)) (not (member sym sx:custom)))
      (setq sx:custom (cons sym sx:custom))))
  (if sx:custom (sx:log (strcat "custom UGO blocks: " (sx:str sx:custom))))
  (sx:try (function (lambda ()
    (sx:import-blocks sx:lib
      (append (list sx:bQF sx:bBUS sx:bHEAD sx:bXT sx:bXT2 sx:bOut sx:bT sx:bPS sx:bMK) (mapcar 'cadr sx:ugo) sx:custom))))
    "import-blocks")
  (sx:lib-close)
  (setq sx:blkcache nil)
  (foreach n (append (list sx:bQF sx:bBUS sx:bHEAD sx:bXT sx:bXT2 sx:bOut sx:bT sx:bPS sx:bMK) (mapcar 'cadr sx:ugo))
    (sx:log (strcat "block " n (if (tblsearch "BLOCK" n) " OK" " MISSING"))))
  (if (not (tblsearch "BLOCK" sx:bQF))
    (progn (princ (strcat "\n" (sx:ru '(1042 1085 1080 1084 1072 1085 1080 1077 58 32 1085 1077 32 1085 1072 1081 1076 1077 1085 1099 32 1074 1072 1096 1080 32 1073 1083 1086 1082 1080 32 40 81 70 49 44 32 88 84 46 46 46 41 32 8212 32 1085 1072 1088 1080 1089 1086 1074 1072 1085 1099 32 1079 1072 1084 1077 1085 1080 1090 1077 1083 1080 46 32 1055 1088 1086 1074 1077 1088 1100 1090 1077 32 1087 1091 1090 1100 32 1082 32 1073 1080 1073 1083 1080 1086 1090 1077 1082 1077 32 48 50 95 1054 1076 1085 1086 1083 1080 1085 1077 1081 1082 1072 46 100 119 103 44 32 1089 1084 46 32 83 88 95 108 111 103 46 116 120 116))))
           (setq sx:errs (cons "blocks: QF1 missing" sx:errs))))
  (sx:stg "table")   (sx:try (function (lambda () (sx:draw-table recs))) "table")
  (sx:stg "sheets")  (sx:try (function (lambda () (sx:draw-sheets draw panel))) "sheets")
  (sx:stg "groups")  (sx:try (function (lambda () (sx:draw-groups draw))) "groups")
  (sx:stg "modules") (sx:try (function (lambda () (sx:draw-modules draw))) "modules")
  (sx:stg "rows")    (sx:try (function (lambda () (sx:draw-rows draw))) "rows")
  (sx:stg "done")
  (vla-endundomark sx:doc)
  (setvar "CMDECHO" oldcmd)
  (princ (strcat "\n" (sx:ru '(1043 1086 1090 1086 1074 1086 58 32 1089 1093 1077 1084 1072 32 1087 1086 1089 1090 1088 1086 1077 1085 1072 32 1085 1072 32 1089 1083 1086 1103 1093 32 83 88 95 42 46 32 1052 1077 1089 1090 58)) " " (itoa (length recs))))
  (if sx:errs
    (progn
      (princ (strcat "\nErrors: " (itoa (length sx:errs))))
      (foreach e (reverse sx:errs) (princ (strcat "\n  " e)))))
  (sx:logdump)
  (princ))

(princ (strcat "\n" (sx:ru '(83 88 95 83 99 104 101 109 97 32 118 51 46 49 51 32 1079 1072 1075 1088 1091 1078 1077 1085 32 40 1074 1072 1096 1080 32 1073 1083 1086 1082 1080 41 58 32 83 88 68 82 65 87 32 8212 32 1087 1086 1089 1090 1088 1086 1080 1090 1100 32 1089 1093 1077 1084 1091 44 32 83 88 67 76 69 65 82 32 8212 32 1089 1090 1077 1088 1077 1090 1100 46))))
(princ "\n[SX_Schema v3.13] SXDRAW / SXUPDATE / SXCLEAR")
(princ)
