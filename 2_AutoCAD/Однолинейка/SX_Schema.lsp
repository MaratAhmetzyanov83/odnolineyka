;;; ============================================================================
;;;  SX_Schema.lsp  v3.23 —  однолинейная схема из Excel на ВАШИХ блоках
;;;
;;;  Команды:
;;;    SXDRAW    — окно: файл Excel, лист «Однолинейка», щит, точка вставки,
;;;                библиотека блоков (DWG) -> строит схему целиком.
;;;    SXUPDATE  — перестроить щит целиком по сохранённым в чертеже настройкам.
;;;    SXUPDSEL  — перестроить только выделенные места (столбцы) схемы: выберите
;;;                любые объекты нужных мест; связанные автоматы групп и модули
;;;                перерисуются вместе с ними; итоги щита обновляются всегда.
;;;                При выборе ввода обновляются и шины/границы его листа.
;;;                Работает по меткам объектов (xdata SX_SCHEMA), которые
;;;                ставят SXDRAW / SXUPDATE начиная с v3.15. Выделена шина или
;;;                граница листа — перестраивается весь этот лист.
;;;    SXUPDSHEET — перестроить целиком лист(ы) схемы, на которых выделены объекты
;;;                (шины, границы, все места, автоматы, модули; на 1-м листе и боковик).
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
(setq sx:ver "v3.23")


;;; ---------------------------------------------------------------- настройки
(if (null *sx-sheet*) (setq *sx-sheet* 17))   ; мест (столбцов) на один лист A3
(if (null *sx-step*)  (setq *sx-step* 20.0))  ; шаг столбца, мм
(setq *sx-range* "A3:CK1000")                  ; данные на листе «Однолинейка» (BY = «УГО (рус)», BZ = «слаботочн.», CE:CI = узел КМ, CJ:CK = ДСУП)
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
;;; ДСУП (v3.23, образец «Однолинейка ЩР-2», лист 5): линии РЕ / КУП идут от той же шины PE (sx:vPE),
;;; без автомата и клемм. Точка РЕ после КУП — отвод от линии КУП: горизонталь sx:vDsStep с фасками 2 мм.
(setq sx:vDsStep (+ 119.23 sx:dLow)   ; отвод от КУП (ниже шины PE, выше нижней границы щита)
      sx:vDsOut 72.16                 ; верх вывода точки РЕ
      sx:vKupTop 71.35 sx:vKupIns 68.23 sx:kupSc 1.66 sx:kupDx 0.03)   ; блок «_КУП»: верх, вставка, масштаб (поворот 90°)
(setq sx:xtb sx:vXTbot)   ; низ клеммы текущей строки (у двухуровневой клеммы выше на 2 мм)
;;; ввод: геометрия кабельной петли и отступ щита из образца 09.10.2026.
;;; Питающий щит слева — пользовательские блоки, его аппарат не выводится из QF ввода.
(setq sx:inDx 80.0 sx:inLow 71.70 sx:inFeed 159.66 sx:inLeft 45.93 sx:inBus 6.63)
;;; ввод по образцу «ЩР-1.1» (v3.21): над вводным аппаратом — рамка питающего щита (шина, автомат фидера),
;;; кабель фидера вниз в QS, от QS петля вниз-вправо-вверх к началу шины щита.
;;; Первый ввод листа — рамка от верха листа (наша граница обходит её), остальные — ниже основной шины.
(setq sx:fdL 25.0 sx:fdR 11.0       ; рамка питающего щита: влево / вправо от столбца ввода
      sx:fdD1 40.0 sx:fdD2 36.0     ; длина автомата фидера (шина -> низ): 1-й ввод листа / следующие
      sx:fdQsBot 123.5 sx:fdLoop 121.5   ; низ QS и петля к шине (выше шины PE и клемм)
      sx:fdTxt1 7.0 sx:fdTxt2 11.0) ; подписи кабеля фидера (повёрнуты) справа от столбца

;;; столбцы листа «Однолинейка» (A = 1)
(setq sx:cPanel 1 sx:cPos 2 sx:cBusName 4 sx:cQFtxt 8 sx:cDev1 9
      sx:cCab 17 sx:cLine 18 sx:cName 19 sx:cPow 20 sx:cI 21 sx:cPh 22
      sx:cGrp 33 sx:cMod 34 sx:cModel 35 sx:cQFown 36 sx:cCh 15 sx:cXTtxt 16
      sx:cJoin 40 sx:cXT 51
      sx:cDraw 56 sx:cBus2 57 sx:cRcd 58 sx:cChT 59 sx:cDevC1 60 sx:cSym 63 sx:cHasLine 64
      sx:cUgoRu 77    ; BY: название УГО по-русски (пусто = у линии нет кода УГО)
      sx:cLow 78      ; BZ: 1 = слаботочная линия (DALI, KNX, датчики): клемма без PE
      ;; узел КМ (АПС), patch_km.py: CE имя шины KM-n, CG 1,1 x Iр шины, CH контактор, CI проверка
      sx:cKM 83 sx:cKMcur 85 sx:cKMmod 86 sx:cKMchk 87
      ;; ДСУП, patch_dsup.py: CJ 1 = линия от шины PE (РЕ, КУП), CK = от какой КУП (пусто = от шины)
      sx:cCode 38 sx:cDsup 88 sx:cDsupFrom 89)
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

;;; «Ввод.1» / код «Ввод» в AL: отдельная питающая линия, не отходящее УГО BOX.
(defun sx:input-p (d / code line)
  (setq code (strcase (sx:s (sx:c d 38)))
        line (strcase (sx:s (sx:c d sx:cLine))))
  (or (= code (strcase (sx:ru '(1042 1074 1086 1076))))
      (wcmatch line (strcat (strcase (sx:ru '(1042 1074 1086 1076))) ".*"))))

;;; Сдвиг левого края только если перед вводом оставлены свободные места, как в образце.
(defun sx:sheet-input (draw s / ds q us)
  (setq ds (vl-remove-if-not '(lambda (d) (= (sx:sheet (sx:i d)) s)) draw)
        q (car (vl-remove-if-not 'sx:input-p ds)) us (* s *sx-sheet* *sx-step*))
  (if (and q (>= (- (sx:u q) us) (+ sx:inDx 10.0))
           (not (vl-some '(lambda (d) (and (< (sx:u d) (sx:u q)) (sx:content-p d))) ds)))
    q))

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
        sx:bKUP    (sx:ru '(95 1050 1059 1055))   ; «_КУП» — коробка уравнивания потенциалов (ДСУП)
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
          (list "PS" "SX_PSU" 1.0 T 12.0)
          (list "PK" "SX_PSK" 1.0 T 12.0)
          (list "DIM" "SX_DIM4" 1.0 T 12.0)
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

;;; геометрия «преобразователя»: выводы, квадрат 9x9, диагональ, «~» слева сверху;
;;; dim = T — диммер: вместо «=» стрелка регулирования через знак и число каналов «4»
(defun sx:conv-ents (dim)
  (append
    (list (bL 0.0 0.0 0.0 -1.5) (bL 0.0 -10.5 0.0 -12.0))
    (bR -4.5 -10.5 4.5 -1.5)
    (list (bA -2.9 -3.9 0.8 0.0 pi) (bA -1.3 -3.9 0.8 pi (* 2 pi)))
    (if dim
      (list (bL -3.8 -9.8 3.8 -2.2) (bL 3.8 -2.2 2.3 -2.4) (bL 3.8 -2.2 3.6 -3.7)
            (bT 2.4 -8.9 2.2 "4" nil 0.0))
      (list (bL -4.5 -10.5 4.5 -1.5) (bL 1.0 -7.4 3.4 -7.4) (bL 1.0 -8.4 3.4 -8.4)))
    (list (bAtt "POS" -5.6 -4.6 2.5 "R" 0.0 nil) (bAtt "N" -5.6 -8.4 2.0 "R" 0.0 nil))))

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
  ;; выключатель-разъединитель QS в стиле QF1 (нож и подписи как у QF1; у неподвижного контакта
  ;; кружок и черта, как в виде QSD; без расцепителя и УЗО). Вставка — как QF1: низ на -54.31,
  ;; верхний вывод (-36.27) до точки подключения тянет программа.
  (sx:defblock "SX_QS"
    (list (bL 0.0 -54.31 0.0 -42.79) (bL 0.0 -42.79 -3.51 -36.72)
          (bC 0.0 -36.77 0.5) (bL -0.5 -36.27 0.5 -36.27)
          (bT -6.21 -27.4 2.5 "QS" "L" 0.0)
          (bAtt "N" -3.15 -27.4 2.5 "L" 0.0 nil) (bAtt "P" -3.4 -30.43 2.5 "R" 0.0 nil)
          (bAtt "A" -1.7 -33.28 2.5 "R" 0.0 nil)) force)
  (sx:defblock "SX_T" (list (bL 0.0 0.0 0.0 -2.5) (bC 0.0 -5.0 2.5) (bC 0.0 -8.8 2.5) (bL 0.0 -11.3 0.0 -13.8)
                            (bAtt "N" -6.2 -5.5 2.5 "L" 0.0 nil) (bAtt "W" -10.3 -9.1 2.5 "L" 0.0 nil)
                            (bAtt "V" -8.8 -12.9 2.5 "L" 0.0 nil)) force)
  (sx:defblock "SX_PS" (append (list (bL 0.0 0.0 0.0 2.5) (bL 0.0 11.2 0.0 13.7)) (bR -2.2 2.5 2.2 11.2)
                               (list (bT 0.0 6.85 2.0 "=" "M" 0.0) (bAtt "N" -4.8 8.3 2.5 "L" 0.0 nil))) force)
  ;; блок питания (ГОСТ 2.721: преобразователь ~/=), БП KNX, диммер 4-канальный: вставка сверху, высота 12
  (sx:defblock "SX_PSU" (sx:conv-ents nil) force)
  (sx:defblock "SX_PSK" (append (sx:conv-ents nil) (list (bT 5.6 -6.9 2.0 "KNX" nil 0.0))) force)
  (sx:defblock "SX_DIM4" (sx:conv-ents T) force)
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
  ;; клемма цепи управления узла КМ: один оранжевый кружок
  (sx:defblock "SX_XTC" (list (bCc 0.0 -1.0 1.0 30) (bAtt "N" 1.7 0.1 2.0 "L" 0.0 nil)) force)
  (sx:defblock "SX_XTPE" (list (bCc 0.0 0.0 1.0 3) (bLc 1.003 0.753 -0.753 -1.003 3) (bLc 0.753 1.003 -1.003 -0.753 3)) force)
  ;; кабельный вывод (если блока «Вывод» нет): линия вниз со стрелкой, вставка в верхней точке
  (sx:defblock "SX_U_OUT" (list (bLc 0.0 0.0 0.0 -9.0 0) (bLc 0.0 -9.0 -1.2 -6.6 0) (bLc 0.0 -9.0 1.2 -6.6 0)) T)
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
      ;; ячейки таблицы до строки «номер линии» (62); строка УГО (62..75) без рамки —
      ;; символ УГО стоит над таблицей; марка кабеля — просто текстом, без ячейки
      (list (bL 10.0 0.0 10.0 62.0))
      (mapcar '(lambda (y) (bL -10.0 y 10.0 y)) '(0.0 30.0 38.0 46.0 54.0 62.0))
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
;;; запасная библиотека: 02_Однолинейка.dwg в папке выбранной (если выбрана другая)
(defun sx:lib-fallback (path / f)
  (if (and path (/= path "") (findfile path)
           (setq f (findfile (strcat (vl-filename-directory (findfile path)) "\\" (sx:ru '(48 50 95 1054 1076 1085 1086 1083 1080 1085 1077 1081 1082 1072 46 100 119 103)))))
           (/= (strcase f) (strcase (findfile path))))
    f))

;;; какие из names есть в библиотеке или в запасной
(defun sx:lib-found (path names / res rest fb)
  (setq res (sx:lib-found1 path names)
        rest (vl-remove-if '(lambda (n) (member n res)) names))
  (if (and rest (setq fb (sx:lib-fallback path)))
    (progn (sx:lib-close) (setq res (append res (sx:lib-found1 fb rest))) (sx:lib-close)))
  res)

(defun sx:lib-found1 (path names / src blks b res)
  (if (and names path (/= path "") (findfile path) (setq src (sx:lib-open path)))
    (progn
      (setq blks (vla-get-blocks src))
      (foreach n names
        (setq b (vl-catch-all-apply 'vla-item (list blks n)))
        (if (not (vl-catch-all-error-p b))
          (progn (setq res (cons n res)) (vl-catch-all-apply 'vlax-release-object (list b)))))))
  (reverse res))

;;; скопировать недостающие блоки из DWG-библиотеки
(defun sx:import-blocks (path names / src objs b sa r blks fb)
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
  (sx:lib-close)
  ;; чего нет в выбранной библиотеке — из запасной (02_Однолинейка.dwg рядом)
  (setq names (vl-remove-if '(lambda (n) (tblsearch "BLOCK" n)) names))
  (if (and names (setq fb (sx:lib-fallback path)))
    (progn (sx:log (strcat "lib: fallback " fb)) (sx:import-blocks fb names))))

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

;;; Итоги берутся из вычисленной таблицы, включая коэффициенты спроса.
;;; Пустые таблицы (даже с заполненным cos и ошибками фаз) не рисуются.
(defun sx:summary-values (vals panel / r c row val found metrics)
  (setq r 0)
  (foreach row vals
    (setq c 0)
    (foreach val row
      (if (and (not found) (> r 0) (>= c 3)
               (= (sx:s val) (sx:ru '(1056 1091 1089 1090 61)))
               (= (sx:s (nth (- c 3) (nth (1- r) vals))) panel))
        (progn
          (setq metrics (list (nth (1+ c) row)
                         (nth (1+ c) (nth (1+ r) vals))
                         (nth (1+ c) (nth (+ r 2) vals))
                         (nth (1+ c) (nth (+ r 3) vals))
                         (nth (1+ c) (nth (+ r 4) vals))
                         (nth (+ c 4) row)
                         (nth (+ c 4) (nth (1+ r) vals))
                         (nth (+ c 4) (nth (+ r 2) vals))))
          (if (not (vl-some '(lambda (v) (not (numberp v))) metrics))
            (setq found metrics)
            (if (car metrics) (sx:log (strcat "summary: invalid values for " panel))))))
      (setq c (1+ c)))
    (setq r (1+ r)))
  found)

(defun sx:read-summary (panel / sh rg vals)
  (if (setq sh (cdr (assoc (sx:ru '(1053 1072 1075 1088 1091 1079 1082 1072 95 1097 1080 1090 1086 1074)) sx:sheets)))
    (progn
      (setq rg (vlax-get-property sh 'Range "A1:AQ1000"))
      (setq vals (vl-catch-all-apply
                   '(lambda () (vlax-safearray->list
                      (vlax-variant-value (vlax-get-property rg 'Value2)))) nil))
      (sx:rel rg)
      (if (vl-catch-all-error-p vals) (error (vl-catch-all-error-message vals)))
      (sx:summary-values (mapcar '(lambda (row) (mapcar 'sx:v row)) vals) panel))))

(defun sx:xl-release ()
  (sx:free-book)
  (foreach wb sx:wb-opened
    (vl-catch-all-apply 'vlax-invoke-method (list wb 'Close :vlax-false))
    (sx:rel wb))
  (if (and sx:xl sx:xl-created) (vl-catch-all-apply 'vlax-invoke-method (list sx:xl 'Quit)))
  (if sx:xl (sx:rel sx:xl))
  (setq sx:xl nil sx:xl-created nil sx:wb nil sx:wb-opened nil sx:sheets nil))

;;; ---------------------------------------------------------------- метки и выборочное обновление
;;; каждый построенный объект получает xdata SX_SCHEMA: (щит, кусок схемы)
;;;   H — боковик, S|лист — шины и границы листа, T|i — столбец таблицы места i,
;;;   G|i — автомат группы в месте i с перемычками и N, M|модуль|лист — контур модуля,
;;;   U — итоги щита, K — шина KNX, R|i — канал, клемма, линия и УГО места i (i = № места - 1)
(setq sx:app "SX_SCHEMA" sx:part nil sx:onlyP nil sx:onlyG nil sx:onlyM nil sx:onlyS nil sx:tagpanel "")
(defun sx:wantP (i) (or (not sx:part) (member i sx:onlyP)))
(defun sx:wantG (i) (or (not sx:part) (member i sx:onlyG)))
(defun sx:wantM (k) (or (not sx:part) (member k sx:onlyM)))
(defun sx:wantS (k) (or (not sx:part) (member k sx:onlyS)))
;; лист схемы (с 0), к которому относится кусок
(defun sx:key-sheet (k / i)
  (cond ((wcmatch k "T|*,R|*,G|*") (sx:sheet (sx:key-n k)))
        ((wcmatch k "S|*") (sx:key-n k))
        ((wcmatch k "M|*") (setq i (vl-string-position 124 k nil T)) (atoi (substr k (+ i 2))))
        ((member k '("H" "U")) 0)))
;; SXUPDSHEET: куски -> листы целиком ("S|лист")
(defun sx:keys->sheets (keys / res s)
  (foreach k keys
    (if (and (setq s (sx:key-sheet k)) (not (member (strcat "S|" (itoa s)) res)))
      (setq res (cons (strcat "S|" (itoa s)) res))))
  res)
(defun sx:xd-set (e key)
  ;; Преобразование ename тоже защищено: служебные записи не являются VLA-объектами.
  (vl-catch-all-apply
    '(lambda ()
       (vla-setxdata (vlax-ename->vla-object e)
         (vlax-safearray-fill (vlax-make-safearray vlax-vbInteger '(0 . 2)) '(1001 1000 1000))
         (vlax-safearray-fill (vlax-make-safearray vlax-vbVariant '(0 . 2)) (list sx:app sx:tagpanel key)))) nil))
;; пометить всё, что создано после mark (атрибуты и концы вставок не трогаем)
(defun sx:tag-new (mark key / e ed)
  (regapp sx:app)
  (setq e (if mark (entnext mark) (entnext)))
  (while e
    (setq ed (entget e))
    (if (and ed (= (cdr (assoc 410 ed)) "Model")
             (not (member (cdr (assoc 0 ed)) '("ATTRIB" "SEQEND" "VERTEX" "BLOCK" "ENDBLK" "ATTDEF"))))
      (sx:xd-set e key))
    (setq e (entnext e))))
;; sx:try + метка куска схемы
(defun sx:ttry (key fn label / first count result err ob)
  ;; Только новые объекты ModelSpace. entnext может заходить в определения блоков.
  (setq first (vla-get-count sx:ms))
  (sx:try fn label)
  (regapp sx:app)
  (setq count (vla-get-count sx:ms))
  (while (< first count)
    (setq result (vl-catch-all-apply
      '(lambda ()
         (setq ob (vla-item sx:ms first))
         (sx:xd-set (vlax-vla-object->ename ob) key)) nil))
    (if (vl-catch-all-error-p result)
      (progn
        (setq err (strcat "tag " label ": " (vl-catch-all-error-message result)))
        (sx:log err) (setq sx:errs (cons err sx:errs))))
    (setq first (1+ first))))
;; (щит кусок) из метки объекта или nil
(defun sx:xd-get (e / xd)
  (if (setq xd (cdr (assoc -3 (entget e (list sx:app)))))
    (mapcar 'cdr (vl-remove-if-not '(lambda (x) (= (car x) 1000)) (cdr (car xd))))))
(defun sx:key-n (k) (atoi (substr k 3)))
;; ключ модуля строки (как в sx:draw-modules) или nil
(defun sx:modkey (d)
  (if (and (sx:has d sx:cMod) (sx:c d sx:cCh) (/= (sx:s (sx:c d sx:cChT)) "="))
    (strcat (sx:s (sx:c d sx:cMod)) "|" (itoa (sx:sheet (sx:i d))))))
;; выбор объектов: (щиты куски непомеченных) или nil
(defun sx:pick-keys ( / ss k e xd pans keys bare)
  (setq ss (ssget "_I" '((8 . "SX_*"))))
  (if (null ss)
    (progn (princ (strcat "\n" (sx:ru '(1042 1099 1073 1077 1088 1080 1090 1077 32 1086 1073 1098 1077 1082 1090 1099 32 1089 1093 1077 1084 1099 44 32 1082 1086 1090 1086 1088 1099 1077 32 1085 1091 1078 1085 1086 32 1087 1077 1088 1077 1089 1090 1088 1086 1080 1090 1100 58 32))))
           (setq ss (ssget '((8 . "SX_*"))))))
  (sssetfirst nil nil)
  (if ss
    (progn
      (setq k (sslength ss) bare 0)
      (repeat k
        (setq e (ssname ss (setq k (1- k))))
        (if (setq xd (sx:xd-get e))
          (progn
            (if (not (member (car xd) pans)) (setq pans (cons (car xd) pans)))
            (if (not (member (cadr xd) keys)) (setq keys (cons (cadr xd) keys))))
          (setq bare (1+ bare))))
      (list pans keys bare))))
;; что перестроить по выбранным кускам: ставит sx:onlyP/G/M, возвращает куски к удалению
(defun sx:sel-units (keys draw / pl gl ml sl dk gid q mk i s insl)
  ;; ввод задаёт левый край шин/границ: обновлять вместе с его листом.
  (foreach d draw
    (if (and (sx:input-p d)
             (vl-some '(lambda (k) (and (wcmatch k "T|*,R|*,G|*") (= (sx:key-n k) (sx:i d)))) keys))
      (setq insl (cons (sx:sheet (sx:i d)) insl))))
  ;; лист целиком: все его места
  (foreach k keys
    (if (and (wcmatch k "S|*") (not (member (sx:key-n k) sl)))
      (progn
        (setq sl (cons (sx:key-n k) sl) i (* (sx:key-n k) *sx-sheet*))
        (repeat *sx-sheet* (if (not (member i pl)) (setq pl (cons i pl))) (setq i (1+ i))))))
  (foreach k keys
    (cond ((wcmatch k "T|*,R|*,G|*") (if (not (member (sx:key-n k) pl)) (setq pl (cons (sx:key-n k) pl))))
          ((wcmatch k "M|*")
           (setq ml (cons (substr k 3) ml))
           (foreach d draw
             (if (and (= (sx:modkey d) (substr k 3)) (not (member (sx:i d) pl))) (setq pl (cons (sx:i d) pl)))))))
  (foreach k keys (if (and (wcmatch k "G|*") (not (member (sx:key-n k) gl))) (setq gl (cons (sx:key-n k) gl))))
  ;; автомат группы и модуль перерисовываются, если затронуто любое их место
  (foreach d draw
    (if (member (sx:i d) pl)
      (progn
        (if (and (sx:has d sx:cQFown) (not (member (sx:i d) gl))) (setq gl (cons (sx:i d) gl)))
        (if (/= (setq gid (sx:s (sx:c d sx:cGrp))) "")
          (foreach q draw
            (if (and (= (sx:s (sx:c q sx:cQFown)) gid) (not (member (sx:i q) gl))) (setq gl (cons (sx:i q) gl)))))
        (if (and (setq mk (sx:modkey d)) (not (member mk ml))) (setq ml (cons mk ml))))))
  ;; Для ввода обновить геометрию шин, сохраняя остальные места листа.
  (foreach s insl (if (not (member s sl)) (setq sl (cons s sl))))
  (setq sx:onlyP pl sx:onlyG gl sx:onlyM ml sx:onlyS sl)
  (foreach s sl (setq dk (cons (strcat "S|" (itoa s)) dk)))
  (if (member 0 sl) (setq dk (cons "H" dk)))
  (foreach i pl (setq dk (cons (strcat "T|" (itoa i)) (cons (strcat "R|" (itoa i)) dk))))
  (foreach i gl (setq dk (cons (strcat "G|" (itoa i)) dk)))
  (foreach mk ml (setq dk (cons (strcat "M|" mk) dk)))
  (if ml (setq dk (cons "K" dk)))
  (cons "U" dk))
;; удалить помеченные куски щита; число удалённых
(defun sx:del-keys (panel keys / ss k e xd n)
  (setq n 0)
  (if (setq ss (ssget "_X" (list (list -3 (list sx:app)))))
    (progn
      (setq k (sslength ss))
      (repeat k
        (setq e (ssname ss (setq k (1- k))) xd (sx:xd-get e))
        (if (and xd (= (car xd) panel) (member (cadr xd) keys)) (progn (entdel e) (setq n (1+ n)))))))
  n)

;;; Однократный переход с ручного образца: взять под управление только его
;;; две точные линии кабеля и левую границу. Остальные пользовательские объекты
;;; (питающий щит, аппарат, УГО отходящих линий) не затрагиваются.
(defun sx:path-equal (pts expected)
  (and (= (length pts) (length expected))
       (not (vl-some '(lambda (a b) (> (distance a b) 0.15)) pts expected))))
(defun sx:adopt-inputs (draw / ss k e ed pts u xs cable1 cable2 bound key)
  (if (setq ss (ssget "_X" '((0 . "LINE,LWPOLYLINE") (410 . "Model"))))
    (foreach d draw
      (if (and (sx:input-p d) (sx:wantP (sx:i d)))
        (progn
          (setq u (sx:u d) xs (- u sx:inDx)
                cable1 (list (sx:pt xs sx:inLow) (sx:pt xs sx:inFeed))
                cable2 (list (sx:pt xs sx:inLow) (sx:pt u sx:inLow) (sx:pt u sx:vOutTop))
                bound (list (sx:pt (- u sx:inLeft) sx:vTop) (sx:pt (- u sx:inLeft) sx:vBottom))
                k (sslength ss))
          (repeat k
            (setq e (ssname ss (setq k (1- k))) ed (entget e) key nil)
            (setq pts (if (= (cdr (assoc 0 ed)) "LINE")
                        (list (cdr (assoc 10 ed)) (cdr (assoc 11 ed)))
                        (mapcar '(lambda (p) (append (cdr p) '(0.0)))
                                (vl-remove-if-not '(lambda (p) (= (car p) 10)) ed))))
            (cond
              ((and (= (cdr (assoc 8 ed)) "0")
                    (or (sx:path-equal pts cable1) (sx:path-equal pts cable2)))
               (vla-put-layer (vlax-ename->vla-object e) sx:L)
               (setq key (strcat "R|" (itoa (sx:i d)))))
              ((and (= (cdr (assoc 8 ed)) sx:B) (null (sx:xd-get e)) (sx:path-equal pts bound))
               (setq key (strcat "S|" (itoa (sx:sheet (sx:i d)))))))
            (if key (sx:xd-set e key))))))))

;;; Итоги под первым листом: блок SX_SUMMARY 120 x 40 мм, восемь атрибутов.
;;; При обновлении сохраняем положение; старую таблицу заменяем блоком.
(defun sx:summary-capture ( / ob e xd pt target)
  (setq sx:sumpt (sx:pt 7.56437443444702 -5.3734190440593)
        sx:sumscale '(1.0 1.0 1.0) sx:sumang sx:ang)
  (setq target sx:sumpt)
  (vlax-for ob sx:ms
    (if (or (and (= (vla-get-objectname ob) "AcDbTable")
                 (= (vla-get-rows ob) 5) (= (vla-get-columns ob) 6))
            (and (= (vla-get-objectname ob) "AcDbBlockReference")
                 (= (strcase (vla-get-effectivename ob)) "SX_SUMMARY")))
      (progn
        (setq e (vlax-vla-object->ename ob) xd (sx:xd-get e)
              pt (vlax-safearray->list (vlax-variant-value (vla-get-insertionpoint ob))))
        (if (or (equal xd (list sx:tagpanel "U"))
                (and (null xd) (< (distance pt target) 0.15)))
          (progn
            (setq sx:sumpt pt)
            (if (= (vla-get-objectname ob) "AcDbBlockReference")
              (setq sx:sumscale (list (vla-get-xscalefactor ob) (vla-get-yscalefactor ob) (vla-get-zscalefactor ob))
                    sx:sumang (vla-get-rotation ob)))
            (vla-put-layer ob sx:TB)
            (sx:xd-set e "U")))))))

(defun sx:summary-num (v / oldz text)
  ;; Два знака после запятой, как в Excel, независимо от DIMZIN чертежа.
  (setq oldz (getvar "DIMZIN")) (setvar "DIMZIN" 0)
  (setq text (vl-string-translate "." "," (rtos v 2 2)))
  (setvar "DIMZIN" oldz)
  text)

(defun sx:summary-block ( / ents y x r labels tags label)
  (setq ents (bR 0.0 -40.0 120.0 0.0))
  (foreach y '(-8.0 -16.0 -24.0 -32.0)
    (setq ents (append ents (list (bL 0.0 y 120.0 y)))))
  (setq ents (append ents (list (bL 20.0 0.0 20.0 -40.0))))
  (foreach x '(40.0 60.0 80.0 100.0)
    (setq ents (append ents (list (bL x 0.0 x -24.0)))))
  (setq labels (list (sx:ru '(1056 1091 1089 1090 61)) (sx:ru '(1056 1088 61)) (sx:ru '(73 1088 61))
                     (sx:ru '(67 111 115 40 1092 41 61)) (sx:ru '(1050 1089 61)))
        tags '("PUST" "PR" "IR" "COS" "KS") r 0)
  (foreach label labels
    (setq y (- -4.0 (* r 8.0))
          ents (append ents (list (bT 10.0 y 2.5 label "M" 0.0)
                                   (bAtt (nth r tags) (if (< r 3) 30.0 70.0) y 2.5 "M" 0.0 nil))))
    (if (< r 3)
      (setq ents (append ents
        (list (bT 50.0 y 2.5 (if (= r 2) (sx:ru '(1040)) (sx:ru '(1082 1042 1090))) "M" 0.0)
              (bT 70.0 y 2.5 (strcat "L" (itoa (1+ r)) "p=") "M" 0.0)
              (bAtt (strcat "L" (itoa (1+ r))) 90.0 y 2.5 "M" 0.0 nil)
              (bT 110.0 y 2.5 (sx:ru '(1082 1042 1090)) "M" 0.0)))))
    (setq r (1+ r)))
  (sx:defblock "SX_SUMMARY" ents nil))

(defun sx:draw-summary ( / ob pairs tags tag k a)
  (if sx:summary
    (progn
      (sx:summary-block)
      (setq ob (vla-insertblock sx:ms (vlax-3d-point sx:sumpt) "SX_SUMMARY"
                 (nth 0 sx:sumscale) (nth 1 sx:sumscale) (nth 2 sx:sumscale) sx:sumang))
      ;; Сразу пометить объект: следующий запуск удалит и незавершённый блок.
      (vla-put-layer ob sx:TB) (sx:xd-set (vlax-vla-object->ename ob) "U")
      (setq tags '("PUST" "PR" "IR" "COS" "KS" "L1" "L2" "L3") pairs nil k 0)
      (foreach tag tags
        (setq pairs (cons (cons tag (sx:summary-num (nth k sx:summary))) pairs) k (1+ k)))
      (foreach a (vlax-invoke ob 'GetAttributes)
        (vla-put-textstring a (cdr (assoc (strcase (vla-get-tagstring a)) pairs))))
      (vla-update ob))))

;;; ---- таблицы по источникам (v3.21): рядом с итогами щита, по одной на источник с нагрузкой.
;;; Цифры — лист «Режимы», нормальный режим (Руст, Кс, cos, Рр, Iр); категория — «Сеть_щита» I69:K71
;;; (блок «ДГУ и ИБП», patch_categories.py). Рисуются, если источников с нагрузкой два и больше или есть ДГУ / ИБП.
(defun sx:read-sources (panel / sh rg vals cell net idx k res nm ust cat)
  (setq res nil sx:cats nil)
  (if (and (setq sh (cdr (assoc (sx:ru '(1056 1077 1078 1080 1084 1099)) sx:sheets)))
           (setq net (cdr (assoc (sx:ru '(1057 1077 1090 1100 95 1097 1080 1090 1072)) sx:sheets))))
    (progn
      (setq rg (vlax-get-property sh 'Range "B4:G21"))
      (setq vals (vl-catch-all-apply
                   '(lambda () (vlax-safearray->list (vlax-variant-value (vlax-get-property rg 'Value2)))) nil))
      (sx:rel rg)
      (if (and (not (vl-catch-all-error-p vals))
               (= (sx:s (sx:v (nth 1 (car vals)))) panel))     ; «Режимы» C4 — этот щит
        (progn
          (setq vals (mapcar '(lambda (rw) (mapcar 'sx:v rw)) vals))
          (defun cell (r c) (nth c (nth (- r 4) vals)))         ; r — строка листа, c — 0 = B, 1 = C ...
          ;; категории шин щита: «Сеть_щита» B6:B8 (щит), I69:K71 (Ввод, ДГУ, ИБП)
          (setq rg (vlax-get-property net 'Range "B6:K71"))
          (setq net (vl-catch-all-apply
                      '(lambda () (vlax-safearray->list (vlax-variant-value (vlax-get-property rg 'Value2)))) nil))
          (sx:rel rg)
          (if (not (vl-catch-all-error-p net))
            (progn
              (setq net (mapcar '(lambda (rw) (mapcar 'sx:v rw)) net) k 0)
              (repeat 3 (if (and (null idx) (= (sx:s (nth 0 (nth k net))) panel)) (setq idx k)) (setq k (1+ k)))))
          (setq k 1)
          (repeat 5
            (setq nm (sx:s (cell 11 k)) ust (cell 14 k))
            (if (and (numberp ust) (> ust 0.0))
              (progn
                (setq cat (if (and idx (listp net))
                            (sx:s (nth (+ 7 idx) (nth (cond ((wcmatch (strcase nm) (strcase (sx:ru '(1044 1043 1059 42)))) 64)
                                                             ((wcmatch (strcase nm) (strcase (sx:ru '(1048 1041 1055 42)))) 65)
                                                             (t 63))
                                                       net)))
                            ""))
                (setq res (append res (list (list nm ust (cell 15 k) (cell 16 k) (cell 18 k) (cell 21 k) cat))))))
            (setq k (1+ k)))
          (setq sx:cats (sx:read-cats sh))))))
  (if (or (> (length res) 1)
          (vl-some '(lambda (r) (not (wcmatch (strcase (car r)) (strcase (sx:ru '(1042 1074 1086 1076 42)))))) res))
    res))

;;; «Режимы» B103:G106 «По категориям»: (категория Руст Кс cos Рр Iр «») по категориям с нагрузкой.
;;; Iр = Sр / (√3 · 0,38 кВ), cos = Рр / Sр, Кс = Рр / Руст. Нужны две категории и больше.
(defun sx:read-cats (sh / rg vals res ust pr sr)
  (setq rg (vlax-get-property sh 'Range "B103:G106"))
  (setq vals (vl-catch-all-apply
               '(lambda () (vlax-safearray->list (vlax-variant-value (vlax-get-property rg 'Value2)))) nil))
  (sx:rel rg)
  (if (not (vl-catch-all-error-p vals))
    (foreach rw (mapcar '(lambda (r) (mapcar 'sx:v r)) vals)
      (setq ust (nth 2 rw) pr (nth 3 rw) sr (nth 5 rw))
      (if (and (numberp ust) (> ust 0.0) (numberp pr) (numberp sr) (> sr 0.0))
        (setq res (append res (list (list (sx:s (nth 0 rw)) ust (/ pr ust) (/ pr sr) pr (/ sr (* (sqrt 3.0) 0.38)) "")))))))
  (if (> (length res) 1) res))

;;; блок таблицы источника 60 x 40 мм: заголовок над таблицей, Руст, Рр, Iр, cos, Кс
(defun sx:source-block ( / ents y r labels tags)
  (setq ents (append (bR 0.0 -40.0 60.0 0.0)
                     (list (bL 0.0 -8.0 60.0 -8.0) (bL 0.0 -16.0 60.0 -16.0) (bL 0.0 -24.0 60.0 -24.0)
                           (bL 0.0 -32.0 60.0 -32.0) (bL 20.0 0.0 20.0 -40.0) (bL 45.0 0.0 45.0 -24.0)
                           (bAtt "TITLE" 0.0 1.5 2.5 "L" 0.0 nil))))
  (setq labels (list (sx:ru '(1056 1091 1089 1090 61)) (sx:ru '(1056 1088 61)) (sx:ru '(73 1088 61))
                     (sx:ru '(67 111 115 40 1092 41 61)) (sx:ru '(1050 1089 61)))
        tags '("PUST" "PR" "IR" "COS" "KS") r 0)
  (foreach lab labels
    (setq y (- -4.0 (* r 8.0))
          ents (append ents (list (bT 10.0 y 2.5 lab "M" 0.0)
                                  (bAtt (nth r tags) (if (< r 3) 32.5 40.0) y 2.5 "M" 0.0 nil))))
    (if (< r 3)
      (setq ents (append ents (list (bT 52.5 y 2.5 (if (= r 2) (sx:ru '(1040)) (sx:ru '(1082 1042 1090))) "M" 0.0)))))
    (setq r (1+ r)))
  (sx:defblock "SX_SOURCE" ents T))

(defun sx:draw-sources ( / k x ob pairs src ttl)
  (if (and (or sx:sources sx:cats) sx:summary)
    (progn
      (sx:source-block)
      ;; над итогами щита — подпись «щит в целом»
      (setq ob (entmake (list '(0 . "TEXT") (cons 8 sx:TB) (cons 10 (list (car sx:sumpt) (+ (cadr sx:sumpt) 1.5) 0.0))
                              (cons 40 2.5) (cons 1 (strcat sx:tagpanel (sx:ru '(32 8212 32 1074 1077 1089 1100 32 1097 1080 1090))))
                              (cons 50 0.0) (cons 7 sx:style))))
      (if ob (sx:xd-set (entlast) "U"))
      (setq k 0)
      (foreach src (append sx:sources (mapcar '(lambda (c) (cons 'cat c)) sx:cats))
        ;; категория: заголовок «I особая категория — итого»
        (setq ttl (if (= (car src) 'cat)
                    (strcat (nth 1 src) (sx:ru '(32 1082 1072 1090 1077 1075 1086 1088 1080 1103 32 8212 32 1080 1090 1086 1075 1086)))
                    (strcat (nth 0 src) (sx:ru '(32 8212 32)) (nth 6 src)
                            (if (/= (nth 6 src) "") (sx:ru '(32 1082 1072 1090 1077 1075 1086 1088 1080 1103)) "")))
              src (if (= (car src) 'cat) (cdr src) src)
              x (+ (car sx:sumpt) (* (+ 125.0 (* k 65.0)) (car sx:sumscale))))
        (setq ob (vla-insertblock sx:ms (vlax-3d-point (list x (cadr sx:sumpt) 0.0)) "SX_SOURCE"
                   (nth 0 sx:sumscale) (nth 1 sx:sumscale) (nth 2 sx:sumscale) 0.0))
        (vla-put-layer ob sx:TB) (sx:xd-set (vlax-vla-object->ename ob) "U")
        (setq pairs (list (cons "TITLE" ttl)
                          (cons "PUST" (sx:summary-num (sx:n (nth 1 src)))) (cons "KS" (sx:summary-num (sx:n (nth 2 src))))
                          (cons "COS" (sx:summary-num (sx:n (nth 3 src)))) (cons "PR" (sx:summary-num (sx:n (nth 4 src))))
                          (cons "IR" (sx:summary-num (sx:n (nth 5 src))))))
        (foreach a (vlax-invoke ob 'GetAttributes)
          (vla-put-textstring a (cdr (assoc (strcase (vla-get-tagstring a)) pairs))))
        (vla-update ob)
        (setq k (1+ k))))))

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

(defun sx:draw-table (recs / nm cab k kn kk ln cur)
  (if (and (sx:wantS 0) (sx:blk-p sx:bHEAD))
    (sx:ttry "H" (function (lambda () (sx:insv sx:bHEAD sx:vHeadDx 0.0 1.0 0.0 nil sx:TB))) "head"))
  (foreach d recs
   (if (sx:wantP (sx:i d))
   (sx:ttry (strcat "T|" (itoa (sx:i d))) (function (lambda ()
    (setq nm (sx:wrap (sx:s (sx:c d sx:cName)) 20 3)
          cab (sx:s (sx:c d sx:cCab))
          k (vl-string-search " L=" cab))
    (if (= (length nm) 1) (setq nm (list "" (car nm))))
    (setq nm (list (sx:nth0 0 nm) (sx:nth0 1 nm) (sx:nth0 2 nm))
          ln (sx:s (sx:c d sx:cLine)) cur (sx:num (sx:c d sx:cI)))
    ;; узел КМ (АПС): три места — сигнал АПС, реле АПС, контактор
    (setq kn (car (vl-remove-if-not '(lambda (n) (and (sx:kma-p n) (<= 0 (- (sx:i d) (sx:i n)) 2))) recs)))
    (if (and kn (= ln ""))
      (progn
        (setq kk (- (sx:i d) (sx:i kn)))
        (cond ((= kk 0) (setq nm (list "" (sx:ru '(1057 1080 1075 1085 1072 1083 32 1040 1055 1057)) "") ln (sx:ru '(1040 1055 1057))))
              ((= kk 1) (setq nm (list "" (sx:ru '(1056 1077 1083 1077 32 1040 1055 1057)) "") ln (sx:ru '(1056 1077 1083 1077))))
              (t (setq nm (list (sx:ru '(1050 1086 1085 1090 1072 1082 1090 1086 1088)) (sx:s (sx:c kn sx:cKMmod)) "")
                       ln (sx:kma-name kn) cur (sx:num (/ (sx:n (sx:c kn sx:cKMcur)) 1.1)))))))
    (sx:ins "SX_TCOL" (sx:u d) 0.0
            (list (cons "NAME1" (nth 0 nm)) (cons "NAME2" (nth 1 nm)) (cons "NAME3" (nth 2 nm))
                  (cons "CUR" cur) (cons "POW" (sx:num (sx:c d sx:cPow)))
                  (cons "PH" (sx:s (sx:c d sx:cPh))) (cons "LINE" ln)
                  (cons "CAB1" (vl-string-translate "." "," (if k (substr cab 1 k) cab)))
                  (cons "CAB2" (vl-string-translate "." "," (if k (substr cab (+ k 2)) ""))))
            sx:TB)))
    (strcat "table " (sx:rowinfo d))))))

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

(defun sx:draw-sheets (draw panel / sheets s us ue onsheet b2 segs k input left busus fd g xr notch q bus2us)
  (setq sheets '())
  (foreach d draw
    (if (and (or (sx:content-p d) (sx:dsup-p d)) (not (member (sx:sheet (sx:i d)) sheets)))
      (setq sheets (cons (sx:sheet (sx:i d)) sheets))))
  (sx:log (strcat "sheets=" (sx:str sheets)))
  (foreach s sheets
   (if (sx:wantS s)
   (sx:ttry (strcat "S|" (itoa s)) (function (lambda ()
    (setq us (* s *sx-sheet* *sx-step*) ue (* (1+ s) *sx-sheet* *sx-step*))
    (setq onsheet (vl-remove-if-not '(lambda (d) (= (sx:sheet (sx:i d)) s)) draw)
          fd (sx:sheet-feed draw s))
    (if fd
      ;; ввод с питающим щитом: шина начинается от петли ввода, граница обходит рамку питающего щита
      (progn
        (setq g (cadr fd) left (nth 1 g) xr (+ (nth 2 g) 3.5) notch (- (nth 7 g) 6.0) input fd
              q (car (vl-remove-if-not '(lambda (d) (and (sx:feed-p d) (not (sx:is1 d sx:cBus2)))) onsheet))
              busus (- (nth 3 (if q (sx:feed-geo q draw) g)) 5.0)
              q (car (vl-remove-if-not '(lambda (d) (and (sx:feed-p d) (sx:is1 d sx:cBus2))) onsheet))
              bus2us (if q (- (nth 3 (sx:feed-geo q draw)) 5.0)))
        (sx:bus busus ue sx:vBusL sx:vBusN "L1,L2,L3")
        (sx:text (+ xr 1.0) (+ sx:vTop 1.5) 2.5 panel sx:T)
        (sx:line xr sx:vTop (- ue 2) sx:vTop sx:B)
        (sx:line xr sx:vTop xr notch sx:B)
        (sx:line left notch xr notch sx:B)
        (sx:line left notch left sx:vBottom sx:B)
        (sx:line left sx:vBottom (- ue 2) sx:vBottom sx:B))
      (progn
        (setq input (sx:sheet-input draw s)
              left (if input (- (sx:u input) sx:inLeft) (+ us 2.0))
              busus (if input (+ left sx:inBus -5.0) us) bus2us nil)
        (sx:bus busus ue sx:vBusL sx:vBusN "L1,L2,L3")
        (sx:text left (+ sx:vTop 1.5) 2.5 panel sx:T)
        (sx:line left sx:vTop (- ue 2) sx:vTop sx:B)
        (sx:line left sx:vBottom (- ue 2) sx:vBottom sx:B)   ; граница щита снизу
        (if input (sx:line left sx:vTop left sx:vBottom sx:B))))
    (setq b2 (mapcar '(lambda (d) (list (sx:s (sx:c d sx:cBusName)) (sx:u d)))
                     (vl-remove-if-not '(lambda (d) (and (sx:has d sx:cQFown) (sx:is1 d sx:cBus2))) onsheet)))
    ;; узел КМ на листе: его шина начинается от силового контакта (третий элемент T)
    (foreach d onsheet
      (if (and (sx:kma-p d) (sx:has d sx:cQFown)) (setq b2 (cons (list (sx:kma-name d) (sx:kma-x d) T) b2))))
    ;; вторые шины — до конца листа; если на листе несколько разных, каждая до начала следующей
    (setq segs '())
    (foreach sg (vl-sort b2 '(lambda (a b) (< (cadr a) (cadr b))))
      (if (or (null segs) (/= (car (car segs)) (car sg)))
        (setq segs (cons sg segs))))
    (setq segs (reverse segs) k 0)
    (foreach sg segs
      (sx:bus (cond ((caddr sg) (- (cadr sg) 5.0))
                    ((= k 0) (if bus2us bus2us busus))
                    (t (- (cadr sg) 9.0)))
              (if (< (1+ k) (length segs))
                (- (cadr (nth (1+ k) segs)) (if (caddr (nth (1+ k) segs)) 6.0 11.0))
                ue)
              sx:vBus2L sx:vBus2N
              (if (sx:kma-find (car sg) draw) (strcat "L1,L2,L3 (" (car sg) ")") (car sg)))
      (setq k (1+ k)))
    ;; шина PE — одна на уровне sx:vPE: для клемм и для линий ДСУП
    (if (vl-some '(lambda (d) (or (sx:is1 d sx:cXT) (sx:dsup-p d))) onsheet)
      (sx:line (if input (+ left 7.0) (+ us 3)) sx:vPE (- ue 1) sx:vPE sx:PE))))
    (strcat "sheet " (itoa (1+ s)))))))

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
    ;; KMA — узел КМ (АПС): не аппарат под автоматом, его рисует sx:draw-kma
    (if (not (member code '("" "KMA"))) (setq res (append res (list (list code (sx:s (sx:c q (+ sx:cDev1 k))))))))
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
    ((member (car dv) '("PS" "PK"))
     (setq sx:pcount (1+ sx:pcount) n (list (strcat "U" (itoa sx:pcount)) (cadr dv))))
    ((= (car dv) "DIM")
     (setq sx:dcount (1+ sx:dcount) n (list (strcat "A" (itoa sx:dcount)) (cadr dv))))
    (t (setq n (list (cadr dv)))))
  (if (not (sx:blk-p blk))
    (setq blk (cdr (assoc blk (list (cons sx:bT "SX_T") (cons sx:bPS "SX_PS") (cons sx:bMK "SX_MKX")
                                    (cons "SX_KM" "SX_KM") (cons "SX_DEV" "SX_DEV"))))))
  (setq ob (sx:insv blk x (if (nth 3 spec) top (- top (nth 4 spec))) (if (= blk "SX_MKX") 0.6111 (nth 2 spec)) 0.0 n sx:BL))
  (if (and (member (car dv) '("T" "PS" "PK")) (not (wcmatch (strcase blk) "SX_*"))) (sx:red ob))
  ob)

(defun sx:draw-qf (q u side nbyp / devs H top parts vIns busL ob name spec r xs xd vTop)
  (setq devs (sx:dev-list q) H 0.0)
  (foreach dv devs (setq H (+ H (nth 4 (sx:dev-spec (car dv))))))
  (setq busL (if (sx:is1 q sx:cBus2) sx:vBus2L sx:vBusL)
        parts (sx:qf-parts q)
        name (sx:s (sx:c q sx:cQFown)))
  (if (and side devs (> (- (if (sx:is1 q sx:cBus2) sx:vBus2N sx:vBusN) 3.0) (+ sx:vSideRet sx:sideLead H)))
    (setq vIns (+ sx:vSideB sx:qfBody))
    ;; аппараты под автоматом: автомат на своём уровне, аппараты — в свободное место до клеммы;
    ;; поднять автомат только на то, что не поместилось (у групп с модулем места нет — как раньше)
    (setq side nil vIns (+ sx:vQFins (if sx:gHasCh H (max 0.0 (- H sx:devRoom))))))
  ;; автомат: точка подключения блока — ровно на шину
  (if (sx:blk-p sx:bQF)
    (progn
      (setq ob (sx:insv sx:bQF u vIns 1.0 0.0 (list (sx:digits name) (nth 0 parts) (nth 1 parts) (nth 2 parts)) sx:BL))
      ;; Обычные отходящие автоматы — ваш вид QF_L+N, включая 3P.
      (sx:dynset ob sx:pVis (if (sx:is1 q sx:cRcd) "QFD" "QF_L+N"))
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
        (setq top (- top (nth 4 spec))))
      ;; низ последнего аппарата — отсюда спуск к клемме / УГО
      (setq sx:qflow (cons (cons u (min sx:vComb top)) sx:qflow)))))

;;; ---------------------------------------------------------------- узел КМ (АПС), patch_km.py
;;; Строка-место: E «авто», I «КМ АПС» (BH = KMA). Место узла — автомат цепи управления, +1 — размыкающий
;;; контакт реле АПС (в строке УГО), +2 — катушка КМ. Силовой контакт — между +2 и +3: от основной шины
;;; к второй шине «L1,L2,L3 (KM-n)»; на ней линии с D = KM-n (CE узла).
(setq sx:kmLoop 68.0 sx:kmCoilB 173.0 sx:kmCoilT 177.0)
(defun sx:kma-p (d) (= (sx:s (sx:c d sx:cDevC1)) "KMA"))
(defun sx:kma-name (d) (sx:s (sx:c d sx:cKM)))
(defun sx:kma-x (d) (+ (sx:u d) (* 2.5 *sx-step*)))   ; силовой контакт
(defun sx:kma-find (name lst)
  (if (/= name "") (car (vl-remove-if-not '(lambda (d) (and (sx:kma-p d) (= (sx:kma-name d) name))) lst))))

(defun sx:draw-kma (q / u1 u2 u3 xk lo nm vm)
  (setq u1 (sx:u q) u2 (+ u1 *sx-step*) u3 (+ u1 (* 2 *sx-step*)) xk (sx:kma-x q)
        lo (sx:qf-low u1) nm (sx:kma-name q) vm (/ (+ sx:kmCoilB sx:kmCoilT) 2.0))
  ;; цепь управления: автомат -> клемма -> вниз в строку УГО -> контакт реле АПС -> клемма -> катушка
  (sx:line u1 lo u1 sx:kmLoop sx:L)
  (sx:ticks u1 (/ (+ lo sx:vXTtop) 2.0) 1 sx:L)
  (sx:hline u1 (- u2 3.0) sx:kmLoop sx:L)
  ;; размыкающий контакт (ГОСТ 2.755), горизонтально: неподвижный упор справа, нож под ним
  (sx:line (+ u2 3.0) sx:kmLoop (+ u2 3.0) (- sx:kmLoop 2.0) sx:L)
  (sx:line (- u2 3.0) sx:kmLoop (+ u2 3.6) (- sx:kmLoop 2.0) sx:L)
  (sx:hline (+ u2 3.0) u3 sx:kmLoop sx:L)
  (sx:line u3 sx:kmLoop u3 sx:kmCoilB sx:L)
  (foreach x (list u1 u3) (sx:insv "SX_XTC" x sx:vXTtop 1.0 0.0 (list "XT") sx:BL))
  ;; катушка: A1 снизу, A2 сверху -> N основной шины
  (sx:rect (- u3 4.0) sx:kmCoilB (+ u3 4.0) sx:kmCoilT sx:L)
  (sx:text (- u3 3.4) (+ sx:kmCoilB 1.0) 2.0 nm sx:T)
  (sx:text (+ u3 0.6) (- sx:kmCoilB 2.4) 1.8 "A1" sx:T)
  (sx:text (+ u3 0.6) (+ sx:kmCoilT 0.6) 1.8 "A2" sx:T)
  (sx:line u3 sx:kmCoilT u3 sx:vBusN sx:LN)
  (sx:dot u3 sx:vBusN sx:LN)
  ;; силовой контакт: основная шина -> шина KM-n; N — рядом, без разрыва
  (sx:dot xk sx:vBusL sx:L)
  (sx:line xk sx:vBusL xk (+ sx:vBus2L 10.0) sx:L)
  (sx:insv "SX_KM" xk (+ sx:vBus2L 10.0) 1.0 0.0 (list "") sx:BL)
  (sx:text (+ xk 1.2) (+ sx:vBusN 1.0) 2.0 nm sx:T)
  (sx:line (+ xk 2.5) sx:vBusN (+ xk 2.5) sx:vBus2N sx:LN)
  (sx:dot (+ xk 2.5) sx:vBusN sx:LN)
  (sx:dot (+ xk 2.5) sx:vBus2N sx:LN)
  ;; механическая связь катушки и контакта
  (sx:line (+ u3 4.0) vm (- xk 1.5) vm sx:LN)
  (sx:line (- xk 1.5) vm (- xk 1.5) (+ sx:vBus2L 5.0) sx:LN))

(setq sx:devRoom 24.0)   ; мм под автоматом для аппаратов (до клеммы XT остаётся место для засечек)
(defun sx:qf-low (u) (cond ((cdr (assoc u sx:qflow))) (t sx:vComb)))
(defun sx:draw-groups (draw / gid u busN members feeds chF xtF chU xtU xts xN hi side devs lowv lo)
  (setq sx:tcount 0 sx:pcount 0 sx:dcount 0 sx:qflow nil)
  (foreach q draw
   (if (and (sx:has q sx:cQFown) (sx:wantG (sx:i q)))
   (sx:ttry (strcat "G|" (itoa (sx:i q))) (function (lambda ()
    (if (sx:feed-p q)
      (sx:draw-feed q draw)
    (if (sx:has q sx:cQFown)
      (progn
        (setq gid (sx:s (sx:c q sx:cQFown)) u (sx:u q)
              busN (if (sx:is1 q sx:cBus2) sx:vBus2N sx:vBusN)
              devs (sx:dev-list q))
        ;; БП/трансформатор сбоку, если соседний справа столбец (на том же листе) без автомата
        (setq lowv (and devs (not (vl-some '(lambda (dv) (not (member (car dv) '("T" "PS" "PK")))) devs))))
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
        (setq sx:gHasCh (if chU T nil))
        (sx:draw-qf q u side (and chU (not lowv)))
        (if (sx:kma-p q) (sx:draw-kma q))
        ;; входы каналов: перемычки IN-IN с фасками
        (if chU (sx:chain (cons u chU) sx:vIN0 sx:L))
        ;; клеммы без канала: спуск от автомата и перемычки по верху клемм
        (if xtU
          (progn
            (if (not (member u chU))
              (progn (setq lo (sx:qf-low u))
                     (sx:line u lo u sx:vXTtop sx:L)
                     (sx:ticks u (min sx:vTickQ (/ (+ lo sx:vXTtop) 2.0)) (sx:nph q) sx:L)))
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
            (sx:hline xN hi sx:vXTN sx:LN))))))))
    (strcat "group " (sx:rowinfo q))))))

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
  ;; шина KNX считается по всем модулям щита (и при выборочном обновлении)
  (foreach item segs
    (setq knx (cons (list (- (apply 'min (mapcar 'sx:u (cdr item))) 4.3)
                          (+ (apply 'max (mapcar 'sx:u (cdr item))) 6.9)) knx)))
  (foreach item segs
   (if (sx:wantM (car item))
   (sx:ttry (strcat "M|" (car item)) (function (lambda ()
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
    (sx:textr (- umin 4.7) (+ sx:vBoxTop 0.6) 2.2 "KNX" sx:KNX 90.0)
    (sx:text (- umin 6.2) (+ sx:vBoxBot 0.9) 2.5 (sx:s (sx:c (car ds) sx:cMod)) sx:KNX)
    (sx:text (+ umin 3.0) (- sx:vBoxTop 3.4) 2.2 (sx:s (sx:c (car ds) sx:cModel)) sx:KNX)))
    (strcat "module " (car item)))))
  ;; шина KNX — одна линия через все модули щита, между листами — разрыв со стрелками
  (if (and knx (or (not sx:part) sx:onlyM))
    (sx:ttry "K" (function (lambda ()
      (sx:hline (apply 'min (mapcar 'car knx)) (apply 'max (mapcar 'cadr knx)) sx:vKNX sx:KNX)))
      "knx bus")))

;;; у строки нет кода УГО: BK = "BOX", а название УГО (BY) пустое (или 0)
(defun sx:ugo-noby (d sym)
  (and (= sym "BOX") (member (sx:s (sx:c d sx:cUgoRu)) '("" "0")) T))

;;; кабельный вывод: линия от клеммы к верху символа + блок «Вывод» (или простой значок SX_U_OUT)
(defun sx:draw-outlet (d u / blk ob)
  (setq blk (sx:blk sx:bOut "SX_U_OUT"))
  (sx:line u sx:xtb u sx:vOutTop sx:L)
  (setq ob (sx:insv blk u sx:vOutTop 1.0 0.0 nil sx:BL))
  (sx:red ob)
  ob)

;;; Ввод от питающего щита слева: вниз, под рамкой щита, вверх к XT.
;;; Кабель/фаза/ток остаются в том же столбце таблицы. УГО отходящей линии нет.
(defun sx:draw-input (d u / xs)
  ;; ввод со своим аппаратом рисует sx:draw-feed (вместе с группами)
  (if (not (sx:feed-p d))
    (progn
      (setq xs (- u sx:inDx))
      (sx:line xs sx:inFeed xs sx:inLow sx:L)
      (sx:hline xs u sx:inLow sx:L)
      (sx:line u sx:inLow u sx:xtb sx:L))))

;;; ---------------------------------------------------------------- ввод с питающим щитом (v3.21)
(defun sx:feed-p (d) (and (sx:input-p d) (sx:has d sx:cQFown)))

;;; «Сеть_щита»: по каждому вводу щита (n QF-фидера автомат кабель L щит-источник).
;;; Ввод 1 — строка щита 6..8, столбцы H:L (раздел А) или ручной раздел 2; вводы 2, 3 — AL:AP.
(defun sx:read-feeds (panel / sh rg vals cell idx k r c res src mark sec cores len in ch row)
  (setq res nil)
  (if (setq sh (cdr (assoc (sx:ru '(1057 1077 1090 1100 95 1097 1080 1090 1072)) sx:sheets)))
    (progn
      (setq rg (vlax-get-property sh 'Range "A1:AP30"))
      (setq vals (vl-catch-all-apply
                   '(lambda () (vlax-safearray->list (vlax-variant-value (vlax-get-property rg 'Value2)))) nil))
      (sx:rel rg)
      (if (not (vl-catch-all-error-p vals))
        (progn
          (setq vals (mapcar '(lambda (rw) (mapcar 'sx:v rw)) vals))
          (defun cell (r c) (nth (1- c) (nth (1- r) vals)))
          (setq k 1)
          (repeat 3
            (if (and (null idx) (= (sx:s (cell (+ 5 k) 2)) panel)) (setq idx k))
            (setq k (1+ k)))
          (if idx
            (progn
              (setq row (+ 5 idx))
              (setq res (list (list 1 (sx:s (cell row 8)) (sx:s (cell row 9)) (sx:s (cell row 10))
                                    (sx:s (cell row 11)) (sx:s (cell row 12)))))
              ;; раздел 1 пуст — питающая линия вручную (раздел 2, столбец щита C/D/E)
              (if (= (strcat (nth 2 (car res)) (nth 3 (car res)) (nth 5 (car res))) "")
                (progn
                  (setq c (+ 2 idx)
                        src (sx:s (cell 12 c)) mark (sx:s (cell 18 c)) sec (sx:s (cell 19 c))
                        cores (sx:s (cell 20 c)) len (sx:s (cell 22 c)) in (sx:s (cell 24 c)) ch (sx:s (cell 25 c)))
                  (if (= cores "") (setq cores "5"))
                  (if (= ch "") (setq ch (sx:ru '(1057))))
                  (setq res (list (list 1 ""
                                        (if (/= in "") (strcat "3P " ch in (sx:ru '(1040))) "")
                                        (if (/= mark "") (strcat mark (if (/= sec "") (strcat " (" cores (sx:ru '(1093)) sec ")") "")) "")
                                        len src)))))
              (setq k 2)
              (repeat 2
                (setq row (+ 6 (* (1- idx) 2) (- k 2)))
                (setq res (append res (list (list k (sx:s (cell row 38)) (sx:s (cell row 39)) (sx:s (cell row 40))
                                                    (sx:s (cell row 41)) (sx:s (cell row 42))))))
                (setq k (1+ k)))))))))
  res)

;;; номер ввода строки: «Ввод.2» -> 2 (иначе 1)
(defun sx:feed-n (d / line k)
  (setq line (sx:s (sx:c d sx:cLine)) k (vl-string-search "." line))
  (if (and k (> (atoi (substr line (+ k 2))) 0)) (atoi (substr line (+ k 2))) 1))

;;; данные фидера ввода: (QF автомат кабель L источник); кабель — из строки ввода, если в «Сеть_щита» пусто
(defun sx:feed-data (d / f cab k)
  (setq f (cdr (assoc (sx:feed-n d) sx:feeds)))
  ;; вводы 2, 3: пусто = как Ввод 1 («Сеть_щита»)
  (if (or (null f) (= (apply 'strcat f) "")) (setq f (cdr (assoc 1 sx:feeds))))
  (if (null f) (setq f (list "" "" "" "" "")))
  (if (= (nth 2 f) "")
    (progn
      (setq cab (sx:s (sx:c d sx:cCab)) k (vl-string-search " L=" cab))
      (setq f (list (nth 0 f) (nth 1 f) (if k (substr cab 1 k) cab) (nth 3 f) (nth 4 f)))))
  f)

;;; геометрия ввода: (уровень xl xr xb верх-рамки шинаL шинаN низ-фидера верх-QS)
;;; уровень 1 — первый ввод листа (рамка от верха листа), 2 — следующие (рамка под основной шиной).
;;; xb — куда петля от QS поднимается к шине (с неё начинается шина щита).
(defun sx:feed-geo (q draw / s us u ins lev xl xr xb top busL busN fb)
  (setq s (sx:sheet (sx:i q)) us (* s *sx-sheet* *sx-step*) u (sx:u q)
        ins (vl-remove-if-not '(lambda (d) (and (sx:feed-p d) (= (sx:sheet (sx:i d)) s))) draw)
        lev (if (vl-some '(lambda (d) (< (sx:u d) u)) ins) 2 1))
  ;; следующее место свободно — петля через его середину, иначе по границе мест
  (setq xb (if (vl-some '(lambda (d) (and (= (sx:i d) (1+ (sx:i q))) (or (sx:content-p d) (sx:has d sx:cQFown)))) draw)
             (+ u (/ *sx-step* 2.0))
             (+ u *sx-step*)))
  (if (= (sx:sheet (1+ (sx:i q))) s) nil (setq xb (+ u (/ *sx-step* 2.0))))
  (setq xl (max (+ us 1.0) (- u sx:fdL))
        xr (min (+ u sx:fdR) (- xb 3.0)))
  (if (= lev 1)
    (setq top sx:vTop busL sx:vBusL busN sx:vBusN fb (- sx:vBusL sx:fdD1))
    (setq top (- sx:vBusN 6.0) busL sx:vBus2N busN (- sx:vBus2N 5.0) fb (- sx:vBus2N sx:fdD2)))
  (list lev xl xr xb top busL busN fb (- fb 9.0)))

;;; первый ввод листа (рамка от верха): (запись геометрия) или nil
(defun sx:sheet-feed (draw s / q)
  (setq q (car (vl-sort (vl-remove-if-not '(lambda (d) (and (sx:feed-p d) (= (sx:sheet (sx:i d)) s))) draw)
                        '(lambda (a b) (< (sx:u a) (sx:u b))))))
  (if q (list q (sx:feed-geo q draw))))

;;; автомат QF1 между низом bot и точкой подключения top (вид «QF», длина подгоняется)
(defun sx:feed-qf (u bot top vals vis / ob r p sts)
  ;; QS: вид «QS» вашего QF1, если он есть; иначе свой блок SX_QS
  (if (and (= vis "QS") (sx:blk-p sx:bQF) (null sx:qfvis))
    (progn (setq ob (sx:insv sx:bQF u (+ bot sx:qfBody) 1.0 0.0 nil sx:BL))
           (setq p (sx:dynprop ob sx:pVis)
                 sts (if p (vl-catch-all-apply '(lambda () (mapcar 'sx:v (vlax-safearray->list (vlax-variant-value (vla-get-allowedvalues p))))) nil)))
           (setq sx:qfvis (if (or (null sts) (vl-catch-all-error-p sts)) '("?") sts))
           (vla-delete ob)))
  (if (and (= vis "QS") (not (member "QS" sx:qfvis)))
    (progn
      (setq ob (sx:insv "SX_QS" u (+ bot sx:qfBody) 1.0 0.0 vals sx:BL))
      (sx:line u top u (+ bot (- sx:qfBody 36.27)) sx:L)
      (sx:dot u top sx:L)
      ob)
  (if (sx:blk-p sx:bQF)
    (progn
      (setq ob (sx:insv sx:bQF u (+ bot sx:qfBody) 1.0 0.0 vals sx:BL))
      (if (and (null sx:qfvis) (setq p (sx:dynprop ob sx:pVis)))
        (progn
          (setq sts (vl-catch-all-apply '(lambda () (mapcar 'sx:v (vlax-safearray->list (vlax-variant-value (vla-get-allowedvalues p))))) nil))
          (setq sx:qfvis (if (vl-catch-all-error-p sts) '("?") sts))
          (sx:log (strcat "QF1 visibility states: " (sx:str sx:qfvis)))))
      (sx:dynset ob sx:pVis (if (member vis sx:qfvis) vis "QF"))
      (setq r (sx:qf-align ob (cadr (sx:pt u top))))
      (if (numberp r) (sx:line u top u (- r (cadr sx:P0)) sx:L))
      ob)
    (progn
      (sx:insv "SX_QF" u (+ bot sx:qfBody) 1.0 0.0 vals sx:BL)
      (sx:line u top u (- (+ bot sx:qfBody) sx:vQFcontact) sx:L)))))

(defun sx:qf-vals (name txt / toks p a ma)
  (setq toks (sx:split txt) p "" a "" ma "")
  (foreach tk toks
    (cond ((wcmatch (strcase tk) "#P") (setq p tk))
          ((wcmatch (strcase tk) "Q@,Q@@") nil)   ; «QS 3P 80А»: буквы аппарата — не номинал
          ((= a "") (setq a tk))
          (t (setq ma (if (= ma "") tk (strcat ma " " tk))))))
  (list (sx:digits name) p a ma))

(defun sx:draw-feed (q draw / g f u lev xl xr xb top busL busN fb qsTop tgt)
  (setq g (sx:feed-geo q draw) f (sx:feed-data q) u (sx:u q)
        lev (nth 0 g) xl (nth 1 g) xr (nth 2 g) xb (nth 3 g) top (nth 4 g)
        busL (nth 5 g) busN (nth 6 g) fb (nth 7 g) qsTop (nth 8 g))
  (sx:log (strcat "feed " (sx:rowinfo q) " lev=" (itoa lev) " data=" (sx:str f)))
  ;; питающий щит: рамка, имя, шина, автомат фидера
  (sx:rect xl top xr (- fb 3.0) sx:B)
  (sx:text xl (+ top 1.5) 2.5 (nth 4 f) sx:T)
  (sx:bus (- xl 3.0) (- xr 2.0) busL busN "L1,L2,L3")
  (sx:feed-qf u fb busL (sx:qf-vals (nth 0 f) (nth 1 f)) "QF")
  ;; кабель фидера и подписи (марка, длина)
  (sx:line u fb u qsTop sx:L)
  (sx:textr (+ u sx:fdTxt1) (+ sx:fdQsBot 1.0) 2.5 (nth 2 f) sx:T 90.0)
  (if (/= (nth 3 f) "")
    (sx:textr (+ u sx:fdTxt2) (+ sx:fdQsBot 1.0) 2.5 (strcat "L=" (nth 3 f) (sx:ru '(1084))) sx:T 90.0))
  ;; вводной аппарат щита и петля к началу шины
  (sx:feed-qf u sx:fdQsBot qsTop (sx:qf-vals (sx:s (sx:c q sx:cQFown)) (sx:s (sx:c q sx:cQFtxt)))
              (if (or (wcmatch (strcase (sx:s (sx:c q sx:cQFtxt))) "QS*") (wcmatch (strcase (sx:s (sx:c q 7))) "QS*")) "QS" "QF"))   ; вид «QS» — если он есть в вашем блоке QF1
  (setq tgt (if (sx:is1 q sx:cBus2) sx:vBus2L sx:vBusL))
  (sx:line u sx:fdQsBot u sx:fdLoop sx:L)
  (sx:hline u xb sx:fdLoop sx:L)
  (sx:line xb sx:fdLoop xb tgt sx:L))

(defun sx:draw-ugo (d u / sym spec blk ob)
  (if (sx:input-p d)
    (sx:draw-input d u)
    (progn
  (setq sym (sx:s (sx:c d sx:cSym)))
  (if (= sym "") (setq sym "BOX"))
  (if (or (sx:ugo-noby d sym)
          (and (not (assoc sym sx:ugo)) (not (sx:blk-p sym))))
    ;; нет кода УГО или код неизвестен -> красный кабельный вывод
    (sx:draw-outlet d u)
    (progn
      (setq spec (cond ((assoc sym sx:ugo))
                       (t (list sym sym 1.0 0.0 68.10 72.40)))   ; свой блок из таблицы значков
            blk (nth 1 spec))
      (if (not (sx:blk-p blk))
        (setq blk (cond ((= sym "SOCKET") "SX_U_SOCKET") ((= sym "LAMP") "SX_U_LAMP")
                        ((member sym '("MOTOR" "SERVO" "VALVE")) "SX_U_M") (t "SX_U_BOX"))
              spec (list sym blk 1.0 0.0 (nth 4 spec) (nth 5 spec))))
      (if (wcmatch (strcase blk) "SX_U_*")
        ;; своего блока УГО нет (рисовалась бы белая заглушка) -> красный кабельный вывод
        (sx:draw-outlet d u)
        (progn
      (sx:line u sx:xtb u (nth 5 spec) sx:L)
      (setq ob (sx:insv blk u (nth 4 spec) (nth 2 spec) (nth 3 spec)
                        (if (= sym "BOX") (list (sx:s (sx:c d sx:cLine))) (list "")) sx:BL))
      (if (/= sym "SOCKET") (sx:red ob))
      ob)))))))

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

;;; ---------------------------------------------------------------- ДСУП (v3.23)
(defun sx:dsup-p (d) (sx:is1 d sx:cDsup))
;;; КУП (или ШДУП) — коробка уравнивания потенциалов: блок «_КУП», к ней идут отводы точек РЕ
(defun sx:kup-p (d)
  (and (member (strcase (sx:s (sx:c d sx:cCode)))
               (list (sx:ru '(1050 1059 1055)) (sx:ru '(1064 1044 1059 1055))))
       T))

;;; окрасить вставку в зелёный (цвет PE) — для блоков с цветом ПоБлоку
(defun sx:green (ob / e ed)
  (if (and ob (setq e (vlax-vla-object->ename ob)) (setq ed (entget e)))
    (entmod (append (vl-remove-if '(lambda (x) (member (car x) '(62 420))) ed) (list (cons 62 3))))))

;;; линия ДСУП: от шины PE (точка) или отводом от линии своей КУП (CK), внизу КУП или вывод.
;;; sx:dsLast: (имя КУП . столбец последнего отвода) — следующий отвод той же КУП примыкает к нему.
(setq sx:dsLast nil)
(defun sx:draw-dsup (d u draw / from par pu dir lastu kup bot ob)
  (setq from (strcase (sx:s (sx:c d sx:cDsupFrom)))
        kup (sx:kup-p d)
        bot (if (and kup (sx:blk-p sx:bKUP)) sx:vKupTop sx:vDsOut))
  (if (/= from "")
    (setq par (car (vl-remove-if-not
                     '(lambda (q) (and (sx:dsup-p q) (= (strcase (sx:s (sx:c q sx:cLine))) from))) draw))))
  (if (and par (/= (sx:u par) u))
    (progn
      (setq pu (sx:u par) dir (if (> u pu) 1.0 -1.0)
            lastu (cdr (assoc from sx:dsLast)))
      ;; вниз-вверх: от вывода до фаски, фаска 2x2 к горизонтали
      (sx:line u bot u (- sx:vDsStep 2.0) sx:PE)
      (sx:line u (- sx:vDsStep 2.0) (- u (* 2.0 dir)) sx:vDsStep sx:PE)
      (if (and lastu (< (* dir (- pu lastu)) 0) (< (* dir (- lastu u)) 0))
        ;; между КУП и этим столбцом уже есть отвод той же КУП — горизонталь до его фаски
        (sx:hline (- u (* 2.0 dir)) (- lastu (* 2.0 dir)) sx:vDsStep sx:PE)
        (progn
          (sx:hline (- u (* 2.0 dir)) (+ pu (* 2.0 dir)) sx:vDsStep sx:PE)
          (sx:line (+ pu (* 2.0 dir)) sx:vDsStep pu (- sx:vDsStep 2.0) sx:PE)))
      (setq sx:dsLast (cons (cons from u) (vl-remove-if '(lambda (x) (= (car x) from)) sx:dsLast))))
    (progn
      (if (/= from "") (sx:log (strcat "DSUP: KUP not found for place " (itoa (1+ (sx:i d))) " -> PE bus")))
      (sx:line u bot u sx:vPE sx:PE)
      (sx:dot u sx:vPE sx:PE)))
  (if (and kup (sx:blk-p sx:bKUP))
    (setq ob (sx:insv sx:bKUP (+ u sx:kupDx) sx:vKupIns sx:kupSc 90.0 nil sx:BL))
    (setq ob (sx:insv (sx:blk sx:bOut "SX_U_OUT") u sx:vDsOut 1.0 0.0 nil sx:BL)))
  (sx:green ob)
  ob)

(defun sx:draw-rows (draw / prev u bf xt two blkn ob xseen)
  (setq prev nil xseen nil sx:dsLast nil)
  (foreach d draw
   (if (not (sx:wantP (sx:i d)))
    ;; место не перестраивается: только состояние для следующих строк (перемычки «=», повтор XT)
    (progn
      (if (sx:is1 d sx:cXT)
        (progn (setq xt (sx:s (sx:c d sx:cXTtxt)))
               (if (/= xt "") (setq xseen (cons (strcase xt) xseen)))
               (setq prev d)))
      ;; отвод ДСУП не перестраивается, но следующий отвод той же КУП примыкает к нему
      (if (and (sx:dsup-p d) (sx:has d sx:cDsupFrom))
        (setq sx:dsLast (cons (cons (strcase (sx:s (sx:c d sx:cDsupFrom))) (sx:u d))
                              (vl-remove-if '(lambda (x) (= (car x) (strcase (sx:s (sx:c d sx:cDsupFrom))))) sx:dsLast)))))
   (sx:ttry (strcat "R|" (itoa (sx:i d))) (function (lambda ()
    (setq u (sx:u d) bf (sx:s (sx:c d sx:cChT)) sx:xtb sx:vXTbot)
    ;; ДСУП: без автомата и клемм, от шины PE или от своей КУП
    (if (sx:dsup-p d) (sx:draw-dsup d u draw))
    (if (and (sx:input-p d) (sx:is1 d sx:cHasLine) (not (sx:is1 d sx:cXT)) (not (sx:c d sx:cCh)))
      (progn (setq sx:xtb sx:vIN0) (sx:draw-input d u)))
    ;; линия без клеммы и канала (б/к: DALI, PE...) со своим автоматом: спуск от автомата прямо к УГО
    (if (and (not (sx:input-p d)) (sx:is1 d sx:cHasLine) (not (sx:is1 d sx:cXT)) (not (sx:c d sx:cCh))
             (sx:has d sx:cQFown))
      (progn (setq sx:xtb (sx:qf-low u))
             (sx:ticks u (/ (+ sx:xtb sx:vXTtop) 2.0) (sx:nph d) sx:L)
             (sx:draw-ugo d u)
             (setq sx:xtb sx:vXTbot)))
    (if (and (sx:c d sx:cCh) (/= bf "="))
      (progn
        (sx:ins "SX_CH" u sx:vChBase (list (cons "CH" (sx:s (sx:c d sx:cCh)))) sx:BL)
        (cond
          ((sx:is1 d sx:cXT) (sx:line u sx:vChBase u sx:vXTtop sx:L))
          ;; канал без клеммы (напр. датчики протечки): сразу линия к УГО
          ((sx:is1 d sx:cHasLine)
           (sx:line u sx:vChBase u sx:vXTbot sx:L)
           (sx:draw-ugo d u)))))
    (if (and (sx:is1 d sx:cXT) (not (sx:feed-p d)))   ; у ввода с питающим щитом клеммы нет
      (progn
        ;; двухуровневая клемма (XT без земли) - только когда на клемму приходят два проводника:
        ;; параллельная линия «=» или повтор номера XT; обычная клемма - XT с кружком PE
        (setq xt (sx:s (sx:c d sx:cXTtxt))
              two (or (and (= bf "=") prev) (and (/= xt "") (member (strcase xt) xseen))
                      (sx:is1 d sx:cLow))   ; слаботочная линия — клемма без PE
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
    (strcat "row " (sx:rowinfo d))))))

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
    (if (and (sx:is1 d sx:cHasLine) (not (sx:input-p d)))
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
(defun sx:validate-layout (recs draw / res qfids id pl lo hi key seenk prevx kn)
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
  ;; (2а) узел КМ: проверка из книги (CI); линия шины KM-n не правее узла
  (foreach d draw
    (if (and (sx:kma-p d) (not (wcmatch (sx:s (sx:c d sx:cKMchk)) "ok*")))
      (setq res (cons (strcat (sx:ru '(1091 1079 1077 1083 32)) (sx:kma-name d) " (" (itoa (1+ (sx:i d))) "): "
                              (if (= (sx:s (sx:c d sx:cKMchk)) "") "CI = ? (patch_km.py)" (sx:s (sx:c d sx:cKMchk))))
                      res)))
    (if (and (sx:is1 d sx:cBus2) (setq kn (sx:kma-find (sx:s (sx:c d sx:cBusName)) draw))
             (< (sx:i d) (+ (sx:i kn) 3)))
      (setq res (cons (strcat (sx:ru '(1091 1079 1077 1083 32)) (sx:kma-name kn) ": " (itoa (1+ (sx:i d))) " < " (itoa (+ (sx:i kn) 4)))
                      res))))
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
            (if (and (not (sx:has d sx:cPow)) (not (sx:dsup-p d))) (setq reasons (append reasons (list (sx:ru '(1085 1077 1090 32 1084 1086 1097 1085 1086 1089 1090 1080))))))))
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
(defun c:SXUPDSEL () (sx:main 'sel))
(defun c:SXUPDSHEET () (sx:main 'sheet))

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

(defun sx:main (upd / *error* data recs panel pos i row draw oldcmd maxi sym probs xrow ridx pick dkeys ndel)
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
        sx:bbcache nil sx:cccache nil sx:blkcache nil sx:qfvis nil sx:bbctx "" sx:pecache nil sx:nopos nil
        sx:part (member upd '(sel sheet)) sx:onlyP nil sx:onlyG nil sx:onlyM nil sx:onlyS nil)
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
  ;; SXUPDSEL: сначала выбор, Excel открывается только если есть что обновлять
  (if sx:part
    (progn
      (sx:stg "pick")
      (setq pick (sx:pick-keys))
      (cond
        ((null pick) (setvar "CMDECHO" oldcmd) (exit))
        ((null (cadr pick))
         (princ (strcat "\n" (sx:ru '(1057 1088 1077 1076 1080 32 1074 1099 1073 1088 1072 1085 1085 1099 1093 32 1085 1077 1090 32 1087 1086 1084 1077 1095 1077 1085 1085 1099 1093 32 1086 1073 1098 1077 1082 1090 1086 1074 46 32 1054 1076 1080 1085 32 1088 1072 1079 32 1074 1099 1087 1086 1083 1085 1080 1090 1077 32 83 88 85 80 68 65 84 69 32 40 118 51 46 49 53 32 1080 32 1085 1086 1074 1077 1077 41 32 8212 32 1087 1086 1089 1083 1077 32 1101 1090 1086 1075 1086 32 1084 1086 1078 1085 1086 32 1086 1073 1085 1086 1074 1083 1103 1090 1100 32 1095 1072 1089 1090 1103 1084 1080 46))))
         (setvar "CMDECHO" oldcmd) (exit))
        ((> (length (car pick)) 1)
         (princ (strcat "\n" (sx:ru '(1042 1099 1073 1088 1072 1085 1099 32 1086 1073 1098 1077 1082 1090 1099 32 1088 1072 1079 1085 1099 1093 32 1097 1080 1090 1086 1074 32 8212 32 1074 1099 1073 1077 1088 1080 1090 1077 32 1086 1076 1080 1085 32 1097 1080 1090 46))))
         (setvar "CMDECHO" oldcmd) (exit)))
      (if (= upd 'sheet) (setq pick (list (car pick) (sx:keys->sheets (cadr pick)) (caddr pick))))
      (sx:log (strcat "pick: panel=" (car (car pick)) " keys=" (sx:str (cadr pick)) " untagged=" (itoa (caddr pick))))))
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
  ;; SXUPDSEL: щит — из меток выбранных объектов
  (if sx:part
    (if (member (car (car pick)) sx:panels)
      (setq sx:pi (vl-position (car (car pick)) sx:panels))
      (progn (princ (strcat "\n" (sx:ru '(1065 1080 1090 32 1080 1079 32 1084 1077 1090 1086 1082 32 1085 1077 32 1085 1072 1081 1076 1077 1085 32 1074 32 1090 1072 1073 1083 1080 1094 1077 58 32)) (car (car pick))))
             (sx:xl-release) (setvar "CMDECHO" oldcmd) (exit))))
  (setq panel (nth sx:pi sx:panels) sx:tagpanel (if panel panel ""))
  (if panel (setenv "SX_LastPanel" panel))
  (setenv "SX_LastXls" sx:path)
  (sx:save-params panel)
  ;; система координат: вдоль X, вверх Y
  (setq sx:P0 sx:inspt sx:R '(1.0 0.0) sx:D '(0.0 1.0) sx:ang 0.0)
  (setq sx:style (sx:ensure-style))
  (sx:log (strcat "style=" sx:style))
  (sx:stg "read excel")
  (setq data (sx:read-data))
  (setq sx:summary (sx:read-summary panel))
  (setq sx:sources (vl-catch-all-apply 'sx:read-sources (list panel)))
  (if (vl-catch-all-error-p sx:sources) (setq sx:sources nil sx:cats nil))
  (sx:log (strcat "sources=" (sx:str sx:sources) " cats=" (sx:str sx:cats)))
  (setq sx:feeds (vl-catch-all-apply 'sx:read-feeds (list panel)))
  (if (vl-catch-all-error-p sx:feeds) (setq sx:feeds nil))
  (sx:log (strcat "feeds=" (sx:str sx:feeds)))
  (sx:log (strcat "summary=" (sx:str sx:summary)))
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
  (if sx:part
    (progn
      (setq dkeys (sx:sel-units (cadr pick) draw))
      (sx:log (strcat "part: sheets=" (sx:str sx:onlyS) " places=" (sx:str sx:onlyP) " qf=" (sx:str sx:onlyG) " modules=" (sx:str sx:onlyM)))))
  (if recs (sx:log (strcat "first " (sx:rowinfo (car recs)))))
  ;; проверка перед рисованием: все проблемы разом, до очистки и построения
  (sx:stg "validate")
  (sx:roles)
  (setq probs (if sx:part
                  (if (vl-some '(lambda (d) (sx:wantP (sx:i d))) draw)
                    (sx:validate panel (vl-remove-if-not '(lambda (d) (sx:wantP (sx:i d))) recs)
                                       (vl-remove-if-not '(lambda (d) (sx:wantP (sx:i d))) draw)))
                  (sx:validate panel recs draw)))
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
  (if (/= (vla-get-name sx:doc) (getvar "DWGNAME"))
    (progn (sx:log "Drawing switched before generation; no objects changed.") (setvar "CMDECHO" oldcmd) (exit)))
  (vla-startundomark sx:doc)
  ;; Adopt the matching manual sample before clearing generated geometry.
  (sx:layers) (sx:summary-capture)
  (if (or sx:part sx:do-clear) (sx:adopt-inputs draw))
  (sx:stg "clear")
  (if sx:part
    (setq ndel (sx:del-keys panel dkeys))
    (if sx:do-clear (sx:clear)))
  (sx:del-keys panel '("U"))
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
      (append (list sx:bQF sx:bBUS sx:bHEAD sx:bXT sx:bXT2 sx:bOut sx:bT sx:bPS sx:bMK sx:bKUP) (mapcar 'cadr sx:ugo) sx:custom))))
    "import-blocks")
  (sx:lib-close)
  (setq sx:blkcache nil)
  (foreach n (append (list sx:bQF sx:bBUS sx:bHEAD sx:bXT sx:bXT2 sx:bOut sx:bT sx:bPS sx:bMK sx:bKUP) (mapcar 'cadr sx:ugo))
    (sx:log (strcat "block " n (if (tblsearch "BLOCK" n) " OK" " MISSING"))))
  (if (not (tblsearch "BLOCK" sx:bQF))
    (progn (princ (strcat "\n" (sx:ru '(1042 1085 1080 1084 1072 1085 1080 1077 58 32 1085 1077 32 1085 1072 1081 1076 1077 1085 1099 32 1074 1072 1096 1080 32 1073 1083 1086 1082 1080 32 40 81 70 49 44 32 88 84 46 46 46 41 32 8212 32 1085 1072 1088 1080 1089 1086 1074 1072 1085 1099 32 1079 1072 1084 1077 1085 1080 1090 1077 1083 1080 46 32 1055 1088 1086 1074 1077 1088 1100 1090 1077 32 1087 1091 1090 1100 32 1082 32 1073 1080 1073 1083 1080 1086 1090 1077 1082 1077 32 48 50 95 1054 1076 1085 1086 1083 1080 1085 1077 1081 1082 1072 46 100 119 103 44 32 1089 1084 46 32 83 88 95 108 111 103 46 116 120 116))))
           (setq sx:errs (cons "blocks: QF1 missing" sx:errs))))
  (sx:stg "table")   (sx:try (function (lambda () (sx:draw-table recs))) "table")
  (if (or (not sx:part) sx:onlyS)
    (progn (sx:stg "sheets") (sx:try (function (lambda () (sx:draw-sheets draw panel))) "sheets")))
  (sx:stg "groups")  (sx:try (function (lambda () (sx:draw-groups draw))) "groups")
  (sx:stg "modules") (sx:try (function (lambda () (sx:draw-modules draw))) "modules")
  (sx:stg "rows")    (sx:try (function (lambda () (sx:draw-rows draw))) "rows")
  (sx:stg "summary") (sx:try 'sx:draw-summary "summary")
  (sx:stg "sources") (sx:try 'sx:draw-sources "sources")
  (sx:stg "done")
  (vla-endundomark sx:doc)
  (setvar "CMDECHO" oldcmd)
  (if sx:part
    (princ (strcat "\n" (sx:ru '(1055 1077 1088 1077 1089 1090 1088 1086 1077 1085 1086 32 1074 1099 1073 1086 1088 1086 1095 1085 1086 58 32 1083 1080 1089 1090 1086 1074 32)) (itoa (length sx:onlyS))
                   (sx:ru '(44 32 1084 1077 1089 1090 32)) (itoa (length (vl-remove-if-not '(lambda (d) (sx:wantP (sx:i d))) recs)))
                   (sx:ru '(44 32 1072 1074 1090 1086 1084 1072 1090 1086 1074 32)) (itoa (length sx:onlyG))
                   (sx:ru '(44 32 1084 1086 1076 1091 1083 1077 1081 32)) (itoa (length sx:onlyM))
                   (sx:ru '(32 40 1091 1076 1072 1083 1077 1085 1086 32 1089 1090 1072 1088 1099 1093 32 1086 1073 1098 1077 1082 1090 1086 1074 58 32)) (itoa ndel) ")"))
  (princ (strcat "\n" (sx:ru '(1043 1086 1090 1086 1074 1086 58 32 1089 1093 1077 1084 1072 32 1087 1086 1089 1090 1088 1086 1077 1085 1072 32 1085 1072 32 1089 1083 1086 1103 1093 32 83 88 95 42 46 32 1052 1077 1089 1090 58)) " " (itoa (length recs)))))
  (if sx:errs
    (progn
      (princ (strcat "\nErrors: " (itoa (length sx:errs))))
      (foreach e (reverse sx:errs) (princ (strcat "\n  " e)))))
  (sx:logdump)
  (princ))

(princ (strcat "\n" (sx:ru '(83 88 95 83 99 104 101 109 97 32 118 51 46 49 55 32 1079 1072 1075 1088 1091 1078 1077 1085 32 40 1074 1072 1096 1080 32 1073 1083 1086 1082 1080 41 58 32 83 88 68 82 65 87 32 8212 32 1087 1086 1089 1090 1088 1086 1080 1090 1100 32 1089 1093 1077 1084 1091 44 32 83 88 67 76 69 65 82 32 8212 32 1089 1090 1077 1088 1077 1090 1100 46))))
(princ "\n[SX_Schema v3.23] SXDRAW / SXUPDATE / SXUPDSEL / SXUPDSHEET / SXCLEAR")
(princ)
