;;; ============================================================================
;;;  SX_Layout.lsp  v0.3 for v5 — компоновка электрощита (вид спереди) из той же книги Excel
;;;  Книги: v5 (.xlsm), v4, v3.12 — листы «Однолинейка» и «Каталог».
;;;
;;;  Команды:
;;;    SXLAYOUT  — окно: книга Excel, щит, база блоков компоновки (DWG), корпус
;;;                (подбор или вручную), запас, точка -> строит компоновку 1:1 (мм).
;;;    SXLUPDATE — перестроить с параметрами, запомненными в чертеже (без окна).
;;;    SXLCLEAR  — удалить построенную компоновку (слои SXL_*).
;;;
;;;  Нужен SX_Schema.lsp (чтение Excel, библиотека блоков, русский текст): если он
;;;  не загружен, ищется в путях поддержки, иначе спрашивается один раз.
;;;
;;;  Что берётся из книги:
;;;    «Однолинейка» (выбранный щит): автоматы (G обозначение, H текст), аппараты
;;;      (I-K), модули (AH обозначение, AI модель), клеммы XT (P; трёхуровневых BW,
;;;      доп. проходных BX);
;;;    «Каталог»: ширина в модулях (L), блок компоновки (Q, необязательно; число в Q — ток канала
;;;      модулей книги v5, это не блок), ручные
;;;      позиции «Вручную / запас» (O) — клеммы, шины, вводные аппараты, которых нет в однолинейке.
;;;  Порядок на рейках — по типам: клеммы, шины, модули, автоматы, дифавтоматы, аппараты.
;;;  Блок ставится по габаритам: левый край вплотную к предыдущему аппарату,
;;;  середина по высоте — на ось рейки. Нет блока — прямоугольник по ширине из каталога.
;;;
;;;  Все строки-литералы — латиница; русский текст хранится кодами (sx:ru).
;;; ============================================================================

(vl-load-com)
(setq sxl:ver "v0.3")

;;; ---------------------------------------------------------------- настройки
(setq *sxl-mod*    18.0)    ; ширина модуля DIN, мм (S201 в базе = 18.0)
(setq *sxl-gap*    18.0)    ; промежуток между типами на одной рейке, мм (если тип не с новой рейки)
(if (null *sxl-reserve*) (setq *sxl-reserve* 20))   ; запас по длине реек при подборе корпуса, %
(if (null *sxl-newrail*) (setq *sxl-newrail* T))    ; каждый тип с новой рейки (если реек хватает)
(setq *sxl-th*     10.0)    ; высота обозначений над аппаратами, мм
(setq *sxl-thb*    20.0)    ; высота заголовка и таблицы, мм (печать 1:10 -> 2 мм)
(setq *sxl-dimsc*  10.0)    ; масштаб размеров (DIMSCALE) под печать 1:10
(setq *sxl-skip*   "D-*,SBF*,UK-E*")   ; модели «Каталога», которые не рисуются: торцевые изоляторы, перемычки, упоры
(setq *sxl-pBox*   125.0)   ; шаг реек в модульном боксе, мм
(setq *sxl-railH*  35.0)    ; высота рейки, если блока рейки нет, мм

;;; столбцы листа «Однолинейка» (A = 1)
(setq sxl:cPanel 1 sxl:cQF 7 sxl:cQFtxt 8 sxl:cDev1 9 sxl:cXT 16
      sxl:cMod 34 sxl:cModel 35 sxl:cPoles 46 sxl:cRcd 58 sxl:cDevC1 60 sxl:c3L 75 sxl:cPass 76)
;;; столбцы листа «Каталог» (A = 1), данные с 4-й строки
(setq sxl:kCat 1 sxl:kMan 2 sxl:kModel 3 sxl:kName 7 sxl:kPoles 9 sxl:kWidth 12
      sxl:kManual 15 sxl:kBlock 17 sxl:kOboz 21 sxl:kKey 22)

;;; ---------------------------------------------------------------- SX_Schema.lsp
;;; русский текст до загрузки SX_Schema (то же преобразование, что sx:cp)
(defun sxl:ru0 (codes)
  (apply 'strcat
    (mapcar '(lambda (c)
               (chr (cond ((< c 128) c)
                          ((and (getvar "LISPSYS") (>= (getvar "LISPSYS") 1)) c)
                          ((and (>= c 1040) (<= c 1103)) (+ (- c 1040) 192))
                          ((= c 1025) 168) ((= c 1105) 184) ((= c 171) 171) ((= c 187) 187)
                          (t 63))))
            codes)))

;;; загружен ли SX_Schema.lsp с нужными функциями (v3.12+)
(defun sxl:schema-ok ()
  (vl-every '(lambda (s) (eval s))
            '(sx:ru sx:open-wb sx:read-data sx:cell sx:xl-release sx:rel sx:lib-open sx:lib-close sx:import-blocks
              sx:join-sep sx:logdir sx:fill-list sx:dynset sx:fit sx:bbox sx:layer sx:ensure-style sx:load-panels)))

;;; где может лежать SX_Schema.lsp: пути поддержки, прошлый выбор, папка его .dcl,
;;; соседняя папка «Однолинейка» (рядом с «Компоновка» и «Блоки»)
(defun sxl:schema-cands ( / res sib f)
  (setq sib (strcat "\\" (R0 "Однолинейка") "\\SX_Schema.lsp"))
  (foreach f (list (findfile "SX_Schema.lsp") (getenv "SX_SchemaLsp")) (if f (setq res (cons f res))))
  (foreach f (list (findfile "SX_Schema.dcl") (getenv "SX_DclFile"))
    (if (and f (findfile f)) (setq res (cons (strcat (vl-filename-directory (findfile f)) "\\SX_Schema.lsp") res))))
  (foreach f (list (findfile "SX_Layout.dcl") (getenv "SXL_DclFile") (getenv "SXL_LibDwg") (getenv "SX_LibDwg"))
    (if (and f (findfile f))
      (setq res (cons (strcat (vl-filename-directory (vl-filename-directory (findfile f))) sib) res))))
  (vl-remove-if-not 'sxl:schema-file-p (reverse res)))

;;; файл существует и называется именно SX_Schema.lsp (не SX_Layout.lsp и не копия)
(defun sxl:schema-file-p (f)
  (and f (findfile f) (= (strcase (strcat (vl-filename-base f) (vl-filename-extension f))) "SX_SCHEMA.LSP")))

(defun sxl:need-schema ( / f)
  (if (not (sxl:schema-ok))
    (progn
      (foreach f (sxl:schema-cands)
        (if (not (sxl:schema-ok))
          (progn (load (findfile f) nil) (if (sxl:schema-ok) (setenv "SX_SchemaLsp" (findfile f))))))
      (if (and (not (sxl:schema-ok))
               (setq f (getfiled (R0 "SXLAYOUT: укажите файл SX_Schema.lsp (однолинейка, папка 2_AutoCAD/Однолинейка)")
                                 "SX_Schema.lsp" "lsp" 0)))
        (if (sxl:schema-file-p f)
          (progn (load f nil) (if (sxl:schema-ok) (setenv "SX_SchemaLsp" f)))
          (princ (strcat "\n" (R0 "Выбран не тот файл: нужен SX_Schema.lsp, а не ") (vl-filename-base f) (vl-filename-extension f)))))))
  (if (sxl:schema-ok)
    T
    (progn
      (princ (strcat "\n" (R0 "SXLAYOUT: не загружен SX_Schema.lsp v3.12 или новее (2_AutoCAD/Однолинейка). ")
                     (R0 "Загрузите его через APPLOAD (или добавьте папку в пути поддержки) и повторите.")))
      nil)))

;;; ---------------------------------------------------------------- справочники (после загрузки SX_Schema)
(defun sxl:init ()
  (setq sxl:bRail (R "Дин.рейка")
        sxl:pLen  (R "Расстояние1"))
  ;; корпуса: основной источник — файл SX_Корпуса.csv рядом с программой (sxl:load-encl).
  ;; Этот короткий список — только запасной, если файла нет.
  ;; Формат записи: (имя тип Ш В Г рядов модулей_в_ряду секций шаг_реек короб поле серия артикул установка)
  ;;   тип "MOD"   — модульный щиток: рядов x модулей (Ш, В, Г — габарит щитка)
  ;;   тип "PLATE" — навесной корпус с монтажной платой: рейки шагом «шаг», между ними кабель-каналы «короб»
  ;;   тип "CAB"   — напольный шкаф: так же, «секций» — монтажных полей по ширине
  (setq sxl:encl0
    (list (list (R "ЩРн-п-36 (2x18)") "MOD" 432.0 464.0 160.0 2 18 1 125.0 0.0 0.0 "RAM Base" "RB36W6TGD" (R "навесной"))
          (list "ST 600x400x250" "PLATE" 400.0 600.0 250.0 0 0 1 150.0 40.0 40.0 "ST" "R5ST0649" (R "навесной"))
          (list "CQE N 2000x800x400" "CAB" 800.0 2000.0 400.0 0 0 1 205.0 60.0 30.0 "CQE N" "" (R "напольный")))
        sxl:encl sxl:encl0)
  (setq sxl:inst-list (list (R "любая") (R "навесной") (R "встраиваемый") (R "напольный")))
  ;; разделы «Каталога» -> тип на компоновке
  (setq sxl:catmap
    (list (cons (R "Автомат") "QF") (cons (R "Дифавтомат") "QFD") (cons (R "Аппарат") "DEV")
          (cons (R "Модуль") "MOD") (cons (R "Клемма") "XT") (cons (R "Шины") "BUS")))
  ;; порядок типов на рейках (по типам)
  (setq sxl:order '("XT" "BUS" "MOD" "QF" "QFD" "DEV" "OTH"))
  ;; высота аппарата без блока, мм
  (setq sxl:hdef '(("XT" . 100.0) ("BUS" . 45.0) ("MOD" . 90.0) ("QF" . 85.0) ("QFD" . 85.0) ("DEV" . 85.0) ("OTH" . 85.0)))
  ;; обозначения аппаратов после автомата (код столбцов BH-BJ -> буквы)
  (setq sxl:devdes '(("KM" . "KM") ("T" . "TV") ("PS" . "G"))))

;;; ---------------------------------------------------------------- база корпусов (SX_Корпуса.csv)
;;; Столбцы (разделитель «;», первая строка — заголовок, строки с # — комментарий):
;;;   0 Серия; 1 Артикул; 2 Наименование; 3 Тип (MOD/PLATE/CAB); 4 Установка (навесной/встраиваемый/напольный);
;;;   5 Ширина; 6 Высота; 7 Глубина; 8 Рядов; 9 Модулей в ряду; 10 Секций; 11 Шаг реек; 12 Короб; 13 Поле
;;; Пустые 10-13 -> по типу: MOD шаг 125; PLATE шаг 150, короб 40, поле 40; CAB шаг 205, короб 60, поле 30.
(if (null sxl:maxW) (setq sxl:maxW 0.0 sxl:maxH 0.0 sxl:maxD 0.0 sxl:inst 0 sxl:series ""))
(if (null sxl:encname) (setq sxl:encname ""))

(defun sxl:split-sc (s / res k)
  (setq res '())
  (while (setq k (vl-string-search ";" s))
    (setq res (cons (vl-string-trim " \"\t" (substr s 1 k)) res) s (substr s (+ k 2))))
  (reverse (cons (vl-string-trim " \"\t\r" s) res)))
(defun sxl:fnum (s / v)
  (if (and (= (type s) 'STR) (setq v (distof (vl-string-translate "," "." s) 2))) v 0.0))
(defun sxl:nz (x d) (if (> x 0) x d))

(defun sxl:csv-file ( / f d nm)
  (setq nm (R "SX_Корпуса.csv"))
  (cond ((setq f (findfile nm)) f)
        ((and (setq d (sxl:dcl-file)) (setq f (findfile (strcat (vl-filename-directory d) "\\" nm)))) f)))

;;; внимание: символы AutoLISP не различают регистр — файл fh, а не h (иначе его затрёт высота H)
(defun sxl:load-encl ( / f fh l c res kind W H D nm)
  (setq res '())
  (if (and (setq f (sxl:csv-file)) (setq fh (open f "r")))
    (progn
      (while (setq l (read-line fh))
        (setq l (vl-string-left-trim (chr 65279) l) c (sxl:split-sc l))
        (if (and (>= (length c) 8) (not (wcmatch (car c) "`#*"))
                 (member (setq kind (strcase (nth 3 c))) '("MOD" "PLATE" "CAB"))
                 (> (setq W (sxl:fnum (nth 5 c))) 0) (> (setq H (sxl:fnum (nth 6 c))) 0))
          (progn
            (setq D (sxl:fnum (nth 7 c))
                  c (append c '("" "" "" "" "" "" "")))
            (setq nm (strcat (nth 0 c) " " (nth 1 c) " " (sxl:r0 H) "x" (sxl:r0 W) "x" (sxl:r0 D)
                             (if (= kind "MOD") (strcat " (" (sxl:r0 (sxl:fnum (nth 8 c))) "x" (sxl:r0 (sxl:fnum (nth 9 c))) ")") "")))
            (setq res (cons (list nm kind W H D
                                  (fix (sxl:fnum (nth 8 c))) (fix (sxl:fnum (nth 9 c)))
                                  (fix (sxl:nz (sxl:fnum (nth 10 c)) 1))
                                  (sxl:nz (sxl:fnum (nth 11 c)) (cdr (assoc kind '(("MOD" . 125.0) ("PLATE" . 150.0) ("CAB" . 205.0)))))
                                  (if (= kind "MOD") 0.0 (sxl:nz (sxl:fnum (nth 12 c)) (if (= kind "CAB") 60.0 40.0)))
                                  (if (= kind "MOD") 0.0 (sxl:nz (sxl:fnum (nth 13 c)) (if (= kind "CAB") 30.0 40.0)))
                                  (nth 0 c) (nth 1 c) (nth 4 c))
                            res)))))
      (close fh)
      (sx:log (strcat "encl csv " f " rows=" (itoa (length res))))))
  (if res
    ;; по площади, при равной — по глубине: подбор берёт самый маленький подходящий
    (setq sxl:encl (vl-sort (reverse res)
                            '(lambda (a b) (if (equal (* (nth 2 a) (nth 3 a)) (* (nth 2 b) (nth 3 b)) 1.0)
                                             (< (nth 4 a) (nth 4 b))
                                             (< (* (nth 2 a) (nth 3 a)) (* (nth 2 b) (nth 3 b)))))))
    (progn (setq sxl:encl sxl:encl0)
           (sxl:warn (R "не найден файл SX_Корпуса.csv рядом с программой: взят короткий запасной список корпусов")))))

;;; ограничения из окна: Ш/В/Г не более (0 = нет), установка (0 = любая), серия ("" = все)
(defun sxl:fits-limits (e)
  (and (or (<= sxl:maxW 0) (<= (nth 2 e) (+ sxl:maxW 0.01)))
       (or (<= sxl:maxH 0) (<= (nth 3 e) (+ sxl:maxH 0.01)))
       (or (<= sxl:maxD 0) (<= (nth 4 e) (+ sxl:maxD 0.01)))
       (or (= sxl:inst 0) (= (nth 13 e) (nth sxl:inst sxl:inst-list)))
       (or (= sxl:series "") (= (nth 11 e) sxl:series))))
(defun sxl:filtered () (vl-remove-if-not 'sxl:fits-limits sxl:encl))
(defun sxl:series-list ( / res)
  (foreach e sxl:encl (if (not (member (nth 11 e) res)) (setq res (cons (nth 11 e) res))))
  (cons (R "все") (reverse res)))
(defun sxl:limits-text ( / s)
  (setq s "")
  (if (> sxl:maxW 0) (setq s (strcat s (R " Ш<=") (sxl:r0 sxl:maxW))))
  (if (> sxl:maxH 0) (setq s (strcat s (R " В<=") (sxl:r0 sxl:maxH))))
  (if (> sxl:maxD 0) (setq s (strcat s (R " Г<=") (sxl:r0 sxl:maxD))))
  (if (> sxl:inst 0) (setq s (strcat s " " (nth sxl:inst sxl:inst-list))))
  (if (/= sxl:series "") (setq s (strcat s " " sxl:series)))
  s)

;;; ---------------------------------------------------------------- утилиты
(defun sxl:col (row c) (nth (1- c) row))
(defun sxl:cs (row c) (vl-string-trim " " (sx:s (sxl:col row c))))
(defun sxl:num (row c) (sx:n (sxl:col row c)))
(defun sxl:warn (s) (if (not (member s sxl:warns)) (setq sxl:warns (cons s sxl:warns))) (sx:log (strcat "WARN " s)) s)
(defun sxl:r0 (x) (rtos x 2 0))
(defun sxl:r1 (x) (rtos x 2 1))
;;; первая цифра в тексте как число (для «2P», «4P», «XT12» -> 12)
(defun sxl:int-in (s / k ch d res)
  (setq k 1 d "" res nil)
  (while (<= k (strlen s))
    (setq ch (substr s k 1))
    (cond ((wcmatch ch "#") (setq d (strcat d ch)))
          ((/= d "") (setq k (strlen s))))
    (setq k (1+ k)))
  (if (/= d "") (atoi d)))
(defun sxl:digits-tail (s / k)
  ;; номер в конце обозначения: «QFD12» -> ("QFD" 12); без номера -> (s nil)
  (setq k (strlen s))
  (while (and (> k 0) (wcmatch (substr s k 1) "#")) (setq k (1- k)))
  (if (< k (strlen s)) (list (substr s 1 k) (atoi (substr s (1+ k)))) (list s nil)))

;;; «QFD1, QFD2, QFD3, QF5» -> «QFD1-QFD3, QF5»
(defun sxl:condense (lst / res cur p n first last)
  (setq res '() cur nil)
  (foreach s lst
    (setq p (sxl:digits-tail s))
    (if (and cur (cadr p) (= (car p) (car cur)) (= (cadr p) (1+ (nth 2 cur))))
      (setq cur (list (car cur) (cadr cur) (cadr p)))
      (progn
        (if cur (setq res (cons cur res)))
        (setq cur (if (cadr p) (list (car p) (cadr p) (cadr p)) (list s nil nil))))))
  (if cur (setq res (cons cur res)))
  (sx:join-sep
    (mapcar '(lambda (c)
               (cond ((null (cadr c)) (car c))
                     ((= (cadr c) (nth 2 c)) (strcat (car c) (itoa (cadr c))))
                     ((= (1+ (cadr c)) (nth 2 c)) (strcat (car c) (itoa (cadr c)) ", " (car c) (itoa (nth 2 c))))
                     (t (strcat (car c) (itoa (cadr c)) "-" (car c) (itoa (nth 2 c))))))
            (reverse res))
    ", "))

;;; ---------------------------------------------------------------- Excel: «Каталог»
(defun sxl:read-range (sh addr / rg vals)
  (setq rg (vl-catch-all-apply 'vlax-get-property (list sh 'Range addr)))
  (if (not (vl-catch-all-error-p rg))
    (progn
      (setq vals (vl-catch-all-apply
                   '(lambda () (vlax-safearray->list (vlax-variant-value (vlax-get-property rg 'Value2))))
                   nil))
      (sx:rel rg)
      (if (not (vl-catch-all-error-p vals)) (mapcar '(lambda (r) (mapcar 'sx:v r)) vals)))))

(defun sxl:read-catalog ( / p sh)
  (setq sxl:cat nil sxl:manQF "" sxl:xt3 "" sxl:xtP "")
  (if (setq p (assoc (R "Каталог") sx:sheets))
    (progn
      (setq sh (cdr p))
      (setq sxl:cat (vl-remove-if '(lambda (r) (= (sxl:cs r sxl:kModel) "")) (sxl:read-range sh "A4:V500")))
      (setq sxl:manQF (sx:s (sx:cell sh "S4"))
            sxl:xt3   (sx:s (sx:cell sh "S7"))
            sxl:xtP   (sx:s (sx:cell sh "S8"))))
    (sxl:warn (R "в книге нет листа «Каталог»: ширины и ручные позиции не взяты")))
  (sx:log (strcat "catalog rows=" (itoa (length sxl:cat)) " manQF=" sxl:manQF " xt3=" sxl:xt3 " xtP=" sxl:xtP)))

;;; строка каталога для автомата по обозначению «2P С16А 30мА» (столбец U); свой производитель первым
(defun sxl:cat-qf (txt / hits own)
  (setq hits (vl-remove-if-not
               '(lambda (r) (and (member (cdr (assoc (sxl:cs r sxl:kCat) sxl:catmap)) '("QF" "QFD"))
                                 (= (sxl:cs r sxl:kOboz) txt)))
               sxl:cat))
  (setq own (vl-remove-if-not '(lambda (r) (= (sxl:cs r sxl:kMan) sxl:manQF)) hits))
  (cond (own (car own)) (hits (car hits))))

;;; строка каталога по модели (столбец C)
(defun sxl:cat-model (m / up)
  (setq up (strcase m))
  (vl-some '(lambda (r) (if (= (strcase (sxl:cs r sxl:kModel)) up) r)) sxl:cat))

;;; ---------------------------------------------------------------- блоки
;;; имена блоков: текущий чертёж + база (без копирования)
(defun sxl:collect-names (lib / res e src n)
  (setq res '() e (tblnext "BLOCK" T))
  (while e
    (if (not (wcmatch (cdr (assoc 2 e)) "`**")) (setq res (cons (cdr (assoc 2 e)) res)))
    (setq e (tblnext "BLOCK")))
  (if (and lib (/= lib "") (findfile lib) (setq src (sx:lib-open lib)))
    (vlax-for b (vla-get-blocks src)
      (setq n (vla-get-name b))
      (if (and (= (vla-get-islayout b) :vlax-false) (= (vla-get-isxref b) :vlax-false)
               (not (wcmatch n "`**")) (not (member n res)))
        (setq res (cons n res)))))
  res)

(defun sxl:avail (n / up)
  (if (and n (/= n ""))
    (progn (setq up (strcase n))
           (vl-some '(lambda (x) (if (= (strcase x) up) x)) sxl:names))))

;;; только буквы и цифры, верхний регистр (сравнение моделей с именами блоков)
(defun sxl:norm (s / k ch r)
  (setq s (strcase s) r "" k 1)
  (while (<= k (strlen s))
    (setq ch (substr s k 1))
    (if (wcmatch ch "@,#") (setq r (strcat r ch)))
    (setq k (1+ k)))
  r)

;;; блок модуля по модели: «ZIOMB24V2» -> «ZIO-MB24_ATT_DEF» (самое длинное совпадение начала)
(defun sxl:fuzzy (model / nm best bl stem sn)
  (setq nm (sxl:norm model) best nil bl 0)
  (if (>= (strlen nm) 4)
    (foreach n sxl:names
      (if (not (wcmatch (strcase n) "*_ENVOLVENTE,*_SERIGRAFIA,*_ENCLOSURE,*_SERIGRAPH,*_LOGO"))
        (progn
          (setq stem (if (wcmatch (strcase n) "*_ATT_DEF") (substr n 1 (- (strlen n) 8)) n)
                sn (sxl:norm stem))
          (if (and (>= (strlen sn) 4) (> (strlen sn) bl)
                   (or (= sn nm) (= (substr nm 1 (strlen sn)) sn) (= (substr sn 1 (min (strlen sn) (strlen nm))) nm)))
            (setq best n bl (strlen sn)))))))
  best)

;;; имя блока компоновки из «Каталога» (Q); число — не блок (в книге v5 в Q ток канала модулей: диммер, DALI)
(defun sxl:blk (crow / q)
  (setq q (if crow (sxl:cs crow sxl:kBlock) ""))
  (if (distof q) "" q))

;;; блок для аппарата: столбец Q каталога -> правило по типу -> модель -> модель_ATT_DEF -> похожее имя (модули)
(defun sxl:resolve (cat model crow code poles / q cands res)
  (setq q (sxl:blk crow))
  (setq cands
    (append
      (list q)
      (cond ((= cat "QF")  (list (strcat "S20" (itoa (max 1 (min 4 poles))))))
            ((= cat "QFD") (list (if (> poles 2) "DS203" "DS201")))
            ((and (= cat "DEV") (= code "MK")) (list "MR-5-1"))
            (t nil))
      (list model (strcat model "_ATT_DEF"))))
  (setq res (vl-some 'sxl:avail cands))
  (if (and (/= q "") (not (sxl:avail q)))
    (sxl:warn (strcat (R "блок «") q (R "» из «Каталога» (Q) не найден ни в чертеже, ни в базе"))))
  (if (and (null res) (member cat '("MOD" "OTH" "DEV"))) (setq res (sxl:fuzzy model)))
  res)

;;; ---------------------------------------------------------------- порядок прорисовки внутри блоков
;;; В базе у S201, DS201, ZIO-... белая заливка/маскировка стоит в блоке последней и убрана назад таблицей
;;; ACAD_SORTENTS. CopyObjects копирует таблицу, но её ссылки остаются на объекты базы — заливка оказывается
;;; поверх линий (белый прямоугольник). Поэтому полный порядок берём из блока в базе (по номеру объекта в блоке)
;;; и задаём блоку в чертеже заново. Блоки без такой таблицы в базе не трогаем (их порядок копируется верно).
(defun sxl:ents (blk / res) (vlax-for e blk (setq res (cons e res))) (reverse res))

(defun sxl:set-order (tb ord / xd st sa)
  (setq xd (vla-getextensiondictionary tb)
        st (vl-catch-all-apply 'vla-item (list xd "ACAD_SORTENTS")))
  (if (vl-catch-all-error-p st) (setq st (vl-catch-all-apply 'vla-addobject (list xd "ACAD_SORTENTS" "AcDbSortentsTable"))))
  (if (and (not (vl-catch-all-error-p st)) (> (length ord) 1))
    (progn
      (setq sa (vlax-make-safearray vlax-vbobject (cons 0 (1- (length ord)))))
      (vlax-safearray-fill sa ord)
      (not (vl-catch-all-error-p (vl-catch-all-apply 'vla-setrelativedraworder (list st sa)))))))

;;; порядок из базы: список номеров объектов блока (0..n-1) в порядке прорисовки, или nil
(defun sxl:src-order (sb / st objs r sh full)
  (if (= (vla-get-hasextensiondictionary sb) :vlax-true)
    (progn
      (setq st (vl-catch-all-apply 'vla-item (list (vla-getextensiondictionary sb) "ACAD_SORTENTS")))
      (if (not (vl-catch-all-error-p st))
        (progn
          (setq r (vl-catch-all-apply 'vla-getfulldraworder (list st 'objs :vlax-false)))
          (if (and (not (vl-catch-all-error-p r)) objs)
            (progn
              (if (= (type objs) 'VARIANT) (setq objs (vlax-variant-value objs)))
              (setq sh (mapcar 'vla-get-handle (sxl:ents sb))
                    full (mapcar '(lambda (o) (vl-position (vla-get-handle o) sh)) (vlax-safearray->list objs)))
              (if (and (= (length full) (length sh)) (not (member nil full))) full))))))))

(defun sxl:fix-draworder (names lib / src sblks tblks sb tb idx tents k)
  (setq k 0 tblks (vla-get-blocks sx:doc))
  (if (and lib (/= lib "") (findfile lib)) (setq src (sx:lib-open lib)))
  (if src (setq sblks (vla-get-blocks src)))
  (foreach n names
    (setq tb (vl-catch-all-apply 'vla-item (list tblks n)))
    (if (not (vl-catch-all-error-p tb))
      (progn
        (setq tents (sxl:ents tb) idx nil)
        (if sblks
          (progn
            (setq sb (vl-catch-all-apply 'vla-item (list sblks n)))
            (if (and (not (vl-catch-all-error-p sb)) (= (vla-get-count sb) (length tents)))
              (setq idx (sxl:src-order sb)))))
        (if (and idx (sxl:set-order tb (mapcar '(lambda (i) (nth i tents)) idx))) (setq k (1+ k))))))
  (sx:lib-close)
  (sx:log (strcat "draworder fixed in blocks: " (itoa k)))
  k)

;;; все блоки, вложенные в names (ZIO-MB24_ATT_DEF -> ..._ENVOLVENTE, ..._SERIGRAFIA)
(defun sxl:with-nested (names / res todo n tb)
  (setq todo names res '())
  (while todo
    (setq n (car todo) todo (cdr todo))
    (if (and n (not (member n res)) (not (wcmatch n "`**")))
      (progn
        (setq res (cons n res))
        (setq tb (vl-catch-all-apply 'vla-item (list (vla-get-blocks sx:doc) n)))
        (if (not (vl-catch-all-error-p tb))
          (vlax-for e tb
            (if (= (vla-get-objectname e) "AcDbBlockReference")
              (setq todo (cons (vla-get-effectivename e) todo))))))))
  res)

;;; габариты блока (вставка в 0,0 и удаление), кэш
(defun sxl:measure (name / hit ob bb)
  (if (setq hit (assoc name sxl:bbs))
    (cdr hit)
    (progn
      (setq ob (vl-catch-all-apply 'vla-insertblock
                 (list sx:ms (vlax-3d-point '(0.0 0.0 0.0)) name 1.0 1.0 1.0 0.0)))
      (if (not (vl-catch-all-error-p ob))
        (progn
          (setq bb (vl-catch-all-apply 'sx:bbox (list ob)))
          (vla-delete ob)
          (if (vl-catch-all-error-p bb) (setq bb nil))))
      (setq sxl:bbs (cons (cons name bb) sxl:bbs))
      bb)))

;;; ---------------------------------------------------------------- состав щита
;;; деталь: (модель наименование блок ширина_по_умолч высота_по_умолч ширина высота)
;;; позиция: (тип обозначение (детали) вид)  вид: "DIN" на рейке / "PNL" на монтажной панели
(defun sxl:part (cat model name crow code poles wmods)
  (list model name (sxl:resolve cat model crow code poles)
        (* wmods *sxl-mod*) (cdr (assoc cat sxl:hdef)) nil nil))
(defun sxl:xt-part (model / crow w)
  (setq crow (sxl:cat-model model)
        w (cond ((wcmatch model "*2.5*,*2,5*") 5.2) ((wcmatch model "*-4-*,*-4") 6.2)
                ((wcmatch model "*-6-*,*-6") 8.2) ((wcmatch model "*-10-*,*-10") 10.2)
                ((wcmatch model "*-16-*,*-16") 12.2) (t 6.2)))
  (list model (if crow (sxl:cs crow sxl:kName) model) (sxl:resolve "XT" model crow nil 0) w 100.0 nil nil))

(defun sxl:cat-width (crow dflt / w)
  (setq w (if crow (sxl:num crow sxl:kWidth) 0))
  (if (> w 0) w dflt))

(defun sxl:build (prow / its txt des crow poles cat rcd k code n dname mods seen m3 mp xt n3 np parts
                      part qty wm sortq cnt)
  (setq its '() cnt '())
  ;; автоматы: строка с текстом автомата (H) — первая строка группы
  (foreach row prow
    (if (/= (setq txt (sxl:cs row sxl:cQFtxt)) "")
      (progn
        (setq des (sxl:cs row sxl:cQF) crow (sxl:cat-qf txt)
              poles (cond ((sxl:int-in (sxl:cs row sxl:cPoles))) ((sxl:int-in txt)) (t 1))
              rcd (cond (crow (= (cdr (assoc (sxl:cs crow sxl:kCat) sxl:catmap)) "QFD"))
                        ((= (sxl:num row sxl:cRcd) 1))
                        (t (wcmatch txt (strcat "*" (R "мА") "*"))))
              cat (if rcd "QFD" "QF"))
        (if (null crow) (sxl:warn (strcat (R "автомат «") txt (R "» не найден в «Каталоге» (столбец U): ширина по числу полюсов"))))
        (setq its (cons (list cat des (list (sxl:part cat (if crow (sxl:cs crow sxl:kModel) txt)
                                                      (if crow (sxl:cs crow sxl:kName) txt)
                                                      crow nil poles (sxl:cat-width crow poles)))
                              "DIN" poles) its)))))
  ;; аппараты после автомата (I-K, коды BH-BJ)
  (foreach row prow
    (setq k 0)
    (repeat 3
      (setq txt (sxl:cs row (+ sxl:cDev1 k)) code (sxl:cs row (+ sxl:cDevC1 k)))
      (if (/= txt "")
        (progn
          (setq crow (sxl:cat-model txt)
                dname (cdr (assoc code sxl:devdes)) des "")
          (if dname
            (progn (setq n (1+ (cond ((cdr (assoc dname cnt))) (t 0))))
                   (setq cnt (cons (cons dname n) (vl-remove-if '(lambda (x) (= (car x) dname)) cnt))
                         des (strcat dname (itoa n)))))
          (setq mods (sxl:cat-width crow (cond ((= code "MK") 1) ((= code "KM") 2) (t 2))))
          (setq its (cons (list "DEV" des (list (sxl:part "DEV" (if crow (sxl:cs crow sxl:kModel) txt)
                                                          (if crow (sxl:cs crow sxl:kName) txt)
                                                          crow code 0 mods))
                                "DIN" 0) its))))
      (setq k (1+ k))))
  ;; модули (по обозначению AH, модель AI)
  (setq seen '())
  (foreach row prow
    (setq des (sxl:cs row sxl:cMod) txt (sxl:cs row sxl:cModel))
    (if (and (/= des "") (/= txt "") (not (member des seen)))
      (progn
        (setq seen (cons des seen) crow (sxl:cat-model txt))
        (setq its (cons (list "MOD" des (list (sxl:part "MOD" txt (if crow (sxl:cs crow sxl:kName) txt)
                                                        crow nil 0 (sxl:cat-width crow 4)))
                              "DIN" 0) its)))))
  ;; клеммы XT: на строку — трёхуровневых BW, доп. проходных BX (как в спецификации)
  (setq m3 (if (/= sxl:xt3 "") sxl:xt3 "VPR-2.5-3L-GY") mp (if (/= sxl:xtP "") sxl:xtP "VPR-2.5-QUATTRO-GY"))
  (setq xt '())
  (foreach row prow
    (if (/= (setq des (sxl:cs row sxl:cXT)) "")
      (progn
        (setq n3 (fix (sxl:num row sxl:c3L)) np (fix (sxl:num row sxl:cPass)) parts '())
        (if (= (+ n3 np) 0) (setq n3 1))
        (repeat n3 (setq parts (cons (sxl:xt-part m3) parts)))
        (repeat np (setq parts (cons (sxl:xt-part mp) parts)))
        (setq xt (cons (list "XT" des (reverse parts) "DIN" 0) xt)))))
  ;; по номеру клеммы (XT1, XT2 ...), при равных — порядок строк
  (setq xt (reverse xt))
  (setq xt (mapcar 'cdr (vl-sort (mapcar '(lambda (x) (cons (cond ((sxl:int-in (cadr x))) (t 0)) x)) xt)
                                 '(lambda (a b) (< (car a) (car b))))))
  (setq its (append (reverse its) xt))
  ;; ручные позиции «Каталога» (столбец O «Вручную / запас»)
  (foreach crow sxl:cat
    (setq qty (fix (sxl:num crow sxl:kManual))
          txt (sxl:cs crow sxl:kModel)
          cat (cdr (assoc (sxl:cs crow sxl:kCat) sxl:catmap)))
    (if (and (> qty 0) (or cat (/= (sxl:blk crow) "")))
      (cond
        ((wcmatch (strcase txt) *sxl-skip*) (sx:log (strcat "skip manual " txt)))
        ((= cat "XT")
         (setq parts '())
         (repeat qty (setq parts (cons (sxl:xt-part txt) parts)))
         (setq its (append its (list (list "XT" "" parts "DIN" 0)))))
        (t
         (setq poles (cond ((sxl:int-in (sxl:cs crow sxl:kPoles))) (t 1))
               wm (sxl:cat-width crow (if (member cat '("QF" "QFD")) poles 4))
               part (sxl:part (if cat cat "OTH") txt (sxl:cs crow sxl:kName) crow nil poles wm))
         (if (and (null (nth 2 part)) (= (sxl:num crow sxl:kWidth) 0) (not (member cat '("QF" "QFD"))))
           (sxl:warn (strcat (R "нет блока и нет ширины в «Каталоге» (L) у «") txt (R "»: принято ") (itoa wm) (R " мод."))))
         (repeat qty
           (setq its (append its (list (list (if cat cat "OTH") "" (list part) (if cat "DIN" "PNL") poles)))))))))
  ;; по типам; автоматы — многополюсные первыми, внутри — порядок однолинейки
  (setq sortq '())
  (foreach c sxl:order
    (if (= c "QF")
      (foreach p '(4 3 2 1)
        (foreach it its
          (if (and (= (car it) "QF") (= (min 4 (max 1 (nth 4 it))) p)) (setq sortq (cons it sortq)))))
      (foreach it its (if (= (car it) c) (setq sortq (cons it sortq))))))
  (reverse sortq))

;;; ширина/высота деталей: блок есть — по габаритам блока, иначе по каталогу/умолчанию
(defun sxl:finalize (its / bb)
  (mapcar
    '(lambda (it)
       (list (nth 0 it) (nth 1 it)
             (mapcar '(lambda (p)
                        (setq bb (if (nth 2 p) (sxl:measure (nth 2 p))))
                        (if bb
                          (list (nth 0 p) (nth 1 p) (nth 2 p) (nth 3 p) (nth 4 p)
                                (- (car (cadr bb)) (car (car bb))) (- (cadr (cadr bb)) (cadr (car bb))))
                          (list (nth 0 p) (nth 1 p) nil (nth 3 p) (nth 4 p) (nth 3 p) (nth 4 p))))
                     (nth 2 it))
             (nth 3 it) (nth 4 it)))
    its))

(defun sxl:iw (it) (apply '+ (mapcar '(lambda (p) (nth 5 p)) (nth 2 it))))
(defun sxl:ih (it) (apply 'max (mapcar '(lambda (p) (nth 6 p)) (nth 2 it))))

;;; позиция шире рейки (длинный ряд клемм) -> куски не длиннее cap
(defun sxl:split-wide (its cap / res cur w pw)
  (foreach it its
    (if (<= (sxl:iw it) cap)
      (setq res (cons it res))
      (progn
        (setq cur '() w 0.0)
        (foreach p (nth 2 it)
          (setq pw (nth 5 p))
          (if (and cur (> (+ w pw) cap))
            (setq res (cons (list (nth 0 it) (nth 1 it) (reverse cur) (nth 3 it) (nth 4 it)) res) cur '() w 0.0))
          (setq cur (cons p cur) w (+ w pw)))
        (if cur (setq res (cons (list (nth 0 it) (nth 1 it) (reverse cur) (nth 3 it) (nth 4 it)) res))))))
  (reverse res))

;;; ---------------------------------------------------------------- корпус
;;; геометрия от левого верхнего угла (y вниз отрицательный):
;;; W H n rows secx(список) raillen railoff startoff usable railY(список) ducts(список прямоугольников) rowtop(функция)
;;; enc = (имя тип Ш В Г рядов модулей секций шаг короб поле серия артикул установка)
(defun sxl:geom (enc / kind W H n p d m rows usable secx railY ducts sw k i top)
  (setq kind (nth 1 enc) W (float (nth 2 enc)) H (float (nth 3 enc)))
  (if (= kind "MOD")
    ;; модульный щиток: рядов x модулей, рейки по центру по ширине, ряды поровну по высоте
    (progn
      (setq rows (max 1 (nth 5 enc)) usable (* (nth 6 enc) *sxl-mod*) p (sxl:nz (nth 8 enc) *sxl-pBox*))
      (if (> (* rows p) (- H 40.0)) (setq p (/ (- H 40.0) rows)))
      (setq top (/ (- H (* rows p)) 2.0) m (min 30.0 top) railY '() k 0)
      (repeat rows (setq railY (cons (- (+ top (* p (+ k 0.5)))) railY) k (1+ k)))
      (list (cons 'W W) (cons 'H H) (cons 'n 1) (cons 'rows rows) (cons 'secx (list (/ (- W usable) 2.0)))
            (cons 'railoff -8.0) (cons 'raillen (+ usable 16.0)) (cons 'startoff 0.0) (cons 'usable usable)
            (cons 'railY (reverse railY)) (cons 'ducts nil) (cons 'p p) (cons 'top top) (cons 'm m) (cons 'd 0.0)))
    ;; корпус с монтажной платой / шкаф: поле от края, кабель-каналы по краям секций и между рейками
    (progn
      (setq n (max 1 (nth 7 enc)) p (nth 8 enc) d (nth 9 enc) m (nth 10 enc)
            sw (/ (- W (* 2 m) (* (1+ n) d)) n)
            rows (max 0 (fix (/ (- H (* 2 m) d) p)))
            secx '() railY '() ducts '() i 0 k 0)
      (repeat n (setq secx (cons (+ m d (* i (+ sw d))) secx) i (1+ i)))
      (setq secx (reverse secx))
      (repeat rows (setq railY (cons (- (+ m (* k p) d (/ (- p d) 2.0))) railY) k (1+ k)))
      ;; кабель-каналы: вертикальные по краям секций, горизонтальные над каждой рейкой и под последней
      (setq i 0)
      (repeat (1+ n)
        (setq ducts (cons (list (+ m (* i (+ sw d))) (- m) (+ m (* i (+ sw d)) d) (- (+ m (* rows p) d))) ducts) i (1+ i)))
      (foreach x0 secx
        (setq k 0)
        (repeat (1+ rows)
          (setq ducts (cons (list x0 (- (+ m (* k p))) (+ x0 sw) (- (+ m (* k p) d))) ducts) k (1+ k))))
      (list (cons 'W W) (cons 'H H) (cons 'n n) (cons 'rows rows) (cons 'secx secx)
            (cons 'railoff 5.0) (cons 'raillen (- sw 10.0)) (cons 'startoff 10.0) (cons 'usable (- sw 20.0))
            (cons 'railY (reverse railY)) (cons 'ducts (reverse ducts)) (cons 'p p) (cons 'top m) (cons 'm m) (cons 'd d)))))
(defun sxl:g (geo key) (cdr (assoc key geo)))

;;; раскладка по рейкам (рейки по рядам: ряд 1 всех секций, затем ряд 2 ...)
;;; -> (размещения ((позиция рейка x) ...) не_поместилось занято_реек)
(defun sxl:pack (its nrails cap newrail / r x prev res over w)
  (setq r 0 x 0.0 prev nil res '() over '())
  (foreach it its
    (setq w (sxl:iw it))
    (if (and prev (/= (car it) prev) (> x 0.0))
      (if newrail (setq r (1+ r) x 0.0) (setq x (+ x *sxl-gap*))))
    (if (and (> x 0.0) (> (+ x w) (+ cap 0.01))) (setq r (1+ r) x 0.0))
    (if (or (>= r nrails) (> w (+ cap 0.01)))
      (setq over (cons it over))
      (setq res (cons (list it r x) res) x (+ x w)))
    (setq prev (car it)))
  (list (reverse res) (reverse over) (if res (1+ (apply 'max (mapcar 'cadr res))) 0)))

;;; аппараты не на рейке (монтажная панель): под последней занятой рейкой, слева направо
(defun sxl:pack-panel (pits geo used / n row y0 ybot x0 x1 x y hrow res ok)
  (setq ok T res '())
  (if pits
    (progn
      (setq n (sxl:g geo 'n) row (if (> used 0) (1+ (/ (1- used) n)) 0)
            y0 (- (+ (sxl:g geo 'top) (* row (sxl:g geo 'p)) (sxl:g geo 'd) 10.0))
            ybot (- (- (sxl:g geo 'H) (sxl:g geo 'm) 10.0))
            x0 (+ (sxl:g geo 'm) (sxl:g geo 'd) 10.0) x1 (- (sxl:g geo 'W) (sxl:g geo 'm) (sxl:g geo 'd) 10.0)
            x x0 y y0 hrow 0.0)
      (foreach it pits
        (if (> (+ x (sxl:iw it)) x1) (setq x x0 y (- y hrow 20.0) hrow 0.0))
        (if (< (- y (sxl:ih it)) ybot) (setq ok nil))
        (setq res (cons (list it x y) res) x (+ x (sxl:iw it) 20.0) hrow (max hrow (sxl:ih it))))
      (setq sxl:pnlrow row)))
  (list ok (reverse res)))

;;; попытка уложить в корпус: (geo размещения не_поместилось занято панель запас newrail) или nil
(defun sxl:try (enc dits pits res newrail force / geo cap nr pk pp its)
  (setq geo (sxl:geom enc)
        cap (* (sxl:g geo 'usable) (- 1.0 (/ res 100.0)))
        nr (* (sxl:g geo 'n) (sxl:g geo 'rows)))
  (if (> cap 0)
    (progn
      (setq its (sxl:split-wide dits cap) pk (sxl:pack its nr cap newrail) pp (sxl:pack-panel pits geo (nth 2 pk)))
      (if (or force (and (null (cadr pk)) (car pp)))
        (list geo (car pk) (cadr pk) (nth 2 pk) (cadr pp) res newrail (car pp))))))

(defun sxl:choose (dits pits / cands r opts hit)
  ;; варианты по убыванию «хорошести»: с запасом и типами с новой рейки -> с запасом подряд -> без запаса
  (setq opts (append (if *sxl-newrail* (list (list *sxl-reserve* T))) (list (list *sxl-reserve* nil))
                     (if *sxl-newrail* (list (list 0 T))) (list (list 0 nil))))
  ;; выбранный в окне корпус — только он; иначе все корпуса базы, прошедшие ограничения (по возрастанию)
  (setq cands (if (/= sxl:encname "")
                (vl-remove-if-not '(lambda (e) (= (car e) sxl:encname)) sxl:encl)
                (sxl:filtered)))
  (setq sxl:encsel nil)
  (if cands
    (progn
      ;; сначала с запасом: самый маленький корпус, в нём по возможности типы с новой рейки, иначе подряд;
      ;; не вышло ни в одном — то же без запаса
      (setq hit (vl-some '(lambda (res)
                            (vl-some '(lambda (e)
                                        (vl-some '(lambda (o) (if (and (= (car o) res) (setq r (sxl:try e dits pits (car o) (cadr o) nil)))
                                                                (cons e r)))
                                                 opts))
                                     cands))
                         (list *sxl-reserve* 0)))
      (if (null hit) (setq hit (cons (last cands) (sxl:try (last cands) dits pits 0 nil T))))
      (if (cdr hit) (setq sxl:encsel (car hit)))
      (cdr hit))))

;;; ---------------------------------------------------------------- рисование
(defun sxl:P (x y) (list (+ (car sxl:inspt) x) (+ (cadr sxl:inspt) y) 0.0))
(defun sxl:line (x1 y1 x2 y2 lay)
  (entmake (list '(0 . "LINE") (cons 8 lay) (cons 10 (sxl:P x1 y1)) (cons 11 (sxl:P x2 y2)))))
(defun sxl:rect (x1 y1 x2 y2 lay / col)
  (entmake (list '(0 . "LWPOLYLINE") '(100 . "AcDbEntity") (cons 8 lay) '(100 . "AcDbPolyline")
                 '(90 . 4) '(70 . 1)
                 (cons 10 (sxl:P x1 y1)) (cons 10 (sxl:P x2 y1)) (cons 10 (sxl:P x2 y2)) (cons 10 (sxl:P x1 y2)))))
;;; just: "L" влево, "C" по центру, "ML" середина-влево, "MC" середина-центр
(defun sxl:text (x y h s lay rot just / p)
  (setq s (sx:s s) p (sxl:P x y))
  (if (/= s "")
    (entmake (append (list '(0 . "TEXT") (cons 8 lay) (cons 10 p) (cons 40 h) (cons 1 s)
                           (cons 50 (* pi (/ rot 180.0))) (cons 7 sx:style))
                     (cond ((= just "C")  (list '(72 . 1) (cons 11 p)))
                           ((= just "ML") (list '(72 . 0) '(73 . 2) (cons 11 p)))
                           ((= just "MC") (list '(72 . 1) '(73 . 2) (cons 11 p)))
                           (t nil))))))

(defun sxl:layers ()
  (setq sxl:LF (sx:layer "SXL_FRAME" 7 nil) sxl:LR (sx:layer "SXL_RAIL" 8 nil)
        sxl:LD (sx:layer "SXL_DUCT" 8 nil)  sxl:LV (sx:layer "SXL_DEV" 7 nil)
        sxl:LT (sx:layer "SXL_TEXT" 7 nil)  sxl:LM (sx:layer "SXL_DIM" 7 nil)
        sxl:LB (sx:layer "SXL_TABLE" 7 nil) sxl:LO (sx:layer "SXL_OVER" 1 nil)))

(defun sxl:clear ( / ss k)
  (if (setq ss (ssget "_X" '((8 . "SXL_*"))))
    (progn (setq k (sslength ss)) (repeat k (entdel (ssname ss (setq k (1- k))))))))
(defun c:SXLCLEAR () (sxl:clear) (princ))

;;; рейка: ваш динамический блок «Дин.рейка» (длина — Расстояние1) или прямоугольник
(defun sxl:rail (x y len / ob bb)
  (if (and (sxl:avail sxl:bRail) (tblsearch "BLOCK" sxl:bRail) (setq bb (sxl:measure sxl:bRail)))
    (progn
      (setq ob (vla-insertblock sx:ms (vlax-3d-point (sxl:P (- x (car (car bb))) (- y (/ (+ (cadr (car bb)) (cadr (cadr bb))) 2.0))))
                                sxl:bRail 1.0 1.0 1.0 0.0))
      (vla-put-layer ob sxl:LR)
      (if (not (sx:dynset ob sxl:pLen len))
        (sx:fit ob sxl:pLen 0 (car (sxl:P (+ x len) y)))))
    (sxl:rect x (+ y (/ *sxl-railH* 2.0)) (+ x len) (- y (/ *sxl-railH* 2.0)) sxl:LR)))

;;; деталь: блок по габаритам (левый край в x, середина на y) или прямоугольник с подписью
(defun sxl:draw-part (p x y lay / bb ob w h)
  (setq w (nth 5 p) h (nth 6 p))
  (if (and (nth 2 p) (setq bb (sxl:measure (nth 2 p))))
    (progn
      (setq ob (vla-insertblock sx:ms (vlax-3d-point (sxl:P (- x (car (car bb))) (- y (/ (+ (cadr (car bb)) (cadr (cadr bb))) 2.0))))
                                (nth 2 p) 1.0 1.0 1.0 0.0))
      (vla-put-layer ob lay))
    (progn
      (sxl:rect x (+ y (/ h 2.0)) (+ x w) (- y (/ h 2.0)) lay)
      (if (>= w 9.0)
        (if (< w 40.0)
          (sxl:text (+ x (/ w 2.0)) y (min 4.0 (* 0.4 w)) (car p) lay 90.0 "MC")
          (sxl:text (+ x (/ w 2.0)) y (min 6.0 (* 0.12 w)) (car p) lay 0.0 "MC"))))))

;;; позиция целиком + обозначение над ней
(defun sxl:draw-item (it x y lay / xx w h th s)
  (setq xx x)
  (foreach p (nth 2 it)
    (sxl:draw-part p xx y lay)
    (setq xx (+ xx (nth 5 p))))
  (setq w (sxl:iw it) h (sxl:ih it) s (nth 1 it))
  ;; у клемм — только номер (XT12 -> 12): иначе подписи соседних клемм сливаются
  (if (and (= (car it) "XT") (cadr (sxl:digits-tail s))) (setq s (itoa (cadr (sxl:digits-tail s)))))
  (if (/= s "")
    (if (< w 60.0)
      (progn (setq th (min *sxl-th* (* 0.8 w)))
             (sxl:text (+ x (/ w 2.0)) (+ y (/ h 2.0) 3.0) th s sxl:LT 90.0 "ML"))
      (sxl:text (+ x (/ w 2.0)) (+ y (/ h 2.0) 3.0) *sxl-th* s sxl:LT 0.0 "C"))))

;;; таблица позиций справа от корпуса: обозначения | наименование | кол-во
(defun sxl:table (x0 y0 rows / th rh w1 w2 w3 y k)
  (setq th (* 0.6 *sxl-thb*) rh (* 1.6 th) w1 0.0 w2 0.0 w3 (* 5.0 th))
  (foreach r rows
    (setq w1 (max w1 (* 0.9 th (strlen (car r)))) w2 (max w2 (* 0.85 th (strlen (cadr r))))))
  (setq w1 (+ (max w1 (* 8.0 th)) th) w2 (+ (max w2 (* 20.0 th)) th) y y0)
  (sxl:text x0 (+ y0 (* 0.8 rh)) *sxl-thb* (R "Перечень аппаратов компоновки") sxl:LB 0.0 "L")
  (setq rows (cons (list (R "Обозначение") (R "Наименование") (R "Кол.")) rows) k 0)
  (foreach r rows
    (sxl:line x0 y (+ x0 w1 w2 w3) y sxl:LB)
    (sxl:text (+ x0 (* 0.5 th)) (- y (/ rh 2.0)) th (car r) sxl:LB 0.0 "ML")
    (sxl:text (+ x0 w1 (* 0.5 th)) (- y (/ rh 2.0)) th (cadr r) sxl:LB 0.0 "ML")
    (sxl:text (+ x0 w1 w2 (/ w3 2.0)) (- y (/ rh 2.0)) th (caddr r) sxl:LB 0.0 "MC")
    (setq y (- y rh) k (1+ k)))
  (sxl:line x0 y (+ x0 w1 w2 w3) y sxl:LB)
  (foreach xx (list x0 (+ x0 w1) (+ x0 w1 w2) (+ x0 w1 w2 w3)) (sxl:line xx y0 xx y sxl:LB)))

;;; строки таблицы: по наименованию детали; обозначения позиций сжаты (QFD1-QFD28)
(defun sxl:table-rows (placed / acc key hit des)
  (setq acc '())
  (foreach it placed
    (foreach p (nth 2 it)
      (setq key (if (/= (nth 1 p) "") (nth 1 p) (car p)) hit (assoc key acc))
      (if hit
        (setq acc (subst (list key (if (and (/= (nth 1 it) "") (not (member (nth 1 it) (cadr hit))))
                                     (append (cadr hit) (list (nth 1 it))) (cadr hit))
                               (1+ (caddr hit)))
                         hit acc))
        (setq acc (append acc (list (list key (if (/= (nth 1 it) "") (list (nth 1 it))) 1)))))))
  (mapcar '(lambda (a) (list (sxl:condense (cadr a)) (car a) (itoa (caddr a)))) acc))

(defun sxl:dim (x1 y1 x2 y2 xl yl rot / d)
  (setq d (vl-catch-all-apply 'vla-adddimrotated
            (list sx:ms (vlax-3d-point (sxl:P x1 y1)) (vlax-3d-point (sxl:P x2 y2)) (vlax-3d-point (sxl:P xl yl)) rot)))
  (if (not (vl-catch-all-error-p d))
    (progn (vla-put-layer d sxl:LM) (vl-catch-all-apply 'vla-put-scalefactor (list d *sxl-dimsc*)))))

(defun sxl:draw (panel r / geo pl over used ppl W H n rows secx railY x0 y it row sec placed fill tot pnlrow ox)
  (setq geo (nth 0 r) pl (nth 1 r) over (nth 2 r) used (nth 3 r) ppl (nth 4 r)
        W (sxl:g geo 'W) H (sxl:g geo 'H) n (sxl:g geo 'n) rows (sxl:g geo 'rows)
        secx (sxl:g geo 'secx) railY (sxl:g geo 'railY)
        pnlrow (if ppl sxl:pnlrow rows))
  ;; корпус
  (sxl:rect 0.0 0.0 W (- H) sxl:LF)
  (sxl:rect (sxl:g geo 'm) (- (sxl:g geo 'm)) (- W (sxl:g geo 'm)) (- (- H (sxl:g geo 'm))) sxl:LF)
  ;; кабель-каналы: вертикальные все, горизонтальные — только над зоной монтажной панели
  (foreach dc (sxl:g geo 'ducts)
    (if (or (<= (- (nth 2 dc) (car dc)) (+ (sxl:g geo 'd) 0.01))
            (>= (cadr dc) (- (+ (sxl:g geo 'top) (* pnlrow (sxl:g geo 'p)) 0.01))))
      (sxl:rect (car dc) (cadr dc) (nth 2 dc) (nth 3 dc) sxl:LD)))
  ;; рейки (кроме зоны монтажной панели)
  (setq row 0)
  (foreach y railY
    (if (< row pnlrow)
      (foreach x0 secx (sxl:rail (+ x0 (sxl:g geo 'railoff)) y (sxl:g geo 'raillen))))
    (setq row (1+ row)))
  ;; аппараты на рейках
  (setq placed '() tot 0.0)
  (foreach e pl
    (setq it (car e) row (/ (cadr e) n) sec (rem (cadr e) n))
    (sxl:draw-item it (+ (nth sec secx) (sxl:g geo 'startoff) (caddr e)) (nth row railY) sxl:LV)
    (setq placed (cons it placed) tot (+ tot (sxl:iw it))))
  ;; монтажная панель
  (foreach e ppl
    (setq it (car e))
    (sxl:draw-item it (cadr e) (- (caddr e) (/ (sxl:ih it) 2.0)) sxl:LV)
    (setq placed (cons it placed)))
  ;; не поместилось: под корпусом, красным
  (if over
    (progn
      (setq ox 0.0 y (- (+ H 250.0)))
      (sxl:text 0.0 (+ y 140.0) *sxl-thb* (R "НЕ ПОМЕСТИЛОСЬ В КОРПУС:") sxl:LO 0.0 "L")
      (foreach it over
        (sxl:draw-item it ox y sxl:LO)
        (setq ox (+ ox (sxl:iw it) 5.0) placed (cons it placed)))))
  ;; заголовок, заполнение, размеры, таблица
  (setq fill (if (> used 0) (* 100.0 (/ tot (* (sxl:g geo 'usable) n rows))) 0.0))
  (sxl:text 0.0 (+ (* 2.2 *sxl-thb*) 10.0) *sxl-thb*
            (strcat panel "  " (car sxl:encsel) (R " (ВxШxГ, мм)")) sxl:LT 0.0 "L")
  (sxl:text 0.0 10.0 (* 0.6 *sxl-thb*)
            (strcat (R "Реек занято: ") (itoa used) (R " из ") (itoa (* n rows))
                    (R ";  заполнение реек: ") (sxl:r0 fill) "%" (R ";  запас при подборе: ") (itoa (nth 5 r)) "%"
                    (if (/= (sxl:limits-text) "") (strcat (R ";  ограничения:") (sxl:limits-text)) ""))
            sxl:LT 0.0 "L")
  (sxl:dim 0.0 (- H) W (- H) 0.0 (- (+ H 80.0)) 0.0)
  (sxl:dim 0.0 0.0 0.0 (- H) -80.0 0.0 (/ pi 2.0))
  (sxl:table (+ W 150.0) 0.0
             (cons (list "" (strcat (R "Корпус ") (car sxl:encsel)) "1") (sxl:table-rows (reverse placed))))
  (list used fill (length over)))

;;; ---------------------------------------------------------------- окно
;;; окно: пути поддержки -> прошлый выбор -> папка «Компоновка» рядом с «Однолинейка» / «Блоки» -> спросить
(defun sxl:dcl-file ( / f sib)
  (setq sib (strcat "\\" (R "Компоновка") "\\SX_Layout.dcl"))
  (cond ((setq f (findfile "SX_Layout.dcl")) f)
        ((and (getenv "SXL_DclFile") (findfile (getenv "SXL_DclFile"))) (getenv "SXL_DclFile"))
        ((setq f (vl-some '(lambda (p)
                             (if (and p (findfile p))
                               (findfile (strcat (vl-filename-directory (vl-filename-directory (findfile p))) sib))))
                          (list (getenv "SX_SchemaLsp") (findfile "SX_Schema.lsp") (getenv "SXL_LibDwg") sxl:lib)))
         (setenv "SXL_DclFile" f) f)
        ((setq f (getfiled (R "SXLAYOUT: укажите файл окна SX_Layout.dcl (папка 2_AutoCAD/Компоновка)") "SX_Layout.dcl" "dcl" 0))
         (setenv "SXL_DclFile" f) f)))

(defun sxl:pt-info ()
  (if sxl:inspt
    (strcat (R "Левый верхний угол корпуса: ") (sxl:r1 (car sxl:inspt)) "; " (sxl:r1 (cadr sxl:inspt)))
    (R "Точка вставки не указана")))

;;; ограничения из полей окна -> переменные; пересобрать список корпусов (выбор сохраняется, если корпус остался)
(defun sxl:dlg-limits ( / lst fl k sl)
  (setq sxl:maxW (sxl:fnum (get_tile "maxw")) sxl:maxH (sxl:fnum (get_tile "maxh")) sxl:maxD (sxl:fnum (get_tile "maxd"))
        sxl:inst (atoi (get_tile "inst"))
        sl (sxl:series-list) k (atoi (get_tile "series"))
        sxl:series (if (> k 0) (nth k sl) ""))
  (setq fl (sxl:filtered)
        lst (cons (R "Подобрать автоматически (самый маленький подходящий)") (mapcar 'car fl))
        k (if (/= sxl:encname "") (vl-position sxl:encname (mapcar 'car fl))))
  (if (null k) (setq sxl:encname ""))
  (sx:fill-list "encl" lst (if k (1+ k) 0))
  (set_tile "enclinfo" (strcat (R "Подходит корпусов: ") (itoa (length fl)) (R " из ") (itoa (length sxl:encl))
                               (if (null fl) (R " - ослабьте ограничения") ""))))

(defun sxl:dlg-encl (val / k)
  (setq k (atoi val) sxl:encname (if (> k 0) (car (nth (1- k) (sxl:filtered))) "")))

(defun sxl:dlg-save ( / v)
  (setq v (distof (get_tile "reserve") 2))
  (if (and v (>= v 0) (< v 90)) (setq *sxl-reserve* (fix v)))
  (setq *sxl-newrail* (= (get_tile "newrail") "1")
        sxl:do-clear (= (get_tile "clear") "1")
        sxl:lib (get_tile "lib"))
  (sxl:dlg-limits)
  (if (/= (get_tile "path") sx:path) (setq sx:path-typed (get_tile "path"))))

(defun sxl:dialog ( / dcl id res path p)
  (if (null (setq dcl (sxl:dcl-file))) (exit))
  (setq id (load_dialog dcl) res 5)
  (if (< id 0) (exit))
  (setq sxl:status "")
  (while (> res 1)
    (if (not (new_dialog "sxl_main" id)) (exit))
    (set_tile "path" sx:path)
    (sx:fill-list "panel" sx:panels sx:pi)
    (set_tile "maxw" (if (> sxl:maxW 0) (sxl:r0 sxl:maxW) ""))
    (set_tile "maxh" (if (> sxl:maxH 0) (sxl:r0 sxl:maxH) ""))
    (set_tile "maxd" (if (> sxl:maxD 0) (sxl:r0 sxl:maxD) ""))
    (sx:fill-list "inst" sxl:inst-list sxl:inst)
    (sx:fill-list "series" (sxl:series-list)
                  (cond ((and (/= sxl:series "") (vl-position sxl:series (sxl:series-list)))) (t 0)))
    (sxl:dlg-limits)
    (foreach k '("maxw" "maxh" "maxd" "inst" "series") (action_tile k "(sxl:dlg-limits)"))
    (action_tile "encl" "(sxl:dlg-encl $value)")
    (set_tile "reserve" (itoa *sxl-reserve*))
    (set_tile "newrail" (if *sxl-newrail* "1" "0"))
    (set_tile "ptinfo" (sxl:pt-info))
    (set_tile "clear" (if sxl:do-clear "1" "0"))
    (set_tile "lib" (if sxl:lib sxl:lib ""))
    (set_tile "status" sxl:status)
    (action_tile "panel" "(setq sx:pi (atoi $value))")
    (action_tile "browse" "(sxl:dlg-save) (done_dialog 3)")
    (action_tile "libbrowse" "(sxl:dlg-save) (done_dialog 4)")
    (action_tile "pick" "(sxl:dlg-save) (done_dialog 2)")
    (action_tile "accept" "(sxl:dlg-save) (done_dialog 1)")
    (action_tile "cancel" "(done_dialog 0)")
    (setq sx:path-typed nil sxl:status "")
    (setq res (start_dialog))
    (if (and sx:path-typed (/= sx:path-typed sx:path) (/= sx:path-typed ""))
      (progn
        (if (not (sx:open-wb sx:path-typed)) (setq sxl:status (R "Не удалось открыть файл Excel.")))
        (if (= res 1) (setq res 5))))
    (cond
      ((= res 3)
       (setq path (getfiled (R "Файл Excel с однолинейкой") (if (/= sx:path "") sx:path "") "xlsx;xlsm;xls" 0))
       (if path (if (sx:open-wb path) (setenv "SX_LastXls" path) (setq sxl:status (R "Не удалось открыть файл Excel.")))))
      ((= res 4)
       (if (setq path (getfiled (R "База блоков компоновки (DWG)") (if sxl:lib sxl:lib "") "dwg" 0))
         (setq sxl:lib path)))
      ((= res 2)
       (if (setq p (getpoint (strcat "\n" (R "Левый верхний угол корпуса: "))))
         (setq sxl:inspt (trans p 1 0))))
      ((= res 1)
       (cond
         ((null sx:wb) (setq sxl:status (R "Выберите файл Excel.") res 5))
         ((null sxl:inspt) (setq sxl:status (R "Укажите точку вставки (кнопка «Указать точку»).") res 5))
         ((null sx:panels) (setq sxl:status (R "В книге не найден лист «Однолинейка» со щитами.") res 5))))))
  (unload_dialog id)
  (= res 1))

;;; ---------------------------------------------------------------- параметры в чертеже
(defun sxl:save-params (panel)
  (vlax-ldata-put "SX_Layout" "xls" sx:path)
  (vlax-ldata-put "SX_Layout" "panel" panel)
  (vlax-ldata-put "SX_Layout" "lib" (if sxl:lib sxl:lib ""))
  (vlax-ldata-put "SX_Layout" "encl" sxl:encname)
  (vlax-ldata-put "SX_Layout" "limits" (list sxl:maxW sxl:maxH sxl:maxD sxl:inst sxl:series))
  (vlax-ldata-put "SX_Layout" "reserve" *sxl-reserve*)
  (vlax-ldata-put "SX_Layout" "newrail" (if *sxl-newrail* 1 0))
  (if sxl:inspt (vlax-ldata-put "SX_Layout" "pt" sxl:inspt)))

(defun sxl:load-params ( / e lm)
  (setq sxl:inspt (vlax-ldata-get "SX_Layout" "pt")
        sxl:lib (cond ((vlax-ldata-get "SX_Layout" "lib")) ((getenv "SXL_LibDwg"))))
  (if (vlax-ldata-get "SX_Layout" "reserve") (setq *sxl-reserve* (vlax-ldata-get "SX_Layout" "reserve")))
  (if (vlax-ldata-get "SX_Layout" "newrail") (setq *sxl-newrail* (= (vlax-ldata-get "SX_Layout" "newrail") 1)))
  (if (and (setq lm (vlax-ldata-get "SX_Layout" "limits")) (= (length lm) 5))
    (setq sxl:maxW (nth 0 lm) sxl:maxH (nth 1 lm) sxl:maxD (nth 2 lm) sxl:inst (nth 3 lm) sxl:series (nth 4 lm)))
  ;; корпус по имени; если его уже нет в базе — подбор
  (setq e (vlax-ldata-get "SX_Layout" "encl"))
  (setq sxl:encname (if (and (= (type e) 'STR) (member e (mapcar 'car sxl:encl))) e "")))

(defun sxl:open-saved ( / xls pan)
  (setq xls (vlax-ldata-get "SX_Layout" "xls") pan (vlax-ldata-get "SX_Layout" "panel"))
  (cond
    ((or (null xls) (null pan) (null sxl:inspt))
     (princ "\nSXLUPDATE: no saved settings in this drawing - run SXLAYOUT once.") nil)
    ((not (findfile xls)) (princ (strcat "\nSXLUPDATE: file not found: " xls)) nil)
    ((not (sx:open-wb xls)) (princ (strcat "\nSXLUPDATE: cannot open " xls)) nil)
    ((not (member pan sx:panels)) (princ (strcat "\nSXLUPDATE: panel not found: " pan)) nil)
    (T (setq sx:pi (vl-position pan sx:panels)) T)))

;;; ---------------------------------------------------------------- команда
(if (null sxl:do-clear) (setq sxl:do-clear T))

(defun sxl:logdump ( / fn f)
  (setq fn (strcat (sx:logdir) "\\SXL_log.txt"))
  (if (setq f (open fn "w"))
    (progn (foreach s (reverse sx:logl) (write-line s f)) (close f) (princ (strcat "\nLog: " fn)))))

;;; mode: "dlg" — окно, "upd" — по сохранённым параметрам, "batch" — параметры уже заданы (для проверки)
(defun sxl:main (mode / *error* oldcmd panel data prow its dits pits res st)
  ;; SX_Schema уже проверен в команде (sxl:run); обработчик всё равно не зовёт отсутствующие функции
  (defun *error* (msg)
    (if sx:doc (vl-catch-all-apply 'vla-endundomark (list sx:doc)))
    (if oldcmd (setvar "CMDECHO" oldcmd))
    (if sx:xl-release (vl-catch-all-apply 'sx:xl-release nil))
    (if sx:lib-close (vl-catch-all-apply 'sx:lib-close nil))
    (if (not (member msg '("Function cancelled" "quit / exit abort")))
      (progn (if sx:log (sx:log (strcat "FATAL at [" (sx:s sx:stage) "]: " msg)))
             (princ (strcat "\nSXLAYOUT error at [" (if sx:s (sx:s sx:stage) "") "]: " msg))
             (if sx:logdir (vl-catch-all-apply 'sxl:logdump nil))))
    (princ))
  (sxl:init)
  (setq sx:logl nil sx:errs nil sxl:warns nil sxl:bbs nil sx:blkcache nil)
  (sx:stg "init")
  (sx:log (strcat "SX_Layout " sxl:ver " / SX_Schema " sx:ver " mode=" mode))
  (sxl:load-encl)
  (setq sx:doc (vla-get-activedocument (vlax-get-acad-object))
        sx:ms (vla-get-modelspace sx:doc)
        oldcmd (getvar "CMDECHO"))
  (setvar "CMDECHO" 0)
  (cond
    ((= mode "upd")
     (sxl:load-params)
     (if (not (sxl:open-saved)) (progn (sx:xl-release) (setvar "CMDECHO" oldcmd) (exit))))
    ((= mode "dlg")
     (sxl:load-params)
     (sx:stg "open last xlsx")
     (cond ((and (vlax-ldata-get "SX_Layout" "xls") (findfile (vlax-ldata-get "SX_Layout" "xls")))
            (sx:open-wb (vlax-ldata-get "SX_Layout" "xls")))
           ((and (getenv "SX_LastXls") (findfile (getenv "SX_LastXls")))
            (sx:open-wb (getenv "SX_LastXls"))))
     (if (and (vlax-ldata-get "SX_Layout" "panel") (member (vlax-ldata-get "SX_Layout" "panel") sx:panels))
       (setq sx:pi (vl-position (vlax-ldata-get "SX_Layout" "panel") sx:panels)))
     (sx:stg "dialog")
     (if (not (sxl:dialog)) (progn (sx:xl-release) (setvar "CMDECHO" oldcmd) (exit)))))
  (setq panel (nth sx:pi sx:panels))
  (sxl:save-params panel)
  (if (and sxl:lib (/= sxl:lib "")) (setenv "SXL_LibDwg" sxl:lib))
  (setq sx:style (sx:ensure-style))
  ;; книга
  (sx:stg "read excel")
  (setq data (sx:read-data))
  (sxl:read-catalog)
  (sx:xl-release)
  (setq prow (vl-remove-if-not '(lambda (row) (= (sx:s (car row)) panel)) data))
  (sx:log (strcat "panel=" (sx:s panel) " rows=" (itoa (length prow)) " lib=" (sx:str sxl:lib)))
  ;; блоки: что есть в чертеже и в базе
  (sx:stg "blocks")
  (setq sxl:names (sxl:collect-names sxl:lib))
  (sx:log (strcat "block names=" (itoa (length sxl:names))))
  (sx:stg "build")
  (setq its (sxl:build prow))
  (sx:log (strcat "items=" (itoa (length its))))
  (if (null its)
    (progn (sx:lib-close) (setvar "CMDECHO" oldcmd)
           (princ (strcat "\n" (R "Нет аппаратов для компоновки у щита ") (sx:s panel)))
           (exit)))
  ;; скопировать из базы нужные блоки
  (sx:stg "import")
  (vla-startundomark sx:doc)
  (sx:try (function (lambda ()
    (sx:import-blocks sxl:lib
      (append (list sxl:bRail)
              (vl-remove nil (apply 'append (mapcar '(lambda (it) (mapcar '(lambda (p) (nth 2 p)) (nth 2 it))) its)))))))
    "import-blocks")
  (sx:lib-close)
  (setq sx:blkcache nil)
  ;; порядок прорисовки внутри блоков — как в базе (заливка за линиями)
  (sx:stg "draworder")
  (sx:try (function (lambda ()
    (sxl:fix-draworder
      (sxl:with-nested
        (append (list sxl:bRail)
                (vl-remove nil (apply 'append (mapcar '(lambda (it) (mapcar '(lambda (p) (nth 2 p)) (nth 2 it))) its)))))
      sxl:lib)))
    "draworder")
  (setq its (sxl:finalize its))
  (foreach it its
    (foreach p (nth 2 it)
      (sx:log (strcat "part " (car it) " " (nth 1 it) " | " (car p) " | blk=" (sx:str (nth 2 p))
                      " w=" (sxl:r1 (nth 5 p)) " h=" (sxl:r1 (nth 6 p))))))
  (setq dits (vl-remove-if-not '(lambda (it) (= (nth 3 it) "DIN")) its)
        pits (vl-remove-if-not '(lambda (it) (= (nth 3 it) "PNL")) its))
  ;; корпус и раскладка
  (sx:stg "choose")
  (setq res (sxl:choose dits pits))
  (if (null res)
    (progn (vla-endundomark sx:doc) (setvar "CMDECHO" oldcmd) (sx:lib-close)
           (princ (strcat "\n" (R "SXLAYOUT: нет корпуса, проходящего ограничения") (sxl:limits-text)
                          (R ". Ослабьте ограничения в окне или выберите корпус вручную.")))
           (sxl:logdump)
           (exit)))
  (sx:log (strcat "encl=" (car sxl:encsel) " reserve=" (itoa (nth 5 res)) " newrail=" (sx:str (nth 6 res))
                  " used=" (itoa (nth 3 res)) " over=" (itoa (length (nth 2 res)))))
  (if (and (= (nth 5 res) 0) (> *sxl-reserve* 0))
    (sxl:warn (strcat (R "запас ") (itoa *sxl-reserve*) (R "% не выдержан: аппараты уложены без запаса"))))
  (if (nth 2 res)
    (sxl:warn (strcat (R "не поместилось в корпус позиций: ") (itoa (length (nth 2 res)))
                      (R " (нарисованы под корпусом красным) - выберите корпус больше"))))
  (if (not (nth 7 res)) (sxl:warn (R "аппараты монтажной панели не помещаются под рейками")))
  ;; рисуем
  (sx:stg "draw")
  (if sxl:do-clear (sxl:clear))
  (sxl:layers)
  (setq st (sxl:draw panel res))
  (vl-catch-all-apply 'vla-regen (list sx:doc acAllViewports))
  (vla-endundomark sx:doc)
  (setvar "CMDECHO" oldcmd)
  (princ (strcat "\n" (R "Компоновка: ") (sx:s panel) ", " (car sxl:encsel)
                 (R ", реек занято ") (itoa (car st)) (R ", заполнение ") (sxl:r0 (cadr st)) "%"))
  (if sxl:warns
    (progn
      (princ (strcat "\n" (R "Замечания: ") (itoa (length sxl:warns))))
      (foreach w (reverse sxl:warns) (princ (strcat "\n  " w)))))
  (sxl:logdump)
  (princ))

(defun sxl:run (mode) (if (sxl:need-schema) (sxl:main mode)) (princ))
(defun c:SXLAYOUT () (sxl:run "dlg"))
(defun c:SXLUPDATE () (sxl:run "upd"))

;;; проверка без окна: (sxl:batch "C:\\...\\book.xlsx" "ЩР" "C:\\...\\base.dwg" "" '(0.0 0.0 0.0) '(0 0 0 0 ""))
;;; enc — имя корпуса из базы или "" (подбор); lim — (Ш В Г установка серия), 0/"" = без ограничения
(defun sxl:batch (xls panel lib enc pt lim)
  (if (sxl:need-schema)
    (progn
      (sxl:init)
      (if (sx:open-wb xls)
        (progn
          (setq sx:pi (cond ((vl-position panel sx:panels)) (t 0))
                sxl:lib lib sxl:encname (if (= (type enc) 'STR) enc "") sxl:inspt pt)
          (if (= (length lim) 5)
            (setq sxl:maxW (float (nth 0 lim)) sxl:maxH (float (nth 1 lim)) sxl:maxD (float (nth 2 lim))
                  sxl:inst (nth 3 lim) sxl:series (nth 4 lim)))
          (sxl:main "batch"))
        (princ (strcat "\nSXLAYOUT batch: cannot open " xls)))))
  (princ))

(princ "\n[SX_Layout v0.3 for v5] SXLAYOUT / SXLUPDATE / SXLCLEAR")
(princ)
