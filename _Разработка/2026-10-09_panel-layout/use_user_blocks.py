from pathlib import Path
p=Path(__file__).with_name('SX_Panel.source.lsp');s=p.read_text(encoding='utf-8')
def fn(name,body):
    global s
    start=s.index('(defun '+name+' ')
    i=start;depth=0;string=False;comment=False
    while i<len(s):
        c=s[i]
        if comment:
            if c=='\n':comment=False
        elif string:
            if c=='\\':i+=1
            elif c=='"':string=False
        elif c==';':comment=True
        elif c=='"':string=True
        elif c=='(':depth+=1
        elif c==')':
            depth-=1
            if depth==0:break
        i+=1
    s=s[:start]+body+s[i+1:]

fn('cp:apply-sizes',r'''(defun cp:apply-sizes ( / mapping metrics name scale rot)
  (setq cp:devices (mapcar '(lambda (d)
    (setq mapping (cdr (assoc (nth 1 d) (cp:g 'sizes))))
    (if mapping
      (progn
        (setq name (car mapping) scale (cadr mapping) rot (caddr mapping)
              metrics (cp:measure name scale rot))
        (list (nth 0 d) (nth 1 d) (nth 2 d)
              (/ (- (nth 2 metrics) (nth 0 metrics)) (cp:g 'pitch))
              (- (nth 3 metrics) (nth 1 metrics)) (nth 5 d) (nth 6 d) name scale rot))
      d)) cp:devices)))''')
fn('cp:size-fill',r'''(defun cp:size-fill ( / item size name)
  (setq item (nth cp:size-i cp:size-items) size (cdr (assoc (car item) cp:size-data)) name (car size))
  (set_tile "model" (cadr item))
  (set_tile "block" (itoa (if (member name cp:blocknames) (vl-position name cp:blocknames) 0)))
  (set_tile "scale" (cp:s (if (cadr size) (cadr size) 1.0)))
  (set_tile "rotation" (cp:s (if (caddr size) (caddr size) 0.0))))''')
fn('cp:size-store',r'''(defun cp:size-store ( / key name scale rot size)
  (setq key (car (nth cp:size-i cp:size-items)) name (nth (atoi (get_tile "block")) cp:blocknames)
        scale (cp:num (get_tile "scale")) rot (cp:num (get_tile "rotation")))
  (if (and name (/= name "") scale (> scale 0) rot)
    (progn
      (setq size (cons key (list name scale rot)))
      (setq cp:size-data (if (assoc key cp:size-data) (subst size (assoc key cp:size-data) cp:size-data) (cons size cp:size-data))) T)
    (progn (set_tile "status" "Выберите ваш блок, положительный масштаб и угол поворота.") nil)))''')
fn('cp:size-accept',r'''(defun cp:size-accept ()
  (if (cp:size-store)
    (if (vl-some '(lambda (p) (or (null (cadr p)) (= (cadr p) ""))) cp:size-data)
      (set_tile "status" "Назначьте ваш блок каждой позиции. Пустые позиции отмечены знаком ?." )
      (done_dialog 1))))''')
fn('cp:sizes-dialog',r'''(defun cp:sizes-dialog ( / id res key d size name label)
  (setq cp:size-items nil cp:size-data nil cp:size-i 0)
  (foreach d cp:devices
    (setq key (nth 1 d) size (cdr (assoc key (cp:g 'sizes))))
    (if (null size) (setq size (list (cp:default-block d) 1.0 0.0)))
    (if (not (assoc key cp:size-items))
      (progn
        (setq label (strcat (if (car size) "" "? ") (car d) " — " (nth 2 d)))
        (setq cp:size-items (append cp:size-items (list (list key label))))
        (setq cp:size-data (cons (cons key size) cp:size-data)))))
  ;; Рейки, клеммы, шины и кабель-каналы используют ту же привязку к вашим блокам.
  (foreach p '(("@DIN" "DIN-рейка" "Дин" 1.0 0.0)
               ("@XT" "Клемма" "VPR-2.5-3L-GY" 1.0 0.0)
               ("@N" "Шина N" "Шина N" 1.0 90.0)
               ("@PE" "Шина PE" "Шина N" 1.0 90.0))
    (setq key (car p) size (cdr (assoc key (cp:g 'sizes))))
    (if (null size) (setq size (cddr p)))
    (setq cp:size-items (append cp:size-items (list (list key (cadr p))))
          cp:size-data (cons (cons key size) cp:size-data)))
  (setq id (load_dialog (findfile "SX_Panel.dcl")))
  (new_dialog "cp_sizes" id)
  (start_list "items") (mapcar '(lambda (p) (add_list (cadr p))) cp:size-items) (end_list)
  (start_list "block") (mapcar 'add_list cp:blocknames) (end_list)
  (set_tile "items" "0") (cp:size-fill)
  (action_tile "items" "(cp:size-select $value)")
  (action_tile "apply" "(cp:size-store)")
  (action_tile "accept" "(cp:size-accept)") (action_tile "cancel" "(done_dialog 0)")
  (setq res (start_dialog)) (unload_dialog id)
  (if (= res 1) (progn (cp:put 'sizes cp:size-data) (cp:apply-sizes) T)))''')
fn('cp:device',r'''(defun cp:device (pos left / d x y w h)
  (setq d (car pos) x (+ left (nth 2 pos)) y (cp:yc (cadr pos)) w (* (nth 3 d) (cp:g 'pitch)) h (nth 4 d))
  (cp:place-block (list (nth 7 d) (nth 8 d) (nth 9 d)) x y "CP_DEVICES" (car d))
  (cp:text (+ x (/ w 2.0)) (+ y (/ h 2.0) 8.0) 4.0 (car d) (- w 2.0) "CP_TEXT"))''')

addon=r'''
;;; Импортируются именно пользовательские определения блоков, без их перерисовки.
(defun cp:default-library ( / base)
  (if (setq base (findfile "SX_Panel.lsp"))
    (findfile (strcat (vl-filename-directory base) "/../Блоки/SX_Компоновка_пользовательские.dwg"))))
(defun cp:import-library ( / path dbx src blks dest names objs b arr result)
  (setq path (cp:g 'library) dest (vla-get-blocks cp:doc))
  (if (and path (/= path "") (findfile path))
    (progn
      (vlax-for d (vla-get-documents (vlax-get-acad-object))
        (if (= (strcase (vla-get-fullname d)) (strcase path)) (setq src d)))
      (if (null src)
        (progn
          (setq dbx (vla-getinterfaceobject (vlax-get-acad-object) (strcat "ObjectDBX.AxDbDocument." (substr (getvar "ACADVER") 1 2))))
          (vla-open dbx path) (setq src dbx)))
      (setq blks (vla-get-blocks src))
      (vlax-for b blks
        (if (and (not (wcmatch (vla-get-name b) "`**")) (= (vla-get-isxref b) :vlax-false)
                 (= (vla-get-islayout b) :vlax-false) (not (tblsearch "BLOCK" (vla-get-name b))))
          (setq objs (cons b objs))))
      (if objs
        (progn
          (setq arr (vlax-make-safearray vlax-vbObject (cons 0 (1- (length objs)))))
          (vlax-safearray-fill arr objs)
          (setq result (vl-catch-all-apply 'vla-copyobjects (list src arr dest)))))
      (if dbx (cp:rel dbx))
      (if (and result (vl-catch-all-error-p result)) (error (vl-catch-all-error-message result)))))
  (setq cp:blocknames (list ""))
  (vlax-for b dest
    (if (and (not (wcmatch (vla-get-name b) "`**,CP_*,SX_*"))
             (= (vla-get-isxref b) :vlax-false) (= (vla-get-islayout b) :vlax-false))
      (setq cp:blocknames (cons (vla-get-name b) cp:blocknames))))
  (setq cp:blocknames (vl-sort cp:blocknames '<)))
(defun cp:default-block (d / parts name)
  (setq name (nth 2 d))
  (cond
    ((tblsearch "BLOCK" name) name)
    ((= (nth 5 d) "QF")
     (setq parts (cp:qparts name))
     (cond ((and (caddr parts) (= (car parts) "2P")) "DS201")
           ((and (caddr parts) (member (car parts) '("3P" "4P"))) "DS203")
           ((= (car parts) "1P") "S201") ((= (car parts) "3P") "S203")))
    ((= (nth 5 d) "KM") "ESB40")
    ((= (nth 5 d) "PS") "БП DIN")))
(defun cp:measure (name scale rot / key cached blocks tmp ob p1 p2 result)
  (setq key (list name scale rot))
  (if (setq cached (assoc key cp:measure-cache)) (cdr cached)
    (progn
      (if (not (tblsearch "BLOCK" name)) (error (strcat "Нет пользовательского блока: " name)))
      (setq blocks (vla-get-blocks cp:doc) tmp (vla-add blocks (cp:p 0.0 0.0) (strcat "CP_MEASURE_" (itoa (abs (getvar "MILLISECS"))))))
      (setq ob (vla-insertblock tmp (cp:p 0.0 0.0) name scale scale scale (* pi (/ rot 180.0))))
      (vla-getboundingbox ob 'p1 'p2)
      (setq p1 (vlax-safearray->list p1) p2 (vlax-safearray->list p2) result (list (car p1) (cadr p1) (car p2) (cadr p2)))
      (vla-delete ob) (vla-delete tmp)
      (if (or (< (- (nth 2 result) (nth 0 result)) 0.01) (< (- (nth 3 result) (nth 1 result)) 0.01)) (error (strcat "Неверные габариты блока: " name)))
      (setq cp:measure-cache (cons (cons key result) cp:measure-cache)) result)))
(defun cp:mapping (key) (cdr (assoc key (cp:g 'sizes))))
(defun cp:role-metrics (key / m)
  (if (setq m (cp:mapping key)) (cp:measure (car m) (cadr m) (caddr m))
    (error (strcat "Не назначен блок: " key))))
(defun cp:role-size (key / bb)
  (setq bb (cp:role-metrics key)) (list (- (nth 2 bb) (nth 0 bb)) (- (nth 3 bb) (nth 1 bb))))
(defun cp:place-block (mapping x y lay id / name scale rot bb ob att)
  (setq name (car mapping) scale (cadr mapping) rot (caddr mapping) bb (cp:measure name scale rot))
  (setq ob (vla-insertblock cp:space (cp:p (- x (nth 0 bb)) (- y (/ (+ (nth 1 bb) (nth 3 bb)) 2.0))) name scale scale scale (* pi (/ rot 180.0))))
  (vla-put-layer ob lay)
  (if (and id (= (vla-get-hasattributes ob) :vlax-true))
    (foreach att (vlax-invoke ob 'GetAttributes)
      (if (member (strcase (vla-get-tagstring att)) (list "ID" "QF" "XT" "TAG" "NAME" "1" (strcase "КМ"))) (vla-put-textstring att id))))
  ob)
'''
s=s.replace(';;; Геометрия блока.',addon+'\n;;; Геометрия блока.')
# Drawer uses source apparatus, terminal, rail and bus definitions; only framing,
# annotations, reserve and cable-channel boundaries are created as geometry.
fn('cp:draw-geometry',r'''(defun cp:draw-geometry ( / w h left right row y ducty usable pos xti x tr ty barsx yy n rsize bsize i)
  (setq w (cp:g 'w) h (cp:g 'h) left (/ (- w cp:railw) 2.0) right (+ left cp:railw)
        usable (* cp:railw (- 1.0 (/ (cp:g 'reserve) 100.0))) rsize (cp:role-size "@DIN"))
  (cp:rect 0.0 0.0 w h "CP_CASE") (cp:rect 15.0 15.0 (- w 15.0) (- h 15.0) "CP_CASE")
  (cp:text (/ w 2.0) (+ h 30.0) 8.0 (strcat "Компоновка щита " (cp:g 'panel)) w "CP_TEXT")
  (cp:text (/ w 2.0) (+ h 12.0) 4.0 (strcat "Корпус " (cp:s w) " x " (cp:s h) " x " (cp:s (cp:g 'depth)) " мм; вид изнутри") w "CP_TEXT")
  (setq row 0)
  (repeat cp:totalrows
    (setq y (cp:yc row) i 0)
    (repeat cp:dinsegments
      (cp:place-block (cp:mapping "@DIN") (+ (/ (- w cp:dinwidth) 2.0) (* i (car rsize))) y "CP_DIN" nil)
      (setq i (1+ i)))
    (cp:text 29.0 y 3.5 (strcat "DIN " (itoa (1+ row))) 25.0 "CP_TEXT")
    (if (and (< row (cp:g 'nrows)) (> (cp:g 'reserve) 0))
      (progn (cp:rect (+ left usable) (- y 40.0) right (+ y 40.0) "CP_DUCT")
             (cp:text (/ (+ left usable right) 2.0) y 4.0 "Резерв" (- cp:railw usable 4.0) "CP_TEXT")))
    (if (< row (1- cp:totalrows))
      (progn (setq ducty (- y (/ cp:rowpitch 2.0)))
             (cp:rect left (- ducty (/ (cp:g 'duct) 2.0)) right (+ ducty (/ (cp:g 'duct) 2.0)) "CP_DUCT")
             (cp:text (/ w 2.0) ducty 3.5 "Кабель-канал" (- cp:railw 10.0) "CP_TEXT")))
    (setq row (1+ row)))
  (foreach pos cp:positions (cp:device pos left))
  (setq xti 0)
  (foreach id cp:terminals
    (setq tr (fix (/ xti cp:xtcap)) ty (cp:yc (+ (cp:g 'nrows) tr)) x (+ left (* (rem xti cp:xtcap) (cp:g 'xtw))))
    (cp:place-block (cp:mapping "@XT") x ty "CP_DEVICES" id)
    (setq ob (cp:text (+ x (/ (cp:g 'xtw) 2.0)) (+ ty (/ (cp:g 'xth) 2.0) 14.0) 2.5 id 28.0 "CP_TEXT"))
    (vla-put-rotation ob (/ pi 2.0)) (setq xti (1+ xti)))
  (if (and (cp:g 'bars) (> cp:xtrows 0))
    (progn (setq barsx (- right 145.0) y (cp:yc (1- cp:totalrows)))
      (foreach p '(("@N" "N" "CP_N" 18.0) ("@PE" "PE" "CP_PE" -18.0))
        (setq yy (+ y (nth 3 p)) bsize (cp:role-size (car p)))
        (cp:place-block (cp:mapping (car p)) barsx yy (nth 2 p) (cadr p))
        (cp:text (+ barsx (/ (car bsize) 2.0)) (+ yy (/ (cadr bsize) 2.0) 8.0) 4.0 (cadr p) (car bsize) (nth 2 p)))))
  (cp:dim 0.0 0.0 w 0.0 (/ w 2.0) -28.0) (cp:dim 0.0 0.0 0.0 h -28.0 (/ h 2.0))
  (cp:text (/ w 2.0) -55.0 4.0 "Аппараты, рейки и клеммы: пользовательские блоки. Масштаб блока сохраняется." (+ w 50.0) "CP_TEXT")
  (cp:draw-list (+ w 45.0) h))''')
# Add dimensions measured from block definitions before plan.
s=s.replace('(setq cp:problem nil cp:positions nil cp:railw', '''(setq cp:problem nil cp:positions nil)
  (setq rsize (cp:role-size "@DIN") xsize (cp:role-size "@XT"))
  (cp:put 'xtw (car xsize)) (cp:put 'xth (cadr xsize))
  (setq cp:railw''')
s=s.replace('(defun cp:plan ( / rw usable packed xtcap xtrows n maxh minpitch required pitch)', '(defun cp:plan ( / rw usable packed xtcap xtrows n maxh minpitch required pitch rsize xsize)')
s=s.replace('  (cond\n    ((> (+ cp:railw', '''  (setq cp:dinsegments (max 1 (fix (+ (/ cp:railw (car rsize)) 0.999999))) cp:dinwidth (* cp:dinsegments (car rsize)))
  (cond
    ((> (+ cp:dinwidth''')
s=s.replace('(cp:s (+ cp:railw (* 2 (cp:g \'margin))))','(cp:s (+ cp:dinwidth (* 2 (cp:g \'margin))))')
s=s.replace('  (if (null (assoc \'path cfg))', '  (if (null (assoc \'library cfg)) (setq cfg (cons (cons \'library (cp:default-library)) cfg)))\n  (if (null (assoc \'path cfg))')
s=s.replace('  (if good\n    (progn (cp:put', '  (if good\n    (progn (cp:put \'library (get_tile "library")) (cp:put')
s=s.replace('    (set_tile "file" (cp:s (cp:g \'path)))', '    (set_tile "file" (cp:s (cp:g \'path)))\n    (set_tile "library" (cp:s (cp:g \'library)))')
s=s.replace('cp:doc (vla-get-activedocument (vlax-get-acad-object)) cp:newdef nil', 'cp:doc (vla-get-activedocument (vlax-get-acad-object)) cp:newdef nil cp:measure-cache nil cp:pendingnew nil')
s=s.replace('    (if (and cp:newdef (null new))', '    (if (and cp:pendingnew (null new)) (vl-catch-all-apply \'vla-delete (list cp:pendingnew)))\n    (if (and cp:newdef (null new))')
s=s.replace('          (setq ok (if edit (cp:dialog) T))', '          (setq ok (if edit (cp:dialog) T) cancel (not ok))\n          (if ok (cp:import-library))')
s=s.replace('(vl-some \'(lambda (d) (null (nth 3 d))) cp:devices)', '(vl-some \'(lambda (d) (null (nth 7 d))) cp:devices)')
s=s.replace('(vla-put-layer ob "CP_CASE") (cp:tag ob)', '(setq cp:pendingnew ob)\n  (vla-put-layer ob "CP_CASE") (cp:tag ob)')
# close the run function (one unmatched opening in the initial implementation)
s=s.replace('  (princ))\n(defun c:CPDRAW', '  (princ)))\n(defun c:CPDRAW')
s=s.replace(';;; Габаритные обозначения на DIN-рейках, не заводские изображения аппаратов.', ';;; Аппараты/рейки/клеммы вставляются из пользовательских блоков с сохранением геометрии.')
s=s.replace(';;; Аппарат: (id, ключ габаритов, наименование, ширина DIN, высота, тип, строка).',';;; Аппарат: (id, ключ, модель, ширина DIN, высота, тип, строка, блок, масштаб, угол).')
p.write_text(s,encoding='utf-8')

d=Path(__file__).with_name('SX_Panel.source.dcl');s=d.read_text(encoding='utf-8')
s=s.replace('    : popup_list { key = "panel";', '    : edit_box { key = "library"; label = "DWG с вашими блоками:"; edit_width = 55; }\n    : popup_list { key = "panel";')
s=s.replace('  label = "Габариты аппаратов для компоновки";', '  label = "Ваши блоки для компоновки щита";')
s=s.replace('Ширина из каталога DIN. Пустые значения задайте по паспорту аппарата.', 'Выберите ваш блок для каждого типа. Габариты измеряются по блоку; геометрия сохраняется.')
s=s.replace('    : edit_box { key = "mods"; label = "Ширина, мод. DIN:"; edit_width = 7; }\n    : edit_box { key = "height"; label = "Высота габарита, мм:"; edit_width = 7; }', '    : popup_list { key = "block"; label = "Блок:"; width = 32; }\n    : edit_box { key = "scale"; label = "Масштаб:"; edit_width = 5; }\n    : edit_box { key = "rotation"; label = "Угол, °:"; edit_width = 5; }')
s=s.replace('    : edit_box { key = "devh"; label = "Высота условного габарита аппарата:";', '    : edit_box { key = "devh"; label = "Минимальная высота зоны аппаратов:";')
s=s.replace('  : boxed_column {\n    label = "Клеммы — отдельная нижняя рейка (при необходимости несколько)";\n    : row {\n      : edit_box { key = "xtw"; label = "Ширина клеммы:"; edit_width = 6; }\n      : edit_box { key = "xth"; label = "Длина вдоль корпуса:"; edit_width = 6; }\n    }', '  : boxed_column {\n    label = "Клеммы — отдельная нижняя рейка, размеры берутся из вашего блока";\n    : edit_box { key = "xtw"; label = "Ширина (заменится габаритом блока):"; edit_width = 6; }\n    : edit_box { key = "xth"; label = "Длина (заменится габаритом блока):"; edit_width = 6; }')
d.write_text(s,encoding='utf-8')
print('Converted layout to user blocks')
