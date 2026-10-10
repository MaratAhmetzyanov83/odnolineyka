from pathlib import Path
import re

p=Path(__file__).parent
f=p/'SX_Panel.source.lsp'
s=f.read_text(encoding='utf-8')
def replace_def(name, text):
    global s
    start=s.index('(defun '+name+' ')
    end=s.find('\n(defun ',start+1)
    if end<0: raise ValueError(name)
    s=s[:start]+text.rstrip()+'\n'+s[end:]

replace_def('cp:plan', '''(defun cp:plan ( / usable packed maxh)
  (setq cp:problem nil cp:positions nil
        cp:railw (* (cp:g 'slots) (cp:g 'pitch))
        usable (* cp:railw (- 1.0 (/ (cp:g 'reserve) 100.0)))
        packed (cp:pack cp:devices usable (cp:g 'pitch)))
  (cond ((null cp:devices) (setq cp:problem "У выбранного щита нет аппаратов для компоновки."))
        ((caddr packed) (setq cp:problem (caddr packed)))
        ((> (cadr packed) (cp:g 'nrows))
         (setq cp:problem (strcat "Нужно рядов: " (itoa (cadr packed)) ". Задано: " (itoa (cp:g 'nrows)) ".")))
        (t (setq cp:positions (car packed))))
  (setq cp:totalrows (cp:g 'nrows)
        maxh (apply 'max (cons (cp:g 'devh) (mapcar '(lambda (d) (nth 4 d)) cp:devices)))
        cp:rowpitch (+ maxh (cp:g 'gap))
        cp:rowtop (+ (/ maxh 2.0) (* (1- cp:totalrows) cp:rowpitch)))
  (cp:put 'w cp:railw)
  (cp:put 'h (+ maxh (* (1- cp:totalrows) cp:rowpitch)))
  (null cp:problem))''')
replace_def('cp:defaults', '''(defun cp:defaults ( / cfg last)
  (setq cfg '((nrows . 3) (slots . 24) (pitch . 18.0) (reserve . 20.0)
              (gap . 60.0) (devh . 90.0)))
  (if (setq last (vlax-ldata-get "SX_PANEL" "last")) (setq cfg last))
  (if (null (assoc 'gap cfg)) (setq cfg (cons '(gap . 60.0) cfg)))
  (if (null (assoc 'library cfg)) (setq cfg (cons (cons 'library (cp:default-library)) cfg)))
  (if (null (assoc 'path cfg)) (setq cfg (cons (cons 'path (getenv "SX_LastXls")) cfg)))
  (if (null (assoc 'panel cfg)) (setq cfg (cons (cons 'panel (getenv "SX_LastPanel")) cfg)))
  cfg)''')
s=s.replace("'(w h depth nrows slots pitch reserve margin duct devh xtw xth)","'(nrows slots pitch reserve gap devh)")
s=s.replace('(cp:put \'bars (= (get_tile "bars") "1")) ', '')
start=s.index('    (start_list "caseblock")');end=s.index('    (start_list "panel")',start);s=s[:start]+s[end:]
start=s.index('    (set_tile "bars"');end=s.index('    (action_tile "browse"',start);s=s[:start]+s[end:]
start=s.index('  ;; Рейки, клеммы, шины');end=s.index('  (setq id (load_dialog',start);s=s[:start]+s[end:]
start=s.index('    (setq txt (cp:s (cp:at r 16)))');end=s.index('  cp:devices)',start);s=s[:start]+'    )\n'+s[end:]
# Remove cabinet selector and helper mapping functions, leaving apparatus geometry.
for name in ['cp:case-choice','cp:role-metrics','cp:role-size']:
    replace_def(name,'')
start=s.index('  (foreach name cp:blocknames');end=s.index('\n(defun cp:ensure-block',start);s=s[:start]+'  cp:blocknames)\n'+s[end:]
# Exclude cabinet/panel/rail/support blocks from apparatus chooser.
s=s.replace('"`**,*_G###"','"`**,CP_*,*_G###,DKC_*,R5*,DIN_*,SX_DYN_CQEN*,SX_DYN_PLASTRON*,SX_DYN_DIN*"')
s=s.replace('"`**,CP_*,*_G###"','"`**,CP_*,*_G###,DKC_*,R5*,DIN_*,SX_DYN_CQEN*,SX_DYN_PLASTRON*,SX_DYN_DIN*"')
replace_def('cp:draw-geometry', '''(defun cp:draw-geometry ( / w h pos)
  (setq w (cp:g 'w) h (cp:g 'h))
  (cp:text (/ w 2.0) (+ h 22.0) 6.0 (strcat "Компоновка аппаратов " (cp:g 'panel)) w "CP_TEXT")
  (foreach pos cp:positions (cp:device pos 0.0))
  (cp:draw-list (+ w 35.0) h))''')
s=s.replace('(vla-put-layer ob "CP_CASE") (cp:tag ob)', '(vla-put-layer ob "CP_DEVICES") (cp:tag ob)')
s=s.replace('"Укажите левый нижний угол корпуса щита: "','"Укажите точку вставки компоновки аппаратов: "')
s=s.replace(' ", клемм " (itoa (length cp:terminals))', '')
start=s.index('  (setq ob (vla-addtext cp:space (strcat "Клеммы:');end=s.index('\n\n;;; Родительский',start);s=s[:start]+'  )'+s[end:]
replace_def('cp:layers', '''(defun cp:layers ()
  (foreach p '(("CP_DEVICES" 7) ("CP_TEXT" 7))) (cp:layer (car p) (cadr p))))'''.replace('7))) (cp:', '7)) (cp:'))
s=s.replace(';;; Аппараты/рейки/клеммы вставляются из пользовательских блоков с сохранением геометрии.',
            ';;; Только модульные аппараты из пользовательской базы; корпус/рейки/клеммы не рисуются.')
s=s.replace('((wcmatch (strcase tk) "*30*") (setq diff T))','((wcmatch (strcase tk) (strcase "*мА")) (setq diff T))')
f.write_text(s,encoding='utf-8')
d=p/'SX_Panel.source.dcl'
dcl=d.read_text(encoding='utf-8')
start=dcl.index('  : boxed_column {\n    label = "Корпус');end=dcl.index('  : text { key = "status"',start)
dcl=dcl[:start]+'''  : boxed_column {
    label = "Размещение модульных аппаратов по рядам";
    : row {
      : edit_box { key = "nrows"; label = "Количество рядов:"; edit_width = 5; }
      : edit_box { key = "slots"; label = "Модулей в ряду:"; edit_width = 5; }
    }
    : row {
      : edit_box { key = "pitch"; label = "Шаг модуля, мм:"; edit_width = 5; }
      : edit_box { key = "reserve"; label = "Свободное место, %:"; edit_width = 5; }
    }
    : edit_box { key = "gap"; label = "Зазор между аппаратами соседних рядов, мм:"; edit_width = 6; }
    : edit_box { key = "devh"; label = "Минимальная высота аппаратов, мм:"; edit_width = 6; }
    : text { label = "Вставляются только ваши блоки аппаратов и их обозначения."; }
  }
'''+dcl[end:]
d.write_text(dcl,encoding='utf-8')
print('Modular-only source ready')
