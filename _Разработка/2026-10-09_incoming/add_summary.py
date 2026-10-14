from pathlib import Path

p = Path(__file__).resolve().parents[2] / '2_AutoCAD/Однолинейка/SX_Schema.lsp'
s = p.read_text(encoding='utf-8')

def ru(t):
    return "(sx:ru '(" + ' '.join(str(ord(c)) for c in t) + '))'

def replace(a, b):
    global s
    assert s.count(a) == 1, a[:120]
    s = s.replace(a, b)

reader = ''';;; Итоги берутся из вычисленной таблицы, включая коэффициенты спроса.
;;; Пустые таблицы (даже с заполненным cos и ошибками фаз) не рисуются.
(defun sx:summary-values (vals panel / r c row prev found metrics)
  (setq r 0)
  (foreach row vals
    (setq c 0)
    (foreach val row
      (if (and (not found) (> r 0) (>= c 3)
               (= (sx:s val) @PUST@)
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
  (if (setq sh (cdr (assoc @SHEET@ sx:sheets)))
    (progn
      (setq rg (vlax-get-property sh 'Range "A1:AQ1000"))
      (setq vals (vl-catch-all-apply
                   '(lambda () (vlax-safearray->list
                      (vlax-variant-value (vlax-get-property rg 'Value2)))) nil))
      (sx:rel rg)
      (if (vl-catch-all-error-p vals) (error (vl-catch-all-error-message vals)))
      (sx:summary-values (mapcar '(lambda (row) (mapcar 'sx:v row)) vals) panel))))

'''.replace('@PUST@', ru('Руст=')).replace('@SHEET@', ru('Нагрузка_щитов'))
replace('(defun sx:xl-release ()', reader + '(defun sx:xl-release ()')

renderer = ''';;; Итоги под первым листом: штатная таблица AutoCAD 120 x 40 мм.
;;; При обновлении сохраняем точку и стиль уже имеющейся таблицы.
(defun sx:summary-capture ( / ob e xd pt target)
  (setq sx:sumpt (sx:pt 7.56437443444702 -5.3734190440593) sx:sumstyle nil)
  (setq target sx:sumpt)
  (vlax-for ob sx:ms
    (if (and (= (vla-get-objectname ob) "AcDbTable")
             (= (vla-get-rows ob) 5) (= (vla-get-columns ob) 6))
      (progn
        (setq e (vlax-vla-object->ename ob) xd (sx:xd-get e)
              pt (vlax-safearray->list (vlax-variant-value (vla-get-insertionpoint ob))))
        (if (or (equal xd (list sx:tagpanel "U"))
                (and (null xd) (< (distance pt target) 0.15)))
          (progn
            (setq sx:sumpt pt sx:sumstyle (vla-get-stylename ob))
            (vla-put-layer ob sx:TB)
            (sx:xd-set e "U")))))))

(defun sx:summary-num (v / oldz text)
  ;; Два знака после запятой, как в Excel, независимо от DIMZIN чертежа.
  (setq oldz (getvar "DIMZIN")) (setvar "DIMZIN" 0)
  (setq text (vl-string-translate "." "," (rtos v 2 2)))
  (setvar "DIMZIN" oldz)
  text)

(defun sx:draw-summary ( / ob rows r c row)
  (if sx:summary
    (progn
      (setq ob (vla-addtable sx:ms (vlax-3d-point sx:sumpt) 5 6 8.0 20.0))
      ;; Сразу пометить объект: следующий запуск удалит и незавершённую таблицу.
      (vla-put-layer ob sx:TB) (sx:xd-set (vlax-vla-object->ename ob) "U")
      (vla-put-regeneratetablesuppressed ob :vlax-true)
      (if sx:sumstyle (vla-put-stylename ob sx:sumstyle))
      (vla-put-titlesuppressed ob :vlax-true)
      (vla-put-headersuppressed ob :vlax-true)
      (vla-unmergecells ob 0 4 0 5)
      (vla-put-horzcellmargin ob 0.5) (vla-put-vertcellmargin ob 0.5)
      (setq rows
        (list
          (list @PUST@ (sx:summary-num (nth 0 sx:summary)) @KW@ "L1p=" (sx:summary-num (nth 5 sx:summary)) @KW@)
          (list @PR@ (sx:summary-num (nth 1 sx:summary)) @KW@ "L2p=" (sx:summary-num (nth 6 sx:summary)) @KW@)
          (list @IR@ (sx:summary-num (nth 2 sx:summary)) @A@ "L3p=" (sx:summary-num (nth 7 sx:summary)) @KW@)
          (list @COS@ (sx:summary-num (nth 3 sx:summary)) "" "" "" "")
          (list @KC@ (sx:summary-num (nth 4 sx:summary)) "" "" "" "")))
      (setq r 0)
      (foreach row rows
        (setq c 0)
        (foreach text row
          (vla-settext ob r c text)
          (vla-setcelltextstyle ob r c sx:style)
          (vla-setcelltextheight ob r c 2.5)
          (vla-setcellalignment ob r c 5)
          (setq c (1+ c)))
        (vla-setrowheight ob r 8.0)
        (setq r (1+ r)))
      (setq c 0)
      (repeat 6 (vla-setcolumnwidth ob c 20.0) (setq c (1+ c)))
      (vla-mergecells ob 3 3 1 5) (vla-mergecells ob 4 4 1 5)
      (vla-put-regeneratetablesuppressed ob :vlax-false)
      (vla-update ob))))

'''
for token, text in {'PUST':'Руст=', 'PR':'Рр=', 'IR':'Iр=', 'KW':'кВт', 'A':'А', 'COS':'Cos(ф)=', 'KC':'Кс='}.items():
    renderer = renderer.replace('@' + token + '@', ru(text))
replace(';;; ---------------------------------------------------------------- очистка', renderer + ';;; ---------------------------------------------------------------- очистка')
replace('((= k "H") 0)', '((member k \'("H" "U")) 0)')
replace('  dk)\n;; удалить помеченные', '  (cons "U" dk))\n;; удалить помеченные')
replace('(defun sx:draw-outlet (d u / blk)', '(defun sx:draw-outlet (d u / blk ob)')
replace('  (sx:insv blk u sx:vOutTop 1.0 0.0 nil sx:BL))', '  (setq ob (sx:insv blk u sx:vOutTop 1.0 0.0 nil sx:BL))\n  (sx:red ob)\n  ob)')
replace(';; нет кода УГО или код неизвестен -> кабельный вывод (не красная заглушка)', ';; нет кода УГО или код неизвестен -> красный кабельный вывод')
replace('  (setq data (sx:read-data))', '  (setq data (sx:read-data))\n  (setq sx:summary (sx:read-summary panel))\n  (sx:log (strcat "summary=" (sx:str sx:summary)))')
replace('  (if (or sx:part sx:do-clear) (progn (sx:layers) (sx:adopt-inputs draw)))', '  (sx:layers) (sx:summary-capture)\n  (if (or sx:part sx:do-clear) (sx:adopt-inputs draw))')
replace('  (sx:stg "layers") (sx:layers)', '  (sx:del-keys panel \'("U"))\n  (sx:stg "layers") (sx:layers)')
replace('  (sx:stg "done")', '  (sx:stg "summary") (sx:try \'sx:draw-summary "summary")\n  (sx:stg "done")')
replace(';;;   K — шина KNX, R|i', ';;;   U — итоги щита, K — шина KNX, R|i')
p.write_bytes(s.replace('\n', '\r\n').encode('utf-8'))
print('Summary and red outlets added')
