from pathlib import Path
p=Path('2_AutoCAD/Однолинейка/SX_Schema.lsp')
s=p.read_text(encoding='utf-8')
def replace(old,new):
    global s
    assert s.count(old)==1,(old[:90],s.count(old))
    s=s.replace(old,new)
replace('(defun sx:sel-units (keys draw / pl gl ml sl dk gid q mk i s)',
        '(defun sx:sel-units (keys draw / pl gl ml sl dk gid q mk i s insl)')
replace('      (setq keys (cons (strcat "S|" (itoa (sx:sheet (sx:i d)))) keys))))',
        '      (setq insl (cons (sx:sheet (sx:i d)) insl))))')
replace('  (setq sx:onlyP pl sx:onlyG gl sx:onlyM ml sx:onlyS sl)',
        '  ;; Для ввода обновить геометрию шин, сохраняя остальные места листа.\n  (foreach s insl (if (not (member s sl)) (setq sl (cons s sl))))\n  (setq sx:onlyP pl sx:onlyG gl sx:onlyM ml sx:onlyS sl)')
replace('    (setq u (sx:u d) bf (sx:s (sx:c d sx:cChT)) sx:xtb sx:vXTbot)',
        '    (setq u (sx:u d) bf (sx:s (sx:c d sx:cChT)) sx:xtb sx:vXTbot)\n    (if (and (sx:input-p d) (sx:is1 d sx:cHasLine) (not (sx:is1 d sx:cXT)) (not (sx:c d sx:cCh)))\n      (progn (setq sx:xtb sx:vIN0) (sx:draw-input d u)))')
# Only the required-block validation skips an outlet symbol for incoming lines.
start=s.index('(defun sx:need-blocks ')
end=s.index('(defun sx:join-sep ',start)
section=s[start:end]
assert section.count('(if (sx:is1 d sx:cHasLine)')==1
s=s[:start]+section.replace('(if (sx:is1 d sx:cHasLine)', '(if (and (sx:is1 d sx:cHasLine) (not (sx:input-p d)))')+s[end:]
replace('(princ "\\n[SX_Schema v3.15] SXDRAW / SXUPDATE / SXUPDSEL / SXUPDSHEET / SXCLEAR")',
        '(princ "\\n[SX_Schema v3.16] SXDRAW / SXUPDATE / SXUPDSEL / SXUPDSHEET / SXCLEAR")')
# Preserve ASCII string literals for LISPSYS=0 as well.
replace('118 51 46 49 51 32 1079 1072 1075', '118 51 46 49 54 32 1079 1072 1075')
p.write_text(s,encoding='utf-8',newline='\r\n')
