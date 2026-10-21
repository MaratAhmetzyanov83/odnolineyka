# $1 = исходный DWG (копия файла пользователя) -> kp/W.dwg с новой ГРЩ-ЖД
S="C:/Users/user/AppData/Local/Temp/claude/C--Users-user-Desktop-02-------------CAD-Revit-odnolineyka/2b874a90-bad1-4e69-aeab-bd65990a1f89/scratchpad"
cd "$S"
cp "$1" kp/W.dwg
py -I runacc.py "$(cygpath -w "$S/kp/W.dwg")" "$(cygpath -w "$S/kp/clean.scr")" 400 | tr -d '\r' | grep -E "CLEANED=" | head -1
py -I gen_grsh.py kp/grsh.lsp
py -I runacc.py "$(cygpath -w "$S/kp/W.dwg")" "$(cygpath -w "$S/kp/w.scr")" 400 | tr -d '\r' | grep -E "GRSH-DONE|ошибка" | head -2
cd kp
py -I - <<'PY' 2>&1 | grep -v FIELD || true
import ezdxf
d=ezdxf.readfile('W.dxf'); l=d.paperspace('ГРЩ-ЖД')
for v in l.query('VIEWPORT'):
    if v.dxf.view_center_point.x>1000:
        v.dxf.view_center_point=(31000+1564.0,3033.48+2100.0); v.dxf.view_height=4100.0
for e in l.query('MTEXT'): e.text=e.text.replace('1:8','1:10')
d.saveas('T.dxf')
PY
rm -f T.dwg
py -I ../runacc.py "$(cygpath -w "$S/kp/T.dxf")" "$(cygpath -w "$S/kp/t.scr")" 400 > /dev/null
py -I ../runacc.py "$(cygpath -w "$S/kp/W.dwg")" "$(cygpath -w "$S/kp/imp.scr")" 400 | tr -d '\r' | grep -E "VPN=" | head -1
