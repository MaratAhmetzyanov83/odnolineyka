
S="C:/Users/user/AppData/Local/Temp/claude/C--Users-user-Desktop-02-------------CAD-Revit-odnolineyka/2b874a90-bad1-4e69-aeab-bd65990a1f89/scratchpad"
cd "$S"
py -I gen_grsh.py kp/grsh.lsp
cp kp/K.dwg kp/W.dwg
py -I runacc.py "$(cygpath -w "$S/kp/W.dwg")" "$(cygpath -w "$S/kp/w.scr")" 400 | tr -d '\r' | grep -iE "GRSH-DONE|ошибка" | head -3
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
true
rm -f T.dwg
py -I ../runacc.py "$(cygpath -w "$S/kp/T.dxf")" "$(cygpath -w "$S/kp/t.scr")" 400 > /dev/null
ls -la T.dwg
py -I ../runacc.py "$(cygpath -w "$S/kp/W.dwg")" "$(cygpath -w "$S/kp/imp.scr")" 400 | tr -d '\r' | grep -E "VPN=" | head -2
py -I ../render.py W.dxf W_new.png Model 31000,3033,34720,7233 200 2>&1 | grep -v FIELD || true
