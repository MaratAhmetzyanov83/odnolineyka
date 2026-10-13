"""Проверка SXDRAW без окна: свой скрытый AutoCAD и свой Excel, SXUPDATE по КОПИИ книги -> DXF + журнал.
Картинка: py 5_Проверка/dxf_png.py out.dxf out.png -60 280 -5 235 130  py run_sx.py <книга.xlsx> <щит> <выход.dxf>"""
import win32com.client as w, os, sys, time, shutil, tempfile
sys.stdout.reconfigure(encoding='utf-8')
xls, panel, out = sys.argv[1], sys.argv[2], sys.argv[3]
T = os.path.join(tempfile.gettempdir(), 'sxdraw_test')
os.makedirs(T, exist_ok=True)
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SRC = os.path.join(ROOT, '2_AutoCAD', 'Однолинейка', 'SX_Schema.lsp')
lsp = os.path.join(T, 'sx.lsp'); shutil.copy(SRC, lsp)
lib = os.path.join(T, 'lib.dwg')
if not os.path.exists(lib):
    shutil.copy(os.path.join(ROOT, '2_AutoCAD', 'Блоки', '02_Однолинейка.dwg'), lib)

def ru(s):
    return "(sx:ru '(" + " ".join(str(ord(c)) for c in s) + "))"
def lp(p):
    return p.replace('\\', '/')

a = w.DispatchEx('AutoCAD.Application')
try:
    a.Visible = False
    for i in range(90):
        try:
            a.Documents; break
        except Exception:
            time.sleep(2)
    doc = a.Documents.Add()
    time.sleep(2)
    doc.SetVariable('USERI1', 0)
    cmds = [
        f'(load "{lp(lsp)}")',
        '(defun sx:ask-go () T)',
        '(setq sx:xl (vlax-create-object "Excel.Application") sx:xl-created T)',
        f'(vlax-ldata-put "SX_Schema" "xls" "{lp(xls)}")',
        f'(vlax-ldata-put "SX_Schema" "panel" {ru(panel)})',
        '(vlax-ldata-put "SX_Schema" "pt" (list 0.0 0.0 0.0))',
        f'(vlax-ldata-put "SX_Schema" "lib" "{lp(lib)}")',
        '(progn (sx:main T) (setvar "USERI1" 77) (princ))',
    ]
    for c in cmds:
        doc.SendCommand(c + '\n')
    for i in range(300):
        try:
            if doc.GetVariable('USERI1') == 77: break
        except Exception:
            pass
        time.sleep(1)
    print('done flag', doc.GetVariable('USERI1'))
    if os.path.exists(out): os.remove(out)
    doc.SaveAs(out, 61)
    print('saved', out)
    doc.Close(False)
finally:
    try:
        a.Quit()
    except Exception as e:
        print('quit', e)
log = os.path.join(os.path.dirname(xls), 'SX_log.txt')
if os.path.exists(log):
    print(open(log, encoding='cp1251', errors='replace').read()[-3000:])
