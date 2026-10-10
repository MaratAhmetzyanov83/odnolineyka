from pathlib import Path
import re

root=Path(__file__).resolve().parents[2]
out=root/'2_AutoCAD/Компоновка'
out.mkdir(parents=True,exist_ok=True)
src=Path(__file__).with_name('SX_Panel.source.lsp').read_text(encoding='utf-8')
def convert(m):
    s=m.group(0)[1:-1]
    if not any(ord(c)>127 for c in s):return m.group(0)
    s=s.replace('\\n','\n')
    assert '\\' not in s, s
    return "(cp:ru '("+' '.join(str(ord(c)) for c in s)+'))'
src=re.sub(r'"(?:[^"\\]|\\.)*"',convert,src)
(out/'SX_Panel.lsp').write_bytes(src.replace('\n','\r\n').encode('utf-8'))
dcl=Path(__file__).with_name('SX_Panel.source.dcl').read_text(encoding='utf-8')
(out/'SX_Panel.dcl').write_bytes(dcl.replace('\n','\r\n').encode('cp1251'))
print('Built',out)
