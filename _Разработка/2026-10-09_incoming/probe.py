import pathlib,time,win32com.client as w
d=w.GetActiveObject('AutoCAD.Application').ActiveDocument
d.SetVariable('LOGFILEMODE',1)
exec((pathlib.Path(__file__).parent/'dump_acad.py').read_text())
time.sleep(1)
p=d.GetVariable('LOGFILENAME')
print(p)
try: print(pathlib.Path(p).read_text(encoding='cp1251',errors='replace')[-6000:])
except Exception as e: print(e)
d.SetVariable('LOGFILEMODE',0)
