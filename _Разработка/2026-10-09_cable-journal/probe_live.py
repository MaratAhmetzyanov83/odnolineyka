from pathlib import Path
import sys, json, shutil, time
import win32com.client as w
sys.stdout.reconfigure(encoding='utf-8',line_buffering=True)
p=Path(__file__).parent
dest=p/'КЖ_проверка.dwg'
if not dest.exists(): shutil.copy2(p/'Образец_КЖ_ЩР-1.dwg',dest)
a=w.GetActiveObject('AutoCAD.Application')
d=next((x for x in a.Documents if x.FullName.lower()==str(dest.resolve()).lower()),None)
if d is None: d=a.Documents.Open(str(dest.resolve()))
for attempt in range(40):
    try:
        d=next(x for x in a.Documents if x.FullName.lower()==str(dest.resolve()).lower())
        print('READY',d.ModelSpace.Count); break
    except Exception: time.sleep(.3)
d.Activate();time.sleep(.5)
for t in d.ModelSpace:
    if t.ObjectName=='AcDbTable':
        print('TABLE',t.Handle)
        for name in ['Rows','Columns','BreaksEnabled','BreakSpacing','TableBreakHeight','TableBreakFlowDirection','RepeatTopLabels','RepeatBottomLabels','TitleSuppressed','HeaderSuppressed','HorzCellMargin','VertCellMargin']:
            try: print(name,getattr(t,name))
            except Exception as e: print(name,str(e))
        print('first heights',[(i,t.GetRowHeight(i),t.GetCellTextHeight(i,2),t.GetCellTextStyle(i,2)) for i in range(min(t.Rows,4))])
first=min((l for l in d.Layouts if not l.ModelType),key=lambda l:l.TabOrder)
print('FIRST',first.Name)
expr='(progn (setq kjprobe (open "'+(p/'viewport.txt').resolve().as_posix()+'" "w")) (foreach kjh (list "2375" "191C" "5AAC4") (prin1 (entget (handent kjh)) kjprobe) (write-line "" kjprobe)) (close kjprobe) (princ))'
d.SendCommand(expr+'\n')
print('Sent probe')
