from pathlib import Path
import time,win32com.client as w

def run(script):
    for attempt in range(30):
        try:
            a=w.GetActiveObject('AutoCAD.Application')
            docs=a.Documents
            candidates=[docs.Item(i) for i in range(docs.Count)]
            d=next(q for q in candidates if q.Name=='Компоновка_проверка_из_базы.dwg')
            d.Activate();time.sleep(.3)
            active=a.ActiveDocument
            if active.Name!='Компоновка_проверка_из_базы.dwg': continue
            active.SendCommand('(load "'+Path(script).as_posix()+'")\n')
            print('Requested',Path(script).name)
            return
        except Exception as exc:
            last=exc
            time.sleep(.5)
    raise RuntimeError('AutoCAD busy; no command sent') from last
