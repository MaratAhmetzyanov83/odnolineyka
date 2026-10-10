from pathlib import Path
from collections import Counter
import win32com.client as w
import sys
sys.stdout.reconfigure(encoding='utf-8',line_buffering=True)
p=Path(__file__).parent;root=p.parents[1]
a=w.GetActiveObject('AutoCAD.Application')
src=a.GetInterfaceObject('ObjectDBX.AxDbDocument.25')
src.Open(str(root/'База блоков компоновки (1).dwg'))
dest=next(d for d in a.Documents if d.Name=='Компоновка_проверка_из_базы.dwg')
for label,doc in [('source',src),('test',dest)]:
    for name in ['S201','S203','AV_1P_18','AV_3P_54']:
        try:
            b=doc.Blocks.Item(name)
            counts=Counter((o.ObjectName,o.Color,o.Layer) for o in b)
            print(label,name,b.Count,counts)
        except Exception as e: print(label,name,str(e))
