import win32com.client as w,pythoncom,sys
import time
sys.stdout.reconfigure(encoding='utf-8',line_buffering=True)
a=w.GetActiveObject('AutoCAD.Application');ds=a.Documents
d=next(ds.Item(i) for i in range(ds.Count) if ds.Item(i).Name=='КЖ_проверка.dwg')
d.Activate()
time.sleep(1)
t=d.ModelSpace.AddTable(w.VARIANT(pythoncom.VT_ARRAY|pythoncom.VT_R8,(0.,-500.,0.)),5,13,8.,30.)
print('BEFORE',t.Rows,t.Height)
t.TitleSuppressed=True;t.HeaderSuppressed=True
print('AFTER',t.Rows,t.Height)
t.SetText(0,0,'TEST');t.SetText(0,1,'COL2')
print('ROWS',[[t.GetText(i,j) for j in range(2)] for i in range(t.Rows)])
print('MERGE',t.IsMergedCell(0,0))
t.Delete()
