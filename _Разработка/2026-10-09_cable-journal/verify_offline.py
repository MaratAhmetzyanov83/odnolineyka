"""Read/calculate sources and edit private DBX files, without controlling the UI."""
from pathlib import Path
import importlib.util,sys,hashlib,json,copy
import pythoncom
sys.stdout.reconfigure(encoding='utf-8',line_buffering=True)
sys.stderr.reconfigure(encoding='utf-8')
p=Path(__file__).resolve().parent;root=p.parents[1]
spec=importlib.util.spec_from_file_location('kj',root/'2_AutoCAD/Кабельный_журнал/cable_journal.py')
kj=importlib.util.module_from_spec(spec);spec.loader.exec_module(kj)
pythoncom.CoInitialize()
a=kj.autocad()
assert kj.documents(a)
source=root/'1_Таблицы_Excel/Объект ЖД/Однолинейка_ЩР-1.xlsx'
hash_before=hashlib.sha256(source.read_bytes()).hexdigest()
books=kj.read_books([source])
assert len(books[0]['rows'])==59
export=p/'Результат/КЖ_ЩР-1.dwg'
db=kj.dbx_open(a,export)
manual=['ВВГнг(А)-LS','5х16','37','Испытано','09.10.2026']
tables=[t for t in db.ModelSpace if t.ObjectName=='AcDbTable' and kj.clean_mtext(t.GetText(0,0))!='Маркировка кабеля']
first=min(tables,key=lambda t:t.InsertionPoint[0])
mark=kj.clean_mtext(first.GetText(0,0))
for j,v in enumerate(manual,8):first.SetText(0,j,v)
copy_path=p/'Результат/КЖ_Ручные_графы_проверка.dwg'
db.SaveAs(str(copy_path))
db=None
actual,_=kj.saved_actual(a,copy_path)
assert [kj.clean_mtext(v) for v in actual[mark]]==manual
changed=copy.deepcopy(books[0]['rows']);changed.reverse()
new=[r[:8]+[kj.clean_mtext(v) for v in actual.get(r[0],r[8:])] for r in changed]
assert next(r for r in new if r[0]==mark)[8:]==manual
pages=kj.page_chunks(changed,[8.]*59,25,200.)
assert [len(page) for page in pages]==[25,25,9]
assert [len(page) for page in kj.page_chunks(changed,[12.]*59,25,200.)]==[16,16,16,11]
try:kj.page_chunks(changed,[201.]+[8.]*58,25,200.)
except ValueError:pass
else:raise AssertionError('oversized row accepted')
# Existing DWG matches all Excel project cells, header geometry and page targets.
db=kj.dbx_open(a,export)
tables=sorted((t for t in db.ModelSpace if t.ObjectName=='AcDbTable' and kj.clean_mtext(t.GetText(0,0))!='Маркировка кабеля'),key=lambda t:t.InsertionPoint[0])
data=[[kj.clean_mtext(t.GetText(i,j)) for j in range(8)] for t in tables for i in range(t.Rows)]
assert data==[r[:8] for r in books[0]['rows']]
assert [t.Height for t in tables]==[200.,200.,72.]
layouts=sorted((l for l in db.Layouts if not l.ModelType),key=lambda l:l.TabOrder)
assert len(layouts)==3
for i,l in enumerate(layouts):
    vp=[v for v in l.Block if v.ObjectName=='AcDbViewport' and v.DisplayLocked]
    assert len(vp)==1 and vp[0].Target[0]==i*420. and vp[0].CustomScale==1.
assert hashlib.sha256(source.read_bytes()).hexdigest()==hash_before
result={'passed':True,'lines':59,'pages':3,'manual_fields':5,'project_cell_comparisons':59*8,
        'checks':['native saved DWG data equals Excel','native table heights fit frame','viewport targets and scale',
                  'offline as-built extraction all five fields','manual preservation after reorder','pagination with long rows',
                  'oversized row rejection','source SHA256 unchanged','COM proxy collection access'],
        'limitations':['Visual Windows check stopped with Escape','Full create-save-update integration of the final build not rerun after Escape']}
(p/'verification.json').write_text(json.dumps(result,ensure_ascii=False,indent=2),encoding='utf-8')
print(json.dumps(result,ensure_ascii=False,indent=2))
pythoncom.CoUninitialize()
