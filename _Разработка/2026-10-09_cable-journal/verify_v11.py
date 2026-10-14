from pathlib import Path
import importlib.util,sys,json,copy,hashlib,time
import pythoncom
sys.stdout.reconfigure(encoding='utf-8',line_buffering=True)
sys.stderr.reconfigure(encoding='utf-8')
p=Path(__file__).resolve().parent;root=p.parents[1]
spec=importlib.util.spec_from_file_location('kj',root/'2_AutoCAD/Кабельный_журнал/cable_journal.py')
kj=importlib.util.module_from_spec(spec);spec.loader.exec_module(kj)
settings=json.loads((Path.home()/'AppData/Roaming/CableJournal/settings.json').read_text(encoding='utf-8'))
source=Path(settings['source']);sha=hashlib.sha256(source.read_bytes()).hexdigest()
pythoncom.CoInitialize()
assert kj.is_summary_row(['Всего кабель',None,None,-2146826246]+[None]*9)
assert not kj.is_summary_row(['Итого.1','ЩР','Потребитель','ВВГ']+[None]*9)
assert not kj.is_summary_row(['Всего кабель','ЩР','Потребитель','ВВГ']+[None]*9)
assert not kj.is_summary_row(['R.1','ЩР','Потребитель',-2146826246]+[None]*9)
book=kj.read_books([source])[0]
assert any(r[0]=='Всего кабель' and '#Н/Д' in r for r in book['rows'])
assert book['cable_count']==19 and book['summary_count']==7
a=kj.autocad()
out=p/'Проверка_v1.1'/'КЖ_ЩC-ВК.01.dwg'
template=root/'2_AutoCAD/Кабельный_журнал/КЖ_шаблон.dwg'
created=kj.export_book(a,book,template,out)
first=book['rows'][0][0]
# Only a private generated test drawing is edited, through ObjectDBX.
db=kj.dbx_open(a,out)
for tb in kj.items(db.ModelSpace):
 if tb.ObjectName!='AcDbTable':continue
 for i in range(tb.Rows):
  if kj.clean_mtext(tb.GetText(i,0))==first:
   for j,v in enumerate(['Проверка марки','5х16','37','Испытано','09.10.2026'],8):tb.SetText(i,j,v)
modified=out.with_name('native_manual.dwg');db.SaveAs(str(modified));db=None
kj.gc.collect();kj.os.replace(modified,out)
changed=copy.deepcopy(book);changed['rows'].reverse()
next(r for r in changed['rows'] if r[0]==first)[7]='99'
updated=kj.export_book(a,changed,template,out)
actual,_=kj.saved_actual(a,out)
assert [kj.clean_mtext(v) for v in actual[first]]==['Проверка марки','5х16','37','Испытано','09.10.2026']
assert Path(updated['backup']).is_file()
db=kj.dbx_open(a,out)
lines={kj.clean_mtext(t.GetText(i,0)):kj.clean_mtext(t.GetText(i,7)) for t in kj.items(db.ModelSpace)
 if t.ObjectName=='AcDbTable' and kj.clean_mtext(t.GetText(0,0))!='Маркировка кабеля' for i in range(t.Rows)}
assert lines[first]=='99' and len(lines)==len(book['rows'])
db=None;kj.gc.collect()
assert hashlib.sha256(source.read_bytes()).hexdigest()==sha
print('PASS full create/update: totals included with visible formula errors, reordered project rows, new length, five manual fields, backup, unchanged network source.')
(p/'verification_v11.json').write_text(json.dumps({'passed':True,'source':str(source),'lines':len(book['rows']),
 'pages':created['pages'],'create_seconds':created['seconds'],'update_seconds':updated['seconds']},ensure_ascii=False,indent=2),encoding='utf-8')
