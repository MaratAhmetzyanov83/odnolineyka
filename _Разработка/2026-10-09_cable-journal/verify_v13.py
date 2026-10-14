from pathlib import Path
import importlib.util,sys,json,hashlib,copy
import pythoncom
sys.stdout.reconfigure(encoding='utf-8',line_buffering=True)
sys.stderr.reconfigure(encoding='utf-8')
p=Path(__file__).resolve().parent;root=p.parents[1]
spec=importlib.util.spec_from_file_location('kj',root/'2_AutoCAD/Кабельный_журнал/cable_journal.py')
kj=importlib.util.module_from_spec(spec);spec.loader.exec_module(kj)
NA=-2146826246
def row(*v):return list(v)+[None]*(13-len(v))
lookup='=VLOOKUP(1,range,2,FALSE)'
values=[row('L.1','ЩР','Насос','ВВГ',None,None,None,10),
 row('Итого по кабельному журналу, м',None,None,None,None,'Итого по гофре, м'),
 row('Марка',NA,NA,'Длина, м',None,'Тип','Диаметр','Длина, м'),
 row('ВВГ',NA,NA,256,None,'ПВХ','Ø20',202),
 row('Всего кабель',NA,NA,256,None,'ПНД','Ø20',2),
 row(NA,NA,NA,NA,None,'Всего гофра',None,204),
 row(NA,NA,NA,NA,None,None,None,NA),
 row('Всего',None,None,NA)]
formulas=[['']*13 for _ in values]
for i in range(2,6):
 for j in [0,1,2,3]:
  if kj.is_error(values[i][j]):formulas[i][j]=lookup
formulas[-1][3]='=SUM(S5:S9)'
parsed=kj.parse_rows(values,formulas,'fixture.xlsx')
assert parsed['cable_count']==1 and parsed['summary_count']==6
assert len(parsed['rows'])==7
assert parsed['rows'][2][1:3]==['','']
assert parsed['rows'][5][0:4]==['','','','']
assert parsed['rows'][-1][3]=='#Н/Д'
# A real cable error still fails, even after an appendix header.
try:kj.parse_rows(values[:2]+[row('L.2',NA,NA,'ВВГ')],formulas[:2]+[[lookup,'','',lookup]+['']*9],'bad.xlsx')
except ValueError as e:assert 'Q7, R7' in str(e)
else:raise AssertionError('Real cable lookup error swallowed')
print('PASS parser: totals, headers, leftover lookup errors, empty formula tail, real SUM error and real cable validation.')
settings=json.loads((Path.home()/'AppData/Roaming/CableJournal/settings.json').read_text(encoding='utf-8'))
source=Path(settings['source']);sha=hashlib.sha256(source.read_bytes()).hexdigest()
pythoncom.CoInitialize()
book=kj.read_books([source])[0]
assert book['cable_count']==19 and book['summary_count']==5
assert any(r[0]=='Марка' and r[1:3]==['',''] for r in book['rows'])
assert any(r[0]=='Всего кабель' and r[3]=='256' for r in book['rows'])
assert any(r[5]=='Всего гофра' and r[7]=='204' for r in book['rows'])
assert len(book['rows'])==24
a=kj.autocad();out=Path(settings['output'])/'_Проверка_программы_1.3'/'КЖ_проверка_сети.dwg'
template=root/'2_AutoCAD/Кабельный_журнал/КЖ_шаблон.dwg'
created=kj.export_book(a,book,template,out)
first=book['rows'][0][0]
db=kj.dbx_open(a,out)
for tb in kj.items(db.ModelSpace):
 if tb.ObjectName!='AcDbTable':continue
 for i in range(tb.Rows):
  if kj.clean_mtext(tb.GetText(i,0))==first:
   for j,v in enumerate(['Проверка марки','5х16','37','Испытано','09.10.2026'],8):tb.SetText(i,j,v)
modified=out.with_name('manual.dwg');db.SaveAs(str(modified));db=None;kj.gc.collect();kj.os.replace(modified,out)
changed=copy.deepcopy(book)
changed['rows'][:19]=list(reversed(changed['rows'][:19]))
next(r for r in changed['rows'] if r[0]==first)[7]='99'
updated=kj.export_book(a,changed,template,out)
actual,_=kj.saved_actual(a,out)
assert [kj.clean_mtext(v) for v in actual[first]]==['Проверка марки','5х16','37','Испытано','09.10.2026']
assert Path(updated['backup']).is_file()
assert hashlib.sha256(source.read_bytes()).hexdigest()==sha
print('PASS full create/update on network Excel AND network DWG destination, totals preserved, five manual fields retained, backup, source unchanged.')
(p/'verification_v13.json').write_text(json.dumps({'passed':True,'source':str(source),
 'cable_lines':19,'summary_rows':5,'pages':created['pages'],'create_seconds':created['seconds'],
 'update_seconds':updated['seconds'],'output':str(out)},ensure_ascii=False,indent=2),encoding='utf-8')
