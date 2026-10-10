from verify_program import kj, p, root, a, books
import json,copy,hashlib

target=p/'Результат/КЖ_Регрессия.dwg'
source=Path(books[0]['source']) if False else root/'1_Таблицы_Excel/Объект ЖД/Однолинейка_ЩР-1.xlsx'
before=hashlib.sha256(source.read_bytes()).hexdigest()
base=copy.deepcopy(books[0])
first=base['rows'][0][0]; last=base['rows'][-1][0]
kj.export_book(a,base,p/'Образец_КЖ_ЩР-1.dwg',target)
d=kj.open_document(a,target);kj.retry(lambda:d.Activate())
for tb in kj.items(d.ModelSpace):
    if tb.ObjectName!='AcDbTable':continue
    for i in range(tb.Rows):
        mark=kj.clean_mtext(tb.GetText(i,0))
        if mark==first:
            for j,v in enumerate(['ВВГнг(А)-LS','5х16','37','Испытано','09.10.2026'],8):tb.SetText(i,j,v)
        if mark==last:tb.SetText(i,11,'Архив: удалённая линия')
kj.retry(lambda:d.Save());kj.retry(lambda:d.Close(False))
changed=copy.deepcopy(base)
changed['rows']=list(reversed(changed['rows'][:-1]))
changed['rows'][-1][7]='99'
added=copy.deepcopy(changed['rows'][0]);added[0]='TEST.ADD';changed['rows'].append(added)
r=kj.export_book(a,changed,p/'Образец_КЖ_ЩР-1.dwg',target)
assert first in r['preserved_actual']
assert last in r['removed_cables_actual_archived']
assert kj.Path(r['backup']).is_file()
actual,_=kj.saved_actual(a,target)
assert [kj.clean_mtext(v) for v in actual[first]]==['ВВГнг(А)-LS','5х16','37','Испытано','09.10.2026']
db=kj.dbx_open(a,target)
lines={kj.clean_mtext(tb.GetText(i,0)):kj.clean_mtext(tb.GetText(i,7)) for tb in db.ModelSpace
       if tb.ObjectName=='AcDbTable' and tb.Rows!=3 for i in range(tb.Rows)}
assert lines[first]=='99' and 'TEST.ADD' in lines and last not in lines
db=None
assert hashlib.sha256(source.read_bytes()).hexdigest()==before
prior=hashlib.sha256(target.read_bytes()).hexdigest()
bad=copy.deepcopy(changed);bad['rows'][0][2]='Длинная строка '*3000
try:kj.export_book(a,bad,p/'Образец_КЖ_ЩР-1.dwg',target)
except (ValueError,RuntimeError):pass
else:raise AssertionError('Expected too-tall-row rejection')
assert hashlib.sha256(target.read_bytes()).hexdigest()==prior
print('PASS: reorder/add/remove, updated project lengths, all five manual fields retained, backup, archive, rollback, source unchanged.')
(p/'verification.json').write_text(json.dumps({'passed':True,'lines':r['lines'],'pages':r['pages'],
    'update_seconds':r['seconds'],'checks':['reorder','add','remove','project change','manual fields','backup','removed data archive','failed update rollback','source SHA256 unchanged']},ensure_ascii=False,indent=2),encoding='utf-8')
