import pathlib,json,shutil,hashlib,datetime
P=pathlib.Path(__file__).parent;OUT=P.parent.parent/'outputs';BASE=pathlib.Path(r'C:\Users\marat\Desktop\odnolineyka')
stats=json.loads((P/'validation_v2.json').read_text(encoding='utf8'))
assert stats['audit_errors']==0 and not stats['old_definition_changes'] and stats['native']['BROKEN_LINKS']==0
def sha(p):
 h=hashlib.sha256()
 with p.open('rb') as f:
  for chunk in iter(lambda:f.read(4*1024*1024),b''):h.update(chunk)
 return h.hexdigest()
targets=[BASE/'База блоков компоновки (1).dwg',BASE/'2_AutoCAD/Блоки/SX_Компоновка_пользовательские.dwg'];baseline=sha(P/'v1.dwg')
assert all(sha(t)==baseline for t in targets),'Рабочие файлы изменились во время обработки; требуется сохранить эти изменения отдельно.'
history=BASE/'_Разработка/2026-10-11_quality-catalog';history.mkdir(parents=True,exist_ok=True)
for t in targets:
 backup=history/(t.stem+'_до_исправления.dwg')
 if not backup.exists():shutil.copy2(t,backup)
 assert sha(backup)==baseline
source=P/'final_v2.dwg';digest=sha(source)
shutil.copy2(source,OUT/'База_компоновки_пополненная.dwg')
for t in targets:
 stage=t.parent/(t.stem+'_quality_incoming.dwg');shutil.copy2(source,stage);assert sha(stage)==digest;stage.replace(t)
 assert sha(t)==digest
for dest in [BASE,BASE/'2_AutoCAD/Блоки']:
 shutil.copy2(OUT/'SX_Каталог.lsp',dest/'SX_Каталог.lsp')
for filename in ['Каталог_блоков.csv','Размеры_для_проверки.csv']:
 shutil.copy2(OUT/filename,BASE/'2_AutoCAD/Блоки'/filename)
oldcsv=BASE/'2_AutoCAD/Блоки/Каталог_компоновки.csv'
if oldcsv.exists():shutil.copy2(oldcsv,history/'Каталог_компоновки_до_исправления.csv')
shutil.copy2(OUT/'Каталог_блоков.csv',oldcsv)
oldreadme=BASE/'4_Инструкции/База_блоков_компоновки.md'
if oldreadme.exists():shutil.copy2(oldreadme,history/'База_блоков_компоновки_до_исправления.md')
shutil.copy2(OUT/'Как_пользоваться_базой.md',oldreadme)
for filename in ['Как_пользоваться_базой.md','Проверка_базы.md']:
 shutil.copy2(OUT/filename,BASE/'4_Инструкции'/filename)
for filename in ['quality_convert.py','build_v2.py','navigator_template.lsp','validate_v2.py','publish_v2.py','validation_v2.json','catalog_v2.json','pages_v2.json','views_v2.json','geometry_check.json']:
 shutil.copy2(P/filename,history/filename)
for filename in ['merge_v2.log','verify_v2.log']:shutil.copy2(P.parent/filename,history/filename)
record={'sha256':digest,'size_bytes':source.stat().st_size,'previous_sha256':baseline,'targets':[str(t) for t in targets],'backup_directory':str(history),'time':datetime.datetime.now().isoformat()}
(history/'publication_v2.json').write_text(json.dumps(record,ensure_ascii=False,indent=2),encoding='utf8');print(json.dumps(record,ensure_ascii=False,indent=2),flush=True)
