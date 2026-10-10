from pathlib import Path
import importlib.util,json,os,hashlib,sys,pythoncom
sys.stdout.reconfigure(encoding="utf-8",line_buffering=True)
root=Path.cwd();spec=importlib.util.spec_from_file_location("kj",root/"2_AutoCAD/Кабельный_журнал/cable_journal.py");kj=importlib.util.module_from_spec(spec);spec.loader.exec_module(kj)
def row(*v):return list(v)+[None]*(13-len(v))
values=[row(),row("FIRST","ЩA","Насос","ВВГ",None,None,None,10),row("SECOND","ЩB","Насос 2","ВВГ",None,None,None,20),row("Всего",None,None,30)]
formulas=[[""]*13 for _ in values]
r=kj.parse_rows(values,formulas,"fixture.xlsx",skip_first=True)
assert r["cable_count"]==1 and r["summary_count"]==1 and r["rows"][0][0]=="SECOND"
assert r["excluded_first"]=={"mark":"FIRST","excel_row":6}
assert r["rows"][-1][3]=="30"
r=kj.parse_rows(values,formulas,"fixture.xlsx",panel="ЩB",skip_first=True)
assert r["cable_count"]==1 and r["rows"][0][0]=="SECOND"
r=kj.parse_rows([row("Итого"),*values],[[""]*13,*formulas],"fixture.xlsx",skip_first=True)
assert r["excluded_first"]["mark"]=="FIRST" and r["summary_count"]==2
print("PASS exclusion: first actual cable, blank leading rows, summaries unchanged, panel filter applied after exclusion")
cfg=json.loads((Path(os.environ["APPDATA"])/"CableJournal/settings.json").read_text(encoding="utf-8"));source=Path(cfg["source"]);sha=hashlib.sha256(source.read_bytes()).hexdigest()
pythoncom.CoInitialize()
try:
 full=kj.read_books([source])[0];skipped=kj.read_books([source],skip_first=True)[0]
 assert skipped["rows"]==full["rows"][1:]
 assert skipped["cable_count"]==full["cable_count"]-1==18 and skipped["summary_count"]==5
 first=full["rows"][0][0];second=full["rows"][1][0]
 assert skipped["excluded_first"]["mark"]==first and skipped["excluded_first"]["excel_row"]==5
 app=kj.autocad();out=root/"_Разработка/2026-10-09_cable-journal/Проверка_v1.4/КЖ_проверка.dwg"
 kj.export_book(app,full,cfg["template"],out)
 db=kj.dbx_open(app,out)
 for tb in kj.items(db.ModelSpace):
  if tb.ObjectName!="AcDbTable":continue
  for i in range(tb.Rows):
   if kj.clean_mtext(tb.GetText(i,0)) in (first,second):tb.SetText(i,12,"Проверено 1.4")
 modified=out.with_name("manual.dwg");db.SaveAs(str(modified));db=None;kj.gc.collect();kj.os.replace(modified,out)
 result=kj.export_book(app,skipped,cfg["template"],out)
 actual,_=kj.saved_actual(app,out)
 assert first not in actual and kj.clean_mtext(actual[second][4])=="Проверено 1.4"
 assert result["removed_cables_actual_archived"][first] and Path(result["backup"]).is_file()
 assert result["skip_first"] and result["excluded_first"]["mark"]==first
 assert hashlib.sha256(source.read_bytes()).hexdigest()==sha
 (root/"_Разработка/2026-10-09_cable-journal/verification_v14.json").write_text(json.dumps(result,ensure_ascii=False,indent=2),encoding="utf-8")
 print("PASS actual Excel and DWG create/update: excluded",first,"; 18 lines + 5 summaries; other manual fields retained, excluded manual fields archived; Excel unchanged")
finally:pythoncom.CoUninitialize()
