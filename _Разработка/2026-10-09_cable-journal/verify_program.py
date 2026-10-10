from pathlib import Path
import importlib.util,sys,pythoncom,json,time
sys.stdout.reconfigure(encoding='utf-8',line_buffering=True)
sys.stderr.reconfigure(encoding='utf-8')
p=Path(__file__).resolve().parent
root=p.parents[1]
spec=importlib.util.spec_from_file_location('kj',root/'2_AutoCAD/Кабельный_журнал/cable_journal.py')
kj=importlib.util.module_from_spec(spec);spec.loader.exec_module(kj)
pythoncom.CoInitialize()
a=kj.autocad()
books=kj.read_books([root/'1_Таблицы_Excel/Объект ЖД/Однолинейка_ЩР-1.xlsx'])
report=kj.export_book(a,books[0],p/'Образец_КЖ_ЩР-1.dwg',p/'Результат/КЖ_ЩР-1.dwg')
print(json.dumps(report,ensure_ascii=False,indent=2))
