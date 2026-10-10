import pathlib,json,re,collections,csv
P=pathlib.Path(__file__).parent
data=(P/'native_inventory.txt').read_bytes()
try:text=data.decode('utf-8-sig')
except UnicodeDecodeError:text=data.decode('cp1251')
native={m.group(1):int(m.group(2)) for m in re.finditer(r'\("([^\"]+)" (\d+)\)',text)}
old=json.loads(pathlib.Path(r'C:\Users\marat\Desktop\odnolineyka\_Разработка\2026-10-09_panel-layout\actual_library_inventory.json').read_text(encoding='utf-8'))
rows=json.loads((P/'svg_inventory.json').read_text(encoding='utf-8'));allrows=json.loads((P/'catalog_all.json').read_text(encoding='utf-8'))
missing=[r['name'] for r in old if r['name'] not in native]
changed=[(r['name'],r['count'],native.get(r['name'])) for r in old if native.get(r['name'])!=r['count']]
newmissing=[r['name'] for r in rows if r['name'] not in native]
model=(P/'native_model.txt').read_text();views=json.loads((P/'catalog_views.json').read_text(encoding='utf-8'))
print('NATIVE',len(native),'OLD',len(old),'MISSING',missing,'COUNT CHANGES',changed[:25],'SVG MISSING',newmissing,flush=True)
print(model,flush=True)
assert not missing and not newmissing
assert len([n for n in native if n.startswith('SVG_')])==2542
svgcountdiff=[(r['name'],r['entities'],native.get(r['name'])) for r in rows if r['entities']!=native.get(r['name'])]
print('SVG ENTITY COUNTS',len(svgcountdiff),svgcountdiff[:10],flush=True)
assert not svgcountdiff
assert 'INSERTS=3440' in model and f'VIEWS={len(views)}' in model
assert 'Всего найдено ошибок: 0, исправлено: 0' in (P/'merge.log').read_text(encoding='utf-8')
with (P.parent/'outputs/Каталог_блоков.csv').open('w',encoding='utf-8-sig',newline='') as f:
 fields=['type','manufacturer','label','name','source','width','height','notes'];wr=csv.DictWriter(f,fieldnames=fields,delimiter=';',extrasaction='ignore');wr.writeheader()
 for r in allrows:wr.writerow({**r,'notes':' | '.join(r['notes'])})
stats={'svg_blocks':2542,'original_definitions':len(old),'original_count_changes':changed,'catalog_rows':len(allrows),'model_inserts':3440,'views':len(views),'types':len({r['type'] for r in allrows}),'raster_trace_blocks':sum(any('Растровый' in n for n in r['notes']) for r in rows),'texture_blocks':sum(any('Малая растровая' in n for n in r['notes']) for r in rows),'gradient_blocks':sum(any('Градиенты' in n for n in r['notes']) for r in rows),'derived_dimension_blocks':sum(any('Линейный' in n for n in r['notes']) for r in rows),'audit_errors':0}
(P/'validation.json').write_text(json.dumps(stats,ensure_ascii=False,indent=2),encoding='utf-8')
print(stats,flush=True)
