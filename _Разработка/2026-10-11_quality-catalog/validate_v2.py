import pathlib,json,re,math,collections,csv
P=pathlib.Path(__file__).parent;OUT=P.parent.parent/'outputs'
def read_native(name):
 data=(P/name).read_bytes()
 try:return data.decode('utf-8-sig')
 except UnicodeDecodeError:return data.decode('cp1251')
rows=json.loads((P/'catalog_v2.json').read_text(encoding='utf8'));svgs=json.loads((P/'inventory_v2.json').read_text(encoding='utf8'));pages=json.loads((P/'pages_v2.json').read_text(encoding='utf8'));views=json.loads((P/'views_v2.json').read_text(encoding='utf8'))
old=json.loads(pathlib.Path(r'C:\Users\marat\Desktop\odnolineyka\_Разработка\2026-10-09_panel-layout\actual_library_inventory.json').read_text(encoding='utf8'))
native={m[1]:int(m[2]) for m in re.finditer(r'\("([^\"]+)" (\d+)\)',read_native('native_blocks_v2.txt'))}
changes=[(r['name'],r['count'],native.get(r['name'])) for r in old if r['count']!=native.get(r['name'])]
assert not changes,changes
assert len(svgs)==2542 and all(native.get(r['name'])==r['entities'] for r in svgs)
stats={k:int(v) for k,v in (line.split('=') for line in (P/'native_stats_v2.txt').read_text().splitlines())}
assert stats['INSERTS']==len(rows)-1==3440
assert stats['VIEWS']==len(views) and stats['BROKEN_LINKS']==0 and stats['LINKS']>len(pages)*2
positions=[]
for line in read_native('native_positions_v2.txt').splitlines():
 n,h,x,y,l=line.split('|');positions.append({'name':n,'handle':h,'point':(float(x),float(y)),'layer':l})
byhandle={e['handle']:e for e in positions};byposition=collections.defaultdict(list)
for e in positions:byposition[(round(e['point'][0],3),round(e['point'][1],3))].append(e)
used=set()
for r in rows:
 if r.get('render_mode')=='list':continue
 point=r['insert']
 candidates=[byhandle[r['handle']]] if r.get('handle') else byposition[(round(point[0],3),round(point[1],3))]
 candidates=[e for e in candidates if e['handle'] not in used and math.dist(point,e['point'])<.002 and e['layer']==r['catalog_layer']]
 assert len(candidates)==1,(r['catalog_id'],r['name'],point,candidates)
 if r['source']=='SVG':assert candidates[0]['name']==r['name']
 used.add(candidates[0]['handle'])
for p in pages:
 ox,oy=p['origin'];w,h=p['width'],p['height'];boxes=[]
 for r in p['items']:
  a,b,c,d=r['card_bbox'];assert a>=ox and c<=ox+w and b>=oy-h and d<=oy-170,r['catalog_id']
  if r.get('render_mode')!='list':
   ga,gb,gc,gd=r['catalog_bbox'];assert ga>=a and gc<=c and gb>=b+30 and gd<=d-45,r['catalog_id']
  for aa,bb,cc,dd in boxes:assert c<=aa or a>=cc or d<=bb or b>=dd,(r['catalog_id'],'overlap')
  boxes.append((a,b,c,d))
nav=dict(line.split('=') for line in (P/'nav_test.txt').read_text().splitlines());assert nav['S0001']=='1';assert int(nav['ABB'])==sum(r['manufacturer']=='ABB' for r in rows)
log=(P.parent/'verify_v2.log').read_text(encoding='utf8');assert 'VERIFY_V2_DONE' in log and 'Всего найдено ошибок: 0, исправлено: 0' in log
result={'svg_blocks':len(svgs),'old_named_definitions':len(old),'old_definition_changes':changes,'catalog_rows':len(rows),'pages':len(pages),'types':len(set(r['type'] for r in rows)),'manufacturer_layers':len(set(r['catalog_layer'] for r in rows)),'corrected_dimensions':sum(r.get('dimension_corrected',False) for r in svgs),'gradient_blocks':sum(any('Градиенты' in n for n in r['notes']) for r in svgs),'simplified_texture_blocks':sum(any('текстура' in n for n in r['notes']) for r in svgs),'native':stats,'checked_positions':len(used),'audit_errors':0,'visual_sample_count':26,'gui_dialog_tested':False}
(P/'validation_v2.json').write_text(json.dumps(result,ensure_ascii=False,indent=2),encoding='utf8')
print(json.dumps(result,ensure_ascii=False,indent=2),flush=True)
