import sys,pathlib,json,re,math,collections,csv
P=pathlib.Path(__file__).parent;sys.path.insert(0,str(P.parent/'python'))
import ezdxf
OUT=P.parent.parent/'outputs'
def save(name,data):(P/name).write_text(json.dumps(data,ensure_ascii=False,indent=2),encoding='utf8')
def clean(s):return re.sub(r'\s+',' ',s.replace('\ufeff','').replace('\u200b','')).strip(' _')
def natural(s):return [int(t) if t.isdigit() else t.casefold() for t in re.split(r'(\d+)',s)]
def canon(s):
 s=clean(s)
 return {'chint':'CHINT','dekraft':'DEKraft','meanwell':'Mean Well','mean well':'Mean Well','schneider electric':'Schneider Electric','systeme electric':'Systeme Electric','wirenboard':'Wiren Board','noname(разное)':'Не указан','прoизводитель не определён':'Не определён','производитель не определён':'Не определён','novatekelectro':'Новатек-Электро','keaz':'КЭАЗ','timvisio':'TimVisio'}.get(s.casefold(),s)
TYPES={'АВДТ(Дифф.автомат)':'Дифференциальные автоматы','ВДТ(УЗО)':'УЗО','Выкл. нагрузки Рубильники Держ. Предохран':'Выключатели нагрузки и предохранители','Звонки модульные Розетки':'Звонки и розетки','Реле напряжений,Переключ. Фаз, Расцепители':'Реле напряжения и расцепители','Реле времени,Таймер Вольт_Амперметры':'Таймеры и измерительные приборы','Распределительные блоки Кросс модуля,Подроз.,Корб':'Распределительные блоки и коробки','Сигнальные лампы, Кнопки Индикаторы модульные':'Кнопки и сигнальные лампы','Счетчики, трансформаторы':'Счётчики и трансформаторы','УЗДП, Термостаты Модули защиты контактов':'УЗДП и термостаты','Макросы, шаблоны линий Кабельный журнал':'Условные обозначения и шаблоны'}
def classify(r):
 t=r['type'];s=(r.get('series','')+' '+r['label']).casefold()
 if r.get('technical') or 'PAYLOAD' in r['label']:return 'Служебные детали'
 if t=='Контакторы,Диммеры,Реле Импульсные реле, Фото реле':
  if 'дим' in s:return 'Диммеры'
  if 'импульс' in s:return 'Импульсные реле'
  if 'фотореле' in s or 'фото реле' in s:return 'Фотореле'
  return 'Контакторы и реле'
 if t=='Блоки питания,Датчики Модули':return 'Блоки питания и датчики'
 if t=='Логические реле,АВР Программируемые реле':return 'Автоматизация и модули управления'
 if t=='Щиты Наклейки, Маркировка':
  return 'Маркировка' if any(k in s for k in ['наклейк','маркировк']) else 'Щиты и комплектующие'
 if t=='Шины,Клеммы Изоляторы,Гребенки':
  if 'клем' in s or 'клeм' in s:return 'Клеммы'
  if 'греб' in s:return 'Соединительные гребёнки'
  return 'Шины и изоляторы'
 if t=='Провода, Наконечники Перфоканал,Динрейки':
  if 'наконеч' in s or 'ншви' in s:return 'Наконечники'
  if 'провод' in s or 'перемычк' in s or 'земля' in s:return 'Провода и перемычки'
  return 'DIN-рейки и кабельные каналы'
 if t=='ОПС(УЗИП), Клипсы Преобразователи':return 'УЗИП и преобразователи'
 return TYPES.get(t,t)
def series(r):
 parts=[]
 for p in r.get('series','').split(' / '):
  p=clean(p)
  if p and p.casefold() not in [v.casefold() for v in parts]:parts.append(p)
 if parts:return ' / '.join(parts)
 if r.get('technical'):
  return re.sub(r'_G\d{3}$','',r['label'])
 return 'Исходные блоки DWG'
def ls(s):return '"'+s.replace('\\','/').replace('"','\\"')+'"'
def pt(p):return '(list %.10f %.10f 0.0)'%(p[0],p[1])

def main():
 rows=json.loads((P/'inventory_v2.json').read_text(encoding='utf8'))+json.loads((P.parent/'existing_inventory.json').read_text(encoding='utf8'))
 print('READ FINAL VECTOR BLOCKS',flush=True);doc=ezdxf.readfile(P/'blocks_v2.dxf');print('READ OK',flush=True)
 bounds=[]
 for r in rows:
  if r['source']!='SVG':continue
  xmin=ymin=math.inf;xmax=ymax=-math.inf
  for e in doc.blocks[r['name']]:
   if e.dxftype()=='HATCH':points=(v for path in e.paths if hasattr(path,'vertices') for v in path.vertices)
   elif e.dxftype()=='LWPOLYLINE':points=e.get_points('xy')
   else:continue
   for point in points:
    x,y=point[:2];xmin=min(xmin,x);xmax=max(xmax,x);ymin=min(ymin,y);ymax=max(ymax,y)
  assert all(math.isfinite(v) for v in [xmin,ymin,xmax,ymax]),r['name']
  r['geometry_bounds']=[xmin,ymin,xmax,ymax];r['display_width']=xmax-xmin;r['display_height']=ymax-ymin
  r['display_offset']=[xmin,ymin];bounds.append({'id':r['id'],'bounds':r['geometry_bounds'],'outside_body':max(-xmin,-ymin,xmax-r['width'],ymax-r['height'])})
 save('geometry_check.json',bounds)
 for i,r in enumerate(rows):
  if r['source']!='SVG':
   u=r['label'].upper()
   if 'ARLIGHT' in u:r['manufacturer']='Arlight'
   if u=='LY-K2810412J':
    r['manufacturer']='Letour';r['type']='Логические реле,АВР Программируемые реле';r['notes'].append('Производитель уточнён по LY/K2810412J: https://www.gzletour.com/h-pd-600.html')
   if u.startswith('SR-') or 'ARLIGHT_CONTROL' in u:r['type']='Логические реле,АВР Программируемые реле'
   if '_PSU' in u:r['type']='Блоки питания,Датчики Модули'
   if '_MCB' in u:r['type']='Автоматические выключатели'
   if '_RCBO' in u:r['type']='АВДТ(Дифф.автомат)'
   if '_RCD' in u:r['type']='ВДТ(УЗО)'
   if 'CONTACTOR' in u:r['type']='Контакторы,Диммеры,Реле Импульсные реле, Фото реле'
   if '_METER' in u:r['type']='Счетчики, трансформаторы'
   if 'STOPPER' in u:r['type']='Провода, Наконечники Перфоканал,Динрейки'
  r['label']=clean(r['label']);r['manufacturer']=canon(r['manufacturer']);r['type_original']=r['type'];r['type']=classify(r);r['series']=series(r);r['catalog_id']=f'S{r["id"]:04}' if r['source']=='SVG' else f'D{i-2541:04}'
  if 'PAYLOAD' in r['label']:r['render_mode']='list';r['width']=0;r['height']=0
 groups=collections.defaultdict(list)
 for r in rows:groups[(r['type'],r['manufacturer'],r['series'])].append(r)
 types=sorted(set(r['type'] for r in rows),key=lambda t:(t in ['Прочие аппараты','Служебные детали'],natural(t)))
 typeids={t:f'{i+1:02}' for i,t in enumerate(types)}
 for r in rows:
  r['catalog_layer']='LIB_'+typeids[r['type']]+'_'+re.sub(r'[<>/\\":;?*|=,]','_',r['manufacturer'])
  if r['catalog_layer'] not in doc.layers:doc.layers.new(r['catalog_layer'],dxfattribs={'color':7})
 pages=[]
 for key,items in sorted(groups.items(),key=lambda p:(types.index(p[0][0]),natural(p[0][1]),natural(p[0][2]))):
  items.sort(key=lambda r:(natural(r['label']),r['source'],r['name']))
  page=None
  for r in items:
   cw=max(215,r.get('display_width',r['width'])+40);ch=max(175,r.get('display_height',r['height'])+85)
   if page is None:page={'key':key,'items':[],'width':max(1100,cw+50),'height':max(780,ch+220),'x':25,'y':-180,'rowh':0}
   if page['x']+cw>page['width']-25 and page['x']>25:page['x']=25;page['y']-=page['rowh']+18;page['rowh']=0
   if len(page['items'])>=12 or -page['y']+ch>page['height']-50:
    if page['items']:pages.append(page)
    page={'key':key,'items':[],'width':max(1100,cw+50),'height':max(780,ch+220),'x':25,'y':-180,'rowh':0}
   r['card']=[cw,ch];r['local']=[page['x'],page['y']];page['items'].append(r);page['x']+=cw+18;page['rowh']=max(page['rowh'],ch)
  if page and page['items']:pages.append(page)
 # Fit sparse series pages to their contents instead of leaving a huge empty sheet.
 for p in pages:
  p['width']=max(450,max(r['local'][0]+r['card'][0] for r in p['items'])+25)
  p['height']=max(420,max(-r['local'][1]+r['card'][1] for r in p['items'])+60)
 # Each type has its own column. Pages follow in reading order down the column.
 m=doc.modelspace();doc.styles.new('CAT_ARIAL',dxfattribs={'font':'arial.ttf'})
 for name,color in [('CAT_TITLE',4),('CAT_NAV',3),('CAT_CAPTION',7),('CAT_FRAME',8),('CAT_QA',2),('CAT_SVG',7)]:doc.layers.new(name,dxfattribs={'color':color})
 doc.layers.get('CAT_TITLE').dxf.true_color=0x2875B5
 doc.layers.get('CAT_NAV').dxf.true_color=0x419968
 def text(value,x,y,size=8,width=1000,layer='CAT_CAPTION'):
  value=value.replace('\\','\\\\').replace('{','\\{').replace('}','\\}').replace('\n','\\P')
  return m.add_mtext(value,dxfattribs={'insert':(x,y),'char_height':size,'width':width,'style':'CAT_ARIAL','layer':layer,'attachment_point':1})
 def rect(x,y,w,h):return m.add_lwpolyline([(x,y),(x+w,y),(x+w,y-h),(x,y-h)],close=True,dxfattribs={'layer':'CAT_FRAME'})
 views=[];oldops=[];pagegroups=collections.defaultdict(list);nav=[]
 colwidth=max(2600,max(p['width'] for p in pages)+400)
 for typ in types:
  group=[p for p in pages if p['key'][0]==typ];ox=(types.index(typ)+1)*colwidth;oy=-1800
  maker_counts=collections.Counter(p['key'][1] for p in group);maker_seq=collections.Counter();maker_ids={k:f'{i+1:02}' for i,k in enumerate(sorted(maker_counts,key=natural))}
  for pi,page in enumerate(group,1):
   _,maker,ser=page['key'];maker_seq[maker]+=1
   code=f'{typeids[typ]}_{maker_ids[maker]}_{maker_seq[maker]:03}'
   page['code']=code;page['origin']=[ox,oy];w=page['width'];h=page['height'];scale=min(1.4,max(1,w/1100,h/1000))
   # Header size responds to very large enclosure pages, remaining readable.
   text(f'{typeids[typ]}  {typ}',ox+25,oy-20,21*scale,w-50,'CAT_TITLE').set_hyperlink('',typ+' — указатель',typeids[typ]+'_РАЗДЕЛ')
   text(f'{maker}  /  {ser}',ox+25,oy-65,11*scale,w-50,'CAT_NAV')
   text(f'Вид: {code}   •   {len(page["items"])} блоков   •   Страница {pi}/{len(group)}',ox+25,oy-118,8*scale,w-50)
   text('Оглавление: Ctrl+щелчок   •   Поиск: SXCAT',ox+25,oy-h+25,8*scale,w*.55,'CAT_NAV').set_hyperlink('','Оглавление','00_ОГЛАВЛЕНИЕ');rect(ox,oy,w,h)
   bounds=[ox-20,oy-h-20,ox+w+20,oy+20]
   vname=code+'_'+re.sub(r'[^\wА-Яа-яЁё -]','',maker).replace(' ','_')[:30]
   page['view']=vname;views.append({'name':vname,'bounds':bounds,'title':f'{typ} / {maker} / {ser}'});pagegroups[typ].append(page)
   for r in page['items']:
    x,y=r.pop('local');x+=ox;y+=oy;cw,ch=r['card'];rect(x,y,cw,ch)
    labelsize=max(5.5,min(12,cw/35));labelh=labelsize*1.7*(math.ceil(len(r['label'])/(max(1,int((cw-16)/(labelsize*.52))))))
    # Cards reserve 65 mm for labels, and larger cards use a proportional margin.
    text(f'{r["catalog_id"]}  {r["label"]}',x+8,y-7,labelsize,cw-16)
    dw=r.get('display_width',r['width']);dh=r.get('display_height',r['height']);bx=x+(cw-dw)/2;by=y-ch+35
    r['catalog_bbox']=[bx,by,bx+dw,by+dh];r['card_bbox']=[x,y-ch,x+cw,y];r['page']=code;r['view']=vname
    dim=f'{r["width"]:.1f} × {r["height"]:.1f} мм'
    if r.get('dimension_corrected'):dim+=' *'
    text(dim+'  |  '+('SVG' if r['source']=='SVG' else 'DWG'),x+8,y-ch+25,5.5,cw-16,'CAT_QA' if r.get('dimension_corrected') else 'CAT_CAPTION')
    if r.get('render_mode')=='list':text('Определение служебной сборки\nСохранено в DWG',x+8,y-60,7,cw-16)
    elif r['source']=='SVG':
     dx,dy=r['display_offset'];r['insert']=[bx-dx,by-dy];m.add_blockref(r['name'],r['insert'],dxfattribs={'layer':r['catalog_layer']})
    else:
     dx,dy=r['offset'];r['insert']=[bx-dx,by-dy];oldops.append(r)
    nav.append([r['catalog_id'],r['type'],r['manufacturer'],r['series'],r['label'],r['name'],r['view'],r['card_bbox'],r.get('dimension_corrected',False),r['width'],r['height'],r['source'],r.get('dimension_status','Габарит исходного DWG')])
   oy-=h+180
 # Compact global index, and a directory for every type listing all series/pages.
 text('БАЗА КОМПОНОВКИ',0,260,55,2450,'CAT_TITLE')
 text('Выберите раздел → производителя → серию → блок',0,155,21,2450,'CAT_NAV')
 text('Ctrl+щелчок по разделу — открыть его. Так же работают переходы на страницах.\nSXCAT — поиск блока (после загрузки помощника через APPLOAD).\nБлоки в мм; * у габаритов — размер восстановлен из SVG, нужна проверка.',0,95,14,2450)
 for i,typ in enumerate(types):
  x=(i%2)*1300;y=-45-(i//2)*85;count=sum(r['type']==typ for r in rows);n=len(pagegroups[typ])
  text(f'{typeids[typ]}  {typ}',x,y,25,1240,'CAT_TITLE').set_hyperlink('',typ,typeids[typ]+'_РАЗДЕЛ')
  text(f'{count} блоков • {n} страниц • вид {typeids[typ]}_РАЗДЕЛ',x+43,y-45,14,1200)
 indexbounds=[-30,-45-math.ceil(len(types)/2)*85-30,2570,360]
 views.insert(0,{'name':'00_ОГЛАВЛЕНИЕ','title':'Оглавление','bounds':indexbounds})
 for typ in types:
  plist=pagegroups[typ];ox=(types.index(typ)+1)*colwidth
  for di in range(math.ceil(len(plist)/24)):
   oy=900+di*1250;dh=1100;rect(ox,oy,2200,dh)
   text(typeids[typ]+'  '+typ,ox+25,oy-25,35,2100,'CAT_TITLE')
   text(f'Производитель / серия → вид страницы  •  Указатель {di+1}/{math.ceil(len(plist)/24)}',ox+25,oy-100,18,2100,'CAT_NAV')
   for i,p in enumerate(plist[di*24:(di+1)*24]):
    x=ox+25+(i%2)*1100;y=oy-175-(i//2)*70
    # Long original series names stay searchable; concise captions fit directory rows.
    caption=p['code']+'  '+p['key'][1]+' / '+p['key'][2]
    text(caption[:110]+('…' if len(caption)>110 else ''),x,y,13,1060).set_hyperlink('',p['key'][1]+' / '+p['key'][2],p['view'])
   if di+1<math.ceil(len(plist)/24):text('Следующий указатель →',ox+25,oy-dh+45,16,2100,'CAT_NAV').set_hyperlink('','Далее',typeids[typ]+f'_СЕРИИ_{di+2:02}')
   else:text('← Оглавление',ox+25,oy-dh+45,16,2100,'CAT_NAV').set_hyperlink('','Оглавление','00_ОГЛАВЛЕНИЕ')
   name=typeids[typ]+'_РАЗДЕЛ' if di==0 else typeids[typ]+f'_СЕРИИ_{di+1:02}'
   views.append({'name':name,'title':typ,'bounds':[ox-20,oy-dh-20,ox+2220,oy+20]})
  for maker in sorted({p['key'][1] for p in plist},key=natural):
   sub=[p for p in plist if p['key'][1]==maker]
   # A manufacturer's view points to its first series page; no unreadable full-area zoom.
   b=next(v['bounds'] for v in views if v['name']==sub[0]['view']);views.append({'name':typeids[typ]+'_ПРОИЗВ_'+re.sub(r'[^\wА-Яа-яЁё]','_',maker)[:35],'title':typ+' / '+maker,'bounds':b})
  for i,p in enumerate(plist):
   ox,oy=p['origin'];w=p['width'];h=p['height']
   if i>0:text('← Предыдущая',ox+w*.58,oy-h+25,8,w*.2,'CAT_NAV').set_hyperlink('','Предыдущая страница',plist[i-1]['view'])
   if i+1<len(plist):text('Следующая →',ox+w*.80,oy-h+25,8,w*.18,'CAT_NAV').set_hyperlink('','Следующая страница',plist[i+1]['view'])
 save('catalog_v2.json',rows);save('pages_v2.json',pages);save('views_v2.json',views);save('nav_v2.json',nav)
 print('RENDER CATALOG PREVIEWS',flush=True)
 import matplotlib
 matplotlib.use('Agg')
 import matplotlib.pyplot as plt
 from ezdxf.addons.drawing import RenderContext,Frontend
 from ezdxf.addons.drawing.matplotlib import MatplotlibBackend
 from ezdxf.addons.drawing.config import Configuration,BackgroundPolicy
 def anchor(e):
  if e.dxftype() in ['INSERT','MTEXT']:return e.dxf.insert.x,e.dxf.insert.y
  if e.dxftype()=='LWPOLYLINE':return tuple(e.get_points('xy')[0])
  return (math.inf,math.inf)
 selected=[('Оглавление_каталога.png',views[0])]
 for filename,condition in [('Пример_страницы_аппаратов.png',lambda p:all(r['source']=='SVG' for r in p['items']) and p['key'][0]=='Автоматические выключатели' and p['key'][1]=='ABB' and len(p['items'])>=3),('Пример_страницы_щитов.png',lambda p:all(r['source']=='SVG' for r in p['items']) and p['key'][0]=='Щиты и комплектующие' and p['key'][1]=='ABB' and len(p['items'])>=2)]:
  sample=next((p for p in pages if condition(p)),None)
  if sample:selected.append((filename,next(v for v in views if v['name']==sample['view'])))
 context=RenderContext(doc)
 for filename,v in selected:
  a,b,c,d=v['bounds'];subset=[e for e in m if a<=anchor(e)[0]<=c and b<=anchor(e)[1]<=d]
  fig=plt.figure(figsize=(16,16*(d-b)/(c-a)));ax=fig.add_axes([0,0,1,1]);Frontend(context,MatplotlibBackend(ax),config=Configuration(background_policy=BackgroundPolicy.WHITE)).draw_entities(subset)
  ax.set_aspect('equal',adjustable='box');ax.set_xlim(a,c);ax.set_ylim(b,d);ax.axis('off');fig.savefig(OUT/filename,dpi=130);plt.close(fig)
 print('SAVE CATALOG DXF',flush=True);doc.saveas(P/'catalog_v2.dxf')
 fields=['catalog_id','type','manufacturer','series','label','name','catalog_layer','page','view','source','width','height','dimension_status','nominal_width','nominal_height','notes']
 with (OUT/'Каталог_блоков.csv').open('w',encoding='utf-8-sig',newline='') as f:
  wr=csv.DictWriter(f,fieldnames=fields,delimiter=';',extrasaction='ignore');wr.writeheader()
  for r in rows:wr.writerow({**r,'notes':' | '.join(r['notes'])})
 ops=['(setvar "CMDECHO" 0)','(setvar "OSMODE" 0)','(setvar "FILEDIA" 0)','(setvar "INSUNITS" 4)','(setvar "ATTREQ" 0)','(setvar "ATTDIA" 0)','(setvar "LISPSYS" 1)','(setvar "TRANSPARENCYDISPLAY" 1)','(command "_.-LAYER" "_UNLOCK" "*" "_THAW" "*" "_ON" "*" "")']
 keep=[r['handle'] for r in oldops if r.get('handle')]
 ops.append('(setq catnewpairs nil)')
 ops+=['(setq catkeep \''+'('+' '.join(ls(h) for h in keep)+'))','(setq catss (ssget "_X" \''+'((410 . "Model"))))','(if catss (repeat (setq cati (sslength catss)) (setq cati (1- cati) cate (ssname catss cati)) (if (not (member (cdr (assoc 5 (entget cate))) catkeep)) (entdel cate))))']
 for r in oldops:
  point=pt(r['insert'])
  if r.get('handle'):ops.append(f'(setq cate (handent {ls(r["handle"])})) (if cate (command "_.MOVE" cate "" "_non" (cdr (assoc 10 (entget cate))) "_non" {point}))')
  else:
   ops.append(f'(command "_.-INSERT" {ls(r["name"])} "_non" {point} 1.0 1.0 0.0)')
   ops.append(f'(setq catnewpairs (cons (cons (entlast) {ls(r["catalog_layer"])}) catnewpairs))')
 ops+=['(command "_.-INSERT" '+ls(str(P/'catalog_v2.dwg'))+' "_non" (list 0.0 0.0 0.0) 1.0 1.0 0.0)','(command "_.EXPLODE" (entlast) "")']
 # Organize original sample references after the catalogue layers are imported.
 for r in oldops:
  if r.get('handle'):
   ops.append(f'(setq cate (handent {ls(r["handle"])})) (if cate (entmod (subst (cons 8 {ls(r["catalog_layer"])}) (assoc 8 (entget cate)) (entget cate))))')
 ops.append('(foreach catpair catnewpairs (entmod (subst (cons 8 (cdr catpair)) (assoc 8 (entget (car catpair))) (entget (car catpair)))))')
 for v in views:
  a,b,c,d=v['bounds'];ops.append(f'(command "_.-VIEW" "_W" {ls(v["name"])} "_non" {pt([a,b])} "_non" {pt([c,d])})')
 ops+=['(command "_.-VIEW" "_R" "00_ОГЛАВЛЕНИЕ")','(command "_.AUDIT" "_Y")','(command "_.QSAVE")','(princ "V2_COMPLETE")','(princ)']
 (P/'merge_v2.lsp').write_text('\n'.join(ops),encoding='utf8');(P/'merge_v2.scr').write_text('(load '+ls(str(P/'merge_v2.lsp'))+')\n_QUIT\n_Y\n',encoding='utf8')
 def lisp(value):
  if isinstance(value,str):return '"'+value.replace('\\','\\\\').replace('"','\\"')+'"'
  if isinstance(value,bool):return 'T' if value else 'nil'
  if isinstance(value,(list,tuple)):return '('+' '.join(lisp(v) for v in value)+')'
  return str(value)
 template=(P/'navigator_template.lsp').read_text(encoding='utf8')
 (OUT/'SX_Каталог.lsp').write_text(template.replace('@@DATA@@','(setq sx:catalog \''+lisp(nav)+')'),encoding='utf-8-sig')
 print('CATALOG',len(rows),'pages',len(pages),'types',len(types),'views',len(views),flush=True)
if __name__=='__main__':main()
