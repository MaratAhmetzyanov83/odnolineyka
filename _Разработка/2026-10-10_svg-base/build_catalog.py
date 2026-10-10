import sys,pathlib,json,re,math,collections,csv,hashlib
P=pathlib.Path(__file__).parent;sys.path.insert(0,str(P/'python'))
import ezdxf
from ezdxf.enums import TextEntityAlignment
OUT=P.parent/'outputs'
def writejson(path,data):path.write_text(json.dumps(data,ensure_ascii=False,indent=2),encoding='utf-8')
def canon(n):
 return {'chint':'CHINT','dekraft':'DEKraft','meanwell':'Mean Well','mean well':'Mean Well','schneider electric':'Schneider Electric','systeme electric':'Systeme Electric','wirenboard':'Wiren Board','noname(разное)':'Без указанного производителя','прoизводитель не определён':'Производитель не определён'}.get(n.casefold(),n)
def main():
 rows=json.loads((P/'svg_inventory.json').read_text(encoding='utf-8'))+json.loads((P/'existing_inventory.json').read_text(encoding='utf-8'))
 for r in rows:
  r['manufacturer']=canon(r['manufacturer'])
  u=r['label'].upper()
  if r['source']!='SVG':
   if 'ARLIGHT' in u:r['manufacturer']='Arlight'
   if 'PSU' in u:r['type']='Блоки питания,Датчики Модули'
   if any(k in u for k in ['_MCB','ABB_MCB']):r['type']='Автоматические выключатели'
   if any(k in u for k in ['_RCBO','ABB_RCBO']):r['type']='АВДТ(Дифф.автомат)'
   if 'CONTACTOR' in u:r['type']='Контакторы,Диммеры,Реле Импульсные реле, Фото реле'
   if 'METER' in u:r['type']='Счетчики, трансформаторы'
   if 'STOPPER' in u:r['type']='Провода, Наконечники Перфоканал,Динрейки'
   if 'PAYLOAD' in u:
    r['type']='Служебные части блоков';r['render_mode']='list';r['notes'].append('Служебная сборка: определение сохранено, в каталоге показана запись без повторного изображения всей сборки')
 groups=collections.defaultdict(lambda:collections.defaultdict(list))
 for r in rows:groups[r['type']][r['manufacturer']].append(r)
 types=sorted(groups,key=lambda t:(t in ['Прочие аппараты','Служебные части блоков'],t.casefold()))
 # Local section coordinates: each card contains a 1:1 block and its caption.
 W=max(5600,max(r['width']+80 for r in rows if r.get('render_mode')!='list'));PAD=25
 sections=[];views=[]
 for ti,typ in enumerate(types,1):
  y=-180.;cards=[];headers=[];manviews=[]
  for mi,maker in enumerate(sorted(groups[typ],key=str.casefold),1):
   start=y;headers.append((maker,y));y-=70
   items=sorted(groups[typ][maker],key=lambda r:(r.get('series','').casefold(),r['label'].casefold(),r['source'],r['name']))
   x=PAD;rowh=0
   for r in items:
    cw=max(180,r['width']+2*PAD) if r.get('render_mode')!='list' else 350
    ch=max(160,r['height']+95) if r.get('render_mode')!='list' else 180
    if x+cw>W-PAD and x>PAD:x=PAD;y-=rowh+15;rowh=0
    r['local']=[x,y];r['card']=[cw,ch];cards.append(r);rowh=max(rowh,ch);x+=cw+15
   y-=rowh+90
   manviews.append({'name':f'T{ti:02}_M{mi:02}','title':typ+' / '+maker,'local_bounds':[0,y+25,W,start+35]})
  sections.append({'type':typ,'id':f'T{ti:02}','height':-y+30,'cards':cards,'headers':headers,'manviews':manviews})
 band_y=-3600.;section_x=0;maxheight=0
 for i,sec in enumerate(sections):
  if i and i%4==0:band_y-=maxheight+600;maxheight=0
  sec['origin']=[(i%4)*(W+600),band_y];maxheight=max(maxheight,sec['height'])
 doc=ezdxf.readfile(P/'svg_blocks.dxf');m=doc.modelspace()
 doc.styles.new('CAT_ARIAL',dxfattribs={'font':'arial.ttf'})
 for name,color in [('CAT_TITLE',4),('CAT_MANUFACTURER',3),('CAT_CAPTION',7),('CAT_FRAME',8),('CAT_SVG',7),('CAT_OLD',7)]:doc.layers.new(name,dxfattribs={'color':color})
 def txt(text,x,y,h=8,width=500,layer='CAT_CAPTION'):
  return m.add_mtext(text,dxfattribs={'insert':(x,y),'char_height':h,'width':width,'style':'CAT_ARIAL','layer':layer,'attachment_point':1})
 def rect(x,y,w,h,layer='CAT_FRAME'):
  m.add_lwpolyline([(x,y),(x+w,y),(x+w,y-h),(x,y-h)],close=True,dxfattribs={'layer':layer})
 txt('БАЗА БЛОКОВ КОМПОНОВКИ',0,180,90,4000,'CAT_TITLE')
 txt('Каталог по типам оборудования и производителям • 10.10.2026',0,50,28,4000)
 txt('Размеры в миллиметрах. Аппараты размещены в масштабе 1:1.\nДля вставки скопируйте блок; рамка и подпись находятся отдельно.\nПереход по каталогу: Диспетчер видов → T01, T01_M01 и т. д.\nКомандная строка: _-VIEW → _Restore → имя вида.',0,-80,20,4000)
 txt(f"SVG: {sum(r['source']=='SVG' for r in rows)} • образцы исходной базы: {sum(r['source']!='SVG' for r in rows)} • разделов: {len(sections)}",0,-250,25,4000,'CAT_TITLE')
 for i,sec in enumerate(sections):
  ix=(i%2)*2100;iy=-450-(i//2)*240
  txt(sec['id']+'  '+sec['type'],ix,iy,30,1900,'CAT_TITLE')
  makers=' • '.join(f"M{j:02} {name}" for j,name in enumerate(sorted(groups[sec['type']],key=str.casefold),1))
  txt(makers,ix,iy-85,18,1900)
 oldops=[]
 for sec in sections:
  ox,oy=sec['origin'];txt(sec['id']+'  '+sec['type'],ox,oy,40,W,'CAT_TITLE');rect(ox,oy+55,W,sec['height']+70)
  views.append({'name':sec['id'],'title':sec['type'],'bounds':[ox-40,oy-sec['height']-40,ox+W+40,oy+90]})
  for mi,(maker,y) in enumerate(sec['headers'],1):txt(f"{sec['id']}_M{mi:02}  {maker}",ox+25,oy+y,28,W-50,'CAT_MANUFACTURER')
  for v in sec['manviews']:
   a,b,c,dd=v.pop('local_bounds');v['bounds']=[ox+a,oy+b,ox+c,oy+dd];views.append(v)
  for r in sec['cards']:
   x,y=r.pop('local');x+=ox;y+=oy;cw,ch=r['card'];rect(x,y,cw,ch)
   label=re.sub(r'([{}\\])',r'\\\1',r['label'])
   txt(label,x+8,y-7,6.0,cw-16)
   # Bottom align all geometry so 1:1 dimensions are obvious and captions never overlap.
   bx=x+(cw-r['width'])/2;by=y-ch+38
   r['catalog_bbox']=[bx,by,bx+r['width'],by+r['height']]
   txt(f"{r['width']:.2f} × {r['height']:.2f} мм | {'SVG' if r['source']=='SVG' else 'DWG'}",x+8,y-ch+28,4.5,cw-16)
   if r.get('render_mode')=='list':
    txt('Служебная сборка\nОпределение сохранено в DWG',x+8,y-65,6,cw-16)
   elif r['source']=='SVG':m.add_blockref(r['name'],(bx,by),dxfattribs={'layer':'CAT_SVG'})
   else:
    dx,dy=r['offset'];r['insert']=[bx-dx,by-dy];oldops.append(r)
 # Overview is deliberately a separate named view; the drawing opens on it.
 indexbounds=[-100,-450-math.ceil(len(sections)/2)*240,4200,350]
 views.insert(0,{'name':'00_INDEX','title':'Оглавление','bounds':indexbounds})
 doc.saveas(P/'svg_catalog.dxf');writejson(P/'catalog_views.json',views);writejson(P/'catalog_all.json',rows)
 with (OUT/'Каталог_блоков.csv').open('w',encoding='utf-8-sig',newline='') as f:
  fields=['type','manufacturer','label','name','source','width','height','notes'];wr=csv.DictWriter(f,fieldnames=fields,delimiter=';',extrasaction='ignore');wr.writeheader()
  for r in rows:wr.writerow({**r,'notes':' | '.join(r['notes'])})
 # Native merge preserves all original definitions, extension dictionaries and dynamic parameters.
 def ls(s):return '"'+s.replace('\\','/').replace('"','\\"')+'"'
 def point(v):return '(list %.10f %.10f 0.0)'%(v[0],v[1])
 ops=['(setvar "CMDECHO" 0)','(setvar "OSMODE" 0)','(setvar "FILEDIA" 0)','(setvar "INSUNITS" 4)','(setvar "ATTREQ" 0)','(setvar "ATTDIA" 0)','(setvar "PICKFIRST" 1)','(command "_.-LAYER" "_UNLOCK" "*" "_THAW" "*" "_ON" "*" "")']
 keep=[r['handle'] for r in oldops if r.get('handle')]
 ops+=['(setq catkeep \''+'('+' '.join(ls(h) for h in keep)+'))','(setq catss (ssget "_X" \''+'((410 . "Model"))))','(if catss (repeat (setq cati (sslength catss)) (setq cati (1- cati) cate (ssname catss cati)) (if (not (member (cdr (assoc 5 (entget cate))) catkeep)) (entdel cate))))']
 for r in oldops:
  pt=point(r['insert'])
  if r.get('handle'):
   ops.append(f'(setq cate (handent {ls(r["handle"])})) (command "_.MOVE" cate "" "_non" (cdr (assoc 10 (entget cate))) "_non" {pt})')
  else:ops.append(f'(command "_.-INSERT" {ls(r["name"])} "_non" {pt} 1.0 1.0 0.0)')
 ops+=['(command "_.-INSERT" '+ls(str(P/'svg_catalog.dwg'))+' "_non" (list 0.0 0.0 0.0) 1.0 1.0 0.0)','(command "_.EXPLODE" (entlast) "")']
 for v in views:
  a,b,c,dd=v['bounds'];ops.append(f'(command "_.-VIEW" "_W" {ls(v["name"])} "_non" {point([a,b])} "_non" {point([c,dd])})')
 ops+=['(command "_.-VIEW" "_R" "00_INDEX")','(command "_.AUDIT" "_Y")','(command "_.QSAVE")','(setq catreport (open '+ls(str(P/'native_merge_report.txt'))+' "w"))','(write-line (strcat "NAMED_BLOCKS=" (itoa (length catkeep))) catreport)','(close catreport)','(princ "CATALOG_COMPLETE")','(princ)']
 (P/'merge.lsp').write_text('\n'.join(ops),encoding='utf-8')
 (P/'merge.scr').write_text('(load "'+(P/'merge.lsp').as_posix()+'")\n_QUIT\n_Y\n',encoding='utf-8')
 print('BUILT',len(rows),'ROWS',len(views),'VIEWS',len(oldops),'OLD OPS',flush=True)
if __name__=='__main__':main()
