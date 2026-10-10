import sys,pathlib,json,re,math,collections,time
P=pathlib.Path(__file__).parent;sys.path.insert(0,str(P/'python'))
import ezdxf
from ezdxf import bbox
print('READ SOURCE DXF',flush=True);d=ezdxf.readfile(P/'original.dxf');print('READ OK',flush=True)
named={b.name:b for b in d.blocks if not b.name.startswith('*')}
cache=bbox.Cache();rows=[];seen=set();represented=set()
def effective(name):
 b=d.blocks[name]
 if not name.startswith('*'):return name
 if b.block_record.xdata:
  for app,tags in b.block_record.xdata.data.items():
   if app.lower()=='acdbblockrepbtag':
    for tag in tags:
     if tag.code==1005:
      e=d.entitydb.get(tag.value)
      if e is not None:
       if e.dxftype()=='BLOCK':return e.dxf.name
       if e.dxftype()=='BLOCK_RECORD':return e.dxf.name
 return name
def classify(n):
 u=n.upper();maker='Производитель не определён';typ='Прочие аппараты'
 if any(k in u for k in ['VPR','TUR','KRU','TTUR','CQEN','DKC','R5','PLASTRON']):maker='DKC'
 elif 'EKF' in u:maker='EKF'
 elif u.startswith('ARV') or any(k in u for k in ['SMART-','DALI-504','KNX-104','KNX-204','SR-']):maker='Arlight'
 elif any(k in u for k in ['HDL','M_','MDL64','MDA04']):maker='HDL'
 elif any(k in u for k in ['ZENNIO','ZIO','ZCL','ZDI','ZDIN','DADO']):maker='Zennio'
 elif u in ['S201','S203','DS201','DS203','F204','ABB S201','ESB40'] or 'ABB' in u:maker='ABB'
 elif u=='MR-5-1':maker='F&F'
 elif u=='KNX-20E-640':maker='Mean Well'
 if any(k in u for k in ['VPR','TUR','KRU','TTUR','КЛЕМ','КЛАМ','ШИНА']):typ='Шины,Клеммы Изоляторы,Гребенки'
 elif any(k in u for k in ['CQEN','R5','PLASTRON','ЩИТ']):typ='Щиты Наклейки, Маркировка'
 elif any(k in u for k in ['KBR','_CM_','CROSS']):typ='Распределительные блоки Кросс модуля,Подроз.,Корб'
 elif u.startswith(('DS','DA_')) or 'ДИФ' in u:typ='АВДТ(Дифф.автомат)'
 elif u.startswith('AV_') or u in ['S201','S203','BREAKER','ABB S201']:typ='Автоматические выключатели'
 elif u=='F204' or 'УЗО' in u:typ='ВДТ(УЗО)'
 elif u.startswith(('ARV','KNX-20')) or 'БП' in u:typ='Блоки питания,Датчики Модули'
 elif any(k in u for k in ['CT_','ESB','DIM','DALI','DMX','RGBW']):typ='Контакторы,Диммеры,Реле Импульсные реле, Фото реле'
 elif u=='MR-5-1':typ='УЗДП, Термостаты Модули защиты контактов'
 elif any(k in u for k in ['DIN_TH','ДИН','RAIL']):typ='Провода, Наконечники Перфоканал,Динрейки'
 elif maker in ['HDL','Zennio'] or 'KNX' in u:typ='Логические реле,АВР Программируемые реле'
 technical=bool(re.search(r'_G\d{3}$',u)) or any(k in u for k in ['SERIGRAF','ENVOLVENTE','ENCLOSURE','COLOURS','LOGO']) or n.startswith('*')
 if technical:typ='Служебные части блоков'
 return typ,maker,technical
def row(name,handle=None,ent=None):
 ename=effective(name);b=d.blocks[name];bb=bbox.extents([ent] if ent else b,fast=True,cache=cache)
 if not bb.has_data:mn=(0,0);mx=(20,20)
 else:mn=(bb.extmin.x,bb.extmin.y);mx=(bb.extmax.x,bb.extmax.y)
 if ent:offset=(mn[0]-ent.dxf.insert.x,mn[1]-ent.dxf.insert.y)
 else:offset=(mn[0]-b.block.dxf.base_point.x,mn[1]-b.block.dxf.base_point.y)
 w=mx[0]-mn[0];h=mx[1]-mn[1]
 if not all(math.isfinite(v) and abs(v)<1e6 for v in [w,h,*offset]):
  print('BAD BOUNDS',name,w,h,flush=True);w=h=100;offset=(0,0)
 typ,maker,tech=classify(ename)
 return {'name':name,'label':ename,'type':typ,'manufacturer':maker,'technical':tech,'width':max(w,1),'height':max(h,1),'source':'Исходный DWG','handle':handle,'offset':offset,'definition':ename,'notes':[]}
for i,e in enumerate(d.modelspace().query('INSERT')):
 name=e.dxf.name;eff=effective(name);key=(name,eff)
 if key in seen:continue
 seen.add(key);represented.add(eff)
 try:rows.append(row(name,e.dxf.handle,e))
 except Exception as ex:print('ERROR existing',name,str(ex),flush=True);raise
 if i%50==0:print('OLD INSERTS',i,'rows',len(rows),flush=True)
for i,(name,b) in enumerate(named.items()):
 if name in represented:continue
 rows.append(row(name));represented.add(name)
 if i%100==0:print('OLD DEFINITIONS',i,flush=True)
(P/'existing_inventory.json').write_text(json.dumps(rows,ensure_ascii=False,indent=2),encoding='utf-8')
(P/'old_definitions.json').write_text(json.dumps(sorted(named),ensure_ascii=False),encoding='utf-8')
print('EXISTING SAVED',len(rows),'DEFINITIONS',len(named),collections.Counter(r['type'] for r in rows),flush=True)
