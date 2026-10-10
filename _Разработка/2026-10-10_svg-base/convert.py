import sys,pathlib,json,re,io,math,xml.etree.ElementTree as ET,base64,collections
HERE=pathlib.Path(__file__).parent;sys.path.insert(0,str(HERE/'python'))
import ezdxf
from svgelements import SVG,Shape,Path,SVGText,Image as SVGImage,Move,Close,Line,Color,Point
from PIL import Image,ImageFont
import numpy as np,cv2
ROOT=pathlib.Path(r'C:\Users\marat\Desktop\SVG база')
MAN=json.loads((ROOT/'catalog-manifest.json').read_text(encoding='utf-8'))
DEV={int(d['id']):d for d in MAN['devices']}
NS='{http://www.w3.org/2000/svg}'
def nums(s):return [float(n) for n in re.findall(r'[-+]?(?:\d*\.\d+|\d+\.?\d*)(?:[eE][-+]?\d+)?',str(s))]
def rgb(c):
 if c is None or getattr(c,'value',None) is None:return None
 return (c.red<<16)|(c.green<<8)|c.blue
def flatten(seg,tol):
 a=seg.start;b=seg.end
 if isinstance(seg,(Line,Close)):return [b]
 def recurse(t0,p0,t1,p1,n=0):
  t=(t0+t1)/2;p=seg.point(t); v=p1-p0;den=abs(v)
  samples=[seg.point(t0+(t1-t0)*k) for k in [.25,.5,.75]]
  err=max(abs((q.x-p0.x)*v.y-(q.y-p0.y)*v.x)/den if den else abs(q-p0) for q in samples)
  if n>=12 or err<=tol:return [p1]
  return recurse(t0,p0,t,p,n+1)+recurse(t,p,t1,p1,n+1)
 return recurse(0,a,1,b)
def loops(e,tol):
 p=Path(e);p.reify();out=[];pts=[];closed=False
 for seg in p:
  if isinstance(seg,Move):
   if len(pts)>1:out.append((pts,closed))
   pts=[seg.end];closed=False
  else:
   if not pts:pts=[seg.start]
   pts.extend(flatten(seg,tol));closed=isinstance(seg,Close)
 if len(pts)>1:out.append((pts,closed))
 return out
def clean(f):
 r=ET.parse(f).getroot();notes=[]
 gradients={}
 pattern_colors={};ids={e.get('id'):e for e in r.iter() if e.get('id')}
 for pat in r.iter(NS+'pattern'):
  image=None
  for sub in pat.iter():
   if sub.tag==NS+'image':image=sub;break
   if sub.tag==NS+'use':
    href=sub.get('{http://www.w3.org/1999/xlink}href',sub.get('href',''))
    target=ids.get(href.lstrip('#'))
    if target is not None and target.tag==NS+'image':image=target;break
  if image is not None:
   href=image.get('{http://www.w3.org/1999/xlink}href',image.get('href',''))
   if href.startswith('data:'):
    ar=np.asarray(Image.open(io.BytesIO(base64.b64decode(href.split(',',1)[1]))).convert('RGBA'))
    weights=ar[:,:,3].astype(float);total=weights.sum()
    if total:
     color=np.round((ar[:,:,:3]*weights[:,:,None]).sum(axis=(0,1))/total).astype(int)
     pattern_colors[pat.get('id')]='#%02x%02x%02x'%tuple(color)
 for e in r.iter():
  if e.tag.split('}')[-1] in ['radialGradient','linearGradient','pattern']:
   stops=[q for q in e.iter() if q.tag.endswith('}stop')]
   colors=[]
   for q in stops:
    c=q.get('stop-color') or re.search(r'stop-color\s*:\s*([^;]+)',q.get('style',''))
    if c:
     if not isinstance(c,str):c=c.group(1)
     try:colors.append(Color(c))
     except:pass
   if colors:gradients[e.get('id')]=str(colors[len(colors)//2])
 for parent in list(r.iter()):
  for e in list(parent):
   tag=e.tag.split('}')[-1]
   if tag in ['mask','foreignObject'] or (tag=='svg' and (e.get('width')=='0' or e.get('height')=='0')):parent.remove(e)
   elif tag in ['a','switch']:e.tag=NS+'g'
 for e in r.iter():
  if e.tag.split('}')[-1] in ['use','text','rect','image']:
   if 'x' not in e.attrib:e.set('x','0')
   if 'y' not in e.attrib:e.set('y','0')
  for k,v in list(e.attrib.items()):
   if 'url(#' in v and k in ['fill','stroke','style']:
    for gid,color in {**gradients,**pattern_colors}.items():v=v.replace('url(#'+gid+')',color)
    e.set(k,v)
 if gradients:notes.append('Градиенты заменены однотонной заливкой')
 if pattern_colors:notes.append('Малая растровая текстура в SVG pattern заменена усреднённым цветом')
 # The parser needs absolute font sizes; resolve relative sizes before parsing.
 cssfonts={}
 for stylesheet in r.iter(NS+'style'):
  for selectors,body in re.findall(r'([^{}]+)\{([^{}]+)\}',stylesheet.text or ''):
   size=re.search(r'font-size\s*:\s*([^;]+)',body)
   if size:
    for selector in selectors.split(','):cssfonts[selector.strip()]=size.group(1).strip()
 def fonts(node,parent_size=12.):
  style=node.get('style','');match=re.search(r'font-size\s*:\s*([^;]+)',style)
  classval=next((cssfonts['.'+cl] for cl in reversed(node.get('class','').split()) if '.'+cl in cssfonts),None)
  value=match.group(1) if match else classval or node.get('font-size')
  fs=parent_size
  if value:
   v=nums(value)
   if v:
    fs=v[0]*(parent_size if value.strip().endswith('em') else parent_size/100 if value.strip().endswith('%') else 96/72 if value.strip().endswith('pt') else 1)
    node.set('font-size',str(fs))
    if match:node.set('style',style[:match.start(1)]+str(fs)+style[match.end(1):])
    elif classval:node.set('style',style+';font-size:'+str(fs)+'px')
  if node.tag.endswith('}style') and node.text:
   node.text=re.sub(r'font-size\s*:\s*([.\d]+)em',lambda m:'font-size:'+str(float(m.group(1))*parent_size)+'px',node.text)
  for key,val in list(node.attrib.items()):
   if key in ['transform','dx','dy','x','y'] and 'em' in val:
    node.set(key,re.sub(r'([-+]?[.\d]+)em',lambda m:str(float(m.group(1))*fs),val))
  for child in node:fonts(child,fs)
 fonts(r)
 # SVG dy on sibling tspans is cumulative; materialize baselines before parsing.
 font=ImageFont.truetype(r'C:\Windows\Fonts\arial.ttf',512)
 def advance(text,size):return font.getlength(text or '')/512*size
 for textnode in list(r.iter(NS+'text')):
  tx=nums(textnode.get('x','0'))[0];ty=nums(textnode.get('y','0'))[0]
  fs=nums(textnode.get('font-size','12'))[0]
  cx=tx+advance(textnode.text,fs);cy=ty
  for span in list(textnode):
   if span.tag!=NS+'tspan':continue
   if span.get('x') is not None:cx=nums(span.get('x'))[0]
   if span.get('y') is not None:cy=nums(span.get('y'))[0]
   cy+=nums(span.get('dy','0'))[0];cx+=nums(span.get('dx','0'))[0]
   span.set('x',str(cx));span.set('y',str(cy));span.set('dy','0');span.set('dx','0')
   spanfs=nums(span.get('font-size',str(fs)))[0]
   cx+=advance(''.join(span.itertext()),spanfs)
   tail=span.tail;span.tail=None
   if tail and tail.strip():
    after=ET.Element(NS+'tspan',{'x':str(cx),'y':str(cy),'font-size':str(fs),'style':'font-size:'+str(fs)+'px'})
    after.text=tail;idx=list(textnode).index(span);textnode.insert(idx+1,after);cx+=advance(tail,fs)
 return r,SVG.parse(io.BytesIO(ET.tostring(r)),on_error='raise'),notes
def raster_trace(e,block,xy,notes):
 href=e.values.get('href') or e.values.get('{http://www.w3.org/1999/xlink}href')
 if not href or not href.startswith('data:'):raise ValueError('Недоступный внешний растр: '+str(href)[:80])
 im=Image.open(io.BytesIO(base64.b64decode(href.split(',',1)[1]))).convert('RGBA')
 im.thumbnail((768,768));ar=np.asarray(im);visible=ar[:,:,3]>64
 imq=im.convert('RGB').quantize(colors=16,method=Image.Quantize.MEDIANCUT)
 arr=np.asarray(imq);pal=imq.getpalette();iw,ih=im.size
 # svgelements renders x/y/width/height without applying image transform.
 def pt(x,y):return xy(Point(e.x+x/iw*e.width,e.y+y/ih*e.height)*e.transform)
 for i in np.unique(arr):
  mask=((arr==i)&visible).astype(np.uint8)
  contours,hier=cv2.findContours(mask,cv2.RETR_CCOMP,cv2.CHAIN_APPROX_SIMPLE)
  if hier is None:continue
  color=(pal[3*i]<<16)|(pal[3*i+1]<<8)|pal[3*i+2]
  for j,c in enumerate(contours):
   if hier[0][j][3]!=-1 or cv2.contourArea(c)<.5:continue
   h=block.add_hatch(dxfattribs={'true_color':int(color),'layer':'0'})
   contour=cv2.approxPolyDP(c,.3,True).reshape(-1,2)
   if len(contour)<3:block.delete_entity(h);continue
   h.paths.add_polyline_path([pt(*p) for p in contour],is_closed=True,flags=1)
   child=hier[0][j][2]
   while child!=-1:
    hole=cv2.approxPolyDP(contours[child],.3,True).reshape(-1,2)
    if len(hole)>=3:h.paths.add_polyline_path([pt(*p) for p in hole],is_closed=True,flags=0)
    child=hier[0][child][0]
 notes.append('Растровый фрагмент трассирован: до 16 цветов, до 768 px')
def convert(f,doc):
 did=int(f.name.split('_')[0]);d=DEV[did];r,s,notes=clean(f)
 els=list(s.elements());shapes=[e for e in els if isinstance(e,(Shape,SVGImage))]
 boxes=[]
 for e in shapes:
  if isinstance(e,SVGImage):
   ps=[Point(e.x+x*e.width,e.y+y*e.height)*e.transform for x,y in [(0,0),(1,0),(1,1),(0,1)]]
   b=(min(p.x for p in ps),min(p.y for p in ps),max(p.x for p in ps),max(p.y for p in ps))
  else:b=e.bbox()
  if b and all(math.isfinite(v) for v in b):boxes.append(b)
 if not boxes:raise ValueError('Нет геометрии')
 xmin=min(b[0] for b in boxes);ymin=min(b[1] for b in boxes);xmax=max(b[2] for b in boxes);ymax=max(b[3] for b in boxes)
 w=float(d['width']);h=float(d['height']);dx=xmax-xmin;dy=ymax-ymin
 scale=w/dx if w>0 and dx>1e-9 else h/dy if h>0 and dy>1e-9 else 25.4/96
 sx=w/dx if w>0 and dx>1e-9 else scale
 sy=h/dy if h>0 and dy>1e-9 else scale
 if w<=0 or h<=0 or dx<=1e-9 or dy<=1e-9:
  w=dx*sx;h=dy*sy
  notes.append('Линейный символ или отсутствующий размер: масштаб выведен из доступного габарита; при отсутствии обоих — из SVG (96 px на дюйм)')
 def xy(p):return ((p.x-xmin)*sx,(ymax-p.y)*sy)
 name='SVG_%04d_%s'%(did,re.sub(r'[<>/\\":;?*|=,`]', '_',d['name']))
 name=re.sub(r'\s+',' ',name).strip()[:240];block=doc.blocks.new(name);block.block_record.dxf.units=4
 for e in els:
  if isinstance(e,SVGImage):raster_trace(e,block,xy,notes)
  elif isinstance(e,Shape):
   fill=rgb(e.fill);stroke=rgb(e.stroke)
   if fill is None and stroke is None:continue
   paths=loops(e,.025/max(sx,sy))
   polygons=[[xy(p) for p in ps] for ps,closed in paths if len(ps)>=3]
   if fill is not None and polygons:
    hat=block.add_hatch(dxfattribs={'true_color':fill,'layer':'0'})
    hat.dxf.hatch_style=0
    # Winding orientation lets AutoCAD handle counters and disjoint outlines.
    for ps in polygons:hat.paths.add_polyline_path(ps,is_closed=True,flags=0)
   if stroke is not None:
    for ps,closed in paths:
     v=[xy(p) for p in ps]
     if len(v)>=2:block.add_lwpolyline(v,close=closed,dxfattribs={'true_color':stroke,'layer':'0','lineweight':-3})
  elif isinstance(e,SVGText) and e.text and e.text.strip():
   text=e.text;att=e.values.get('attributes',{});xs=nums(att.get('x',e.values.get('x','0')));ys=nums(att.get('y',e.values.get('y','0')))
   if not xs:xs=[0]
   if not ys:ys=[0]
   fs=float(e.font_size);a=e.transform.a;b=e.transform.b;c=e.transform.c;dd=e.transform.d
   fh=fs*math.hypot(c*sx,dd*sy)*.72
   if fh<.02:continue
   rotation=math.degrees(math.atan2(-b*sy,a*sx));style='SVG_ARIAL'
   common={'height':fh,'rotation':rotation,'style':style,'layer':'0','true_color':rgb(e.fill) or 0}
   if len(xs)>1:
    for j,ch in enumerate(text):
     if ch.isspace():continue
     px=xs[j] if j<len(xs) else xs[-1]+(j-len(xs)+1)*fs*.55
     py=ys[j] if j<len(ys) else ys[0]
     ins=xy(Point(px+e.dx,py+e.dy)*e.transform)
     block.add_text(ch,dxfattribs={**common,'insert':ins})
   else:
    ins=xy(Point(xs[0]+e.dx,ys[0]+e.dy)*e.transform)
    ent=block.add_text(text,dxfattribs={**common,'insert':ins})
    if e.anchor in ['middle','end']:
     ent.dxf.halign=1 if e.anchor=='middle' else 2;ent.dxf.align_point=ins
 if len(block)==0:raise ValueError('Пустой блок')
 block.set_redraw_order([(e.dxf.handle,e.dxf.handle) for e in block])
 parts=f.relative_to(ROOT).parts
 return {'id':did,'name':name,'label':d['name'],'type':parts[0],'manufacturer':parts[1] if len(parts)>2 else 'Не указан','series':' / '.join(parts[2:-1]),'width':w,'height':h,'source':'SVG','file':str(f),'notes':sorted(set(notes)),'entities':len(block)}
def main():
 doc=ezdxf.new('R2018');doc.units=4;doc.styles.new('SVG_ARIAL',dxfattribs={'font':'arial.ttf'})
 rows=[];errors=[];files=sorted(ROOT.rglob('*.svg'))
 if '--sample' in sys.argv:files=files[:12]
 for i,f in enumerate(files):
  try:rows.append(convert(f,doc))
  except Exception as ex:errors.append({'file':str(f),'error':str(ex)});print('ERROR',f.name,str(ex),flush=True)
  if i%50==0:print('CONVERT',i+1,'/',len(files),'OK',len(rows),'ERRORS',len(errors),flush=True)
 (HERE/'svg_inventory.json').write_text(json.dumps(rows,ensure_ascii=False,indent=2),encoding='utf-8')
 (HERE/'svg_errors.json').write_text(json.dumps(errors,ensure_ascii=False,indent=2),encoding='utf-8')
 doc.saveas(HERE/'svg_blocks.dxf');print('SAVED',len(rows),len(errors),flush=True)
if __name__=='__main__':main()
