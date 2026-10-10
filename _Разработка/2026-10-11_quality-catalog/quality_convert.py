import pathlib,sys,json,math,re,copy,io,collections,time
P=pathlib.Path(__file__).parent;sys.path.insert(0,str(P.parent));sys.path.insert(0,str(P.parent/'python'))
import convert as base
import ezdxf
from svgelements import SVG,Shape,SVGText,Image as SVGImage,Point,Matrix,Color
from shapely.geometry import Polygon,LineString,Point as SPoint
from shapely.ops import unary_union,polygonize
from shapely import make_valid
from matplotlib.textpath import TextPath
from matplotlib.font_manager import FontProperties
from PIL import ImageFont
from functools import lru_cache
NS=base.NS

def polys(g):
 if g.is_empty:return []
 if g.geom_type=='Polygon':return [g]
 if hasattr(g,'geoms'):return [p for q in g.geoms for p in polys(q)]
 return []

def region(loops,evenodd=True):
 rings=[]
 for points in loops:
  if len(points)>=3:
   points=list(points)
   if points[0]!=points[-1]:points.append(points[0])
   rings.append(points)
 if not rings:return Polygon()
 if len(rings)==1:
  return make_valid(Polygon(rings[0]))
 candidates=list(polygonize(unary_union([LineString(x) for x in rings])))
 selected=[]
 for g in candidates:
  q=g.representative_point();winding=0;crossings=0
  for ring in rings:
   for (x1,y1),(x2,y2) in zip(ring,ring[1:]):
    if (y1<=q.y<y2) or (y2<=q.y<y1):
     cross=(x2-x1)*(q.y-y1)-(q.x-x1)*(y2-y1)
     if (y2>y1 and cross>0) or (y2<y1 and cross<0):
      crossings+=1;winding+=1 if y2>y1 else -1
  if (crossings%2) if evenodd else winding!=0:selected.append(g)
 return unary_union(selected) if selected else Polygon()

def clips(root):
 ids={e.get('id'):e for e in root.iter() if e.get('id')};n=0
 def walk(node,ancestors):
  nonlocal n
  if node.tag.split('}')[-1] in ['defs','clipPath','mask']:return
  value=node.get('clip-path')
  if not value:
   m=re.search(r'clip-path\s*:\s*([^;]+)',node.get('style',''));value=m[1] if m else ''
  mask=node.get('mask')
  if not mask:
   m=re.search(r'mask\s*:\s*([^;]+)',node.get('style',''));mask=m[1] if m else ''
  matches=[(re.search(r'url\(#([^)]+)\)',value),False),(re.search(r'url\(#([^)]+)\)',mask),True)]
  active=list(ancestors);children=list(node)
  for match,is_mask in matches:
   if not match or match[1] not in ids:continue
   n+=1;key=str(n);active.append(key)
   target=ids[match[1]]
   sentinel=copy.deepcopy(target);sentinel.tag=NS+'g';sentinel.attrib.pop('id',None)
   for sub in sentinel.iter():
    sub.attrib.pop('id',None);sub.attrib.pop('clip-path',None)
    sub.set('data-cad-clipshape',key);sub.set('data-cad-clips','')
    if sub.tag==NS+'use':
     if 'x' not in sub.attrib:sub.set('x','0')
     if 'y' not in sub.attrib:sub.set('y','0')
    if is_mask:sub.set('data-cad-maskshape','1')
    else:sub.set('fill','black');sub.set('stroke','none');sub.set('style','fill:black;stroke:none')
   node.append(sentinel)
  node.set('data-cad-clips',','.join(active))
  for sub in children:walk(sub,active)
 walk(root,[])
 return SVG.parse(io.BytesIO(base.ET.tostring(root)),on_error='raise')

@lru_cache(maxsize=40)
def fontfile(family,weight,style):
 family=family.lower();bold=weight in ['bold','bolder'] or (weight.isdigit() and int(weight)>=600);italic=style in ['italic','oblique']
 if 'calibri' in family:f='calibri'+('z' if bold and italic else 'b' if bold else 'i' if italic else '')+'.ttf'
 elif 'bahnschrift' in family:f='bahnschrift.ttf'
 elif 'times' in family:f='times'+('bi' if bold and italic else 'bd' if bold else 'i' if italic else '')+'.ttf'
 else:f='arial'+('bi' if bold and italic else 'bd' if bold else 'i' if italic else '')+'.ttf'
 return 'C:/Windows/Fonts/'+f

@lru_cache(maxsize=50000)
def glyph(text,font):
 tp=TextPath((0,0),text,size=1,prop=FontProperties(fname=font))
 # Scale before flattening for small curves; ~0.002 em resolution.
 from matplotlib.transforms import Affine2D
 return [x/64 for x in tp.to_polygons(transform=Affine2D().scale(64),closed_only=True)]

def hatch(block,g,color,opacity=1):
 for poly in polys(g):
  if poly.area<1e-8:continue
  h=block.add_hatch(dxfattribs={'true_color':color,'layer':'0'})
  if opacity<1:h.transparency=1-opacity
  h.paths.add_polyline_path(list(poly.exterior.coords)[:-1],is_closed=True,flags=1)
  for hole in poly.interiors:h.paths.add_polyline_path(list(hole.coords)[:-1],is_closed=True,flags=0)

def opacity(e,kind):
 try:return max(0,min(1,float(e.values.get('opacity',1))*float(e.values.get(kind+'-opacity',1))))
 except:return 1

def gradients(raw,root):
 specs={};styles={};rawids={e.get('id'):e for e in raw.iter() if e.get('id')}
 for e in raw.iter():
  if e.tag in [NS+'radialGradient',NS+'linearGradient']:
   stops=[]
   for child in e:
    if child.tag!=NS+'stop':continue
    st=child.get('style','');m=re.search(r'stop-color\s*:\s*([^;]+)',st);v=child.get('stop-color') or (m[1] if m else '#000')
    off=child.get('offset','0');off=float(off.rstrip('%'))/(100 if '%' in off else 1);c=Color(v)
    stops.append((off,(c.red,c.green,c.blue)))
   if stops:specs[e.get('id')]={'kind':e.tag.split('}')[-1],'attrs':e.attrib,'stops':sorted(stops)}
  if e.tag==NS+'style':
   for selectors,body in re.findall(r'([^{}]+)\{([^{}]+)\}',e.text or ''):
    m=re.search(r'fill\s*:\s*url\(#([^)]+)\)',body)
    if m:
     for selector in selectors.split(','):styles[selector.strip()]=m[1]
 for e in root.iter():
  original=rawids.get(e.get('id'))
  if original is None:original=e
  paint=original.get('fill','');m=re.search(r'fill\s*:\s*([^;]+)',original.get('style',''))
  if m:paint=m[1]
  match=re.search(r'url\(#([^)]+)\)',paint);gid=match[1] if match else next((styles['.'+c] for c in reversed(original.get('class','').split()) if '.'+c in styles),None)
  if gid in specs:
   color=specs[gid]['stops'][0][1];value='#%02x%02x%02x'%color
   e.set('data-cad-gradient',gid);e.set('fill',value);e.set('style',e.get('style','')+';fill:'+value)
  elif paint and 'url(' not in paint:e.set('data-cad-gradient','')
 return specs

def gradient_fill(block,g,e,spec,xy,alpha):
 attrs=spec['attrs'];stops=spec['stops'];gm=Matrix(attrs.get('gradientTransform',''));em=Matrix(e.values.get('transform',''))
 def number(k,default):
  value=attrs.get(k,str(default));return float(value.rstrip('%'))/(100 if '%' in value else 1)
 def color(t):
  if t<=stops[0][0]:c=stops[0][1]
  elif t>=stops[-1][0]:c=stops[-1][1]
  else:
   for (a,ca),(b,cb) in zip(stops,stops[1:]):
    if a<=t<=b:
     f=(t-a)/(b-a);c=tuple(round(x+(y-x)*f) for x,y in zip(ca,cb));break
  return (c[0]<<16)|(c[1]<<8)|c[2]
 if attrs.get('gradientUnits','objectBoundingBox')=='objectBoundingBox':
  # Most gradients in these exports use explicit user coordinates. Relative
  # gradients are mapped through the original geometry's local bounding box.
  # CAD g is already normalized; relative gradients fall back to its CAD box.
  xmin,ymin,xmax,ymax=g.bounds
  if spec['kind']=='radialGradient':
   cx=xmin+number('cx',.5)*(xmax-xmin);cy=ymax-number('cy',.5)*(ymax-ymin);r=number('r',.5)
   center=(cx,cy);u=(r*(xmax-xmin),0);v=(0,r*(ymax-ymin))
  else:
   start=(xmin+number('x1',0)*(xmax-xmin),ymax-number('y1',0)*(ymax-ymin));end=(xmin+number('x2',1)*(xmax-xmin),ymax-number('y2',0)*(ymax-ymin))
 else:
  def point(x,y):return xy(Point(x,y)*gm*em)
  if spec['kind']=='radialGradient':
   cx=number('cx',.5);cy=number('cy',.5);r=number('r',.5);center=point(cx,cy);p=point(cx+r,cy);q=point(cx,cy+r);u=(p[0]-center[0],p[1]-center[1]);v=(q[0]-center[0],q[1]-center[1])
  else:start=point(number('x1',0),number('y1',0));end=point(number('x2',1),number('y2',0))
 if spec['kind']=='radialGradient':
  from shapely.affinity import affine_transform
  hatch(block,g,color(1),alpha)
  # Paint nested ellipses outside-to-inside. They remain portable vector fills.
  for i in range(32,0,-1):
   t=i/32;circle=SPoint(0,0).buffer(t,quad_segs=20);ellipse=affine_transform(circle,[u[0],v[0],u[1],v[1],center[0],center[1]])
   hatch(block,g.intersection(ellipse),color((i-.5)/32),alpha)
 else:
  vx=end[0]-start[0];vy=end[1]-start[1];length=math.hypot(vx,vy)
  if length<1e-9:hatch(block,g,color(.5),alpha);return
  ux=vx/length;uy=vy/length;radius=max(g.bounds[2]-g.bounds[0],g.bounds[3]-g.bounds[1],length)*4
  for i in range(-1,33):
   a=-radius if i==-1 else i/32*length;b=radius if i==32 else (i+1)/32*length
   strip=Polygon([(start[0]+ux*t-uy*s,start[1]+uy*t+ux*s) for t,s in [(a,-radius),(b,-radius),(b,radius),(a,radius)]])
   hatch(block,g.intersection(strip),color((i+.5)/32),alpha)

def convert(row,doc):
 f=pathlib.Path(row['file']);root,_,notes=base.clean(f)
 # Keep binary masks as clipping regions. The old cleaner removed all masks.
 raw=base.ET.parse(f).getroot();gradient_specs=gradients(raw,root);defs=base.ET.SubElement(root,NS+'defs')
 for mask in raw.iter(NS+'mask'):defs.append(copy.deepcopy(mask))
 def wrap(node):
  for child in list(node):
   tag=child.tag.split('}')[-1]
   if tag in ['text','path','rect','circle','ellipse','polygon','polyline','image','use'] and (child.get('mask') or child.get('clip-path') or re.search(r'(?:mask|clip-path)\s*:',child.get('style',''))):
    wrapper=base.ET.Element(NS+'g')
    for k in ['mask','clip-path','transform']:
     if child.get(k) is not None:wrapper.set(k,child.attrib.pop(k))
    style=child.get('style','')
    for k in ['mask','clip-path']:
     m=re.search(k+r'\s*:\s*([^;]+)',style)
     if m:wrapper.set(k,m[1]);style=style[:m.start()]+style[m.end():]
    child.set('style',style);idx=list(node).index(child);node.remove(child);node.insert(idx,wrapper);wrapper.append(child)
   wrap(child)
 wrap(root)
 s=clips(root);els=list(s.elements());boxes=[];clipregions={}
 for e in els:
  if isinstance(e,Shape) and e.values.get('data-cad-clipshape'):
   pts=[[tuple(p) for p in ps] for ps,closed in base.loops(e,.05)]
   key=e.values['data-cad-clipshape'];g=region(pts,True)
   # Visio's binary masks describe the allowed region; referenced clip paths
   # often inherit fill:none through defs in svgelements. Their geometry still
   # defines the mask, so do not discard it based on that inherited paint.
   clipregions[key]=unary_union([clipregions[key],g]) if key in clipregions else g
 for e in els:
  if e.values.get('data-cad-clipshape'):continue
  if isinstance(e,SVGImage):
   ps=[Point(e.x+x*e.width,e.y+y*e.height)*e.transform for x,y in [(0,0),(1,0),(1,1),(0,1)]]
   b=(min(p.x for p in ps),min(p.y for p in ps),max(p.x for p in ps),max(p.y for p in ps))
  elif isinstance(e,Shape):
   if (base.rgb(e.fill) is None or opacity(e,'fill')==0) and (base.rgb(e.stroke) is None or opacity(e,'stroke')==0):continue
   b=e.bbox()
  else:continue
  if b and all(math.isfinite(x) for x in b):boxes.append(b)
 if not boxes:raise ValueError('No geometry')
 xmin=min(b[0] for b in boxes);ymin=min(b[1] for b in boxes);xmax=max(b[2] for b in boxes);ymax=max(b[3] for b in boxes)
 dx=xmax-xmin;dy=ymax-ymin;w0=row['width'];h0=row['height']
 # Independent X/Y scaling was distorting the SVG. Always use one scale.
 factor=(w0/h0)/(dx/dy) if w0>0 and h0>0 and dx>1e-9 and dy>1e-9 else 0
 corrected=factor==0 or abs(math.log(factor))>math.log(1.05)
 if corrected:
  scale=25.4/96;status='Размер из SVG; проверить перед рабочей компоновкой'
  notes.append('Пропорции исправлены: ошибочные/неполные габариты каталога заменены размером SVG (96 px/дюйм)')
 else:
  scale=min(w0/dx,h0/dy);status='Габарит каталога, пропорции SVG сохранены'
 w=dx*scale;h=dy*scale
 def xy(p):return ((p.x-xmin)*scale,(ymax-p.y)*scale)
 def transform_geom(g):
  from shapely.affinity import affine_transform
  return affine_transform(g,[scale,0,0,-scale,-xmin*scale,ymax*scale])
 def clip(g,e):
  for key in e.values.get('data-cad-clips','').split(','):
   if key in clipregions:g=g.intersection(transform_geom(clipregions[key]))
  return g
 block=doc.blocks.new(row['name']);block.block_record.dxf.units=4;texts=[]
 for e in els:
  if e.values.get('data-cad-clipshape'):continue
  if isinstance(e,SVGImage):base.raster_trace(e,block,xy,notes)
  elif isinstance(e,Shape):
   fill=base.rgb(e.fill);stroke=base.rgb(e.stroke);fo=opacity(e,'fill');so=opacity(e,'stroke')
   if (fill is None or fo==0) and (stroke is None or so==0):continue
   paths=base.loops(e,.012/max(scale,1e-9));loops=[[xy(p) for p in ps] for ps,closed in paths if len(ps)>=3]
   if fill is not None and fo>0 and loops:
    g=clip(region(loops,e.values.get('fill-rule','nonzero')=='evenodd'),e)
    gid=e.values.get('data-cad-gradient')
    if gid in gradient_specs and not g.is_empty:gradient_fill(block,g,e,gradient_specs[gid],xy,fo)
    else:hatch(block,g,fill,fo)
   if stroke is not None and so>0:
    sw=float(e.stroke_width or 0)*scale
    # SVG parser already applies the geometry transform to stroke_width.
    for ps,closed in paths:
     points=[xy(p) for p in ps]
     if len(points)<2:continue
     if closed and points[0]!=points[-1]:points.append(points[0])
     g=LineString(points);dash=base.nums(e.values.get('stroke-dasharray',''))
     if dash and sum(dash)>0:
      from shapely.ops import substring
      dash=[max(v*scale,.001) for v in dash];dash=dash*2 if len(dash)%2 else dash
      at=0;idx=0;pieces=[]
      while at<g.length:
       end=min(at+dash[idx%len(dash)],g.length)
       if idx%2==0:pieces.append(substring(g,at,end))
       at=end;idx+=1
      geometries=pieces
     else:geometries=[g]
     for line in geometries:
      # Filled stroke geometry gives faithful width, caps and joins, and clipping.
      cap={'round':1,'butt':2,'square':3}.get(e.values.get('stroke-linecap','butt'),2)
      join={'round':1,'miter':2,'bevel':3}.get(e.values.get('stroke-linejoin','miter'),2)
      if sw>1e-5:hatch(block,clip(line.buffer(sw/2,cap_style=cap,join_style=join),e),stroke,so)
      else:
       block.add_lwpolyline(points,close=closed,dxfattribs={'true_color':stroke,'layer':'0'})
  elif isinstance(e,SVGText) and e.text and e.text.strip():
   color=base.rgb(e.fill);alpha=opacity(e,'fill')
   if color is None or alpha<=0:continue
   att=e.values.get('attributes',{});xs=base.nums(att.get('x',e.values.get('x','0'))) or [0];ys=base.nums(att.get('y',e.values.get('y','0'))) or [0]
   fs=float(e.font_size);font=fontfile(e.font_family or 'Arial',str(e.values.get('font-weight','400')),e.values.get('font-style','normal'))
   pil=ImageFont.truetype(font,1024);fulladvance=pil.getlength(e.text)/1024*fs
   anchor=-fulladvance/2 if e.anchor=='middle' else -fulladvance if e.anchor=='end' else 0
   for j,ch in enumerate(e.text):
    if ch.isspace():continue
    px=xs[j] if len(xs)>1 and j<len(xs) else xs[0]+pil.getlength(e.text[:j])/1024*fs
    py=ys[j] if len(ys)>1 and j<len(ys) else ys[0]
    loops=[]
    for poly in glyph(ch,font):
     loops.append([xy(Point(px+e.dx+anchor+u*fs,py+e.dy-v*fs)*e.transform) for u,v in poly])
    hatch(block,clip(region(loops,False),e),color,alpha)
   texts.append(e.text)
 if not len(block):raise ValueError('Empty block')
 block.set_redraw_order([(e.dxf.handle,e.dxf.handle) for e in block])
 if gradient_specs:
  notes=[n for n in notes if 'Градиент' not in n];notes.append('Градиенты переданы векторной заливкой: 32 градации')
 return {**row,'width':w,'height':h,'nominal_width':w0,'nominal_height':h0,'dimension_status':status,'dimension_corrected':corrected,'notes':sorted(set(notes)), 'entities':len(block),'text_content':' '.join(texts)}

def main():
 rows=json.loads((P/'inventory_v1.json').read_text(encoding='utf8'))
 if '--sample' in sys.argv:
  ids=json.loads((P/'sample_ids.json').read_text());rows=[r for r in rows if r['id'] in ids]
 doc=ezdxf.new('R2018');doc.units=4;out=[];errors=[];t=time.time()
 for i,r in enumerate(rows):
  try:out.append(convert(r,doc))
  except Exception as ex:errors.append({'id':r['id'],'error':str(ex)});print('ERROR',r['id'],ex,flush=True)
  if i%40==0:print('CONVERT',i+1,len(rows),'sec',round(time.time()-t),flush=True)
 suffix='_sample' if '--sample' in sys.argv else ''
 (P/f'inventory_v2{suffix}.json').write_text(json.dumps(out,ensure_ascii=False,indent=2),encoding='utf8')
 (P/f'errors_v2{suffix}.json').write_text(json.dumps(errors,ensure_ascii=False,indent=2),encoding='utf8')
 doc.saveas(P/f'blocks_v2{suffix}.dxf');print('DONE',len(out),'errors',len(errors),'seconds',round(time.time()-t),flush=True)
if __name__=='__main__':main()
