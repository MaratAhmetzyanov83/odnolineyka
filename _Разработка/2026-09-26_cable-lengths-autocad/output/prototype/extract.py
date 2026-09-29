import re, math, collections, json
from geo import *
from ezdxf.tools.text import plain_mtext
exec(open("load.py").read())
msp=doc.modelspace()
GRP=re.compile(r'^([A-Za-zА-Яа-яЁё][A-Za-zА-Яа-яЁё\-]*(?:\.\d+)+)')
TOL=60      # мм: допуск касания стрелки выноски
JTOL=60     # мм: допуск стыковки концов полилиний
# основные видовые экраны планов: окно в модели + замороженные слои
views=[]
for lay in doc.layouts:
    if lay.name=="Model": continue
    for vp in lay.query("VIEWPORT"):
        if vp.dxf.id==1: continue
        t=vp.dxf.view_target_point; c=vp.dxf.view_center_point
        h=vp.dxf.view_height; w=h*vp.dxf.width/vp.dxf.height
        if w<20000: continue
        cx,cy=t[0]+c[0],t[1]+c[1]
        views.append(dict(layout=lay.name, vp=vp.dxf.handle, win=(cx-w/2,cy-h/2,cx+w/2,cy+h/2), frozen=set(vp.frozen_layers)))
def level_of(win):
    return "Ур.2" if win[0]>60000 else "Ур.1"
def inwin(p,w): return w[0]<p[0]<w[2] and w[1]<p[1]<w[3]
layer_off={l.dxf.name for l in doc.layers if l.is_off() or l.is_frozen()}
# все полилинии трасс и выноски
polys=[]
for e in msp:
    if e.dxftype()=="LWPOLYLINE" and "Трасс" in e.dxf.layer:
        pts=[(p[0],p[1]) for p in e.get_points('xy')]
        if e.closed: pts.append(pts[0])
        lt=e.dxf.get("linetype","ByLayer")
        if lt=="ByLayer": lt=doc.layers.get(e.dxf.layer).dxf.linetype
        polys.append(dict(h=e.dxf.handle, layer=e.dxf.layer, lt=lt, pts=pts, L=pl_len(pts)))
labels=[]
for e in msp:
    t=e.dxftype()
    if t=="MULTILEADER" and e.context.mtext:
        txt=plain_mtext(e.context.mtext.default_content)
        arrows=[(ln.vertices[0].x, ln.vertices[0].y) for lr in e.context.leaders for ln in lr.lines if ln.vertices]
        m=GRP.match(txt.strip())
        if m: labels.append(dict(h=e.dxf.handle, kind="ML", layer=e.dxf.layer, grp=m.group(1), pos=(e.context.mtext.insert.x,e.context.mtext.insert.y), arrows=arrows, raw=txt.split("\n")[0]))
    elif t=="MTEXT":
        txt=plain_mtext(e.text); m=GRP.match(txt.strip())
        if m: labels.append(dict(h=e.dxf.handle, kind="MT", layer=e.dxf.layer, grp=m.group(1), pos=(e.dxf.insert.x,e.dxf.insert.y), arrows=[], raw=txt.split("\n")[0]))
