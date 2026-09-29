import sys, matplotlib; matplotlib.use("Agg")
import matplotlib.pyplot as plt
from ezdxf.addons.drawing import RenderContext, Frontend
from ezdxf.addons.drawing.matplotlib import MatplotlibBackend
from ezdxf.addons.drawing.config import Configuration
exec(open("load.py").read())
def render(vph, out, x0=None,y0=None,x1=None,y1=None, dpi=200, size=(20,14)):
    vp=None
    for lay in doc.layouts:
        for v in lay.query("VIEWPORT"):
            if v.dxf.handle==vph: vp=v
    t=vp.dxf.view_target_point; c=vp.dxf.view_center_point
    h=vp.dxf.view_height; w=h*vp.dxf.width/vp.dxf.height
    cx,cy=t[0]+c[0],t[1]+c[1]
    if x0 is None: x0,x1,y0,y1=cx-w/2,cx+w/2,cy-h/2,cy+h/2
    frozen=set(vp.frozen_layers)
    fig=plt.figure(figsize=size); ax=fig.add_axes([0,0,1,1])
    ctx=RenderContext(doc)
    for l in doc.layers:
        pass
    msp=doc.modelspace()
    def filt(e):
        if e.dxf.layer in frozen: return False
        lay=doc.layers.get(e.dxf.layer)
        if lay and (lay.is_frozen() or lay.is_off()): return False
        try:
            bb=None
        except: pass
        return True
    ents=[e for e in msp if filt(e)]
    Frontend(ctx, MatplotlibBackend(ax), config=Configuration(background_policy=None) if False else Configuration()).draw_entities(ents)
    ax.set_xlim(x0,x1); ax.set_ylim(y0,y1); ax.set_aspect("equal")
    fig.savefig(out,dpi=dpi); print("saved",out)
if __name__=="__main__":
    a=sys.argv; render(a[1],a[2],*(float(x) for x in a[3:7])) if len(a)>3 else render(a[1],a[2])
