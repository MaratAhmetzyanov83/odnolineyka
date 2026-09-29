import math
from geo import seg_d
def rect_d(p,r):
    dx=max(r[0]-p[0],0,p[0]-r[2]); dy=max(r[1]-p[1],0,p[1]-r[3]); return math.hypot(dx,dy)
def seg_rect_d(a,b,r):
    # 0 если отрезок проходит через прямоугольник
    x0,y0,x1,y1=r
    # Liang–Barsky clip
    dx,dy=b[0]-a[0],b[1]-a[1]; t0,t1=0.0,1.0; ok=True
    for p,q in ((-dx,a[0]-x0),(dx,x1-a[0]),(-dy,a[1]-y0),(dy,y1-a[1])):
        if p==0:
            if q<0: ok=False;break
        else:
            t=q/p
            if p<0:
                if t>t1: ok=False;break
                t0=max(t0,t)
            else:
                if t<t0: ok=False;break
                t1=min(t1,t)
    if ok: return 0.0
    return min(rect_d(a,r),rect_d(b,r),*(seg_d(c,a,b) for c in ((x0,y0),(x0,y1),(x1,y0),(x1,y1))))
def poly_rect_d(pts,r): return min(seg_rect_d(pts[i],pts[i+1],r) for i in range(len(pts)-1))
