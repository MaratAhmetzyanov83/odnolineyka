import math
def seg_d(p,a,b):
    ax,ay=a;bx,by=b;px,py=p
    dx,dy=bx-ax,by-ay; L=dx*dx+dy*dy
    t=0 if L==0 else max(0,min(1,((px-ax)*dx+(py-ay)*dy)/L))
    return math.hypot(px-ax-t*dx,py-ay-t*dy)
def pl_d(p,pts): return min(seg_d(p,pts[i],pts[i+1]) for i in range(len(pts)-1))
def pl_len(pts): return sum(math.dist(pts[i],pts[i+1]) for i in range(len(pts)-1))
