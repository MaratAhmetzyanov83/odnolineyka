import re, math
from geo import *
def sysname(layer): return layer.rsplit(".",1)[0].strip()
def major(layer):
    m=re.match(r"0_UNC_(\d+)", layer); return m.group(1) if m else layer
def run(region_polys, region_labels, TOL=60, JTOL=60, RTOL=400):
    P={p["h"]:p for p in region_polys}
    A={h:{} for h in P}; hits=[]; issues=[]
    ends=lambda h:(P[h]["pts"][0],P[h]["pts"][-1])
    def enddist(a,h): return min(math.dist(a,e) for e in ends(h))
    for lb in region_labels:
        for a in lb["arrows"]:
            cand=[h for h in P if pl_d(a,P[h]["pts"])<=TOL]
            same=[h for h in cand if sysname(P[h]["layer"])==sysname(lb["layer"])] or [h for h in cand if major(P[h]["layer"])==major(lb["layer"])]
            if not same: issues.append((lb["grp"],"нет трассы")); continue
            best=min(same,key=lambda h:enddist(a,h))
            A[best][lb["grp"]]="выноска"; hits.append((lb["grp"],best,a))
    # 1) группа уже имеет свою полилинию -> убрать с общей
    changed=True
    while changed:
        changed=False
        for h,g in A.items():
            if len(g)>1:
                for grp in list(g):
                    if len(g)>1 and any(h2!=h and set(A[h2])=={grp} for h2 in A): del g[grp]; changed=True
    # 2) на полилинии остаётся та группа, чья стрелка ближе к её концу; остальные ищут свободную трассу рядом
    for h,g in A.items():
        if len(g)>1:
            dist={grp:min(enddist(a,h) for gg,hh,a in hits if gg==grp and hh==h) for grp in g}
            keep=min(dist,key=dist.get)
            for grp in list(g):
                if grp==keep: continue
                arrows=[a for gg,hh,a in hits if gg==grp and hh==h]
                free=[(min(math.dist(a,e) for a in arrows for e in ends(h2)),h2) for h2 in A if not A[h2] and sysname(P[h2]["layer"])==sysname(P[h]["layer"])]
                free=[x for x in free if x[0]<=RTOL]
                if free:
                    del g[grp]; A[min(free)[1]][grp]="выноска рядом"
    # 3) цепочки, не через узлы (щит), где сходятся 3+ концов
    allends=[e for h in P for e in ends(h)]
    hub=lambda p: sum(1 for e in allends if math.dist(p,e)<=JTOL)>=3
    changed=True
    while changed:
        changed=False
        for h in P:
            if A[h]: continue
            for h2 in P:
                if h2==h or len(A[h2])!=1 or sysname(P[h2]["layer"])!=sysname(P[h]["layer"]): continue
                if any(math.dist(a,b)<=JTOL and not hub(a) for a in ends(h) for b in ends(h2)):
                    A[h]={next(iter(A[h2])):"цепочка"}; changed=True; break
    return P,A,issues
