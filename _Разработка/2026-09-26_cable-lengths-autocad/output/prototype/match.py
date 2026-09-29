import re, math
from geo import *
def sysname(layer):
    # "0_UNC_2 Розетки. Трассы" -> "0_UNC_2 Розетки"
    return layer.rsplit(".",1)[0].strip()
def major(layer):
    m=re.match(r"0_UNC_(\d+)", layer); return m.group(1) if m else layer
def resolve(assign):
    # полилиния с 2+ группами: если группа уже имеет другие "свои" полилинии — убираем её отсюда
    changed=True
    while changed:
        changed=False
        for h,g in assign.items():
            if len(g)<2: continue
            for grp in list(g):
                own=[h2 for h2,g2 in assign.items() if h2!=h and grp in g2 and len(g2)==1]
                if own and len(g)>1:
                    del g[grp]; changed=True
def run(region_polys, region_labels, TOL=60, JTOL=60, MTTOL=500):
    P={p["h"]:p for p in region_polys}
    assign={h:{} for h in P}   # h -> {grp: reason}
    issues=[]
    for lb in region_labels:
        pts=lb["arrows"]
        if not pts: continue
        for a in pts:
            cand=[(pl_d(a,p["pts"]),h) for h,p in P.items()]
            cand=[(d,h) for d,h in cand if d<=TOL]
            if not cand:
                issues.append(("Выноска ни на что не указывает", lb["grp"], lb["h"])); continue
            same=[(d,h) for d,h in cand if sysname(P[h]["layer"])==sysname(lb["layer"])] or [(d,h) for d,h in cand if major(P[h]["layer"])==major(lb["layer"])]
            if not same:
                issues.append(("Выноска касается трассы другой системы", lb["grp"], lb["h"])); continue
            if len(same)>1:
                # стрелка на пучке: берём полилинию, у которой рядом конец
                ends=[(min(math.dist(a,P[h]["pts"][0]),math.dist(a,P[h]["pts"][-1])),h) for d,h in same]
                ends.sort()
                same=[(0,ends[0][1])]
            assign[same[0][1]].setdefault(lb["grp"],"выноска")
    resolve(assign)
    # текстовые подписи без стрелок — только если полилиния ещё без группы
    for lb in region_labels:
        if lb["arrows"]: continue
        cand=sorted((pl_d(lb["pos"],p["pts"]),h) for h,p in P.items() if sysname(p["layer"])==sysname(lb["layer"]))
        if cand and cand[0][0]<=MTTOL and not assign[cand[0][1]]:
            assign[cand[0][1]][lb["grp"]]="подпись"
    # цепочки: конец одной полилинии = начало следующей
    changed=True
    while changed:
        changed=False
        for h,p in P.items():
            if assign[h]: continue
            for h2,p2 in P.items():
                if h2==h or not assign[h2] or len(assign[h2])!=1: continue
                if sysname(p2["layer"])!=sysname(p["layer"]): continue
                e1=[p["pts"][0],p["pts"][-1]]; e2=[p2["pts"][0],p2["pts"][-1]]
                if any(math.dist(a,b)<=JTOL for a in e1 for b in e2):
                    g=next(iter(assign[h2])); assign[h][g]="цепочка от "+h2; changed=True; break
    return P, assign, issues
