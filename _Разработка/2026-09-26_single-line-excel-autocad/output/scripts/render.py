import pickle, math, sys, re
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.collections import LineCollection

SP = "/tmp/claude-0/-home-claude/5e645175-0e52-5350-b586-c0cb5927e030/scratchpad/"
s = pickle.load(open(SP + "dxf.pkl", "rb"))

def g(e, c, d=None):
    for k, v in e:
        if k == c:
            return v
    return d

def gall(e, c):
    return [v for k, v in e if k == c]

# block table
blocks = {}; base = {}
cur = None
for e in s["BLOCKS"]:
    t = g(e, 0)
    if t == "BLOCK":
        cur = g(e, 2); blocks[cur] = []
        base[cur] = (float(g(e, 10, 0)), float(g(e, 20, 0)))
    elif t == "ENDBLK":
        cur = None
    elif cur is not None:
        blocks[cur].append(e)

def mtext_clean(t):
    t = re.sub(r"\\[A-Za-z][^;\\]*;", "", t)
    t = t.replace("\\P", " ").replace("{", "").replace("}", "")
    return t

def xf(T, x, y):
    a, b, c, d, e_, f = T
    return (a * x + b * y + e_, c * x + d * y + f)

def compose(T, U):
    a, b, c, d, e, f = T
    A, B, C, D, E, F = U
    return (a*A + b*C, a*B + b*D, c*A + d*C, c*B + d*D, a*E + b*F + e, c*E + d*F + f)

segs = []; texts = []

def walk(ents, T, depth=0, skip_attdef=True):
    it = iter(range(len(ents)))
    i = 0
    while i < len(ents):
        e = ents[i]; t = g(e, 0); i += 1
        if (g(e, 60, "0") or "0").strip() == "1":
            continue
        if t in ("ATTRIB", "ATTDEF") and (int((g(e, 70, "0") or "0").strip() or 0) & 1):
            continue
        if t == "LINE":
            p = xf(T, float(g(e, 10)), float(g(e, 20))); q = xf(T, float(g(e, 11)), float(g(e, 21)))
            segs.append((p, q))
        elif t == "LWPOLYLINE":
            xs = [float(v) for v in gall(e, 10)]; ys = [float(v) for v in gall(e, 20)]
            pts = [xf(T, x, y) for x, y in zip(xs, ys)]
            if g(e, 70, "0").strip() in ("1", "129") and pts:
                pts.append(pts[0])
            segs.extend(zip(pts, pts[1:]))
        elif t in ("CIRCLE", "ARC"):
            cx, cy, r = float(g(e, 10)), float(g(e, 20)), float(g(e, 40))
            a0, a1 = (0, 360) if t == "CIRCLE" else (float(g(e, 50)), float(g(e, 51)))
            if a1 < a0: a1 += 360
            n = 16
            pts = [xf(T, cx + r*math.cos(math.radians(a0 + (a1-a0)*k/n)), cy + r*math.sin(math.radians(a0 + (a1-a0)*k/n))) for k in range(n+1)]
            segs.extend(zip(pts, pts[1:]))
        elif t in ("TEXT", "MTEXT", "ATTRIB") or (t == "ATTDEF" and not skip_attdef):
            txt = g(e, 1, "")
            if t == "MTEXT":
                txt = "".join(gall(e, 3)) + txt
                txt = mtext_clean(txt)
            h = float(g(e, 40, 2.5))
            p = xf(T, float(g(e, 10, 0)), float(g(e, 20, 0)))
            sc = math.hypot(T[0], T[2])
            texts.append((p, txt, h * sc))
        elif t == "INSERT":
            name = g(e, 2)
            if name not in blocks or depth > 8:
                continue
            x, y = float(g(e, 10, 0)), float(g(e, 20, 0))
            sx, sy = float(g(e, 41, 1)), float(g(e, 42, 1))
            r = math.radians(float(g(e, 50, 0)))
            bx, by = base[name]
            cs, sn = math.cos(r), math.sin(r)
            # local: translate(-base) -> scale -> rotate -> translate(x,y)
            U = (cs*sx, -sn*sy, sn*sx, cs*sy, 0, 0)
            U = (U[0], U[1], U[2], U[3], x - (U[0]*bx + U[1]*by), y - (U[2]*bx + U[3]*by))
            walk(blocks[name], compose(T, U), depth + 1)

ents = [e for e in s["ENTITIES"] if g(e, 67, "0").strip() == "0"]
walk(ents, (1, 0, 0, 1, 0, 0))
pickle.dump((segs, texts), open(SP + "geom.pkl", "wb"))
print(len(segs), len(texts))

def render(x0, y0, x1, y1, out, dpi=200, width=30, fs_scale=1.0):
    fig_w = width; fig_h = width * (y1 - y0) / (x1 - x0)
    fig = plt.figure(figsize=(fig_w, fig_h)); ax = fig.add_axes([0, 0, 1, 1])
    sel = [(p, q) for p, q in segs if (x0 <= p[0] <= x1 or x0 <= q[0] <= x1) and (y0 <= p[1] <= y1 or y0 <= q[1] <= y1)]
    ax.add_collection(LineCollection(sel, linewidths=0.4, colors="k"))
    ppu = fig_w * 72 / (x1 - x0)  # points per unit
    for (p, txt, h) in texts:
        if x0 <= p[0] <= x1 and y0 <= p[1] <= y1 and txt.strip():
            fs = max(2, h * ppu * fs_scale)
            ax.text(p[0], p[1], txt[:60], fontsize=fs, color="b")
    ax.set_xlim(x0, x1); ax.set_ylim(y0, y1); ax.set_aspect("equal"); ax.axis("off")
    fig.savefig(out, dpi=dpi); plt.close(fig)

if __name__ == "__main__" and len(sys.argv) > 1:
    x0, y0, x1, y1 = map(float, sys.argv[1:5]); out = sys.argv[5]
    dpi = int(sys.argv[6]) if len(sys.argv) > 6 else 100
    render(x0, y0, x1, y1, SP + out, dpi=dpi)
