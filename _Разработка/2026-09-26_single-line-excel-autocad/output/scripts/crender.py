"""Color renderer for comparing DXF drawings (AutoCAD-like dark background).
usage: python3 crender.py <dxf.pkl> x0 y0 x1 y1 out.png [dpi]"""
import pickle, math, sys, re
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.collections import LineCollection

ACI = {1: (255, 0, 0), 2: (255, 255, 0), 3: (0, 255, 0), 4: (0, 255, 255), 5: (0, 0, 255), 6: (255, 0, 255),
       7: (230, 230, 230), 8: (128, 128, 128), 9: (192, 192, 192), 30: (255, 127, 0), 32: (204, 102, 0),
       244: (153, 27, 30), 12: (189, 0, 0), 14: (129, 0, 0), 116: (38, 38, 255), 177: (41, 49, 137), 18: (104, 0, 0)}


def g(e, c, d=None):
    for k, v in e:
        if k == c:
            return v
    return d


def gall(e, c):
    return [v for k, v in e if k == c]


def load(pkl):
    s = pickle.load(open(pkl, "rb"))
    blocks, base = {}, {}
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
    layers = {}
    for e in s.get("TABLES", []):
        if g(e, 0) == "LAYER":
            try:
                col = abs(int(str(g(e, 62, "7")).strip()))
            except ValueError:
                col = 7
            tc = g(e, 420)
            layers[g(e, 2)] = (col, int(tc) if tc else None, (g(e, 6) or "CONTINUOUS").upper())
    return s, blocks, base, layers


def rgb_of(e, layers, parent):
    tc = g(e, 420)
    if tc:
        v = int(str(tc).strip()); return ((v >> 16) & 255, (v >> 8) & 255, v & 255)
    c = g(e, 62)
    c = int(str(c).strip()) if c is not None else 256
    if c == 0:
        return parent
    if c == 256:
        lay = layers.get(g(e, 8, "0"), (7, None, "CONTINUOUS"))
        if lay[1]:
            v = lay[1]; return ((v >> 16) & 255, (v >> 8) & 255, v & 255)
        c = lay[0]
    return ACI.get(c, (230, 230, 230))


def lt_of(e, layers):
    lt = (g(e, 6) or "BYLAYER").upper()
    if lt == "BYLAYER":
        lt = layers.get(g(e, 8, "0"), (7, None, "CONTINUOUS"))[2]
    return lt not in ("CONTINUOUS", "BYBLOCK")


def xf(T, x, y):
    a, b, c, d, e_, f = T
    return (a * x + b * y + e_, c * x + d * y + f)


def compose(T, U):
    a, b, c, d, e, f = T
    A, B, C, D, E, F = U
    return (a * A + b * C, a * B + b * D, c * A + d * C, c * B + d * D, a * E + b * F + e, c * E + d * F + f)


def mtext_clean(t):
    t = re.sub(r"\\[A-Za-z][^;\\]*;", "", t)
    return t.replace("\\P", " ").replace("{", "").replace("}", "")


def collect(pkl):
    s, blocks, base, layers = load(pkl)
    segs, texts = [], []

    def walk(ents, T, parent, depth=0):
        for e in ents:
            t = g(e, 0)
            if (g(e, 60, "0") or "0").strip() == "1":
                continue
            if t in ("ATTDEF",):
                continue
            if t == "ATTRIB" and (int((g(e, 70, "0") or "0").strip() or 0) & 1):
                continue
            col = rgb_of(e, layers, parent); dash = lt_of(e, layers)
            if t == "LINE":
                segs.append((xf(T, float(g(e, 10)), float(g(e, 20))), xf(T, float(g(e, 11)), float(g(e, 21))), col, dash))
            elif t == "LWPOLYLINE":
                xs = [float(v) for v in gall(e, 10)]; ys = [float(v) for v in gall(e, 20)]
                pts = [xf(T, x, y) for x, y in zip(xs, ys)]
                if (g(e, 70, "0") or "0").strip() in ("1", "129") and pts:
                    pts.append(pts[0])
                for p, q in zip(pts, pts[1:]):
                    segs.append((p, q, col, dash))
            elif t in ("CIRCLE", "ARC"):
                cx, cy, r = float(g(e, 10)), float(g(e, 20)), float(g(e, 40))
                a0, a1 = (0, 360) if t == "CIRCLE" else (float(g(e, 50)), float(g(e, 51)))
                if a1 < a0: a1 += 360
                n = 20
                pts = [xf(T, cx + r * math.cos(math.radians(a0 + (a1 - a0) * k / n)), cy + r * math.sin(math.radians(a0 + (a1 - a0) * k / n))) for k in range(n + 1)]
                for p, q in zip(pts, pts[1:]):
                    segs.append((p, q, col, False))
            elif t in ("TEXT", "MTEXT", "ATTRIB"):
                txt = g(e, 1, "")
                if t == "MTEXT":
                    txt = mtext_clean("".join(gall(e, 3)) + txt)
                h = float(g(e, 40, 2.5)); rot = float(g(e, 50, 0) or 0)
                al = (g(e, 72, "0") or "0").strip()
                if al not in ("0", "") and g(e, 11) is not None:
                    p = xf(T, float(g(e, 11, 0)), float(g(e, 21, 0)))
                else:
                    p = xf(T, float(g(e, 10, 0)), float(g(e, 20, 0)))
                sc = math.hypot(T[0], T[2]); rT = math.degrees(math.atan2(T[2], T[0]))
                texts.append((p, txt, h * sc, rot + rT, al, col))
            elif t == "INSERT":
                name = g(e, 2)
                if name not in blocks or depth > 8:
                    continue
                x, y = float(g(e, 10, 0)), float(g(e, 20, 0))
                sx, sy = float(g(e, 41, 1)), float(g(e, 42, 1))
                r = math.radians(float(g(e, 50, 0)))
                bx, by = base[name]
                cs, sn = math.cos(r), math.sin(r)
                U = (cs * sx, -sn * sy, sn * sx, cs * sy, 0, 0)
                U = (U[0], U[1], U[2], U[3], x - (U[0] * bx + U[1] * by), y - (U[2] * bx + U[3] * by))
                walk(blocks[name], compose(T, U), col, depth + 1)

    ents = [e for e in s["ENTITIES"] if (g(e, 67, "0") or "0").strip() == "0"]
    walk(ents, (1, 0, 0, 1, 0, 0), (230, 230, 230))
    return segs, texts


def render(geo, x0, y0, x1, y1, out, dpi=120, width=24, title=None):
    segs, texts = geo
    fig_w = width; fig_h = width * (y1 - y0) / (x1 - x0)
    fig = plt.figure(figsize=(fig_w, fig_h), facecolor=(0.13, 0.15, 0.19))
    ax = fig.add_axes([0, 0, 1, 1]); ax.set_facecolor((0.13, 0.15, 0.19))
    sel = [sg for sg in segs if (x0 <= sg[0][0] <= x1 or x0 <= sg[1][0] <= x1) and (y0 <= sg[0][1] <= y1 or y0 <= sg[1][1] <= y1)]
    for dash in (False, True):
        part = [sg for sg in sel if sg[3] == dash]
        if part:
            ax.add_collection(LineCollection([(p, q) for p, q, c, d in part], linewidths=0.8,
                                             colors=[tuple(v / 255 for v in c) for p, q, c, d in part],
                                             linestyles=(0, (4, 3)) if dash else "solid"))
    ppu = fig_w * 72 / (x1 - x0)
    for (p, txt, h, rot, al, col) in texts:
        if x0 <= p[0] <= x1 and y0 <= p[1] <= y1 and txt.strip():
            fs = max(2, h * ppu * 0.95)
            ha = "center" if al in ("1", "4") else ("right" if al == "2" else "left")
            va = "center" if al == "4" else "baseline"
            ax.text(p[0], p[1], txt[:60], fontsize=fs, color=tuple(v / 255 for v in col), rotation=rot,
                    ha=ha, va=va, rotation_mode="anchor", family="DejaVu Sans")
    if title:
        ax.text(x0 + 2, y1 - 6, title, fontsize=14, color="yellow")
    ax.set_xlim(x0, x1); ax.set_ylim(y0, y1); ax.set_aspect("equal"); ax.axis("off")
    fig.savefig(out, dpi=dpi, facecolor=fig.get_facecolor()); plt.close(fig)


if __name__ == "__main__":
    geo = collect(sys.argv[1])
    x0, y0, x1, y1 = map(float, sys.argv[2:6])
    render(geo, x0, y0, x1, y1, sys.argv[6], int(sys.argv[7]) if len(sys.argv) > 7 else 120)
