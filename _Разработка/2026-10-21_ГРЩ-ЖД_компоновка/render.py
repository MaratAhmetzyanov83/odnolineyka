import sys, ezdxf, matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from ezdxf.addons.drawing import RenderContext, Frontend
from ezdxf.addons.drawing.matplotlib import MatplotlibBackend
from ezdxf.addons.drawing.config import Configuration, BackgroundPolicy, ColorPolicy
src, out, lay = sys.argv[1], sys.argv[2], sys.argv[3]
win = [float(v) for v in sys.argv[4].split(",")] if len(sys.argv)>4 else None
dpi = int(sys.argv[5]) if len(sys.argv)>5 else 150
doc = ezdxf.readfile(src)
sp = doc.modelspace() if lay=="Model" else doc.paperspace(lay)
fs=(16,11) if not win else (12,12*(win[3]-win[1])/(win[2]-win[0]))
fig = plt.figure(figsize=fs); ax = fig.add_axes([0,0,1,1])
cfg = Configuration(background_policy=BackgroundPolicy.WHITE, color_policy=ColorPolicy.BLACK)
Frontend(RenderContext(doc), MatplotlibBackend(ax), config=cfg).draw_layout(sp, finalize=(win is None))
if win:
    ax.set_aspect("equal", adjustable="box"); ax.set_xlim(win[0],win[2]); ax.set_ylim(win[1],win[3])
fig.savefig(out, dpi=dpi)
