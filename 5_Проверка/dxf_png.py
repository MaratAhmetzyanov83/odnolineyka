"""DXF -> PNG (окно x0 x1 y0 y1).  py render.py in.dxf out.png x0 x1 y0 y1 [dpi]"""
import sys, ezdxf, matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from ezdxf.addons.drawing import RenderContext, Frontend
from ezdxf.addons.drawing.matplotlib import MatplotlibBackend
from ezdxf.addons.drawing.config import Configuration, BackgroundPolicy, ColorPolicy

src, out = sys.argv[1], sys.argv[2]
x0, x1, y0, y1 = map(float, sys.argv[3:7])
dpi = int(sys.argv[7]) if len(sys.argv) > 7 else 200
doc = ezdxf.readfile(src)
msp = doc.modelspace()
fig = plt.figure(figsize=((x1 - x0) / 25.4, (y1 - y0) / 25.4))
ax = fig.add_axes([0, 0, 1, 1])
cfg = Configuration(background_policy=BackgroundPolicy.WHITE, color_policy=ColorPolicy.COLOR)
Frontend(RenderContext(doc), MatplotlibBackend(ax), config=cfg).draw_layout(msp, finalize=True)
ax.set_xlim(x0, x1); ax.set_ylim(y0, y1)
fig.savefig(out, dpi=dpi)
