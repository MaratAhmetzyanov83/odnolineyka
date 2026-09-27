import ezdxf, pickle, os
from ezdxf import recover
doc, aud = recover.readfile("plan.dxf")
