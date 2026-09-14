"""Pixel audit of the ten composed listing images: cards must stay inside the content zone between the title band and the footer,
with the 80 px side margins intact. Exit 1 on any violation."""
from PIL import Image
import numpy as np, glob, os, sys
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
LIGHT = np.array((238, 246, 247)); bad = 0
for f in sorted(glob.glob(os.path.join(ROOT, "listing", "images", "*.png"))):
    a = np.asarray(Image.open(f).convert("RGB")).astype(int)
    zone = a[305:1895]
    nonbg = (abs(zone - LIGHT).sum(axis=2) > 40)
    rows = np.where(nonbg.any(axis=1))[0]; cols = np.where(nonbg.any(axis=0))[0]
    top, bot, left, right = rows.min() + 305, rows.max() + 305, cols.min(), cols.max()
    ok = 330 <= top and bot <= 1885 and left >= 78 and right <= 1936   # cards span 80..1920 plus a 14 px drop shadow
    bad += not ok
    print(f"{os.path.basename(f):28s} content y {top:4d}-{bot:4d}  x {left:3d}-{right:4d}  {'OK' if ok else 'VIOLATION'}")
sys.exit(1 if bad else 0)
