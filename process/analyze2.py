# -*- coding: utf-8 -*-
import numpy as np
from PIL import Image

def load(p):
    return np.array(Image.open(p).convert('L'))

def profile(a, axis, t0, t1, thresh=150):
    sub = a[t0:t1, :] if axis == 'x' else a[:, t0:t1]
    # axis 'x': want darkness per column -> sub shape (rows, cols), mean over axis0
    fr = (sub < thresh).mean(axis=0 if axis == 'x' else 1)
    return fr

def peaks(fr, min_h=0.25, gap=3):
    idx = np.where(fr >= min_h)[0]
    groups = []
    if len(idx):
        s = idx[0]; p = idx[0]
        for x in idx[1:]:
            if x - p > gap:
                groups.append((s, p, round(float(fr[s:p+1].max()),3)))
                s = x
            p = x
        groups.append((s, p, round(float(fr[s:p+1].max()),3)))
    return groups

p1 = r'D:\图片\留存记录\文章\火山\202609Seed-2.1-pro约稿\参考资料\图纸\挑选\输入01_别墅一层平面图.png'
a = load(p1)

# building extent approx x 85..490, y 50..895
fx = profile(a, 'x', 50, 895)
print('== vertical dark bands (walls) full-height ==')
for g in peaks(fx, min_h=0.20): print(g)
fy = profile(a, 'y', 85, 490)
print('== horizontal dark bands (walls) full-width ==')
for g in peaks(fy, min_h=0.18): print(g)
