# -*- coding: utf-8 -*-
import numpy as np
from PIL import Image

def load(p):
    im = Image.open(p).convert('L')
    return np.array(im)

def dark_cols(a, y0, y1, thresh=120, min_frac=0.6):
    """columns that are dark for >=min_frac of rows in band"""
    band = a[y0:y1, :]
    frac = (band < thresh).mean(axis=0)
    idx = np.where(frac >= min_frac)[0]
    # group
    groups = []
    if len(idx):
        s = idx[0]; p = idx[0]
        for x in idx[1:]:
            if x - p > 2:
                groups.append((s, p, frac[s:p+1].max()))
                s = x
            p = x
        groups.append((s, p, frac[s:p+1].max()))
    return groups

def dark_rows(a, x0, x1, thresh=120, min_frac=0.6):
    band = a[:, x0:x1]
    frac = (band < thresh).mean(axis=1)
    idx = np.where(frac >= min_frac)[0]
    groups = []
    if len(idx):
        s = idx[0]; p = idx[0]
        for y in idx[1:]:
            if y - p > 2:
                groups.append((s, p, frac[s:p+1].max()))
                s = y
            p = y
        groups.append((s, p, frac[s:p+1].max()))
    return groups

p1 = r'D:\图片\留存记录\文章\火山\202609Seed-2.1-pro约稿\参考资料\图纸\挑选\输入01_别墅一层平面图.png'
a = load(p1)
print('shape', a.shape)

# vertical grid lines show in top margin y 5..50 and bottom margin y 890..935
print('== top margin vertical lines ==')
for g in dark_cols(a, 2, 52, min_frac=0.5): print(g)
print('== bottom margin vertical lines ==')
for g in dark_cols(a, 888, 938, min_frac=0.5): print(g)
# horizontal grid lines in left margin x 2..85
print('== left margin horizontal lines ==')
for g in dark_rows(a, 2, 85, min_frac=0.5): print(g)
print('== right margin horizontal lines ==')
for g in dark_rows(a, 470, 489, min_frac=0.5): print(g)
