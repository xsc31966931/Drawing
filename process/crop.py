# -*- coding: utf-8 -*-
import os
from PIL import Image

SRC1 = r'D:\图片\留存记录\文章\火山\202609Seed-2.1-pro约稿\参考资料\图纸\挑选\输入01_别墅一层平面图.png'
SRC2 = r'D:\图片\留存记录\文章\火山\202609Seed-2.1-pro约稿\参考资料\图纸\挑选\输入02_别墅二层平面图.png'
OUT = r'D:\code\Drawing\_work'
os.makedirs(OUT, exist_ok=True)

def tiles(src, tag, W, H, rows, cols, scale=3):
    im = Image.open(src).convert('RGB')
    w, h = im.size
    rh = h // rows
    cw = w // cols
    ov_r = int(rh * 0.18)
    ov_c = int(cw * 0.18)
    for r in range(rows):
        for c in range(cols):
            t = max(0, r*rh - (ov_r if r > 0 else 0))
            b = min(h, (r+1)*rh + (ov_r if r < rows-1 else 0))
            l = max(0, c*cw - (ov_c if c > 0 else 0))
            rr = min(w, (c+1)*cw + (ov_c if c < cols-1 else 0))
            crop = im.crop((l, t, rr, b))
            crop = crop.resize((crop.width*scale, crop.height*scale), Image.LANCZOS)
            crop.save(os.path.join(OUT, f'{tag}_r{r}c{c}.png'))
            print(tag, r, c, (l,t,rr,b))

tiles(SRC1, 'f1', 490, 940, rows=4, cols=2, scale=3)
tiles(SRC2, 'f2', 490, 920, rows=4, cols=2, scale=3)
print('done')
