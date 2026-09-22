# -*- coding: utf-8 -*-
import os
from PIL import Image
OUT = r'D:\code\Drawing\_work'
for src, tag in [
    (r'D:\图片\留存记录\文章\火山\202609Seed-2.1-pro约稿\参考资料\图纸\挑选\输入01_别墅一层平面图.png','f1'),
    (r'D:\图片\留存记录\文章\火山\202609Seed-2.1-pro约稿\参考资料\图纸\挑选\输入02_别墅二层平面图.png','f2')]:
    im = Image.open(src).convert('RGB')
    w,h = im.size
    im2 = im.resize((w*2,h*2), Image.LANCZOS)
    H = h*2
    for i in range(3):
        y0 = i*H//3
        y1 = min(H, (i+1)*H//3 + 50)
        im2.crop((0,y0,w*2,y1)).save(os.path.join(OUT,f'{tag}_2x_b{i}.png'))
    im2.save(os.path.join(OUT,f'{tag}_2x_full.png'))
print('done')
