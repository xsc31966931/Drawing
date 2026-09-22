# -*- coding: utf-8 -*-
import os
from PIL import Image

OUT = r'D:\code\Drawing\_work'

def edge(src, tag):
    im = Image.open(src).convert('RGB')
    w, h = im.size
    strips = {
        'top': (0, 0, w, 70),
        'bottom': (0, h-75, w, h),
        'left': (0, 0, 95, h),
        'right': (w-70, 0, w, h),
    }
    for name, box in strips.items():
        c = im.crop(box)
        s = 4 if name in ('top', 'bottom') else 3
        c = c.resize((c.width*s, c.height*s), Image.LANCZOS)
        c.save(os.path.join(OUT, f'{tag}_{name}.png'))
        print(tag, name, c.size)

edge(r'D:\图片\留存记录\文章\火山\202609Seed-2.1-pro约稿\参考资料\图纸\挑选\输入01_别墅一层平面图.png', 'f1')
edge(r'D:\图片\留存记录\文章\火山\202609Seed-2.1-pro约稿\参考资料\图纸\挑选\输入02_别墅二层平面图.png', 'f2')
print('done')
