# -*- coding: utf-8 -*-
import os
from PIL import Image

BASE = r"C:\Users\94383\Desktop\闲鱼网页"
ASSETS = os.path.join(BASE, "assets")
names = ["item1", "item2", "item3", "item4", "item5", "item6", "item7", "item8", "item9"]

for n in names:
    p = os.path.join(ASSETS, n + ".jpg")
    im = Image.open(p)
    im.thumbnail((640, 640), Image.LANCZOS)
    im = im.convert("RGB")
    im.save(p, "JPEG", quality=82, optimize=True)
    print(n, im.size, round(os.path.getsize(p) / 1024), "KB")

w, h = 280, 280
sheet = Image.new("RGB", (w * 3 + 20, h * 3 + 20), "white")
for i, n in enumerate(names):
    im = Image.open(os.path.join(ASSETS, n + ".jpg"))
    im.thumbnail((w, h), Image.LANCZOS)
    x = (i % 3) * (w + 10)
    y = (i // 3) * (h + 10)
    sheet.paste(im, (x, y))
out = os.path.join(BASE, "_shots", "imgs_check.jpg")
sheet.save(out, "JPEG", quality=85)
print("montage:", out)
