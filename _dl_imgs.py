# -*- coding: utf-8 -*-
import urllib.request, os, sys

BASE = r"C:\Users\94383\Desktop\闲鱼网页"
ASSETS = os.path.join(BASE, "assets")

urls = {
    "item1": "https://aka.doubaocdn.com/s/ZGbx4uLsUe",
    "item2": "https://aka.doubaocdn.com/s/UvIUP62dbS",
    "item3": "https://aka.doubaocdn.com/s/eKdmGsOIH8",
    "item4": "https://aka.doubaocdn.com/s/fMWkbwWr9q",
    "item5": "https://aka.doubaocdn.com/s/CmAINEdpOU",
    "item6": "https://aka.doubaocdn.com/s/0UgmKROjH0",
    "item7": "https://aka.doubaocdn.com/s/nnay81je3a",
    "item8": "https://aka.doubaocdn.com/s/V86nbhlfSo",
    "item9": "https://aka.doubaocdn.com/s/VWtph7lXvd",
}

try:
    from PIL import Image
    has_pil = True
except Exception:
    has_pil = False

for name, u in urls.items():
    path = os.path.join(ASSETS, name + ".jpg")
    try:
        req = urllib.request.Request(u, headers={"User-Agent": "Mozilla/5.0"})
        data = urllib.request.urlopen(req, timeout=60).read()
        with open(path, "wb") as f:
            f.write(data)
        size_kb = len(data) / 1024
        if has_pil:
            im = Image.open(path)
            im.thumbnail((640, 640), Image.LANCZOS)
            im = im.convert("RGB")
            im.save(path, "JPEG", quality=82, optimize=True)
            newsize = os.path.getsize(path) / 1024
            print("%s ok %sx%s %sKB->%sKB" % (name, im.size[0], im.size[1], round(size_kb), round(newsize)))
        else:
            print("%s ok %sKB (no PIL)" % (name, round(size_kb)))
    except Exception as e:
        print("%s FAIL: %s" % (name, e))

if has_pil:
    try:
        files = [os.path.join(ASSETS, n + ".jpg") for n in sorted(urls.keys())]
        imgs = [Image.open(p) for p in files]
        w, h = 260, 260
        sheet = Image.new("RGB", (w * 3 + 20, h * 3 + 20), "white")
        for i, im in enumerate(imgs):
            im2 = im.copy()
            im2.thumbnail((w, h), Image.LANCZOS)
            x = (i % 3) * (w + 10)
            y = (i // 3) * (h + 10)
            sheet.paste(im2, (x, y))
        out = os.path.join(BASE, "_shots", "imgs_check.jpg")
        os.makedirs(os.path.dirname(out), exist_ok=True)
        sheet.save(out, "JPEG", quality=85)
        print("montage:", out)
    except Exception as e:
        print("montage FAIL:", e)
