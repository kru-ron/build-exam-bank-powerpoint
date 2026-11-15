# crop figures listed in a data module's CROPS; remove coloured pen/highlighter, erase boxes, transparent paper
import sys, os, importlib, fitz
from PIL import Image
from collections import deque
mod, pdf, out = sys.argv[1:4]
sys.path.insert(0, '.'); D = importlib.import_module(mod)
os.makedirs(out, exist_ok=True)
d = fitz.open(pdf)
RD = 170 / 72
only = set(sys.argv[4:])
for key, (pg, half, (x0, y0, x1, y1), erase) in D.CROPS.items():
    if only and key not in only: continue
    p = d[pg - 1]; W, H = p.rect.width, p.rect.height
    ox, oy = W * 0.08, (H * 0.06 if half == 0 else H * 0.48)
    rect = fitz.Rect(ox + x0 / RD, oy + y0 / RD, ox + x1 / RD, oy + y1 / RD)
    pix = p.get_pixmap(dpi=300, clip=rect)
    im = Image.frombytes('RGB', (pix.width, pix.height), pix.samples).convert('HSV')
    rgb = Image.frombytes('RGB', (pix.width, pix.height), pix.samples)
    w, h = rgb.size; hp = im.load(); rp = rgb.load()
    g = Image.new('L', (w, h)); gp = g.load()
    k = 300 / 170
    eb = [(int((a - x0) * k), int((b - y0) * k), int((c - x0) * k), int((e - y0) * k)) for a, b, c, e in erase]
    for y in range(h):
        for x in range(w):
            hh, s, v = hp[x, y]
            r_, g_, b_ = rp[x, y]
            lum = int(0.299 * r_ + 0.587 * g_ + 0.114 * b_)
            if s > 60 and v > 90: lum = 255           # coloured ink / highlighter, not dark text
            if any(a <= x <= c and b <= y <= e for a, b, c, e in eb): lum = 255
            gp[x, y] = lum
    # transparent paper background (flood from border)
    px = g.load(); bg = bytearray(w * h); dq = deque()
    for x in range(w): dq.append((x, 0)); dq.append((x, h - 1))
    for y in range(h): dq.append((0, y)); dq.append((w - 1, y))
    while dq:
        x, y = dq.popleft(); i = y * w + x
        if bg[i] or px[x, y] < 200: continue
        bg[i] = 1
        if x + 1 < w: dq.append((x + 1, y))
        if x > 0: dq.append((x - 1, y))
        if y + 1 < h: dq.append((x, y + 1))
        if y > 0: dq.append((x, y - 1))
    o = Image.new('RGBA', (w, h)); op = o.load()
    for y in range(h):
        for x in range(w):
            v = px[x, y]
            op[x, y] = (255, 255, 255, 0) if bg[y * w + x] else ((255,) * 3 + (255,) if v > 235 else (v, v, v, 255))
    o.save(f'{out}/{key}.png')
    print(key, o.size)
