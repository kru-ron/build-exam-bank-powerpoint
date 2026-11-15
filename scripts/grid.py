import sys
from PIL import Image, ImageDraw
src, dst = sys.argv[1:3]
x0,y0,x1,y1 = map(int, sys.argv[3:7]) if len(sys.argv) > 6 else (0,0,10**5,10**5)
im = Image.open(src).convert('RGB'); im = im.crop((x0, y0, min(x1, im.width), min(y1, im.height)))
d = ImageDraw.Draw(im)
for x in range((x0 // 25 + 1) * 25, x0 + im.width, 25):
    c = (255, 0, 0) if x % 100 == 0 else (255, 180, 180)
    d.line([(x - x0, 0), (x - x0, im.height)], fill=c, width=1)
    if x % 100 == 0: d.text((x - x0 + 2, 2), str(x), fill=(255, 0, 0))
for y in range((y0 // 25 + 1) * 25, y0 + im.height, 25):
    c = (0, 0, 255) if y % 100 == 0 else (180, 180, 255)
    d.line([(0, y - y0), (im.width, y - y0)], fill=c, width=1)
    if y % 100 == 0: d.text((2, y - y0 + 2), str(y), fill=(0, 0, 255))
im.save(dst)
