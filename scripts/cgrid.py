import sys
from PIL import Image, ImageDraw
src, dst = sys.argv[1:3]; step = int(sys.argv[3]) if len(sys.argv) > 3 else 50
im = Image.open(src); bg = Image.new('RGB', im.size, 'white'); bg.paste(im, (0, 0), im if im.mode == 'RGBA' else None)
d = ImageDraw.Draw(bg)
for x in range(0, bg.width, step // 2):
    big = x % step == 0
    d.line([(x, 0), (x, bg.height)], fill=(255, 120, 120) if big else (255, 215, 215))
    if big: d.text((x + 1, 1), str(x), fill=(220, 0, 0))
for y in range(0, bg.height, step // 2):
    big = y % step == 0
    d.line([(0, y), (bg.width, y)], fill=(120, 120, 255) if big else (215, 215, 255))
    if big: d.text((1, y + 1), str(y), fill=(0, 0, 220))
bg.save(dst)
