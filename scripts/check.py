import fitz,sys
from PIL import Image
pdf=sys.argv[1]; pre=sys.argv[2]
d=fitz.open(pdf); print('pages',len(d))
for i,p in enumerate(d):
    ys=[b[3] for b in p.get_text('blocks') if b[3]<785 and b[4].strip()]
    print(i+1,'max text y',round(max(ys),1) if ys else None)
    p.get_pixmap(dpi=100).save(f'{pre}{i+1:02d}.png')
