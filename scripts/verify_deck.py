# ตรวจไฟล์ที่สร้างเสร็จ: ข้อความล้น footer, ตัวเลือกชิดซ้าย, คำว่า "ตอบ", จำนวนรูป/กลุ่ม
# usage: verify_deck.py deck.pptx deck.pdf
import sys, re, zipfile, fitz
pptx, pdf = sys.argv[1:3]
d = fitz.open(pdf); over = []
for i, p in enumerate(d):
    ys = [b[3] for b in p.get_text('blocks') if b[3] < 785 and b[4].strip()]
    if ys and max(ys) > 772: over.append((i + 1, round(max(ys), 1)))
z = zipfile.ZipFile(pptx); bad = opts = 0; t = ''
for n in z.namelist():
    if not re.match(r'ppt/slides/slide\d+\.xml$', n): continue
    x = z.read(n).decode(); t += x
    for para in re.findall(r'<a:p>.*?</a:p>', x, re.S):
        tt = ''.join(re.findall(r'<a:t[^>]*>([^<]*)</a:t>', para))
        if re.match(r'\s*[กขคงจ]\. ', tt):
            opts += 1; ppr = re.search(r'<a:pPr[^>]*>', para).group(0)
            if tt[0].isspace() or 'marL="0"' not in ppr or 'indent="0"' not in ppr: bad += 1
print(f'pages {len(d)} | overflow {over or "none"} | option lines {opts} bad {bad} | "ตอบ " {t.count("ตอบ ")} | '
      f'drawn groups {t.count("</p:grpSp>")} | raster images {sum(1 for n in z.namelist() if n.startswith("ppt/media"))}')
