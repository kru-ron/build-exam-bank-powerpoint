# measure real line counts of each flow paragraph from PowerPoint-exported PDF
import fitz, json, sys, re
pdf, keysf, outf = sys.argv[1:4]
keys = json.load(open(keysf))
def norm(t):
    t = t.replace(chr(0x0E4D), '').replace(chr(0x0E33), chr(0x0E32)).replace('|', '').replace(chr(0xA6), '')
    return re.sub('[' + chr(32) + chr(9) + chr(0xA0) + chr(0x200B) + ']', '', t)
lines = []; raw = []
d = fitz.open(pdf)
for pi, p in enumerate(d):
    for b in p.get_text('dict')['blocks']:
        for l in b.get('lines', []):
            t = ''.join(s['text'] for s in l['spans'])
            x0, y0 = l['bbox'][0], l['bbox'][1]
            if y0 > 785 or not t.strip(): continue
            if abs(l.get('dir', (1, 0))[1]) > 0.1: continue
            if max(sp['size'] for sp in l['spans']) < 13: continue   # ป้ายในรูปวาด (≤ 12pt) ไม่ใช่เนื้อหา 14pt   # ข้อความหมุน (ป้ายแกนในรูปวาด) ไม่ใช่บรรทัดเนื้อหา
            raw.append((pi, y0, x0, t))
# รวมชิ้นข้อความบนบรรทัดเดียวกัน (ข้อความที่จัดด้วยแท็บถูกแยกเป็นหลายชิ้นใน PDF)
raw.sort(key=lambda r: (r[0], r[1], r[2]))
groups = []
for r in raw:
    if groups and groups[-1][0][0] == r[0] and abs(groups[-1][0][1] - r[1]) < 4: groups[-1].append(r)
    else: groups.append([r])
merged = []
for g in groups:
    g.sort(key=lambda r: r[2])
    merged.append((g[0][0], min(r[1] for r in g), g[0][2], ' '.join(r[3] for r in g)))
for pi, y0, x0, t in merged:
    if not (60 <= x0 <= 63 or 96 <= x0 <= 99): continue
    if re.fullmatch(r'(\s*[กขคงจ]\.\s*)+', t): continue   # ป้ายตัวเลือกรูป (อาจหลายป้ายในบรรทัดเดียว)
    lines.append((pi, y0, norm(t), t))
lines.sort()
res, ptr, miss = {}, 0, []
WR = {}
starts = []
for k in keys:
    nk = norm(k)[:12]
    j = ptr
    while j < len(lines) and not (lines[j][2].startswith(nk) or (len(lines[j][2]) >= 3 and nk.startswith(lines[j][2]))): j += 1
    if j == len(lines): miss.append(k); starts.append(None); continue
    starts.append(j); ptr = j + 1
idx = [s for s in starts if s is not None] + [len(lines)]
ii = 0
for k, s in zip(keys, starts):
    if s is None: continue
    nxt = idx[ii + 1]; ii += 1
    # stop at page change
    cnt = 1; texts = [lines[s][3]]
    for m in range(s + 1, nxt):
        if lines[m][0] != lines[s][0]: break
        cnt += 1; texts.append(lines[m][3])
    res[k] = cnt
    WR[k] = texts
json.dump(res, open(outf, 'w'), ensure_ascii=False, indent=0)
json.dump(WR, open(outf + '.wrap.json', 'w'), ensure_ascii=False, indent=0)
print('measured', len(res), 'missing', len(miss), miss[:5])
