# หา "จุดตัดกลางวลี" จากบรรทัดจริงใน PDF แล้วกำหนดให้ขึ้นบรรทัดใหม่ที่ช่องว่างก่อนหน้า
# usage: breaks.py wrap.json breaks.json
import json, sys, os, re
wr = json.load(open(sys.argv[1])); bp = sys.argv[2]
B = json.load(open(bp)) if os.path.exists(bp) else {}
SKIP = set('ํ \t ​|¦\n')
def nmap(s):
    """normalized string + index map back to s"""
    out, idx = [], []
    for i, c in enumerate(s):
        if c in SKIP: continue
        out.append('า' if c == 'ำ' else c); idx.append(i)
    return ''.join(out), idx
def norm(t): return nmap(t)[0]
new = 0
for key, lines in wr.items():
    if len(lines) < 2 or '\t' in key: continue
    base = key.replace('\u00a6', ' ')
    head = key.split('|')[0]                   # ไม่รวมบรรทัดที่มา
    nb, ix = nmap(head)
    pos = 0
    for i in range(len(lines) - 1):
        nxt = norm(lines[i + 1])[:10]
        if not nxt or lines[i + 1].lstrip().startswith('('): break
        # หาจุดเริ่มบรรทัดถัดไป โดยค้นหลังเนื้อหาบรรทัดปัจจุบัน (กันจับคำซ้ำที่อยู่ต้นข้อความ)
        prev = pos
        j = nb.find(nxt, pos + max(1, len(norm(lines[i])) - 12))
        if j <= 0: break
        pos = j
        cut = ix[j]                              # ตำแหน่งในข้อความจริงที่บรรทัดใหม่เริ่ม
        if head[cut - 1] in ' \u00a6': continue   # ตัดที่ช่องว่างอยู่แล้ว ดี
        sp = head.rfind(' ', 0, cut)
        if sp <= 3 or sp < head.rfind('\u00a6', 0, cut): continue                     # ไม่มีช่องว่างให้ใช้ (หรือเป็นช่องหลังเลขข้อ)
        line_start = ix[prev] if prev < len(ix) else 0
        moved = len(norm(head[sp:cut])); linelen = len(norm(head[line_start:cut]))
        if moved > 0.5 * linelen: continue       # ถ้าต้องย้ายเกินครึ่งบรรทัด ไม่ทำ
        f = norm(head[sp + 1:])[:10]
        k = sum(1 for t in range(4, sp) if head[t] in ' \u00a6' and norm(head[t + 1:]).startswith(f))
        if k: f = f'{f}#{k}'                     # วลีซ้ำในย่อหน้า: ระบุว่าเป็นครั้งที่เท่าไร
        lst = B.setdefault(base, [])
        if f not in lst:
            # จุดตัดใหม่อยู่ก่อนจุดตัดเดิม → ล้างจุดตัดเดิมที่อยู่หลังมัน (การจัดบรรทัดเปลี่ยนแล้ว รอบถัดไปจะหาใหม่)
            def where(fr):
                b_, n_ = (fr.split('#')[0], int(fr.split('#')[1])) if '#' in fr else (fr, 0)
                hits = [t for t in range(4, len(head)) if head[t] in ' \u00a6' and norm(head[t + 1:]).startswith(b_)]
                return hits[n_] if n_ < len(hits) else 10 ** 9
            lst[:] = [x for x in lst if where(x) < sp]
            lst.append(f); new += 1
            break                                # ทีละจุดต่อย่อหน้าต่อรอบ (จุดหลังจะหาจากการจัดบรรทัดใหม่)
json.dump(B, open(bp, 'w'), ensure_ascii=False, indent=0)
print('new breaks', new, 'keys with breaks', len(B))
