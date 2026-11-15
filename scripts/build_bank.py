# -*- coding: utf-8 -*-
"""สร้างคลังข้อสอบ PowerPoint ตามสกิล build-exam-bank-powerpoint บน master ของไฟล์ต้นแบบ
usage: python3 build_bank.py data_module template.pptx img_dir out.pptx
"""
import sys, os, re, math, importlib, copy, json
from xml.sax.saxutils import escape
from lxml import etree
from pptx import Presentation
from pptx.util import Pt, Emu
from PIL import ImageFont, Image

data_mod, TEMPLATE, IMG_DIR, OUT = sys.argv[1:5]
sys.path.insert(0, '.'); sys.path.insert(0, os.environ.get('EXAM_DATA', '.'))
D = importlib.import_module(data_mod)

FONT = "TH SarabunPSK"
PAN = 'panose="020B0500040200020003" pitchFamily="34" charset="-34"'
NS = 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
E = 12700  # EMU per pt

# ---- geometry (pt) taken from template textbox ----
BOX_X, BOX_Y, BOX_W = 54.4, 39.1, 486.6
INS_L, INS_T = 7.2, 3.6
TEXT_W = BOX_W - 2 * INS_L
LINE = 17.0           # exact line spacing for every paragraph
BOTTOM = 772.0        # content must end above footer line (y=788.8)
TAB = 36.0            # indent for data/steps

import glob as _g, os as _o
def _font(bold):
    pats = ['/Applications/Microsoft PowerPoint.app/Contents/Resources/DFonts/thsarabun' + ('-bold' if bold else '') + '.ttf',
            _o.path.expanduser('~/Library/Fonts/THSarabun' + (' Bold' if bold else '') + '.ttf'),
            '/Library/Fonts/THSarabun' + (' Bold' if bold else '') + '.ttf',
            'C:/Windows/Fonts/THSarabun' + (' Bold' if bold else '') + '.ttf']
    for p in pats:
        if _o.path.exists(p): return ImageFont.truetype(p, 140)
    raise SystemExit('ไม่พบฟอนต์ TH SarabunPSK — ติดตั้งฟอนต์หรือ PowerPoint ก่อน')
F14 = _font(False)
F14B = _font(True)

def plain(s):
    return re.sub(r'<u>|</u>|[\^_]\{([^}]*)\}', lambda m: m.group(1) or '', s)

def width(s, size=14, bold=False):
    f = F14B if bold else F14
    return f.getlength(plain(s)) * size / 140.0

def n_lines(s, avail, size=14, bold=False):
    """greedy wrap on spaces (Thai lines break at spaces or between words)"""
    words = plain(s).split(' ')
    lines, cur = 1, ''
    f = F14B if bold else F14
    for w in words:
        t = (cur + ' ' + w) if cur else w
        if f.getlength(t) * size / 140.0 * 1.03 <= avail:
            cur = t
        else:
            if cur == '':
                cur = t; continue
            # a long Thai run can still break mid-run; approximate by filling
            lines += 1
            cur = w
            while f.getlength(cur) * size / 140.0 * 1.03 > avail:
                lines += 1
                # chop roughly
                k = int(len(cur) * avail / (f.getlength(cur) * size / 140.0 * 1.03))
                cur = cur[k:]
    return lines

# ---- run / paragraph XML ----
def rpr(size=14, b=False, u=False, base=None, color='000000', hl=None, tag='a:rPr'):
    at = f'lang="th-TH" sz="{int(size*100)}"' + (' b="1"' if b else ' b="0"') + (' u="sng"' if u else '')
    if base: at += f' baseline="{base}"'
    inner = f'<a:solidFill><a:srgbClr val="{color}"/></a:solidFill>'
    if hl: inner += f'<a:highlight><a:srgbClr val="{hl}"/></a:highlight>'
    inner += f'<a:latin typeface="{FONT}" {PAN}/><a:ea typeface="{FONT}" {PAN}/><a:cs typeface="{FONT}" {PAN}/>'
    return f'<{tag} {at} dirty="0">{inner}</{tag}>'

def runs(s, size=14, b=False, color='000000', hl=None):
    out = []
    pos = 0
    u = False
    for m in re.finditer(r'<u>|</u>|\^\{([^}]*)\}|_\{([^}]*)\}', s):
        if m.start() > pos:
            out.append((s[pos:m.start()], u, None))
        tok = m.group(0)
        if tok == '<u>': u = True
        elif tok == '</u>': u = False
        elif tok.startswith('^'): out.append((m.group(1), u, '30000'))
        else: out.append((m.group(2), u, '-25000'))
        pos = m.end()
    if pos < len(s): out.append((s[pos:], u, None))
    x = ''
    for t, uu, base in out:
        if t == '': continue
        sp = ' xml:space="preserve"' if (t[:1] == ' ' or t[-1:] == ' ') else ''
        x += f'<a:r>{rpr(size, b, uu, base, color, hl)}<a:t{sp}>{escape(t)}</a:t></a:r>'
    return x

def para(content_xml, marL=0, algn='l', size=14, spc=LINE, end_b=False):
    ppr = (f'<a:pPr marL="{int(marL*E)}" indent="0" algn="{algn}" defTabSz="{int(TAB*E)}">'
           f'<a:lnSpc><a:spcPts val="{int(spc*100)}"/></a:lnSpc><a:spcBef><a:spcPts val="0"/></a:spcBef>'
           f'<a:spcAft><a:spcPts val="0"/></a:spcAft><a:buNone/></a:pPr>')
    return f'<a:p>{ppr}{content_xml}{rpr(size, end_b, tag="a:endParaRPr")}</a:p>'

def textbox_xml(sid, name, x, y, w, h, paras, autofit=True, fill=None, anchor='t', ins=(INS_L, INS_T, INS_L, INS_T)):
    fillx = f'<a:solidFill><a:srgbClr val="{fill}"/></a:solidFill>' if fill else '<a:noFill/>'
    fit = '<a:spAutoFit/>' if autofit else '<a:noAutofit/>'
    l, t, r, b = ins
    return (f'<p:sp {NS}><p:nvSpPr><p:cNvPr id="{sid}" name="{escape(name)}"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>'
            f'<p:spPr><a:xfrm><a:off x="{int(x*E)}" y="{int(y*E)}"/><a:ext cx="{int(w*E)}" cy="{int(h*E)}"/></a:xfrm>'
            f'<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>{fillx}</p:spPr>'
            f'<p:txBody><a:bodyPr wrap="square" lIns="{int(l*E)}" tIns="{int(t*E)}" rIns="{int(r*E)}" bIns="{int(b*E)}" rtlCol="0" anchor="{anchor}">{fit}</a:bodyPr>'
            f'<a:lstStyle/>{"".join(paras)}</p:txBody></p:sp>')

# ---- presentation scaffolding ----
prs = Presentation(TEMPLATE)
layout = prs.slides[0].slide_layout
# remove template example slides
sldIdLst = prs.slides._sldIdLst
for sldId in list(sldIdLst):
    prs.part.drop_rel(sldId.rId)
    sldIdLst.remove(sldId)

_sid = [100]
def nid():
    _sid[0] += 1; return _sid[0]

def new_slide():
    s = prs.slides.add_slide(layout)
    for ph in list(s.placeholders):
        ph._element.getparent().remove(ph._element)
    return s

def add_xml(slide, xml):
    el = etree.fromstring(xml)
    slide.shapes._spTree.append(el)
    return el

def add_pic(slide, path, x, y, w=None, h=None, name=None):
    pic = slide.shapes.add_picture(path, Emu(int(x*E)), Emu(int(y*E)),
                                   Emu(int(w*E)) if w else None, Emu(int(h*E)) if h else None)
    if name: pic.name = name
    return pic

def add_table(slide, spec, x, y, rowh=21.0, header_fill=None, bold_header=True, name=None):
    rows = spec['rows']
    cols = spec.get('cols')
    if spec.get('auto', True) and not spec.get('fixed'):
        cols = [max(width(str(r[j]), bold=(i == 0)) for i, r in enumerate(rows)) * 1.05 + 10 + 8 for j in range(len(rows[0]))]
        avail = TEXT_W - TAB
        if sum(cols) > avail:
            k = (sum(cols) - avail); jmax = cols.index(max(cols)); cols[jmax] -= k
    left_cols = set(spec.get('left_cols', []))
    tw = sum(cols)
    gf = slide.shapes.add_table(len(rows), len(cols), Emu(int(x*E)), Emu(int(y*E)), Emu(int(tw*E)), Emu(int(rowh*len(rows)*E)))
    if name: gf.name = name
    tbl = gf.table
    # strip theme style
    tblPr = tbl._tbl.tblPr
    for k in ('firstRow', 'bandRow'): tblPr.set(k, '0')
    sid = tblPr.find('{http://schemas.openxmlformats.org/drawingml/2006/main}tableStyleId')
    if sid is not None: sid.text = '{5940675A-B579-460E-94D1-54222C63F5DA}'  # No Style, Table Grid
    for j, cw in enumerate(cols): tbl.columns[j].width = Emu(int(cw*E))
    for i, row in enumerate(rows):
        tbl.rows[i].height = Emu(int(rowh*E))
        for j, val in enumerate(row):
            cell = tbl.cell(i, j)
            hdr = (i == 0)
            algn = 'l' if (j in left_cols and not hdr) else 'ctr'
            tc = cell._tc
            txBody = tc.find('{http://schemas.openxmlformats.org/drawingml/2006/main}txBody')
            for p in txBody.findall('{http://schemas.openxmlformats.org/drawingml/2006/main}p'): txBody.remove(p)
            for line in str(val).split('\n'):
                txBody.append(etree.fromstring(f'<root {NS}>' + para(runs(line, b=(hdr and bold_header)), algn=algn) + '</root>')[0])
            fill = header_fill if (hdr and header_fill) else 'FFFFFF'
            tcPr = tc.get_or_add_tcPr()
            for ch in list(tcPr): tcPr.remove(ch)
            tcPr.set('anchor', 'ctr'); tcPr.set('marL', str(int(5*E))); tcPr.set('marR', str(int(5*E)))
            tcPr.set('marT', str(int(1.8*E))); tcPr.set('marB', str(int(1.8*E)))
            bx = ''.join(f'<a:{s} w="{int(1*E)}" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:srgbClr val="000000"/></a:solidFill><a:prstDash val="solid"/></a:{s}>'
                         for s in ('lnL', 'lnR', 'lnT', 'lnB'))
            for el in etree.fromstring(f'<root {NS}>{bx}<a:solidFill><a:srgbClr val="{fill}"/></a:solidFill></root>'):
                tcPr.append(el)
    return gf, tw, rowh * len(rows)

def black_bar(slide, text, y=BOX_Y, h=30):
    add_xml(slide, textbox_xml(nid(), 'หัวข้อใหญ่', BOX_X, y, BOX_W, h,
                               [para(runs(text, size=20, b=True, color='FFFFFF'), algn='ctr', size=20, spc=24)],
                               autofit=False, fill='000000', anchor='ctr', ins=(7.2, 2, 7.2, 2)))


# ===================== native figure drawing =====================
class Fig:
    def __init__(self): self.s = 1.0; self.items = []
    def scale(self, s): self.s = s
    def _p(self, x, y): return (x * self.s, y * self.s)
    def line(self, x1, y1, x2, y2, arrow=0, dash=0, w=1.0):
        self.items.append(('L', self._p(x1, y1), self._p(x2, y2), arrow, dash, w))
    def rect(self, x, y, w, h, fill=None):
        (a, b), (c, d) = self._p(x, y), self._p(x + w, y + h)
        self.items.append(('R', a, b, c - a, d - b, fill, 'rect'))
    def ell(self, cx, cy, rx, ry, fill=None):
        (a, b) = self._p(cx - rx, cy - ry); (c, d) = self._p(cx + rx, cy + ry)
        self.items.append(('R', a, b, c - a, d - b, fill, 'ellipse'))
    def path(self, pts, closed=0, smooth=0, fill=None, dash=0):
        self.items.append(('P', [self._p(*p) for p in pts], closed, smooth, fill, dash))
    def arc(self, cx, cy, r, a0, a1, n=24):
        pts = [(cx + r * math.cos(math.radians(a0 + (a1 - a0) * i / n)), cy - r * math.sin(math.radians(a0 + (a1 - a0) * i / n))) for i in range(n + 1)]
        self.path(pts)
    def text(self, s, cx, cy, sz=12, al='ctr', rot=0, color='000000'):
        x, y = self._p(cx, cy)
        self.items.append(('T', s, x, y, sz, al, rot, color))
    def bbox(self):
        xs, ys = [], []
        for it in self.items:
            if it[0] == 'L': xs += [it[1][0], it[2][0]]; ys += [it[1][1], it[2][1]]
            elif it[0] == 'R': xs += [it[1], it[1] + it[3]]; ys += [it[2], it[2] + it[4]]
            elif it[0] == 'P': xs += [p[0] for p in it[1]]; ys += [p[1] for p in it[1]]
            elif it[0] == 'T':
                w = width(it[1], it[4]) + 4; h = it[4] * 1.3
                if it[6]: w, h = h, w
                x0 = it[2] - (w / 2 if it[5] == 'ctr' or it[6] else (w if it[5] == 'r' else 0))
                xs += [x0, x0 + w]; ys += [it[3] - h / 2, it[3] + h / 2]
        return min(xs), min(ys), max(xs), max(ys)

def _ln(w=1.0, arrow=0, dash=0):
    return (f'<a:ln w="{int(w*E)}"><a:solidFill><a:srgbClr val="000000"/></a:solidFill>'
            + (f'<a:prstDash val="{dash if isinstance(dash, str) else "dash"}"/>' if dash else '') + '<a:round/>'
            + ('<a:tailEnd type="triangle" w="med" len="med"/>' if arrow else '') + '</a:ln>')

def fig_xml(fig, name, X, Y):
    x0, y0, x1, y1 = fig.bbox()
    ox, oy = X - x0, Y - y0
    e = lambda v: int(round(v * E))
    kids = []
    for it in fig.items:
        k = it[0]
        if k == 'L':
            (a, b), (c, d), arrow, dash, w = it[1], it[2], it[3], it[4], it[5]
            a, b, c, d = a + ox, b + oy, c + ox, d + oy
            fl = (' flipH="1"' if c < a else '') + (' flipV="1"' if d < b else '')
            kids.append(f'<p:cxnSp><p:nvCxnSpPr><p:cNvPr id="{nid()}" name="Line"/><p:cNvCxnSpPr/><p:nvPr/></p:nvCxnSpPr>'
                        f'<p:spPr><a:xfrm{fl}><a:off x="{e(min(a,c))}" y="{e(min(b,d))}"/><a:ext cx="{e(abs(c-a))}" cy="{e(abs(d-b))}"/></a:xfrm>'
                        f'<a:prstGeom prst="line"><a:avLst/></a:prstGeom>{_ln(w, arrow, dash)}</p:spPr></p:cxnSp>')
        elif k == 'R':
            a, b, w, h, fill, prst = it[1] + ox, it[2] + oy, it[3], it[4], it[5], it[6]
            fx = f'<a:solidFill><a:srgbClr val="{fill}"/></a:solidFill>' if fill else '<a:noFill/>'
            kids.append(f'<p:sp><p:nvSpPr><p:cNvPr id="{nid()}" name="Shape"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>'
                        f'<p:spPr><a:xfrm><a:off x="{e(a)}" y="{e(b)}"/><a:ext cx="{e(w)}" cy="{e(h)}"/></a:xfrm>'
                        f'<a:prstGeom prst="{prst}"><a:avLst/></a:prstGeom>{fx}{_ln()}</p:spPr></p:sp>')
        elif k == 'P':
            pts = [(p[0] + ox, p[1] + oy) for p in it[1]]; closed, smooth, fill = it[2], it[3], it[4]
            pdash = it[5] if len(it) > 5 else 0
            segs = []
            if smooth:
                for i in range(len(pts) - 1):
                    p0 = pts[i - 1] if i else pts[i]; p1 = pts[i]; p2 = pts[i + 1]; p3 = pts[i + 2] if i + 2 < len(pts) else p2
                    segs.append(((p1[0] + (p2[0] - p0[0]) / 6, p1[1] + (p2[1] - p0[1]) / 6), (p2[0] - (p3[0] - p1[0]) / 6, p2[1] - (p3[1] - p1[1]) / 6), p2))
            allp = pts + [c for sg in segs for c in sg]   # รวมจุดควบคุมโค้ง กันพิกัดติดลบ
            px0, py0 = min(p[0] for p in allp), min(p[1] for p in allp)
            w, h = max(0.5, max(p[0] for p in allp) - px0), max(0.5, max(p[1] for p in allp) - py0)
            q = lambda p: f'<a:pt x="{e(p[0]-px0)}" y="{e(p[1]-py0)}"/>'
            dd = f'<a:moveTo>{q(pts[0])}</a:moveTo>'
            if smooth:
                dd += ''.join('<a:cubicBezTo>' + q(a) + q(b) + q(c) + '</a:cubicBezTo>' for a, b, c in segs)
            else:
                dd += ''.join(f'<a:lnTo>{q(p)}</a:lnTo>' for p in pts[1:])
            if closed: dd += '<a:close/>'
            fx = f'<a:solidFill><a:srgbClr val="{fill}"/></a:solidFill>' if fill else '<a:noFill/>'
            nf = '' if fill else ' fill="none"'
            kids.append(f'<p:sp><p:nvSpPr><p:cNvPr id="{nid()}" name="Path"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>'
                        f'<p:spPr><a:xfrm><a:off x="{e(px0)}" y="{e(py0)}"/><a:ext cx="{e(w)}" cy="{e(h)}"/></a:xfrm>'
                        f'<a:custGeom><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/><a:rect l="0" t="0" r="r" b="b"/>'
                        f'<a:pathLst><a:path w="{e(w)}" h="{e(h)}"{nf}>{dd}</a:path></a:pathLst></a:custGeom>'
                        f'{fx}{_ln(dash=pdash)}</p:spPr></p:sp>')
        elif k == 'T':
            s, cx, cy, sz, al, rot, col = it[1], it[2] + ox, it[3] + oy, it[4], it[5], it[6], it[7]
            w = width(s, sz) + 4; h = sz * 1.3
            x = cx - w / 2 if (al == 'ctr' or rot) else (cx - w if al == 'r' else cx)
            y = cy - h / 2
            ra = f' rot="{int(rot*60000)}"' if rot else ''
            kids.append(f'<p:sp><p:nvSpPr><p:cNvPr id="{nid()}" name="Label"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>'
                        f'<p:spPr><a:xfrm{ra}><a:off x="{e(x)}" y="{e(y)}"/><a:ext cx="{e(w)}" cy="{e(h)}"/></a:xfrm>'
                        f'<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr>'
                        f'<p:txBody><a:bodyPr wrap="none" lIns="0" tIns="0" rIns="0" bIns="0" anchor="ctr"><a:noAutofit/></a:bodyPr><a:lstStyle/>'
                        f'<a:p><a:pPr algn="{"ctr" if rot else al}"/>{runs(s, sz, color=col)}</a:p></p:txBody></p:sp>')
    W, H = x1 - x0, y1 - y0
    return (f'<p:grpSp {NS}><p:nvGrpSpPr><p:cNvPr id="{nid()}" name="{escape(name)}"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>'
            f'<p:grpSpPr><a:xfrm><a:off x="{e(X)}" y="{e(Y)}"/><a:ext cx="{e(W)}" cy="{e(H)}"/>'
            f'<a:chOff x="{e(X)}" y="{e(Y)}"/><a:chExt cx="{e(W)}" cy="{e(H)}"/></a:xfrm></p:grpSpPr>{"".join(kids)}</p:grpSp>')

FIGS = {}
for _k, _fn in getattr(D, 'FIGS', {}).items():
    _f = Fig(); _fn(_f); FIGS[_k] = _f
def fig_size(key):
    x0, y0, x1, y1 = FIGS[key].bbox(); return x1 - x0, y1 - y0


# ===================== body (v2) =====================
import os
TH = ['ก', 'ข', 'ค', 'ง', 'จ', 'ฉ']
TBL_ROW = 21.0
order = [c for c, _ in D.SUBTOPICS]
by_sub = {c: [] for c in order}
for qq in D.Q:
    by_sub[qq['sub']].append(qq)
BOOKS = list(getattr(D, 'BOOKS', [D.BOOK]))
for qq in D.Q: qq.setdefault('book', D.BOOK)
UID = lambda q: (q['book'], q['no'])
for c in by_sub: by_sub[c].sort(key=lambda d: (BOOKS.index(d['book']), d['no']))
newno = {}
for c in order:
    for i, qq in enumerate(by_sub[c], 1):
        newno[UID(qq)] = (c, i)
SHARED = list(getattr(D, 'SHARED', []))

def is_code(c): return c[:1].isdigit()
def head_label(c, name): return f'{c} {name}' if is_code(c) else name

def ranges(nums):
    nums = sorted(nums); out = []; i = 0
    while i < len(nums):
        j = i
        while j + 1 < len(nums) and nums[j+1] == nums[j] + 1: j += 1
        out.append(f'{nums[i]}–{nums[j]}' if j - i >= 2 else (f'{nums[i]}, {nums[j]}' if j > i else f'{nums[i]}'))
        i = j + 1
    return 'ข้อ ' + ', '.join(out)

CACHE = json.load(open(OUT + '.lines.json')) if os.path.exists(OUT + '.lines.json') else {}
BREAKS = json.load(open(OUT + '.breaks.json')) if os.path.exists(OUT + '.breaks.json') else {}
def _nz(t):
    return ''.join('\u0e32' if c == '\u0e33' else c for c in t if c not in '\u0e4d \t\u00a0\u200b|\u00a6\x00\n')
def apply_breaks(m, frags):
    """แทนช่องว่างที่อยู่ก่อนวลีที่ถูกตัด ด้วยตัวขึ้นบรรทัดใหม่ (\x00)"""
    depth = 0; seen = {}
    want = [(f.split('#')[0], int(f.split('#')[1]) if '#' in f else None) for f in frags]   # 'วลี#n' = เฉพาะครั้งที่ n (นับจาก 0)
    for i, c in enumerate(m):
        if c == '{': depth += 1
        elif c == '}': depth -= 1
        elif c == ' ' and depth == 0 and i > 3:
            rest = _nz(plain(m[i + 1:]))
            hit = False
            for f, n in want:
                if rest.startswith(f):
                    cnt = seen.get(f, 0); seen[f] = cnt + 1
                    if n is None or n == cnt: hit = True
            if hit: m = m[:i] + '\x00' + m[i + 1:]
    return m
def _runs_parts(m, size, b, hl):
    parts = m.split('\x00'); out = []; carry = False
    for k, pt in enumerate(parts):
        if carry: pt = '<u>' + pt
        carry = pt.count('<u>') > pt.count('</u>')
        if carry: pt = pt + '</u>'
        out.append(runs(pt, size, b, hl=hl))
    return f'<a:br>{rpr(size, b)}</a:br>'.join(out)
def P(text, marL=0, size=14, b=False, hl=None, prefix='', tail=None, tabs=None):
    key = prefix + plain(text)
    tail_inline = False
    if tail:
        full = prefix + plain(text) + ' ' + plain(tail)
        if width(full, size, b) * 1.03 > TEXT_W - marL: key = key + '|' + plain(tail)
        else: key = full; tail_inline = True
    m = prefix + text
    frags = BREAKS.get(key)
    if frags:
        m = apply_breaks(m, frags)
        bk = plain(m).replace('\x00', '\u00a6')
        key = bk + (' ' + plain(tail) if tail_inline else ('|' + plain(tail) if tail else ''))
    content = _runs_parts(m, size, b, hl)
    if tail:
        if tail_inline: content += runs(' ' + tail, size, b, hl=hl)
        else: content += f'<a:br>{rpr(size, b)}</a:br>' + runs(tail, size, b, hl=hl)
    nl = CACHE.get(key) or (n_lines(key.replace('|', ' ').replace('\u00a6', ' '), TEXT_W - marL, size, b)
                            + key.count('|') + key.count('\u00a6'))
    xml = para(content, marL=marL, size=size)
    if tabs:
        tl = '<a:tabLst>' + ''.join(f'<a:tab pos="{int(t*E)}" algn="l"/>' for t in tabs) + '</a:tabLst>'
        xml = xml.replace('<a:buNone/></a:pPr>', '<a:buNone/>' + tl + '</a:pPr>', 1)
    return ('p', xml, nl, key)
BLANK = lambda: ('p', para(''), 1, None)

def nlines_of(items): return sum(i[2] for i in items)

def table_cols(spec, avail):
    rows = spec['rows']
    if spec.get('fixed'): return list(spec['cols'])
    cols = [max(max(width(x, bold=(i == 0)) for x in str(r[j] or '').split('\n')) for i, r in enumerate(rows)) * 1.06 + 16 for j in range(len(rows[0]))]
    if sum(cols) > avail:
        extra = sum(cols) - avail
        jmax = cols.index(max(cols)); cols[jmax] -= extra
    return cols

def table_height(spec, avail):
    """row heights: grow a row when a cell wraps"""
    cols = table_cols(spec, avail); hs = []
    for i, r in enumerate(spec['rows']):
        n = max(sum(n_lines(x, cw - 10, bold=(i == 0)) for x in str(v or '').split('\n')) for v, cw in zip(r, cols))
        hs.append(max(TBL_ROW, 4 + n * 17.5))
    return cols, hs

def add_obj_table(items, objs, spec, dx):
    avail = TEXT_W - dx
    cols, hs = table_height(spec, avail)
    th = sum(hs)
    objs.append(dict(kind='table', spec=dict(spec, cols=cols, fixed=True, heights=hs), at=nlines_of(items), dx=dx, dy=4))
    items += [BLANK() for _ in range(math.ceil((th + 10) / LINE))]

def add_obj_pic(items, objs, key, h, dx=TAB):
    if key in FIGS:
        w, h = fig_size(key)
        objs.append(dict(kind='draw', key=key, at=nlines_of(items), dx=dx, dy=4, w=w, h=h))
        items += [BLANK() for _ in range(math.ceil((h + 8) / LINE))]
        return
    im = Image.open(f'{IMG_DIR}/{key}.png'); w = im.width / im.height * h
    if dx + w > TEXT_W:
        w = TEXT_W - dx; h = w * im.height / im.width
    objs.append(dict(kind='pic', key=key, at=nlines_of(items), dx=dx, dy=3, w=w, h=h))
    items += [BLANK() for _ in range(math.ceil((h + 6) / LINE))]

def add_info(items, lines, marL=TAB):
    lines = [re.sub(r' {3,}', '\t', l) for l in lines]
    if any('\t' in l for l in lines):
        cols = [l.split('\t') for l in lines]
        nc = max(len(c) for c in cols)
        wmax = [max((width(c[k]) for c in cols if len(c) > k), default=0) for k in range(nc)]
        stops, acc = [], marL
        for k in range(nc - 1):
            acc += wmax[k] + 18; stops.append(acc)
        for l in lines: items.append(P(l, marL=marL, tabs=stops))
    else:
        for l in lines: items.append(P(l, marL=marL))

def question_block(qq):
    c, n = newno[UID(qq)]
    src = f"({qq['book']}, ข้อ {qq['no']})"
    items, objs = [], []
    # ผู้ใช้กำหนด (15 พ.ย.): รูปต้องอยู่ตามลำดับการอ่าน — ไม่วางรูปข้างขวาตัวเลือก ให้วางใต้ข้อความที่อ้างถึงรูป
    if qq.get('side_img') and not (qq.get('mid_img') or qq.get('top_img')):
        qq = dict(qq, mid_img=qq['side_img']); qq.pop('side_img')
    # ที่มา (หนังสือ, ข้อ) อยู่ท้ายข้อความคำถามส่วนสุดท้าย ก่อนตัวเลือก
    last = 'stem3' if qq.get('stem3') else ('stem2' if qq.get('stem2') else 'stem')
    tl = lambda k: src if k == last else None
    items.append(P(qq['stem'] or '', prefix=f'{n}. ', tail=tl('stem')))
    add_info(items, qq.get('info', []))
    if qq.get('stem_b'): items.append(P(qq['stem_b']))   # ข้อความหลังข้อมูล ก่อนรูป/ตาราง
    add_info(items, qq.get('info_b', []))   # ข้อมูลเยื้องหลัง stem_b ก่อนตาราง
    if qq.get('mid_img') or qq.get('top_img'):
        add_obj_pic(items, objs, qq.get('mid_img') or qq.get('top_img'), qq.get('img_h', 80))
    if qq.get('table'): add_obj_table(items, objs, qq['table'], TAB)
    add_info(items, qq.get('info_after', []))
    if qq.get('table2'): add_obj_table(items, objs, qq['table2'], TAB)
    if qq.get('stem2'): items.append(P(qq['stem2'], tail=tl('stem2')))
    add_info(items, qq.get('info2', []))
    if qq.get('stem3'): items.append(P(qq['stem3'], tail=tl('stem3')))
    if qq.get('img_opts'):
        keys = qq['img_opts']; hh = qq.get('img_h', 58)
        if all(k in FIGS for k in keys): hh = max(fig_size(k)[1] for k in keys)
        cols = qq.get('img_cols', len(keys)); nrow = math.ceil(len(keys) / cols)   # img_cols=2 → ตัวเลือกรูปแบบ 2×2
        rl = math.ceil((hh + 6) / LINE)
        objs.append(dict(kind='imgrow', keys=keys, at=nlines_of(items), h=hh, cols=cols, rl=rl))
        items += [BLANK() for _ in range(rl * nrow)]
        for k, t in enumerate(qq.get('text_opts_after', [])):
            items.append(P(t, prefix=f'{TH[len(keys)+k]}. '))
    if qq.get('opts_table'):
        spec = dict(qq['opts_table']); spec.setdefault('left_cols', [])
        add_obj_table(items, objs, spec, 0)
    if qq.get('opts'):
        start = nlines_of(items)
        for k, t in enumerate(qq['opts']):
            items.append(P(t, prefix=f'{TH[k]}. '))
        if qq.get('side_img'):
            key = qq['side_img']
            if key in FIGS:
                ww, hh = fig_size(key)
                objs.append(dict(kind='draw', key=key, at=start, dx=TEXT_W - ww, dy=2, w=ww, h=hh))
            else:
                im = Image.open(f'{IMG_DIR}/{key}.png')
                hh = qq.get('img_h', 4 * LINE + 10); ww = im.width / im.height * hh
                objs.append(dict(kind='pic', key=key, at=start, dx=TEXT_W - ww, dy=0, w=ww, h=hh))
            need = math.ceil((hh + 4) / LINE) - len(qq['opts'])
            items += [BLANK() for _ in range(max(0, need))]
    if qq.get('subjective') and not qq.get('opts'):
        pass
    return items, objs

def shared_for(qq):
    for sh in SHARED:
        if sh.get('book', D.BOOK) == qq['book'] and sh['first'] <= qq['no'] <= sh['last']: return sh
    return None

def shared_block(sh, grp):
    nums = [newno[UID(g)][1] for g in grp]
    a, b = nums[0], nums[-1]
    txt = sh['text'].format(a=a, b=b)
    if a == b: txt = re.sub(r'ข้อ \d+\s*[–-]\s*\d+', f'ข้อ {a}', txt).replace(f'ข้อ {a} และ {a}', f'ข้อ {a}')
    elif b == a + 1: txt = txt.replace(f'{a}–{b}', f'{a} และ {b}')
    it, objs = [P('คำชี้แจง ' + txt, b=True)], []
    for line in sh.get('lines', []): it.append(P(line, marL=TAB))
    if sh.get('img'): add_obj_pic(it, objs, sh['img'], sh.get('img_h', 90))
    if sh.get('table'): add_obj_table(it, objs, sh['table'], TAB)
    add_info(it, sh.get('lines_after', []))
    for gi, gq in enumerate(grp):
        qi, qo = question_block(gq)
        base = nlines_of(it)
        for o in qo: o['at'] += base
        objs += qo; it += qi
        if gi < len(grp) - 1: it.append(BLANK())
    return it, objs

def heading_items(c, name, count):
    it = [P(f'{head_label(c, name)} จำนวน {count} ข้อ', size=16, b=True, hl='FFF2CC')]
    if count == 0: it.append(P('ไม่พบข้อสอบในไฟล์ที่ได้รับ'))
    return it

blocks = []
for c, name in D.SUBTOPICS:
    qs = by_sub[c]
    if not is_code(c) and not qs: continue
    head = heading_items(c, name, len(qs))
    if not qs: blocks.append((head, [])); continue
    first = True; i = 0
    while i < len(qs):
        qq = qs[i]; sh = shared_for(qq)
        if sh:
            grp = [x for x in qs[i:] if shared_for(x) is sh]
            it, objs = shared_block(sh, grp); i += len(grp)
        else:
            it, objs = question_block(qq); i += 1
        if first:
            base = nlines_of(head)
            for o in objs: o['at'] += base
            it = list(head) + it; first = False
        blocks.append((it, objs))

MAXL = int((BOTTOM - BOX_Y - INS_T) // LINE)
FLOW_KEYS = []
def render_table(slide, spec, x, y, name, header_fill=None):
    rows, cols = spec['rows'], spec['cols']
    hs = spec.get('heights') or [TBL_ROW] * len(rows)
    left_cols = set(spec.get('left_cols', []))
    gf = slide.shapes.add_table(len(rows), len(cols), Emu(int(x*E)), Emu(int(y*E)), Emu(int(sum(cols)*E)), Emu(int(sum(hs)*E)))
    gf.name = name
    tbl = gf.table; tblPr = tbl._tbl.tblPr
    for k in ('firstRow', 'bandRow'): tblPr.set(k, '0')
    for j, cw in enumerate(cols): tbl.columns[j].width = Emu(int(cw*E))
    A = '{http://schemas.openxmlformats.org/drawingml/2006/main}'
    for i, row in enumerate(rows):
        tbl.rows[i].height = Emu(int(hs[i]*E))
        for j, val in enumerate(row):
            hdr = (i == 0) and not spec.get('noheader'); tc = tbl.cell(i, j)._tc
            blank = val is None; val = '' if blank else val
            if isinstance(val, str) and val.startswith('~'): blank = True; val = val[1:]
            algn = 'l' if (j in left_cols and not hdr) else 'ctr'
            txBody = tc.find(A + 'txBody')
            for p in txBody.findall(A + 'p'): txBody.remove(p)
            for _ln in str(val).split('\n'):
                txBody.append(etree.fromstring(f'<root {NS}>' + para(runs(_ln, b=hdr), algn=algn) + '</root>')[0])
            tcPr = tc.get_or_add_tcPr()
            for ch in list(tcPr): tcPr.remove(ch)
            tcPr.set('anchor', 'ctr'); tcPr.set('marL', str(int(5*E))); tcPr.set('marR', str(int(5*E)))
            tcPr.set('marT', str(int(1.8*E))); tcPr.set('marB', str(int(1.8*E)))
            fill = header_fill if (hdr and header_fill) else 'FFFFFF'
            if blank:
                def bordered(ii, jj):
                    if ii < 0 or jj < 0 or ii >= len(rows) or jj >= len(rows[ii]): return False
                    v = rows[ii][jj]
                    return not (v is None or (isinstance(v, str) and v.startswith('~')))
                nb = {'lnL': bordered(i, j - 1), 'lnR': bordered(i, j + 1), 'lnT': bordered(i - 1, j), 'lnB': bordered(i + 1, j)}
                ln_on = lambda t: f'<a:{t} w="{int(1*E)}" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:srgbClr val="000000"/></a:solidFill><a:prstDash val="solid"/></a:{t}>'
                xx = ''.join(ln_on(t) if nb[t] else f'<a:{t}><a:noFill/></a:{t}>' for t in ('lnL', 'lnR', 'lnT', 'lnB'))
                for el in etree.fromstring(f'<root {NS}>' + xx + '<a:noFill/></root>'):
                    tcPr.append(el)
                continue
            bx = ''.join(f'<a:{s} w="{int(1*E)}" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:srgbClr val="000000"/></a:solidFill><a:prstDash val="solid"/></a:{s}>'
                         for s in ('lnL', 'lnR', 'lnT', 'lnB'))
            for el in etree.fromstring(f'<root {NS}>{bx}<a:solidFill><a:srgbClr val="{fill}"/></a:solidFill></root>'):
                tcPr.append(el)
    return gf

def render_flow(slide, items, objs, y0=BOX_Y):
    FLOW_KEYS.extend([i[3] for i in items if i[3]])
    add_xml(slide, textbox_xml(nid(), 'กล่องข้อความหลัก', BOX_X, y0, BOX_W,
                               nlines_of(items) * LINE + 2 * INS_T, [i[1] for i in items]))
    for o in objs:
        top = y0 + INS_T + o['at'] * LINE
        if o['kind'] == 'draw':
            add_xml(slide, fig_xml(FIGS[o['key']], f"รูป {o['key']}", BOX_X + INS_L + o['dx'], top + o['dy']))
        elif o['kind'] == 'pic':
            add_pic(slide, f"{IMG_DIR}/{o['key']}.png", BOX_X + INS_L + o['dx'], top + o['dy'], o['w'], o['h'], name=f"รูป {o['key']}")
        elif o['kind'] == 'table':
            render_table(slide, o['spec'], BOX_X + INS_L + o['dx'], top + o['dy'], 'ตาราง')
        elif o['kind'] == 'imgrow':
            cols = o.get('cols', len(o['keys'])); slot = TEXT_W / cols; top0 = top
            for k, key in enumerate(o['keys']):
                r_, c_ = divmod(k, cols); top = top0 + r_ * o.get('rl', 0) * LINE
                sx = BOX_X + INS_L + c_ * slot
                add_xml(slide, textbox_xml(nid(), f'ป้าย {TH[k]}', sx, top, 24, LINE + 4, [para(runs(f'{TH[k]}.'))], ins=(0, 0, 0, 0)))
                if key in FIGS:
                    add_xml(slide, fig_xml(FIGS[key], f'รูปตัวเลือก {TH[k]}', sx + 22, top + 3)); continue
                im = Image.open(f'{IMG_DIR}/{key}.png'); hh = o['h']; ww = im.width / im.height * hh
                if ww > slot - 26: ww = slot - 26; hh = ww * im.height / im.width
                add_pic(slide, f'{IMG_DIR}/{key}.png', sx + 22, top + 3, ww, hh, name=f'รูปตัวเลือก {TH[k]}')

def paginate(blocks, maxl):
    pages, cur, curobjs, used = [], [], [], 0
    for items, objs in blocks:
        n = nlines_of(items)
        gap = 1 if cur else 0
        if cur and used + gap + n > maxl:
            pages.append((cur, curobjs)); cur, curobjs, used, gap = [], [], 0, 0
        if gap: cur.append(BLANK()); used += 1
        for o in objs:
            o = dict(o); o['at'] += used; curobjs.append(o)
        cur += items; used += n
    if cur: pages.append((cur, curobjs))
    return pages

# ---- title + index ----
s = new_slide()
black_bar(s, f'{D.TOPIC_NO}. {D.TOPIC_NAME}')
y = BOX_Y + 30 + 6
intro = [P(f'คลังข้อสอบเคมี จำนวน {len(D.Q)} ข้อ', size=16, b=True),
         P('อินเด็กซ์อ้างเลขข้อเดิมในหนังสือ ส่วนข้อสอบและเฉลยเริ่มเลขใหม่ในทุกหัวข้อย่อย')]
if getattr(D, 'NOTE_INDEX', None): intro.append(P(D.NOTE_INDEX))
add_xml(s, textbox_xml(nid(), 'คำอธิบาย', BOX_X, y, BOX_W, nlines_of(intro) * LINE + 2 * INS_T, [i[1] for i in intro]))
y += nlines_of(intro) * LINE + 2 * INS_T + 6
rows = [['หัวข้อ'] + BOOKS + ['จำนวนข้อ']]
for c, name in D.SUBTOPICS:
    qs_ = by_sub[c]
    if not is_code(c) and not qs_: continue
    cells = []
    for bk in BOOKS:
        nums = [qq['no'] for qq in qs_ if qq['book'] == bk]
        cells.append(ranges(nums) if nums else '–')
    rows.append([head_label(c, name)] + cells + [str(len(qs_))])
rows.append(['รวม'] + [str(sum(1 for q in D.Q if q['book'] == bk)) for bk in BOOKS] + [str(len(D.Q))]
            if len(BOOKS) > 1 else ['รวม', '', str(len(D.Q))])
w0 = min(max(width(r[0]) for r in rows) + 18, TEXT_W * 0.40)
w2 = width('จำนวนข้อ', bold=True) + 20
wb = (TEXT_W - w0 - w2) / len(BOOKS)
IDX_COLS = [w0] + [wb] * len(BOOKS) + [w2]
_, hs = table_height(dict(rows=rows, cols=IDX_COLS, fixed=True), TEXT_W)
def idx_page(slide, ytop, part_rows, part_h):
    render_table(slide, dict(rows=[rows[0]] + part_rows, cols=IDX_COLS, heights=[hs[0]] + part_h, left_cols=[0] + list(range(1, len(BOOKS) + 1))),
                 BOX_X + INS_L, ytop, 'อินเด็กซ์', header_fill='F2F2F2')
body, bh = rows[1:], hs[1:]
while body:
    room = BOTTOM - y - hs[0]; k = 0; acc = 0
    while k < len(body) and acc + bh[k] <= room: acc += bh[k]; k += 1
    idx_page(s, y, body[:k], bh[:k]); body, bh = body[k:], bh[k:]
    if body:
        s = new_slide()
        add_xml(s, textbox_xml(nid(), 'หัวเรื่อง', BOX_X, BOX_Y, BOX_W, LINE + 2 * INS_T, [P('อินเด็กซ์ต่อ', size=16, b=True)[1]]))
        y = BOX_Y + LINE + 2 * INS_T + 6

qpages = paginate(blocks, MAXL)
for items, objs in qpages:
    render_flow(new_slide(), items, objs)

# ---- answer key ----
def ans_text(i, qq):
    if qq.get('ans_note'): return f'{i}. {qq["ans_note"]}'
    if qq.get('self_ans'): return f'{i}. {TH[qq["ans"]-1]}* (เฉลยเอง: {qq["self_ans"]})'
    if qq.get('ans'): return f'{i}. {TH[qq["ans"]-1]}'
    return f'{i}. รอตรวจ (ไม่มีเฉลยในไฟล์ต้นฉบับ)'
def paginate_answers(maxl):
    # เฉลยไหลต่อเนื่อง: หัวข้อย่อยที่ยาวเกินหน้าจะแบ่งต่อหน้าถัดไปพร้อมหัวเรื่อง "(ต่อ)"
    pages, cur, used = [], [], 0
    def flush():
        nonlocal cur, used
        if cur: pages.append((cur, []))
        cur, used = [], 0
    for c, name in D.SUBTOPICS:
        qs = by_sub[c]
        if not is_code(c) and not qs: continue
        head = P(head_label(c, name), size=16, b=True, hl='FFF2CC')
        body = [P('ไม่พบข้อสอบในไฟล์ที่ได้รับ')] if not qs else [P(ans_text(i, qq)) for i, qq in enumerate(qs, 1)]
        hn = nlines_of([head])
        total = hn + nlines_of(body)
        gap = 1 if cur else 0
        if cur and (used + gap + total > maxl) and (maxl - used - gap < hn + 3):
            flush(); gap = 0
        if gap: cur.append(BLANK()); used += 1
        cur.append(head); used += hn; k = 0
        while k < len(body):
            n = nlines_of([body[k]])
            if used + n > maxl:
                flush()
                h2 = P(head_label(c, name) + ' (ต่อ)', size=16, b=True, hl='FFF2CC')
                cur.append(h2); used += nlines_of([h2])
            cur.append(body[k]); used += n; k += 1
    flush()
    return pages
bar_h = 30; first = True
for items, objs in paginate_answers(int((BOTTOM - BOX_Y - bar_h - 8 - INS_T) // LINE)):
    s = new_slide()
    if first: black_bar(s, 'เฉลย'); render_flow(s, items, objs, y0=BOX_Y + bar_h + 8); first = False
    else: render_flow(s, items, objs)

prs.save(OUT)
log = []
for qq in sorted(D.Q, key=lambda d: (BOOKS.index(d['book']), d['no'])):
    c, n = newno[UID(qq)]
    log.append(dict(src=f"{qq['book']} ข้อ {qq['no']}", sub=c, new=n, ans=ans_text(n, qq), note=qq.get('edit_note')))
json.dump(FLOW_KEYS, open(OUT + '.keys.json', 'w'), ensure_ascii=False)
json.dump(dict(slides=len(prs.slides), qpages=len(qpages), map=log), open(OUT + '.map.json', 'w'), ensure_ascii=False, indent=1)
print('slides', len(prs.slides), 'question pages', len(qpages), 'questions', len(D.Q))
