"""
NeuraSlide v5 — Professional PPT Engine
10 smart layouts · Rich components · Deliverable quality
"""
import io, os, re, tempfile, hashlib, requests
from PIL import Image, ImageDraw, ImageFilter, ImageEnhance
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN

SW, SH = 10.0, 5.625   # slide width/height in inches

# ═══════════════════════════════════════════════════════════════════════════════
#  PROFESSIONAL THEME SYSTEM
# ═══════════════════════════════════════════════════════════════════════════════
THEMES = {
    "🌌 Midnight Executive": {
        "bg":       (7,   7,  19),
        "surface":  (14,  12, 34),
        "raised":   (22,  18, 52),
        "border":   (38,  30, 85),
        "accent":   (124, 58,237),
        "accent2":  (167,139,250),
        "accent3":  (56, 189,248),
        "gold":     (245,158, 11),
        "success":  (16, 185,129),
        "danger":   (239, 68, 68),
        "title":    (248,246,255),
        "body":     (208,203,242),
        "muted":    (138,118,195),
        "dim":      (80,  65,140),
        "ex_bg":    (18,  13, 46),
        "tag_bg":   (124, 58,237),
        "tag_fg":   (248,246,255),
    },
    "🌊 Ocean Deep": {
        "bg":       (2,  18, 42),
        "surface":  (3,  32, 65),
        "raised":   (5,  48, 90),
        "border":   (8,  65,118),
        "accent":   (6, 182,212),
        "accent2":  (103,232,249),
        "accent3":  (34, 211,238),
        "gold":     (251,191, 36),
        "success":  (52, 211,153),
        "danger":   (251,113,133),
        "title":    (236,254,255),
        "body":     (178,234,252),
        "muted":    (98, 188,222),
        "dim":      (55, 120,160),
        "ex_bg":    (2,  24, 55),
        "tag_bg":   (6, 182,212),
        "tag_fg":   (2,  18, 42),
    },
    "🌿 Forest & Moss": {
        "bg":       (6,  26, 12),
        "surface":  (10, 40, 18),
        "raised":   (16, 56, 26),
        "border":   (22, 78, 35),
        "accent":   (34,197, 94),
        "accent2":  (74,222,128),
        "accent3":  (52,211,153),
        "gold":     (250,204, 21),
        "success":  (34,197, 94),
        "danger":   (248,113,113),
        "title":    (240,253,244),
        "body":     (187,247,208),
        "muted":    (110,190,135),
        "dim":      (65, 130, 88),
        "ex_bg":    (8,  32, 14),
        "tag_bg":   (34,197, 94),
        "tag_fg":   (6,  26, 12),
    },
    "🔥 Coral Energy": {
        "bg":       (65,   8,  8),
        "surface":  (95,  12, 12),
        "raised":   (125, 16, 16),
        "border":   (160, 24, 24),
        "accent":   (239, 68, 68),
        "accent2":  (251,146, 60),
        "accent3":  (252,211, 77),
        "gold":     (253,224, 71),
        "success":  (74, 222,128),
        "danger":   (239, 68, 68),
        "title":    (255,245,235),
        "body":     (254,215,195),
        "muted":    (202,140,120),
        "dim":      (145, 85, 72),
        "ex_bg":    (75,  10, 10),
        "tag_bg":   (239, 68, 68),
        "tag_fg":   (255,245,235),
    },
    "⬜ Clean Minimal": {
        "bg":       (10,  10, 22),
        "surface":  (16,  16, 35),
        "raised":   (24,  24, 52),
        "border":   (36,  36, 75),
        "accent":   (99,  91,255),
        "accent2":  (139,148,255),
        "accent3":  (6,  182,212),
        "gold":     (245,158, 11),
        "success":  (52, 211,153),
        "danger":   (248,113,133),
        "title":    (248,248,255),
        "body":     (203,203,240),
        "muted":    (130,130,178),
        "dim":      (78,  78,130),
        "ex_bg":    (14,  14, 32),
        "tag_bg":   (99,  91,255),
        "tag_fg":   (248,248,255),
    },
    "🌸 Berry & Cream": {
        "bg":       (42,   9, 85),
        "surface":  (60,  14,118),
        "raised":   (80,  22,155),
        "border":   (105, 32,195),
        "accent":   (168, 85,247),
        "accent2":  (232,121,249),
        "accent3":  (236, 72,153),
        "gold":     (251,191, 36),
        "success":  (74, 222,128),
        "danger":   (248,113,133),
        "title":    (253,245,255),
        "body":     (228,198,253),
        "muted":    (168,128,215),
        "dim":      (110, 75,165),
        "ex_bg":    (48,  11, 98),
        "tag_bg":   (168, 85,247),
        "tag_fg":   (253,245,255),
    },
    "🌑 Carbon Dark": {
        "bg":       (8,   8,  10),
        "surface":  (18,  18, 22),
        "raised":   (28,  28, 35),
        "border":   (45,  45, 58),
        "accent":   (14, 165,233),
        "accent2":  (56, 189,248),
        "accent3":  (34, 211,238),
        "gold":     (251,191, 36),
        "success":  (52, 211,153),
        "danger":   (248,113,133),
        "title":    (248,250,252),
        "body":     (203,213,225),
        "muted":    (130,145,165),
        "dim":      (75,  90,108),
        "ex_bg":    (12,  12, 16),
        "tag_bg":   (14, 165,233),
        "tag_fg":   (8,   8,  10),
    },
    "🟣 Royal Purple": {
        "bg":       (20,   5, 50),
        "surface":  (32,  10, 75),
        "raised":   (48,  16,108),
        "border":   (68,  26,148),
        "accent":   (192,132,252),
        "accent2":  (216,180,254),
        "accent3":  (196,181,253),
        "gold":     (251,191, 36),
        "success":  (74, 222,128),
        "danger":   (248,113,133),
        "title":    (253,250,255),
        "body":     (233,213,255),
        "muted":    (168,148,210),
        "dim":      (110, 90,155),
        "ex_bg":    (25,   7, 62),
        "tag_bg":   (192,132,252),
        "tag_fg":   (20,   5, 50),
    },
}

# ═══════════════════════════════════════════════════════════════════════════════
#  IMAGE SYSTEM
# ═══════════════════════════════════════════════════════════════════════════════
_CACHE = {}

def _fetch(kw, w=1280, h=720, enabled=True):
    if not enabled: return None
    k = hashlib.md5(f"{kw}{w}{h}".encode()).hexdigest()
    if k in _CACHE: return _CACHE[k]
    try:
        r = requests.get(f"https://source.unsplash.com/{w}x{h}/?{kw.replace(' ',',')}",
                         timeout=15, allow_redirects=True)
        if r.status_code == 200 and len(r.content) > 5000:
            _CACHE[k] = r.content; return r.content
    except: pass
    return None

def _grad(w, h, *colors):
    """Multi-stop gradient."""
    img = Image.new("RGB",(w,h)); d = ImageDraw.Draw(img)
    n = len(colors)
    for y in range(h):
        t = y / h
        seg = min(int(t * (n-1)), n-2)
        lt = t * (n-1) - seg
        c = tuple(int(colors[seg][i]*(1-lt) + colors[seg+1][i]*lt) for i in range(3))
        d.line([(0,y),(w,y)], fill=c)
    buf = io.BytesIO(); img.save(buf,"PNG"); return buf.getvalue()

def _dark(ib, op=0.55, c=(0,0,0)):
    img = Image.open(io.BytesIO(ib)).convert("RGBA")
    ov = Image.new("RGBA", img.size, (*c, int(255*op)))
    out = Image.alpha_composite(img, ov).convert("RGB")
    buf = io.BytesIO(); out.save(buf,"PNG"); return buf.getvalue()

def _tint(ib, c, op=0.25):
    img = Image.open(io.BytesIO(ib)).convert("RGBA")
    ov = Image.new("RGBA", img.size, (*c, int(255*op)))
    out = Image.alpha_composite(img, ov).convert("RGB")
    buf = io.BytesIO(); out.save(buf,"PNG"); return buf.getvalue()

def _blur(ib, r=6):
    img = Image.open(io.BytesIO(ib)).filter(ImageFilter.GaussianBlur(r))
    buf = io.BytesIO(); img.save(buf,"PNG"); return buf.getvalue()

def _brighten(ib, factor=1.15):
    img = ImageEnhance.Brightness(Image.open(io.BytesIO(ib))).enhance(factor)
    buf = io.BytesIO(); img.save(buf,"PNG"); return buf.getvalue()

def _fade(w, h, color, direction="right"):
    img = Image.new("RGBA",(w,h),(0,0,0,0)); d = ImageDraw.Draw(img)
    for i in range(w if direction in("right","left") else h):
        t = i / (w if direction in("right","left") else h)
        if direction in ("right","down"): a = int(255*(1-t))
        else: a = int(255*t)
        if direction in ("right","left"): d.line([(i,0),(i,h)], fill=(*color,a))
        else: d.line([(0,i),(w,i)], fill=(*color,a))
    buf = io.BytesIO(); img.save(buf,"PNG"); return buf.getvalue()

def _place(sl, data, x, y, w, h):
    tmp = tempfile.NamedTemporaryFile(suffix=".png", delete=False)
    tmp.write(data); tmp.flush(); tmp.close()
    sl.shapes.add_picture(tmp.name, Inches(x), Inches(y), Inches(w), Inches(h))
    os.unlink(tmp.name)

def _get_bg(kw, th, w=1280, h=720, dark_op=0.55, tint_op=0.22, imgs=True, enabled=True):
    raw = _fetch(kw, w, h, imgs)
    if raw:
        return _dark(_tint(raw, th["accent"], tint_op), op=dark_op)
    return _grad(w, h, th["bg"], th["surface"], th["raised"])

# ═══════════════════════════════════════════════════════════════════════════════
#  DRAWING PRIMITIVES
# ═══════════════════════════════════════════════════════════════════════════════
def _c(t): return RGBColor(*t)

def _rect(sl, x, y, w, h, fill, lc=None, lw=0.5):
    s = sl.shapes.add_shape(1, Inches(x), Inches(y), Inches(w), Inches(h))
    s.fill.solid(); s.fill.fore_color.rgb = _c(fill)
    if lc: s.line.color.rgb = _c(lc); s.line.width = Pt(lw)
    else: s.line.fill.background()
    return s

def _rrect(sl, x, y, w, h, fill, adj=0.07):
    s = sl.shapes.add_shape(5, Inches(x), Inches(y), Inches(w), Inches(h))
    s.fill.solid(); s.fill.fore_color.rgb = _c(fill); s.line.fill.background()
    s.adjustments[0] = adj; return s

def _circle(sl, cx, cy, r, fill):
    s = sl.shapes.add_shape(9, Inches(cx-r/2), Inches(cy-r/2), Inches(r), Inches(r))
    s.fill.solid(); s.fill.fore_color.rgb = _c(fill); s.line.fill.background()
    return s

def _line(sl, x1, y1, x2, y2, color, width=0.5):
    w = abs(x2-x1) or 0.001; h = abs(y2-y1) or 0.001
    s = sl.shapes.add_shape(1, Inches(min(x1,x2)), Inches(min(y1,y2)), Inches(w), Inches(h))
    s.fill.background(); s.line.color.rgb = _c(color); s.line.width = Pt(width)

def _txt(sl, text, x, y, w, h, size, color, bold=False, italic=False,
         align=PP_ALIGN.LEFT, name="Calibri"):
    if not str(text).strip(): return
    b = sl.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = b.text_frame; tf.word_wrap = True
    p = tf.paragraphs[0]; p.alignment = align
    r = p.add_run()
    r.text = str(text); r.font.size = Pt(size)
    r.font.bold = bold; r.font.italic = italic
    r.font.color.rgb = _c(color); r.font.name = name

def _mltext(sl, text, x, y, w, h, size, color, bold=False, italic=False,
            align=PP_ALIGN.LEFT, spacing=5, name="Calibri"):
    if not str(text).strip(): return
    b = sl.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = b.text_frame; tf.word_wrap = True
    for i, line in enumerate(str(text).split("\n")):
        if not line.strip(): continue
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align; p.space_before = Pt(spacing)
        r = p.add_run()
        r.text = line; r.font.size = Pt(size)
        r.font.bold = bold; r.font.italic = italic
        r.font.color.rgb = _c(color); r.font.name = name

def _bullets(sl, text, x, y, w, h, size, color, dot_c, gap=6, indent=0.22):
    if not str(text).strip(): return
    b = sl.shapes.add_textbox(Inches(x+indent), Inches(y), Inches(w-indent), Inches(h))
    tf = b.text_frame; tf.word_wrap = True
    lines = [l.lstrip("•●▪▸-–— ").strip() for l in str(text).split("\n") if l.strip()]
    rh = size/72 + gap/72 + 0.048
    for i, line in enumerate(lines):
        p = tf.paragraphs[0] if i==0 else tf.add_paragraph()
        p.space_before = Pt(gap)
        r = p.add_run(); r.text = line
        r.font.size = Pt(size); r.font.color.rgb = _c(color); r.font.name = "Calibri"
        # Dot
        dy = y + i*rh + 0.06
        try:
            d = sl.shapes.add_shape(9, Inches(x+0.04), Inches(dy), Inches(0.09), Inches(0.09))
            d.fill.solid(); d.fill.fore_color.rgb = _c(dot_c); d.line.fill.background()
        except: pass

# ── Tag pill ──────────────────────────────────────────────────────────────────
def _tag(sl, text, x, y, th, w=1.8, h=0.26):
    _rrect(sl, x, y, w, h, th["tag_bg"], 0.5)
    _txt(sl, text.upper(), x+0.1, y+0.03, w-0.15, h-0.06,
         7.5, th["tag_fg"], bold=True, align=PP_ALIGN.CENTER)

# ── Page number / footer bar ──────────────────────────────────────────────────
def _footer(sl, num, total, th, topic=""):
    _rect(sl, 0, 5.42, SW, 0.205, th["surface"])
    _rect(sl, 0, 5.42, SW, 0.025, th["accent"])
    if topic: _txt(sl, topic[:52], 0.3, 5.44, 5.5, 0.18, 7.5, th["muted"])
    _txt(sl, f"{num}  /  {total}", SW-1.4, 5.44, 1.1, 0.18, 7.5, th["muted"], align=PP_ALIGN.RIGHT)
    # Small accent dot
    _circle(sl, SW/2, 5.522, 0.05, th["accent"])

def _wm(sl, text):
    if not text: return
    b = sl.shapes.add_textbox(Inches(0.5), Inches(1.8), Inches(9), Inches(2))
    tf=b.text_frame; p=tf.paragraphs[0]; p.alignment=PP_ALIGN.CENTER
    r=p.add_run(); r.text=text.upper()
    r.font.size=Pt(56); r.font.bold=True
    r.font.color.rgb=RGBColor(160,160,160); r.font.name="Calibri"
    b.rotation=-28

# ═══════════════════════════════════════════════════════════════════════════════
#  RICH CONTENT COMPONENTS
# ═══════════════════════════════════════════════════════════════════════════════
def _example_block(sl, text, th, x, y, w, h=1.02):
    if not str(text).strip(): return 0
    _rrect(sl, x, y, w, h, th["ex_bg"], adj=0.05)
    _rect(sl, x, y, 0.055, h, th["accent2"])
    _rrect(sl, x+0.12, y+0.06, 1.45, 0.22, th["accent"])
    _txt(sl, "📌  EXAMPLE", x+0.18, y+0.07, 1.35, 0.2, 7.5, th["tag_fg"], bold=True)
    _mltext(sl, str(text)[:280], x+0.16, y+0.34, w-0.28, h-0.42,
            10.5, th["body"], italic=True, spacing=4)
    return h

def _stat_block(sl, text, th, x, y, w=4.6, h=0.78):
    if not str(text).strip(): return 0
    _rrect(sl, x, y, w, h, th["raised"], adj=0.06)
    _rect(sl, x, y, 0.055, h, th["gold"])
    # Extract number
    nums = re.findall(r'[\d,\.]+\s*[%BMKxX×]+', str(text))
    if nums:
        _txt(sl, nums[0], x+0.14, y+0.06, 1.3, 0.65, 22, th["gold"], bold=True)
        rest = str(text).replace(nums[0],"").strip()
        _mltext(sl, rest[:100], x+1.55, y+0.1, w-1.7, 0.62, 9.5, th["muted"])
    else:
        _mltext(sl, str(text)[:110], x+0.14, y+0.1, w-0.25, 0.62, 10, th["gold"], bold=True)
    return h

def _quote_block(sl, text, th, x, y, w, h=1.18):
    if not str(text).strip(): return 0
    _rrect(sl, x, y, w, h, th["raised"], adj=0.04)
    _rect(sl, x, y, 0.055, h, th["accent2"])
    _txt(sl, "\u201c", x+0.16, y-0.14, 0.7, 0.7, 44, th["accent2"], bold=True)
    _mltext(sl, str(text)[:300], x+0.16, y+0.18, w-0.3, h-0.3,
            12, th["body"], italic=True, align=PP_ALIGN.CENTER, spacing=5)
    return h

def _step_item(sl, n, text, th, x, y, w, accent_c=None):
    ac = accent_c or th["accent"]
    _circle(sl, x+0.22, y+0.24, 0.4, ac)
    nb = sl.shapes.add_textbox(Inches(x+0.02), Inches(y+0.04), Inches(0.4), Inches(0.4))
    ntf=nb.text_frame; np_=ntf.paragraphs[0]; np_.alignment=PP_ALIGN.CENTER
    nr=np_.add_run(); nr.text=str(n)
    nr.font.size=Pt(13); nr.font.bold=True
    nr.font.color.rgb=_c(th["bg"]); nr.font.name="Calibri"
    clean = re.sub(r'^Step\s*\d+\s*[:.)\-]?\s*','',str(text),flags=re.I).strip()
    _txt(sl, clean, x+0.56, y+0.06, w-0.6, 0.42, 13, th["body"])

def _info_card(sl, title, body, th, x, y, w, h, accent_c=None):
    ac = accent_c or th["accent"]
    _rrect(sl, x, y, w, h, th["raised"], adj=0.05)
    _rect(sl, x, y, w, 0.06, ac)
    if title: _txt(sl, title, x+0.1, y+0.1, w-0.2, 0.3, 11, th["accent2"], bold=True)
    if body:  _mltext(sl, body, x+0.1, y+0.46, w-0.2, h-0.55, 10.5, th["body"], spacing=4)

# ═══════════════════════════════════════════════════════════════════════════════
#  SLIDE TEMPLATES  (10 professional layouts)
# ═══════════════════════════════════════════════════════════════════════════════

# ── 1. TITLE ──────────────────────────────────────────────────────────────────
def tpl_title(prs, d, th, wm, num, total, topic, imgs):
    sl = prs.slides.add_slide(prs.slide_layouts[6])
    kw = d.get("image_keyword","technology innovation")
    raw = _fetch(kw, 1280, 720, imgs)
    if raw:
        bg = _dark(_tint(raw, th["accent"], 0.22), 0.5)
        _place(sl, bg, 0, 0, SW, SH)
        # Bottom gradient overlay
        bot = _fade(1280, 260, th["bg"], "up")
        _place(sl, bot, 0, 3.3, SW, 2.325)
    else:
        _place(sl, _grad(1280,720, th["bg"], th["surface"], th["raised"]), 0, 0, SW, SH)

    # Structural elements
    _rect(sl, 0, 0, 0.09, SH, th["accent2"])              # left stripe
    _rect(sl, 0, SH-0.32, SW, 0.32, th["accent"])         # bottom bar

    # Content
    _tag(sl, "PRESENTATION", 0.55, 0.42, th)
    _txt(sl, d.get("title",""), 0.55, 0.85, 8.9, 1.75,
         44, th["title"], bold=True, name="Calibri")
    _rect(sl, 0.55, 3.0, 3.5, 0.045, th["accent2"])       # title underline
    sub = d.get("subtitle","") or d.get("content","")[:110]
    _txt(sl, sub, 0.55, 3.08, 8.0, 0.85, 15, th["body"], italic=True)

    if d.get("stat"):
        _stat_block(sl, d["stat"], th, 0.55, 3.98, 5.5, 0.68)

    _txt(sl, topic[:55], 0.55, SH-0.28, 6, 0.24, 8, th["bg"], bold=True)
    _wm(sl, wm)

# ── 2. AGENDA ─────────────────────────────────────────────────────────────────
def tpl_agenda(prs, d, th, wm, num, total, topic, imgs):
    sl = prs.slides.add_slide(prs.slide_layouts[6])
    _rect(sl, 0, 0, SW, SH, th["bg"])
    _rect(sl, 0, 0, 3.1, SH, th["surface"])
    _rect(sl, 3.1, 0, SW-3.1, SH, th["bg"])

    raw = _fetch(d.get("image_keyword","roadmap plan"), 620, 720, imgs)
    if raw:
        img = _dark(_tint(raw, th["accent"], 0.2), 0.65)
        _place(sl, img, 0, 0, 3.1, SH)

    _rect(sl, 0, 0, 0.075, SH, th["accent2"])
    _tag(sl, "TODAY'S AGENDA", 0.22, 0.45, th, w=2.4)
    _txt(sl, d.get("title",""), 0.22, 0.88, 2.65, 1.4, 22, th["title"], bold=True)
    sub = d.get("subtitle","")
    if sub: _txt(sl, sub[:80], 0.22, 2.4, 2.65, 0.85, 11, th["body"], italic=True)
    _rect(sl, 0, SH-0.32, 3.1, 0.32, th["accent"])

    # Agenda items
    lines = [l.lstrip("•●▪-–— ").strip() for l in d.get("content","").split("\n") if l.strip()]
    palette = [th["accent"],th["accent2"],th["accent3"],th["gold"],th["success"],th["danger"]]
    for i, line in enumerate(lines[:7]):
        yy = 0.3 + i * 0.73
        ac = palette[i % len(palette)]
        # Number box
        _rrect(sl, 3.45, yy, 0.46, 0.46, ac, adj=0.3)
        nb = sl.shapes.add_textbox(Inches(3.45),Inches(yy),Inches(0.46),Inches(0.46))
        ntf=nb.text_frame; np_=ntf.paragraphs[0]; np_.alignment=PP_ALIGN.CENTER
        nr=np_.add_run(); nr.text=str(i+1)
        nr.font.size=Pt(15); nr.font.bold=True
        nr.font.color.rgb=_c(th["bg"]); nr.font.name="Calibri"
        # Text
        _txt(sl, line, 4.08, yy+0.06, 5.55, 0.42, 13.5, th["body"], bold=(i==0))
        # Separator line
        if i < len(lines)-1:
            _rect(sl, 3.45, yy+0.56, 6.15, 0.018, th["border"])

    _footer(sl, num, total, th, topic); _wm(sl, wm)

# ── 3. CONTENT + PHOTO RIGHT ─────────────────────────────────────────────────
def tpl_content_right(prs, d, th, wm, num, total, topic, imgs):
    sl = prs.slides.add_slide(prs.slide_layouts[6])
    _rect(sl, 0, 0, SW, SH, th["bg"])
    _rect(sl, 0, 0, 5.6, SH, th["surface"])

    raw = _fetch(d.get("image_keyword","business"), 820, 720, imgs)
    if raw:
        img = _dark(_tint(raw, th["accent"], 0.1), 0.2)
        _place(sl, img, 5.6, 0, 4.4, SH)
        fade = _fade(140, 900, th["surface"], "right")
        _place(sl, fade, 5.46, 0, 0.34, SH)
    else:
        _rect(sl, 5.6, 0, 4.4, SH, th["raised"])

    _rect(sl, 0, 0, 5.6, 0.06, th["accent"])
    _tag(sl, f"SLIDE {num:02d}", 0.35, 0.1, th)
    _txt(sl, d.get("title",""), 0.35, 0.5, 5.05, 0.78, 24, th["title"], bold=True)
    _rect(sl, 0.35, 1.36, 2.4, 0.042, th["accent"])

    lines = [l for l in d.get("content","").split("\n") if l.strip()]
    n = min(len(lines), 5)
    _bullets(sl, d.get("content",""), 0.35, 1.5, 5.05, n*0.44+0.2, 13.5, th["body"], th["accent2"], gap=7)

    ey = min(1.55 + n*0.46, 3.68)
    used = _example_block(sl, d.get("example",""), th, 0.35, ey, 5.1)
    if d.get("stat"):
        sy = ey + (used or 0) + 0.08
        if sy < 4.82: _stat_block(sl, d["stat"], th, 0.35, sy, 5.1, 0.62)

    _footer(sl, num, total, th, topic); _wm(sl, wm)

# ── 4. BANNER TOP + 2-COL ─────────────────────────────────────────────────────
def tpl_banner(prs, d, th, wm, num, total, topic, imgs):
    sl = prs.slides.add_slide(prs.slide_layouts[6])
    _rect(sl, 0, 0, SW, SH, th["bg"])

    raw = _fetch(d.get("image_keyword","professional"), 1280, 380, imgs)
    if raw:
        img = _dark(_tint(raw, th["accent"], 0.18), 0.42)
        _place(sl, img, 0, 0, SW, 1.9)
    else:
        _rect(sl, 0, 0, SW, 1.9, th["raised"])

    _rect(sl, 0, 1.87, SW, 0.055, th["accent"])
    _tag(sl, f"SLIDE {num:02d}", 0.38, 0.25, th)
    _txt(sl, d.get("title",""), 0.38, 0.6, 9.25, 1.0, 27, th["title"], bold=True)

    lines = [l for l in d.get("content","").split("\n") if l.strip()]
    half = max(1, len(lines)//2)
    _bullets(sl, "\n".join(lines[:half]), 0.3, 2.05, 4.65, 2.15, 13.5, th["body"], th["accent2"], gap=7)
    _bullets(sl, "\n".join(lines[half:]), 5.05, 2.05, 4.65, 2.15, 13.5, th["body"], th["accent2"], gap=7)

    ex = d.get("example","")
    _example_block(sl, ex, th, 0.3, 4.3, 9.4, 0.9)

    _footer(sl, num, total, th, topic); _wm(sl, wm)

# ── 5. BIG STAT HERO ──────────────────────────────────────────────────────────
def tpl_stat(prs, d, th, wm, num, total, topic, imgs):
    sl = prs.slides.add_slide(prs.slide_layouts[6])
    bg = _get_bg(d.get("image_keyword","data analytics"), th, imgs=imgs, dark_op=0.68, tint_op=0.3)
    _place(sl, bg, 0, 0, SW, SH)

    _rect(sl, 0, 0, SW, 0.06, th["accent"])
    _tag(sl, "KEY STATISTIC", 0.4, 0.1, th)
    _txt(sl, d.get("title",""), 0.4, 0.48, 9.2, 0.72, 22, th["title"], bold=True)
    _rect(sl, 0.4, 1.28, 2.2, 0.042, th["accent"])

    stat = d.get("stat","")
    if stat:
        _rrect(sl, 1.5, 1.42, 7.0, 1.56, th["raised"], adj=0.05)
        _rect(sl, 1.5, 1.42, 7.0, 0.065, th["gold"])
        _rect(sl, 1.5, 1.42, 0.065, 1.56, th["gold"])
        nums = re.findall(r'[\d,\.]+\s*[%$BMKkxX×]+', stat)
        big = nums[0] if nums else stat[:20]
        rest = stat.replace(big,"").strip() if nums else ""
        _txt(sl, big, 2.8, 1.52, 4.4, 1.0, 52, th["gold"], bold=True, align=PP_ALIGN.CENTER)
        if rest:
            _txt(sl, rest[:100], 1.7, 2.48, 6.6, 0.42, 11, th["muted"], align=PP_ALIGN.CENTER)

    _bullets(sl, d.get("content",""), 0.4, 3.18, 9.2, 1.1, 13, th["body"], th["accent2"], gap=6)
    _example_block(sl, d.get("example",""), th, 0.4, 4.4, 9.2, 0.78)

    _footer(sl, num, total, th, topic); _wm(sl, wm)

# ── 6. QUOTE ──────────────────────────────────────────────────────────────────
def tpl_quote(prs, d, th, wm, num, total, topic, imgs):
    sl = prs.slides.add_slide(prs.slide_layouts[6])
    bg = _get_bg(d.get("image_keyword","leadership inspiration"), th, imgs=imgs, dark_op=0.62, tint_op=0.3)
    _place(sl, bg, 0, 0, SW, SH)

    _rect(sl, 0, 0, SW, 0.06, th["accent2"])
    _tag(sl, "EXPERT INSIGHT", 0.45, 0.1, th)
    _txt(sl, d.get("title",""), 0.45, 0.48, 9.1, 0.68, 21, th["title"], bold=True)

    q = d.get("quote","") or d.get("content","")
    _quote_block(sl, q, th, 0.55, 1.28, 8.9, 2.0)

    lines = [l for l in d.get("content","").split("\n") if l.strip()]
    if lines:
        _bullets(sl, d.get("content",""), 0.45, 3.42, 9.1, 1.0, 12.5, th["body"], th["accent2"], gap=6)

    _example_block(sl, d.get("example",""), th, 0.45, 4.55, 9.1, 0.65)
    _footer(sl, num, total, th, topic); _wm(sl, wm)

# ── 7. SPLIT / COMPARISON ─────────────────────────────────────────────────────
def tpl_split(prs, d, th, wm, num, total, topic, imgs):
    sl = prs.slides.add_slide(prs.slide_layouts[6])
    _rect(sl, 0, 0, SW, SH, th["bg"])
    _rect(sl, 0, 0, SW, 0.06, th["accent"])

    _tag(sl, "COMPARISON", 0.35, 0.1, th)
    _txt(sl, d.get("title",""), 0.35, 0.46, 9.3, 0.72, 23, th["title"], bold=True)
    _rect(sl, 0.35, 1.26, 9.3, 0.035, th["border"])

    left  = d.get("content_left","") or d.get("content","")
    right = d.get("content_right","")
    if not right:
        lns = [l for l in left.split("\n") if l.strip()]
        mid = max(1, len(lns)//2)
        left = "\n".join(lns[:mid]); right = "\n".join(lns[mid:])

    title_parts = re.split(r'\s+vs\.?\s+', d.get("title",""), flags=re.I)
    lhdr = title_parts[0].strip()[:28] if len(title_parts)>1 else "Option A"
    rhdr = title_parts[1].strip()[:28] if len(title_parts)>1 else "Option B"

    # Left card
    _rrect(sl, 0.28, 1.38, 4.52, 3.6, th["surface"], adj=0.04)
    _rect(sl, 0.28, 1.38, 4.52, 0.065, th["accent"])
    _txt(sl, lhdr, 0.48, 1.5, 4.0, 0.42, 13, th["accent2"], bold=True)
    _bullets(sl, left, 0.48, 2.02, 4.12, 2.68, 12.5, th["body"], th["accent"], gap=7)

    # Right card
    _rrect(sl, 5.2, 1.38, 4.52, 3.6, th["raised"], adj=0.04)
    _rect(sl, 5.2, 1.38, 4.52, 0.065, th["accent2"])
    _txt(sl, rhdr, 5.4, 1.5, 4.0, 0.42, 13, th["accent"], bold=True)
    _bullets(sl, right, 5.4, 2.02, 4.12, 2.68, 12.5, th["body"], th["accent2"], gap=7)

    # VS badge
    _circle(sl, SW/2, 2.68, 0.56, th["accent"])
    vb = sl.shapes.add_textbox(Inches(SW/2-0.28), Inches(2.4), Inches(0.56), Inches(0.56))
    vtf=vb.text_frame; vp=vtf.paragraphs[0]; vp.alignment=PP_ALIGN.CENTER
    vr=vp.add_run(); vr.text="VS"
    vr.font.size=Pt(13); vr.font.bold=True
    vr.font.color.rgb=_c(th["bg"]); vr.font.name="Calibri"

    _example_block(sl, d.get("example",""), th, 0.28, 5.06, 9.44, 0.44)
    _footer(sl, num, total, th, topic); _wm(sl, wm)

# ── 8. PROCESS / TIMELINE ─────────────────────────────────────────────────────
def tpl_process(prs, d, th, wm, num, total, topic, imgs):
    sl = prs.slides.add_slide(prs.slide_layouts[6])
    raw = _fetch(d.get("image_keyword","workflow process"), 1280, 720, imgs)
    if raw:
        bg = _blur(_dark(raw, 0.82), 7)
        _place(sl, bg, 0, 0, SW, SH)
    else:
        _rect(sl, 0, 0, SW, SH, th["bg"])

    _rect(sl, 0, 0, SW, 0.06, th["accent"])
    _tag(sl, "PROCESS", 0.4, 0.1, th)
    _txt(sl, d.get("title",""), 0.4, 0.47, 9.2, 0.7, 23, th["title"], bold=True)
    _rect(sl, 0.4, 1.24, 9.2, 0.035, th["accent"])

    steps_raw = d.get("steps","") or d.get("content","")
    steps = [l.strip() for l in steps_raw.split("\n") if l.strip()]
    palette = [th["accent"],th["accent2"],th["accent3"],th["gold"],th["success"],th["danger"]]

    # Connector line
    if len(steps) > 1:
        _rect(sl, 0.62, 1.54, 0.025, min(len(steps),6)*0.72-0.3, th["border"])

    for i, step in enumerate(steps[:6]):
        yy = 1.38 + i * 0.72
        _step_item(sl, i+1, step, th, 0.38, yy, 9.22, palette[i%len(palette)])
        if i < len(steps)-1:
            _rect(sl, 0.38, yy+0.56, 9.22, 0.018, th["border"])

    _example_block(sl, d.get("example",""), th, 0.38, 4.4, 9.24, 0.88)
    _footer(sl, num, total, th, topic); _wm(sl, wm)

# ── 9. BIG IMAGE (visual impact) ──────────────────────────────────────────────
def tpl_bigimage(prs, d, th, wm, num, total, topic, imgs):
    sl = prs.slides.add_slide(prs.slide_layouts[6])
    raw = _fetch(d.get("image_keyword","abstract"), 1280, 720, imgs)
    if raw:
        bg = _dark(_tint(raw, th["accent"], 0.18), 0.48)
        _place(sl, bg, 0, 0, SW, SH)
        bot = _fade(1280, 300, th["bg"], "up")
        _place(sl, bot, 0, 3.1, SW, 2.525)
    else:
        _place(sl, _grad(1280,720,th["bg"],th["surface"],th["raised"]), 0, 0, SW, SH)

    _rect(sl, 0, 0, 0.08, SH, th["accent2"])
    _tag(sl, f"SLIDE {num:02d}", 0.45, 0.38, th)
    _txt(sl, d.get("title",""), 0.45, 0.82, 9.0, 1.3, 34, th["title"], bold=True)
    _rect(sl, 0.45, 2.2, 3.0, 0.045, th["accent2"])

    sub = d.get("subtitle","") or d.get("content","").split("\n")[0].lstrip("•●▪-–— ").strip() if d.get("content","") else ""
    _txt(sl, sub[:100], 0.45, 2.32, 8.5, 0.6, 14, th["body"], italic=True)

    if d.get("stat"):
        _stat_block(sl, d["stat"], th, 0.45, 3.05, 5.5, 0.72)
    if d.get("quote"):
        _txt(sl, f'\u201c{d["quote"][:100]}\u201d', 0.45, 3.88, 8.5, 0.62, 12, th["muted"], italic=True)

    _footer(sl, num, total, th, topic); _wm(sl, wm)

# ── 10. CARDS (4-up) ──────────────────────────────────────────────────────────
def tpl_cards(prs, d, th, wm, num, total, topic, imgs):
    sl = prs.slides.add_slide(prs.slide_layouts[6])
    _rect(sl, 0, 0, SW, SH, th["bg"])
    _rect(sl, 0, 0, SW, 0.06, th["accent"])
    raw = _fetch(d.get("image_keyword","abstract pattern"), 1280, 720, imgs)
    if raw:
        _place(sl, _blur(_dark(raw, 0.88), 9), 0, 0, SW, SH)
        _rect(sl, 0, 0, SW, 0.06, th["accent"])

    _tag(sl, "KEY POINTS", 0.35, 0.1, th)
    _txt(sl, d.get("title",""), 0.35, 0.48, 9.3, 0.7, 23, th["title"], bold=True)
    _rect(sl, 0.35, 1.26, 2.0, 0.042, th["accent"])

    lines = [l.lstrip("•●▪-–— ").strip() for l in d.get("content","").split("\n") if l.strip()]
    xs = [0.25, 2.6, 4.95, 7.3]; cw, ch = 2.2, 2.85
    palette = [th["accent"],th["accent2"],th["accent3"],th["gold"]]

    for i, line in enumerate(lines[:4]):
        x = xs[i]; ac = palette[i]
        _rrect(sl, x, 1.38, cw, ch, th["raised"], adj=0.04)
        _rect(sl, x, 1.38, cw, 0.065, ac)
        # Number circle
        _circle(sl, x+cw/2, 1.82, 0.52, ac)
        nb = sl.shapes.add_textbox(Inches(x+cw/2-0.26), Inches(1.56), Inches(0.52), Inches(0.52))
        ntf=nb.text_frame; np_=ntf.paragraphs[0]; np_.alignment=PP_ALIGN.CENTER
        nr=np_.add_run(); nr.text=str(i+1)
        nr.font.size=Pt(15); nr.font.bold=True
        nr.font.color.rgb=_c(th["bg"]); nr.font.name="Calibri"
        # Text
        ct = sl.shapes.add_textbox(Inches(x+0.1), Inches(2.2), Inches(cw-0.2), Inches(1.88))
        ctf=ct.text_frame; ctf.word_wrap=True
        cp=ctf.paragraphs[0]; cp.alignment=PP_ALIGN.CENTER
        cr=cp.add_run(); cr.text=line
        cr.font.size=Pt(11); cr.font.color.rgb=_c(th["body"]); cr.font.name="Calibri"

    _example_block(sl, d.get("example",""), th, 0.25, 4.35, 9.5, 0.85)
    _footer(sl, num, total, th, topic); _wm(sl, wm)

# ── 11. CONCLUSION ────────────────────────────────────────────────────────────
def tpl_conclusion(prs, d, th, wm, num, total, topic, imgs):
    sl = prs.slides.add_slide(prs.slide_layouts[6])
    bg = _get_bg(d.get("image_keyword","success victory"), th, imgs=imgs, dark_op=0.62, tint_op=0.3)
    _place(sl, bg, 0, 0, SW, SH)

    _rect(sl, 0, 0, SW, 0.06, th["accent2"])
    _tag(sl, "KEY TAKEAWAYS", 0.45, 0.1, th)
    _txt(sl, d.get("title","Key Takeaways"), 0.45, 0.48, 9.1, 0.82, 32, th["title"], bold=True)
    _rect(sl, 0.45, 1.38, 2.5, 0.048, th["accent2"])

    lines = [l.lstrip("•●▪-–— ").strip() for l in d.get("content","").split("\n") if l.strip()]
    palette = [th["accent"],th["accent2"],th["accent3"],th["gold"],th["success"]]
    for i, line in enumerate(lines[:5]):
        yy = 1.52 + i * 0.72
        _circle(sl, 0.67, yy+0.24, 0.42, palette[i%len(palette)])
        nb = sl.shapes.add_textbox(Inches(0.46), Inches(yy+0.04), Inches(0.42), Inches(0.42))
        ntf=nb.text_frame; np_=ntf.paragraphs[0]; np_.alignment=PP_ALIGN.CENTER
        nr=np_.add_run(); nr.text=str(i+1)
        nr.font.size=Pt(13); nr.font.bold=True
        nr.font.color.rgb=_c(th["bg"]); nr.font.name="Calibri"
        _txt(sl, line, 1.05, yy+0.07, 8.5, 0.5, 14, th["body"])

    q = d.get("quote","")
    if q: _txt(sl, f'\u201c{q[:110]}\u201d', 0.45, 5.05, 9.1, 0.42, 10, th["accent2"], italic=True)

    _footer(sl, num, total, th, topic); _wm(sl, wm)

# ── 12. REFERENCES ────────────────────────────────────────────────────────────
def tpl_references(prs, d, th, wm, num, total, topic, imgs):
    sl = prs.slides.add_slide(prs.slide_layouts[6])
    _rect(sl, 0, 0, SW, SH, th["bg"])
    _rect(sl, 0, 0, SW, 0.06, th["accent"])
    _tag(sl, "REFERENCES & SOURCES", 0.35, 0.1, th, w=2.8)
    _txt(sl, d.get("title","References"), 0.35, 0.48, 9.3, 0.7, 24, th["title"], bold=True)
    _rect(sl, 0.35, 1.24, 1.8, 0.042, th["accent"])
    _bullets(sl, d.get("content",""), 0.35, 1.4, 9.3, 3.7, 13, th["body"], th["accent2"], gap=11)
    _footer(sl, num, total, th, topic); _wm(sl, wm)

# ── CHART SLIDE ───────────────────────────────────────────────────────────────
def tpl_chart(prs, path, title, insight, th, wm, num, total, topic):
    sl = prs.slides.add_slide(prs.slide_layouts[6])
    _rect(sl, 0, 0, SW, SH, th["bg"])
    _rect(sl, 0, 0, SW, 0.06, th["accent"])
    _tag(sl, "DATA VISUALIZATION", 0.35, 0.1, th, w=2.6)
    _txt(sl, title, 0.35, 0.46, 9.3, 0.65, 22, th["title"], bold=True)
    if insight:
        if path and os.path.exists(path):
            sl.shapes.add_picture(path, Inches(0.3), Inches(1.22), Inches(9.4), Inches(3.5))
        _example_block(sl, f"💡  {insight}", th, 0.3, 4.82, 9.4, 0.58)
    else:
        if path and os.path.exists(path):
            sl.shapes.add_picture(path, Inches(0.3), Inches(1.22), Inches(9.4), Inches(4.05))
    _footer(sl, num, total, th, topic); _wm(sl, wm)

# ═══════════════════════════════════════════════════════════════════════════════
#  LAYOUT ROUTER
# ═══════════════════════════════════════════════════════════════════════════════
_CYCLE = ["content_right","banner","cards","bigimage","content_right","banner","content_right","cards"]

_TYPE_MAP = {
    "title":      tpl_title,
    "agenda":     tpl_agenda,
    "statistic":  tpl_stat,
    "quote":      tpl_quote,
    "split":      tpl_split,
    "comparison": tpl_split,
    "process":    tpl_process,
    "timeline":   tpl_process,
    "bigimage":   tpl_bigimage,
    "conclusion": tpl_conclusion,
    "references": tpl_references,
}

def _route(prs, d, th, wm, num, total, topic, imgs, ci):
    t = d.get("type","content")
    fn = _TYPE_MAP.get(t)
    if fn:
        fn(prs, d, th, wm, num, total, topic, imgs); return
    # Content type — cycle layouts
    layout = _CYCLE[ci % len(_CYCLE)]
    if layout == "content_right": tpl_content_right(prs,d,th,wm,num,total,topic,imgs)
    elif layout == "banner":      tpl_banner(prs,d,th,wm,num,total,topic,imgs)
    elif layout == "cards":       tpl_cards(prs,d,th,wm,num,total,topic,imgs)
    elif layout == "bigimage":    tpl_bigimage(prs,d,th,wm,num,total,topic,imgs)
    else:                         tpl_content_right(prs,d,th,wm,num,total,topic,imgs)

# ═══════════════════════════════════════════════════════════════════════════════
#  MAIN BUILD
# ═══════════════════════════════════════════════════════════════════════════════
def build_presentation(slides, theme, watermark_text="", topic="",
                       chart_map=None, fetch_images=True):
    th = THEMES.get(theme, THEMES["🌌 Midnight Executive"])
    chart_map = chart_map or {}
    prs = Presentation()
    prs.slide_width = Inches(SW); prs.slide_height = Inches(SH)

    total = len(slides); ci = 0
    for num, d in enumerate(slides, 1):
        _route(prs, d, th, watermark_text, num, total, topic, fetch_images, ci)
        if d.get("type","content") not in _TYPE_MAP:
            ci += 1
        # Append chart slide if mapped
        if num in chart_map:
            tpl_chart(prs, chart_map[num], f"Chart — {d.get('title','')[:45]}",
                      d.get("example",""), th, watermark_text, num, total, topic)

    buf = io.BytesIO(); prs.save(buf); buf.seek(0)
    return buf.read()
