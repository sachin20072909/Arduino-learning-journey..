#!/usr/bin/env python3
"""
ZODIAC AI — pitch deck builder v2 (PRO)
Cyber-Grid theme: AI-rendered backgrounds, DrawingML neon glow, glass panels,
gradient scrims, CRT scanlines, ghost numerals, progress bar.
"""
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn
from lxml import etree
from PIL import Image

# ---------------------------------------------------------------- palette
BG      = "0A0E14";  PANEL  = "10192A";  PANEL_D = "0B111C"
LINE    = "223046";  INK    = "141F30"
GREEN   = "00FF41";  BLUE   = "00D4FF";  RED  = "FF3B61"
TEXT    = "EAF4FF";  MUTED  = "93A6C4";  FAINT = "5A6A85"
GREEN_D = "0D2A18";  BLUE_D = "0A2431";  RED_D = "2A0F1A"

DISP = "Bahnschrift"        # display / headlines
BODY = "Segoe UI"           # body text
MONO = "Consolas"           # terminal text

SW, SH = 13.333, 7.5
def C(h): return RGBColor(int(h[0:2],16), int(h[2:4],16), int(h[4:6],16))

prs = Presentation()
prs.slide_width  = Inches(SW); prs.slide_height = Inches(SH)
BLANK = prs.slide_layouts[6]

# ---------------------------------------------------------------- xml fx
def _sub(parent, tag, **attrs):
    e = etree.SubElement(parent, qn(tag))
    for k, v in attrs.items(): e.set(k, str(v))
    return e

def add_glow(shape, color, rad_pt=5, alpha=45):
    """neon glow around a shape"""
    spPr = shape._element.spPr
    for e in spPr.findall(qn('a:effectLst')): spPr.remove(e)
    eff = etree.Element(qn('a:effectLst'))
    g = _sub(eff, 'a:glow', rad=Pt(rad_pt))
    c = _sub(g, 'a:srgbClr', val=color); _sub(c, 'a:alpha', val=int(alpha*1000))
    spPr.append(eff)

def add_shadow(shape, blur=0.16, dist=0.045, alpha=55):
    spPr = shape._element.spPr
    for e in spPr.findall(qn('a:effectLst')): spPr.remove(e)
    eff = etree.Element(qn('a:effectLst'))
    sh = _sub(eff, 'a:outerShdw', blurRad=Inches(blur), dist=Inches(dist), dir=5400000, rotWithShape=0)
    c = _sub(sh, 'a:srgbClr', val='000000'); _sub(c, 'a:alpha', val=int(alpha*1000))
    spPr.append(eff)

def set_fill_alpha(shape, pct):
    sf = shape._element.spPr.find(qn('a:solidFill'))
    if sf is not None:
        c = sf.find(qn('a:srgbClr'))
        if c is not None: _sub(c, 'a:alpha', val=int(pct*1000))

def gradient_fill(shape, stops, angle_deg=0):
    """stops: [(pos_pct, hex, alpha_pct), ...] — replaces solid fill"""
    spPr = shape._element.spPr
    for e in spPr.findall(qn('a:solidFill')): spPr.remove(e)
    grad = etree.Element(qn('a:gradFill')); grad.set('rotWithShape', '1')
    lst = _sub(grad, 'a:gsLst')
    for pos, hx, al in stops:
        gs = _sub(lst, 'a:gs', pos=int(pos*1000))
        c = _sub(gs, 'a:srgbClr', val=hx); _sub(c, 'a:alpha', val=int(al*1000))
    _sub(grad, 'a:lin', ang=int(angle_deg*60000))
    ln = spPr.find(qn('a:ln'))
    if ln is not None: ln.addprevious(grad)
    else: spPr.append(grad)

def run_glow(run, color, rad_pt=4, alpha=50):
    rPr = run._r.get_or_add_rPr()
    for e in rPr.findall(qn('a:effectLst')): rPr.remove(e)
    eff = etree.Element(qn('a:effectLst'))
    g = _sub(eff, 'a:glow', rad=Pt(rad_pt))
    c = _sub(g, 'a:srgbClr', val=color); _sub(c, 'a:alpha', val=int(alpha*1000))
    latin = rPr.find(qn('a:latin'))
    if latin is not None: latin.addprevious(eff)
    else: rPr.append(eff)

def run_alpha(run, pct):
    rPr = run._r.get_or_add_rPr()
    sf = rPr.find(qn('a:solidFill'))
    if sf is not None:
        c = sf.find(qn('a:srgbClr'))
        if c is not None: _sub(c, 'a:alpha', val=int(pct*1000))

# ---------------------------------------------------------------- primitives
def rect(s, x, y, w, h, fill=None, line=None, lw=1.0, rounded=False, adj=None, alpha=None):
    shp = s.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE if rounded else MSO_SHAPE.RECTANGLE,
        Inches(x), Inches(y), Inches(w), Inches(h))
    shp.shadow.inherit = False
    if rounded and adj is not None:
        try: shp.adjustments[0] = adj
        except Exception: pass
    if fill is None: shp.fill.background()
    else:
        shp.fill.solid(); shp.fill.fore_color.rgb = C(fill)
        if alpha is not None: set_fill_alpha(shp, alpha)
    if line is None: shp.line.fill.background()
    else:
        shp.line.color.rgb = C(line); shp.line.width = Pt(lw)
    return shp

def R(t, size=12, color=TEXT, bold=False, italic=False, mono=False,
      strike=False, spc=None, alpha=None, glow=None, font=None):
    return (t, dict(size=size, color=color, bold=bold, italic=italic, mono=mono,
                    strike=strike, spc=spc, alpha=alpha, glow=glow, font=font))

def P(runs, align=None, sb=None, sa=None, ls=1.0, hang=None):
    return dict(runs=runs, align=align, sb=sb, sa=sa, ls=ls, hang=hang)

def text(s, x, y, w, h, paras, anchor=MSO_ANCHOR.TOP, wrap=True):
    tb = s.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = tb.text_frame
    tf.word_wrap = wrap; tf.vertical_anchor = anchor
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    for i, p in enumerate(paras):
        par = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        par.alignment = p.get('align') or PP_ALIGN.LEFT
        if p.get('sb') is not None: par.space_before = Pt(p['sb'])
        if p.get('sa') is not None: par.space_after  = Pt(p['sa'])
        par.line_spacing = p.get('ls', 1.0)
        if p.get('hang'):
            pPr = par._p.get_or_add_pPr()
            pPr.set('marL', str(int(Inches(p['hang'])))); pPr.set('indent', str(-int(Inches(p['hang']))))
        for txt, st in p['runs']:
            r = par.add_run(); r.text = txt; f = r.font
            f.size  = Pt(st.get('size', 12))
            f.color.rgb = C(st.get('color', TEXT))
            f.bold  = st.get('bold', False); f.italic = st.get('italic', False)
            f.name  = st.get('font') or (MONO if st.get('mono') else BODY)
            if st.get('strike'): r._r.get_or_add_rPr().set('strike', 'sngStrike')
            if st.get('spc') is not None: r._r.get_or_add_rPr().set('spc', str(st['spc']))
            if st.get('alpha') is not None: run_alpha(r, st['alpha'])
            if st.get('glow'):
                hx, rad, al = st['glow']; run_glow(r, hx, rad, al)
    return tb

def picture_fit(s, path, x, y, maxW, maxH):
    iw, ih = Image.open(path).size
    ar = iw / ih
    w = maxW; h = w / ar
    if h > maxH: h = maxH; w = h * ar
    px = x + (maxW - w) / 2; py = y + (maxH - h) / 2
    pic = s.shapes.add_picture(path, Inches(px), Inches(py), Inches(w), Inches(h))
    try: add_shadow(pic, 0.2, 0.05, 60)
    except Exception: pass
    return pic, px, py, w, h

# ---------------------------------------------------------------- components
def chrome(s, num, tag, bg):
    """slide chrome: bg, ghost numeral, hairlines, footer, progress bar"""
    s.shapes.add_picture(bg, 0, 0, Inches(SW), Inches(SH))
    text(s, 10.15, 0.0, 2.75, 1.8,
         [P([R(f"{num:02d}", size=115, color=TEXT, bold=True, font=DISP, alpha=7)],
            align=PP_ALIGN.RIGHT)])
    r1 = rect(s, 0, 0, SW*0.42, 0.032, fill=GREEN); add_glow(r1, GREEN, 4, 35)
    r2 = rect(s, SW*0.42, 0, SW*0.58, 0.032, fill=BLUE);  add_glow(r2, BLUE, 4, 35)
    rect(s, 0.28, SH-0.66, 0.36, 0.022, fill=BLUE)
    rect(s, 0.28, SH-0.64, 0.022, 0.38, fill=BLUE)
    rect(s, 0.55, 7.05, SW-1.1, 0.012, fill=LINE)
    text(s, 0.55, 7.11, 6.2, 0.22,
         [P([R("ZODIAC AI ", size=8, color=GREEN, mono=True, bold=True),
             R("// CHAT WITH YOUR DATABASE // ", size=8, color=FAINT, mono=True),
             R(tag, size=8, color=BLUE, mono=True)])])
    text(s, SW-1.6, 7.11, 1.05, 0.22,
         [P([R(f"{num:02d}", size=8, color=GREEN, mono=True, bold=True),
             R(" / 11", size=8, color=FAINT, mono=True)], align=PP_ALIGN.RIGHT)])
    seg_w, gap = 0.30, 0.085
    x0 = (SW - (11*seg_w + 10*gap)) / 2
    for i in range(11):
        if i == num-1:
            sg = rect(s, x0+i*(seg_w+gap), 7.30, seg_w, 0.05, fill=GREEN)
            add_glow(sg, GREEN, 5, 65)
        else:
            rect(s, x0+i*(seg_w+gap), 7.30, seg_w, 0.05,
                 fill="16536E" if i < num-1 else INK)

def kicker_title(s, kicker, title_runs):
    k = rect(s, 0.57, 0.51, 0.11, 0.11, fill=GREEN); add_glow(k, GREEN, 5, 60)
    text(s, 0.80, 0.44, 11.5, 0.3,
         [P([R(kicker, size=10.5, color=GREEN, mono=True, spc=170, glow=(GREEN, 3, 40))])])
    text(s, 0.55, 0.76, 12.3, 0.62, [P(title_runs)])
    b1 = rect(s, 0.57, 1.47, 1.7, 0.034, fill=GREEN); add_glow(b1, GREEN, 4, 55)
    b2 = rect(s, 2.27, 1.47, 0.55, 0.034, fill=BLUE);  add_glow(b2, BLUE, 4, 55)

def panel(s, x, y, w, h, title, accent=GREEN, glass=90):
    p = rect(s, x, y, w, h, fill=PANEL, line=LINE, lw=1.0, rounded=True, adj=0.05)
    set_fill_alpha(p, glass); add_shadow(p, 0.18, 0.05, 58)
    top = rect(s, x+0.14, y, w-0.28, 0.045, fill=accent); add_glow(top, accent, 5, 60)
    text(s, x+0.26, y+0.15, w-0.52, 0.3,
         [P([R(title, size=11.5, color=accent, mono=True, bold=True, spc=110)])])
    return y + 0.52

def bullets(s, x, y, w, h, items, size=12.5, marker="▸", mcolor=GREEN, gap=10, mglow=False):
    paras = []
    for it in items:
        if isinstance(it, tuple):
            lead, rest = it
            runs = [R(marker+"  ", size=size, color=mcolor, bold=True, mono=True, glow=(mcolor,3,45) if mglow else None),
                    R(lead, size=size, color=TEXT, bold=True)]
            if rest: runs.append(R(" — " + rest, size=size, color=MUTED))
        else:
            runs = [R(marker+"  ", size=size, color=mcolor, bold=True, mono=True),
                    R(it, size=size, color=TEXT)]
        paras.append(P(runs, sa=gap, ls=1.12, hang=0.26))
    text(s, x, y, w, h, paras)

def chip(s, x, y, w, h, label, color=BLUE, fill=None, size=10, glow=True):
    c = rect(s, x, y, w, h, fill=fill if fill else BG, line=color, lw=1.2, rounded=True, adj=0.5)
    if fill: set_fill_alpha(c, 80)
    if glow: add_glow(c, color, 4, 40)
    text(s, x, y, w, h, [P([R(label, size=size, color=color, mono=True, bold=True)], align=PP_ALIGN.CENTER)],
         anchor=MSO_ANCHOR.MIDDLE)

def bottom_bar(s, label, accent=GREEN, y=6.55):
    b = rect(s, 0.55, y, SW-1.1, 0.46, fill=PANEL, line=LINE, lw=1.0, rounded=True, adj=0.5)
    set_fill_alpha(b, 88); add_shadow(b, 0.14, 0.04, 50)
    n = rect(s, 0.76, y+0.11, 0.06, 0.24, fill=accent); add_glow(n, accent, 4, 65)
    text(s, 1.0, y, SW-2.2, 0.46,
         [P([R(label, size=10.5, color=accent, mono=True, bold=True, spc=55, glow=(accent,3,30))])],
         anchor=MSO_ANCHOR.MIDDLE)

def terminal(s, x, y, w, h, label, accent=GREEN):
    p = rect(s, x, y, w, h, fill=PANEL_D, line=LINE, lw=1.0, rounded=True, adj=0.045)
    set_fill_alpha(p, 93); add_shadow(p, 0.18, 0.05, 58)
    top = rect(s, x+0.14, y, w-0.28, 0.045, fill=accent); add_glow(top, accent, 5, 60)
    for i, dc in enumerate(["FF5F56", "FFBD2E", "27C93F"]):
        d = s.shapes.add_shape(MSO_SHAPE.OVAL, Inches(x+0.18+i*0.18), Inches(y+0.16), Inches(0.11), Inches(0.11))
        d.shadow.inherit = False; d.fill.solid(); d.fill.fore_color.rgb = C(dc); d.line.fill.background()
        add_glow(d, dc, 3, 45)
    text(s, x+0.85, y+0.13, w-1.1, 0.26,
         [P([R("zodiac:~ ", size=10, color=FAINT, mono=True),
             R(label, size=10, color=accent, mono=True, bold=True)])])
    ly = y + 0.62
    while ly < y + h - 0.08:                       # CRT scanlines
        rect(s, x+0.12, ly, w-0.24, 0.016, fill=BLUE, alpha=5)
        ly += 0.135
    return y + 0.56

def display(s, x, y, w, h, txt, size, align=PP_ALIGN.LEFT):
    """neon-sign display text: blue chromatic ghost + white core with green glow"""
    text(s, x+0.07, y+0.05, w, h, [P([R(txt, size=size, color=BLUE, bold=True, font=DISP, alpha=55)], align=align)])
    text(s, x, y, w, h, [P([R(txt, size=size, color="FFFFFF", bold=True, font=DISP, glow=(GREEN, 11, 50))], align=align)])

def notes(s, visual, speaker):
    s.notes_slide.notes_text_frame.text = "VISUAL DIRECTION — " + visual + "\n\nSPEAKER NOTES — " + speaker

def num_badge(s, x, y, n, size=0.42):
    b = rect(s, x, y, size, size, fill=GREEN, rounded=True, adj=0.22)
    add_glow(b, GREEN, 6, 60)
    text(s, x, y, size, size, [P([R(str(n), size=14, color=BG, mono=True, bold=True)], align=PP_ALIGN.CENTER)],
         anchor=MSO_ANCHOR.MIDDLE)

# ============================================================ SLIDE 1 · TITLE
s = prs.slides.add_slide(BLANK)
s.shapes.add_picture("assets/hero-title.jpg", 0, 0, Inches(SW), Inches(SH))
scrim = rect(s, 0, 0, 8.4, SH, fill=BG)
gradient_fill(scrim, [(0, BG, 97), (55, BG, 82), (100, BG, 0)], 0)
r1 = rect(s, 0, 0, SW*0.42, 0.04, fill=GREEN); add_glow(r1, GREEN, 5, 45)
r2 = rect(s, SW*0.42, 0, SW*0.58, 0.04, fill=BLUE);  add_glow(r2, BLUE, 5, 45)
text(s, 0.62, 0.52, 9.0, 0.3,
     [P([R("ZODIAC OS v1.0 // TEXT-TO-SQL AGENT // STATUS: ", size=11, color=MUTED, mono=True),
         R("ONLINE", size=11, color=GREEN, mono=True, bold=True, glow=(GREEN, 4, 55)),
         R("  ▮", size=11, color=GREEN, mono=True, glow=(GREEN, 4, 55))])])
display(s, 0.55, 1.42, 8.6, 1.65, "ZODIAC AI", 92)
text(s, 0.63, 3.18, 8.0, 0.55, [P([R("Chat with Your Database", size=26, color=TEXT, font="Segoe UI Light")])])
b1 = rect(s, 0.65, 3.82, 2.4, 0.034, fill=GREEN); add_glow(b1, GREEN, 4, 55)
b2 = rect(s, 3.05, 3.82, 0.8, 0.034, fill=BLUE);  add_glow(b2, BLUE, 4, 55)
text(s, 0.63, 4.05, 7.5, 1.0, [
    P([R("A conversational agent that talks directly to relational databases —", size=13, color=MUTED)], sa=3, ls=1.18),
    P([R("answering with ", size=13, color=MUTED), R("exact SQL", size=13, color=GREEN, bold=True),
       R(", ", size=13, color=MUTED), R("interactive Plotly charts", size=13, color=BLUE, bold=True),
       R(", ", size=13, color=MUTED), R("live Mermaid diagrams", size=13, color=BLUE, bold=True),
       R(" and", size=13, color=MUTED)], sa=3, ls=1.18),
    P([R("mathematically grounded explanations", size=13, color=GREEN, bold=True),
       R(".", size=13, color=MUTED)], ls=1.18)])
tp = rect(s, 0.63, 5.32, 5.35, 0.54, fill=BG, line=GREEN, lw=1.2, rounded=True, adj=0.5)
set_fill_alpha(tp, 62); add_glow(tp, GREEN, 6, 45)
text(s, 0.63, 5.32, 5.35, 0.54,
     [P([R("> ", size=14.5, color=BLUE, mono=True, bold=True),
         R("NOT GUESSES. GROUND TRUTH.", size=14.5, color=GREEN, mono=True, bold=True, spc=170, glow=(GREEN, 4, 45))],
        align=PP_ALIGN.CENTER)], anchor=MSO_ANCHOR.MIDDLE)
cx = 0.63
for lbl in ["TEXT-TO-SQL", "AGENTIC LOOP", "MULTI-PROVIDER", "8-STEP REASONING"]:
    wch = 0.18 + 0.082*len(lbl)
    chip(s, cx, 6.14, wch, 0.36, lbl, color=BLUE, size=9.5)
    cx += wch + 0.17
text(s, 0.63, 6.92, 8.0, 0.3,
     [P([R("PITCH DECK & TECHNICAL OVERVIEW", size=9, color=FAINT, mono=True, spc=280)])])
notes(s, "Full-bleed cyber-grid scene. Holographic wireframe database cylinder (right) streaming neon-green particles into a floating chat terminal. Electric-blue perspective grid floor, faint 12-star Zodiac constellation overhead. 'ZODIAC' set as a neon sign with chromatic fringing; terminal status bar on top.",
      "Open with the one-liner: 'Every company runs on data — almost nobody can talk to it directly.' Introduce Zodiac AI as a conversational agent between humans and the relational database: ask in plain English, get the exact SQL, the chart, the diagram, and the numbers — every figure traceable to an executed query. Set the hook fast; promise the architecture tour.")

# ==================================================== SLIDE 2 · EXECUTIVE VISION
s = prs.slides.add_slide(BLANK)
chrome(s, 2, "EXECUTIVE VISION", "assets/bg-subtle-blue.jpg")
kicker_title(s, "02 — EXECUTIVE VISION",
             [R("The Wall Between ", size=30, color=TEXT, bold=True, font=DISP),
              R("You and Your Data", size=30, color=BLUE, bold=True, font=DISP, glow=(BLUE, 5, 35))])
seam = rect(s, 6.62, 1.85, 0.035, 4.3, fill=BLUE)
gradient_fill(seam, [(0, BG, 0), (50, BLUE, 95), (100, BG, 0)], 90)
add_glow(seam, BLUE, 5, 50)
py = panel(s, 0.55, 1.78, 5.85, 4.5, "THE OLD WAY — TRADITIONAL BI", accent=RED)
bullets(s, 0.88, py+0.22, 5.25, 3.6, [
    ("Every question becomes a ticket", "analysts queue, answers ship days later"),
    ("Static dashboards", "answer yesterday's questions, not today's"),
    ("SQL fluency is the bottleneck", "the data exists; direct access doesn't"),
    ("Naive LLM chatbots", "hallucinate numbers with total confidence"),
], size=13, marker="✕", mcolor=RED, gap=14)
py = panel(s, 6.98, 1.78, 5.85, 4.5, "THE ZODIAC WAY", accent=GREEN)
bullets(s, 7.31, py+0.22, 5.25, 3.6, [
    ("Ask in plain English", "get the exact SQL, executed live"),
    ("Visual answers", "interactive Plotly charts + live Mermaid diagrams"),
    ("Grounded numbers", "every figure computed from real result sets"),
    ("Zero-latency analytics", "answers at the speed of conversation"),
], size=13, marker="▸", mcolor=GREEN, gap=14, mglow=True)
bottom_bar(s, "ANALYTICS AT THE SPEED OF CONVERSATION — WITHOUT THE HALLUCINATION TAX", accent=GREEN, y=6.5)
notes(s, "Split-screen divided by a pulsing electric-blue lightning seam. Left: desaturated grey-blue mosaic of stale BI dashboards, CSV exports, ticket queues, slightly glitching. Right: vivid neon chat exchange — 'What drove revenue last quarter?' answered with a glowing SQL snippet, mini chart and bullet explanation.",
      "Frame it as economics: the real cost of BI is the latency between question and answer. When latency is days, people stop asking. Generic chatbots feel instant but fabricate — speed without trust. Zodiac collapses the loop to seconds while keeping every number traceable to an executed query. Position carefully: 'We're not replacing analysts — we're unblocking everyone else.'")

# ==================================================== SLIDE 3 · TECH STACK
s = prs.slides.add_slide(BLANK)
chrome(s, 3, "TECH STACK", "assets/bg-subtle-green.jpg")
kicker_title(s, "03 — TECH STACK",
             [R("The Stack — ", size=30, color=TEXT, bold=True, font=DISP),
              R("Engineered for Speed, Built to Swap", size=30, color=BLUE, bold=True, font=DISP, glow=(BLUE, 5, 35))])
layers = [
    ("FRONTEND", GREEN, "React + TanStack Query", "Declarative UI · async server-state · smart caching & live refetch", "async · cached · live"),
    ("BACKEND",  BLUE,  "FastAPI (Python)",       "Async endpoints · Pydantic-validated payloads · native streaming", "fast · typed · streams"),
    ("AI LAYER", GREEN, "Groq + Gemini",          "Dual LLM providers behind a single router interface", "fallback-ready"),
    ("DATA",     BLUE,  "SQLite",                 "Zero-config · file-based · instantly portable — swap for Postgres/MySQL", "ships in the repo"),
]
ly = 1.78
for name, acc, tech, desc, spec in layers:
    p = rect(s, 0.55, ly, 7.5, 1.02, fill=PANEL, line=LINE, lw=1.0, rounded=True, adj=0.12)
    set_fill_alpha(p, 90); add_shadow(p, 0.15, 0.04, 55)
    top = rect(s, 0.69, ly, 7.22, 0.045, fill=acc); add_glow(top, acc, 4, 55)
    tag = rect(s, 0.55, ly, 2.3, 1.02, fill=GREEN_D if acc == GREEN else BLUE_D, rounded=True, adj=0.12)
    set_fill_alpha(tag, 92)
    rect(s, 2.83, ly+0.12, 0.028, 0.78, fill=acc)
    text(s, 0.55, ly, 2.3, 1.02, [P([R(name, size=14.5, color=acc, mono=True, bold=True, spc=110)], align=PP_ALIGN.CENTER)],
         anchor=MSO_ANCHOR.MIDDLE)
    text(s, 3.1, ly+0.15, 5.6, 0.38, [P([R(tech, size=14.5, color=TEXT, bold=True)])])
    text(s, 3.1, ly+0.54, 4.9, 0.42, [P([R(desc, size=9.5, color=MUTED)])])
    text(s, 8.9, ly, 0.62, 1.02, [P([R(spec, size=7.5, color=acc, mono=True)], align=PP_ALIGN.RIGHT)], anchor=MSO_ANCHOR.MIDDLE)
    ly += 1.17
py = panel(s, 8.35, 1.78, 4.43, 4.55, "SYSTEM CROSS-SECTION", accent=BLUE)
picture_fit(s, "assets/stack-visual.jpg", 8.55, py+0.08, 4.03, 3.15)
text(s, 8.55, py+3.4, 4.03, 0.3,
     [P([R("SWAP ANY LAYER — THE AGENT DOESN'T CARE", size=8, color=BLUE, mono=True, spc=60)], align=PP_ALIGN.CENTER)])
bottom_bar(s, "DESIGN PRINCIPLE — EVERY LAYER IS MODULAR. SWAP ANY COMPONENT WITHOUT REWRITING THE AGENT", accent=BLUE, y=6.55)
notes(s, "Vertical layered-architecture diagram rendered as a neon circuit-board cross-section (chip die-shot). Four glowing strata: FRONTEND, API, AI (two chips GROQ/GEMINI through a router switch), DATA (green SQLite cylinder). Data beams travel through vias between layers; monospace spec labels on the right edge. Side panel: 3D render of the four-layer glass stack.",
      "Two reasons for these choices: iteration velocity and architectural portability. React + TanStack handles the chat-plus-live-chart pattern. FastAPI gives async Python with strict validation. SQLite makes the demo infinitely portable. Most important: the AI layer is deliberately abstracted — the agent never talks to a vendor, it talks to the router.")

# ==================================================== SLIDE 4 · MULTI-PROVIDER ARCHITECTURE
s = prs.slides.add_slide(BLANK)
chrome(s, 4, "AI ARCHITECTURE", "assets/bg-subtle-blue.jpg")
kicker_title(s, "04 — MULTI-PROVIDER AI ARCHITECTURE",
             [R("Two Brains, One Loop — ", size=30, color=TEXT, bold=True, font=DISP),
              R("Never Offline", size=30, color=BLUE, bold=True, font=DISP, glow=(BLUE, 5, 35))])
py = panel(s, 0.55, 1.78, 5.35, 4.55, "PROVIDER FALLBACK ROUTER", accent=BLUE)
bullets(s, 0.86, py+0.14, 4.75, 2.1, [
    ("One unified LLM interface", "provider-agnostic by design"),
    ("Primary: Groq", "near-instant Llama-class inference"),
    ("Auto-failover to Gemini", "on rate limits, timeouts, outages"),
    ("Zero single points of failure", "run with one key or both"),
], size=11, marker="▸", mcolor=BLUE, gap=7)
picture_fit(s, "assets/router-visual.jpg", 0.8, py+2.32, 4.85, 1.62)
px = 6.1; pw = 6.68
py2 = panel(s, px, 1.78, pw, 4.55, "THE 8-STEP AGENT LOOP", accent=GREEN)
steps = [("INGEST", "parse prompt + context"), ("INTROSPECT", "pull live schema"),
         ("PLAN", "decompose, choose tools"), ("ACT", "execute next tool call"),
         ("OBSERVE", "inspect results / errors"), ("VALIDATE", "complete & grounded?"),
         ("ITERATE", "loop 3 → 6 until yes"), ("SYNTHESIZE", "SQL + chart + explanation")]
for i, (lbl, d) in enumerate(steps):
    col, row = divmod(i, 4)
    sx = px + 0.3 + col*3.2; sy = py2 + 0.16 + row*0.9
    num_badge(s, sx, sy, i+1)
    text(s, sx+0.56, sy-0.04, 2.55, 0.3, [P([R(lbl, size=12, color=TEXT, bold=True, mono=True, spc=60)])])
    text(s, sx+0.56, sy+0.24, 2.55, 0.26, [P([R(d, size=8.5, color=MUTED, mono=True)])])
text(s, px+0.3, py2+3.8, pw-0.6, 0.3,
     [P([R("▸ steps 3 → 6 repeat until the answer is complete and grounded", size=9.5, color=BLUE, mono=True)])])
bottom_bar(s, "ANSWERS ARE EARNED THROUGH ITERATION — NOT GUESSED ON THE FIRST PASS", accent=GREEN, y=6.55)
notes(s, "Left panel: 3D render of the provider router — central cube node, bright neon-green GROQ path severed mid-flight with red glitch sparks, electric-blue GEMINI path actively carrying particle traffic. Right panel: orbital ring of 8 nodes with a directional energy pulse around an AGENT CONTEXT core; segment 4→7 glows brighter, labeled ITERATE UNTIL GROUNDED.",
      "Two ideas. First, resilience: LLM providers are utilities, and utilities go down — the router treats Groq and Gemini as interchangeable, so when Groq rate-limits mid-session the user never notices. Second, the loop: this is not one-shot prompt → answer. The agent reasons, acts, observes, iterates — if the first query errors or returns partial data, it self-corrects. Answers are earned through iteration.")

# ==================================================== SLIDE 5 · CORE TOOLS 1
s = prs.slides.add_slide(BLANK)
chrome(s, 5, "CORE TOOLS 1/3", "assets/bg-subtle-green.jpg")
kicker_title(s, "05 — CORE TOOLS · 1 OF 3",
             [R("get_schema", size=29, color=GREEN, bold=True, mono=True, glow=(GREEN, 5, 35)),
              R("  &  ", size=29, color=MUTED, bold=True, font=DISP),
              R("execute_query", size=29, color=GREEN, bold=True, mono=True, glow=(GREEN, 5, 35))])
ty = terminal(s, 0.55, 1.78, 5.85, 4.15, "get_schema", accent=GREEN)
text(s, 0.9, ty+0.1, 5.2, 1.85, [
    P([R("sales_db", size=12.5, color=BLUE, mono=True, bold=True, glow=(BLUE, 3, 35))], sa=3),
    P([R(" ├─ ", size=12.5, color=FAINT, mono=True), R("customers", size=12.5, color=TEXT, mono=True)], sa=3),
    P([R(" ├─ ", size=12.5, color=FAINT, mono=True), R("products", size=12.5, color=TEXT, mono=True)], sa=3),
    P([R(" ├─ ", size=12.5, color=FAINT, mono=True), R("orders", size=12.5, color=TEXT, mono=True)], sa=3),
    P([R(" └─ ", size=12.5, color=FAINT, mono=True), R("order_items", size=12.5, color=TEXT, mono=True)], sa=4),
    P([R("   tables · columns · types · relations", size=8.5, color=FAINT, mono=True)])])
bullets(s, 0.9, ty+2.25, 5.2, 1.0, [
    ("Live runtime introspection", "no hardcoded, stale schema assumptions"),
], size=11.5, gap=6)
ty = terminal(s, 6.98, 1.78, 5.85, 4.15, "execute_query", accent=GREEN)
text(s, 7.33, ty+0.1, 5.2, 1.6, [
    P([R("SELECT c.name, SUM(oi.qty*oi.price)", size=10.5, color=GREEN, mono=True)], sa=1),
    P([R("FROM order_items oi", size=10.5, color=GREEN, mono=True)], sa=1),
    P([R("JOIN orders o ON o.id = oi.order_id", size=10.5, color=GREEN, mono=True)], sa=1),
    P([R("JOIN customers c ON c.id = o.customer_id", size=10.5, color=GREEN, mono=True)], sa=1),
    P([R("GROUP BY c.name;", size=10.5, color=GREEN, mono=True)], sa=5),
    P([R("✓ 1,284 rows returned", size=10, color=BLUE, mono=True, bold=True, glow=(BLUE, 3, 40)),
       R("  [READ-ONLY ✓]", size=10, color=GREEN, mono=True, bold=True, glow=(GREEN, 3, 40))])])
bullets(s, 7.33, ty+2.0, 5.2, 0.7, [
    ("Exact SQL", "generated & executed on the real database"),
], size=11.5, gap=5)
blk = rect(s, 7.33, ty+2.72, 5.15, 0.52, fill=RED_D, line=RED, lw=1.1, rounded=True, adj=0.28)
add_glow(blk, RED, 4, 45)
text(s, 7.5, ty+2.72, 4.9, 0.52,
     [P([R("DROP TABLE orders;", size=10, color=RED, mono=True, strike=True),
         R("  ⚠ BLOCKED — WRITE GUARD", size=9.5, color=RED, mono=True, bold=True)])],
     anchor=MSO_ANCHOR.MIDDLE)
bottom_bar(s, "RADICAL TRANSPARENCY — EVERY ANSWER SHIPS WITH THE SQL THAT PRODUCED IT. COPY IT. AUDIT IT. RERUN IT.", accent=GREEN, y=6.15)
notes(s, "Two floating terminals with CRT scanlines. Left (get_schema): glowing schema tree — sales_db → customers/products/orders/order_items — column types cascading in green monospace. Right (execute_query): a SELECT+JOIN executing with a green READ-ONLY badge; beneath it, a red-glowing rejected 'DROP TABLE orders;' stamped BLOCKED — WRITE GUARD.",
      "Trust is the product, and these two tools are its foundation. The agent reads the schema fresh at runtime — it cannot reference a column that doesn't exist — and it executes real SQL, so answers come from the database, not the model's memory. The read-only guard means the worst-case interaction is an empty result set. Users always see the exact SQL behind every number.")

# ==================================================== SLIDE 6 · CORE TOOLS 2
s = prs.slides.add_slide(BLANK)
chrome(s, 6, "CORE TOOLS 2/3", "assets/bg-subtle-blue.jpg")
kicker_title(s, "06 — CORE TOOLS · 2 OF 3",
             [R("generate_chart", size=29, color=BLUE, bold=True, mono=True, glow=(BLUE, 5, 35)),
              R("  &  ", size=29, color=MUTED, bold=True, font=DISP),
              R("generate_flowchart", size=29, color=BLUE, bold=True, mono=True, glow=(BLUE, 5, 35))])
py = panel(s, 0.55, 1.78, 7.8, 2.24, "generate_chart — PLOTLY", accent=BLUE)
bullets(s, 0.88, py+0.12, 7.15, 1.6, [
    ("Specs authored from real query results", "the model writes the definition, the browser renders it"),
    ("Fully interactive", "zoom · hover · pan · filter — exploration, not screenshots"),
    ("Deterministic rendering", "bar · line · scatter · pie, matched to the question — no hallucinated visuals"),
], size=11.5, marker="▸", mcolor=BLUE, gap=7)
py = panel(s, 0.55, 4.24, 7.8, 2.24, "generate_flowchart — MERMAID", accent=BLUE)
bullets(s, 0.88, py+0.12, 7.15, 1.6, [
    ("Mermaid syntax, written live by the agent", "diagrams render inside the chat"),
    ("Schema maps & ER diagrams", "entity relationships on demand"),
    ("Logic & pipeline flows", "the shape of the system, not just the numbers"),
], size=11.5, marker="▸", mcolor=BLUE, gap=7)
py = panel(s, 8.6, 1.78, 4.18, 4.7, "LIVE RENDER", accent=GREEN)
picture_fit(s, "assets/viz-visual-square.jpg", 8.78, py+0.1, 3.82, 3.42)
text(s, 8.78, py+3.66, 3.82, 0.3,
     [P([R("PLOTLY × MERMAID — RENDERED LIVE", size=8, color=GREEN, mono=True, spc=60)], align=PP_ALIGN.CENTER)])
bottom_bar(s, "ONE PROMPT → THE NUMBERS AND THE PICTURE OF THE NUMBERS", accent=BLUE, y=6.62)
notes(s, "Left: two stacked glass panels for the Plotly and Mermaid tools. Right: hologram render — a glowing line chart above a node-graph flow diagram, particles drifting between them. Faint source code (fig = px.bar(...) / graph TD;) morphs into the visuals: the code-to-render transformation made visible.",
      "Text-only answers waste an LLM's real strength: structured generation. Zodiac treats visuals as first-class outputs — the model doesn't describe a chart, it writes the chart definition, and Plotly and Mermaid render it deterministically. Because the spec is generated from executed query results, the visual carries the same grounding as the SQL behind it.")

# ==================================================== SLIDE 7 · CORE TOOLS 3
s = prs.slides.add_slide(BLANK)
chrome(s, 7, "CORE TOOLS 3/3", "assets/bg-subtle-green.jpg")
kicker_title(s, "07 — CORE TOOLS · 3 OF 3",
             [R("explain_data", size=29, color=GREEN, bold=True, mono=True, glow=(GREEN, 5, 35)),
              R(" — Grounded, Not Guessed", size=29, color=BLUE, bold=True, font=DISP, glow=(BLUE, 5, 35))])
pipeline = [("1 · QUERY", GREEN, "against the live database"),
            ("2 · COMPUTE", BLUE, "totals · means · min/max · shares"),
            ("3 · NARRATE", GREEN, "prose around proven numbers")]
pxs = 0.85
for lbl, acc, d in pipeline:
    pc = rect(s, pxs, 1.8, 3.35, 1.0, fill=PANEL, line=acc, lw=1.4, rounded=True, adj=0.14)
    set_fill_alpha(pc, 90); add_glow(pc, acc, 5, 40); add_shadow(pc, 0.15, 0.04, 50)
    text(s, pxs, 1.93, 3.35, 0.4, [P([R(lbl, size=15, color=acc, mono=True, bold=True, glow=(acc, 3, 40))], align=PP_ALIGN.CENTER)])
    text(s, pxs+0.1, 2.34, 3.15, 0.32, [P([R(d, size=8.5, color=MUTED, mono=True)], align=PP_ALIGN.CENTER)])
    if lbl != "3 · NARRATE":
        text(s, pxs+3.38, 1.8, 0.72, 1.0, [P([R("▶▶", size=15, color=BLUE, mono=True, bold=True, glow=(BLUE, 4, 50))], align=PP_ALIGN.CENTER)], anchor=MSO_ANCHOR.MIDDLE)
    pxs += 4.1
bullets(s, 0.88, 3.25, 7.0, 2.9, [
    ("Numeric summaries from executed result sets", "totals, averages, min/max, distributions, shares — computed, not recalled"),
    ("Every figure traces to a query result", "text is generated from computed numbers, never from model memory"),
    ("The math ships with the story", "growth rates, proportions, outliers — derived first, then narrated"),
    ("Anti-hallucination by architecture", "untraceable numbers don't make the answer"),
], size=12.5, gap=13)
py = panel(s, 8.35, 3.05, 4.43, 3.15, "GROUNDING LENS", accent=GREEN)
picture_fit(s, "assets/lens-visual.jpg", 8.53, py+0.08, 4.07, 2.15)
text(s, 8.53, py+2.32, 4.07, 0.3,
     [P([R("EVERY NUMBER TRACED TO A ROW", size=8, color=GREEN, mono=True, spc=60)], align=PP_ALIGN.CENTER)])
q = rect(s, 0.55, 6.5, SW-1.1, 0.5, fill=GREEN_D, line=GREEN, lw=1.1, rounded=True, adj=0.5)
set_fill_alpha(q, 85); add_glow(q, GREEN, 5, 40)
text(s, 0.9, 6.5, SW-1.8, 0.5,
     [P([R("> ", size=12, color=GREEN, mono=True, bold=True),
         R("We don't ask the model to remember your revenue. We ask your database.", size=12, color=GREEN, mono=True, bold=True)])],
     anchor=MSO_ANCHOR.MIDDLE)
notes(s, "A neon magnifying lens hovers over a glowing result-set table; cells illuminate green and lift out, flowing upward into a paragraph where the same numbers appear highlighted — numbers traced to rows. Math glyphs (∑, μ, x̄, Δ%) orbit the lens. In the corner, a ghostly chatbot bubble with a made-up statistic is slashed in red.",
      "This slide separates Zodiac from a wrapped chat API. Generic LLMs interpolate plausible-sounding statistics; Zodiac computes the statistics first and writes the narrative around them. The pipeline is strict: query → compute → narrate. If a number can't be traced to a result set, it doesn't make it into the answer. Say it plainly: 'We don't ask the model to remember your revenue. We ask your database.'")

# ==================================================== SLIDE 8 · DATABASE ARCHITECTURE
s = prs.slides.add_slide(BLANK)
chrome(s, 8, "DATA MODEL", "assets/bg-subtle-blue.jpg")
kicker_title(s, "08 — DATABASE ARCHITECTURE",
             [R("Under the Hood — ", size=30, color=TEXT, bold=True, font=DISP),
              R("The Sample Sales Schema", size=30, color=BLUE, bold=True, font=DISP, glow=(BLUE, 5, 35))])
tables = [
    ("CUSTOMERS", GREEN, "who is buying", "customer_id · name\ncity · signup_date"),
    ("ORDERS",    BLUE,  "the transaction header", "order_id · customer_id\ndate · status"),
    ("ORDER_ITEMS", GREEN, "line-level detail", "order_id · product_id\nqty · price"),
    ("PRODUCTS",  BLUE,  "what is being sold", "product_id · name\ncategory · price"),
]
rels = [("1 — ∞", True), ("1 — ∞", True), ("∞ — 1", False)]
tx = 0.55; tw = 2.72; tg = (SW-1.1-4*tw)/3
for i, (name, acc, d, cols) in enumerate(tables):
    p = rect(s, tx, 1.88, tw, 1.66, fill=PANEL, line=acc, lw=1.3, rounded=True, adj=0.09)
    set_fill_alpha(p, 90); add_shadow(p, 0.16, 0.045, 55); add_glow(p, acc, 4, 22)
    top = rect(s, tx+0.12, 1.88, tw-0.24, 0.045, fill=acc); add_glow(top, acc, 5, 60)
    text(s, tx, 2.04, tw, 0.32, [P([R(name, size=12.5, color=acc, mono=True, bold=True, spc=70)], align=PP_ALIGN.CENTER)])
    text(s, tx+0.1, 2.4, tw-0.2, 0.28, [P([R(d, size=9.5, color=TEXT)], align=PP_ALIGN.CENTER)])
    text(s, tx+0.08, 2.72, tw-0.16, 0.72,
         [P([R(cols.split("\n")[0], size=8, color=MUTED, mono=True)], align=PP_ALIGN.CENTER, sa=2),
          P([R(cols.split("\n")[1], size=8, color=MUTED, mono=True)], align=PP_ALIGN.CENTER)])
    if i < 3:
        lbl, fwd = rels[i]
        text(s, tx+tw-0.05, 2.24, tg+0.1, 0.3,
             [P([R("━━━▶" if fwd else "◀━━━", size=10, color=BLUE, mono=True, glow=(BLUE, 3, 45))], align=PP_ALIGN.CENTER)])
        text(s, tx+tw-0.05, 2.5, tg+0.1, 0.3, [P([R(lbl, size=11, color=BLUE, mono=True, bold=True)], align=PP_ALIGN.CENTER)])
    tx += tw + tg
bullets(s, 0.88, 3.95, 11.6, 1.85, [
    ("Normalized e-commerce model — four tables, one star", "join-rich by design: every interesting question crosses at least one join"),
    ("The schema that stresses text-to-SQL edge cases", "if the agent is correct here, it generalizes"),
    ("Nothing is hardcoded", "the agent maps all relationships at runtime via get_schema"),
], size=12.5, gap=12)
bottom_bar(s, "SCHEMA DISCOVERED LIVE — NEVER BAKED INTO PROMPTS", accent=GREEN, y=6.15)
notes(s, "An ERD reimagined as a neon circuit schematic on a dark PCB. Four table-nodes as glowing chips — CUSTOMERS (green), PRODUCTS (blue), ORDERS (green), ORDER_ITEMS (blue) — connected by luminous traces with pulsing cardinality glyphs (1 — ∞) at each junction. Pin-labels expose key columns. Faint grid and ruler ticks frame it like engineering blueprints.",
      "The demo dataset is deliberately canonical e-commerce: instantly understandable, and join-heavy — exactly where naive text-to-SQL breaks. Every interesting question (top customers, revenue per product, average order value) crosses at least one join. If the agent is correct on this schema, it generalizes. And the schema is never baked into prompts — it's discovered live.")

# ==================================================== SLIDE 9 · SAMPLE PROMPTS
s = prs.slides.add_slide(BLANK)
chrome(s, 9, "SAMPLE PROMPTS", "assets/bg-subtle-green.jpg")
kicker_title(s, "09 — VALID QUERY CATEGORIES & SAMPLE PROMPTS",
             [R("Ask It Anything ", size=30, color=TEXT, bold=True, font=DISP),
              R("(Structured)", size=30, color=BLUE, bold=True, font=DISP, glow=(BLUE, 5, 35))])
cards = [
    ("AGGREGATION", '"What was total revenue\nper month this year?"'),
    ("RANKING",     '"Show me the top 10 customers\nby lifetime value."'),
    ("TREND",       '"Plot quarterly sales growth\nas a line chart."'),
    ("JOIN",        '"Which products are most\nfrequently bought together?"'),
    ("SCHEMA",      '"Diagram how my tables\nrelate to each other."'),
    ("QUALITY",     '"Are there any orders\nwith no line items?"'),
]
cw = (SW-1.1-0.44)/3
for i, (cat, q) in enumerate(cards):
    col, row = divmod(i, 3)
    cx = 0.55 + col*(cw+0.22); cy = 1.85 + row*1.72
    p = rect(s, cx, cy, cw, 1.52, fill=PANEL, line=LINE, lw=1.0, rounded=True, adj=0.09)
    set_fill_alpha(p, 90); add_shadow(p, 0.15, 0.04, 55)
    acc = BLUE if i % 2 == 0 else GREEN
    edge = rect(s, cx, cy+0.14, 0.055, 1.24, fill=acc); add_glow(edge, acc, 4, 55)
    text(s, cx+0.26, cy+0.17, cw-0.45, 0.3,
         [P([R("// ", size=10, color=FAINT, mono=True), R(cat, size=10.5, color=acc, mono=True, bold=True, spc=130)])])
    l1, l2 = q.split("\n")
    text(s, cx+0.26, cy+0.52, cw-0.45, 0.85,
         [P([R(l1, size=12, color=TEXT, italic=True)], ls=1.15, sa=1),
          P([R(l2, size=12, color=TEXT, italic=True)], ls=1.15)])
bottom_bar(s, "EVERY ANSWER = EXACT SQL + INTERACTIVE CHART + GROUNDED EXPLANATION", accent=GREEN, y=5.6)
ap = rect(s, 0.55, 6.28, 4.35, 0.5, fill=PANEL_D, line=LINE, lw=1.0, rounded=True, adj=0.5)
set_fill_alpha(ap, 90); add_shadow(ap, 0.13, 0.04, 50)
text(s, 0.85, 6.28, 4.0, 0.5,
     [P([R("> ", size=13, color=GREEN, mono=True, bold=True),
         R("ask anything", size=13, color=TEXT, mono=True),
         R("_", size=13, color=GREEN, mono=True, bold=True, glow=(GREEN, 4, 60))])],
     anchor=MSO_ANCHOR.MIDDLE)
notes(s, "Dark-mode chat UI with user message bubbles in neon outline, each paired with an accent category chip (AGGREGATION, RANKING, TREND, JOIN, SCHEMA, QUALITY). One exchange expanded to show the full multi-part response: SQL block, mini Plotly chart, explanation paragraph. An empty input field with a blinking cursor: '> ask anything_'",
      "Walk one or two bubbles aloud, then let the audience read the rest. Point out the data-quality category last — it proves the agent can reason about the database, not just retrieve from it. Close on the bundle: a single prompt yields the SQL, a rendered chart, and a grounded explanation. That triple — query, visual, narrative — is the product experience.")

# ==================================================== SLIDE 10 · GETTING STARTED
s = prs.slides.add_slide(BLANK)
chrome(s, 10, "GETTING STARTED", "assets/bg-subtle-blue.jpg")
kicker_title(s, "10 — GETTING STARTED",
             [R("Zero to Query in ", size=30, color=TEXT, bold=True, font=DISP),
              R("Five Minutes", size=30, color=GREEN, bold=True, font=DISP, glow=(GREEN, 5, 35))])
ty = terminal(s, 0.55, 1.78, 7.55, 4.4, "boot-sequence — zodiac-ai", accent=GREEN)
text(s, 0.9, ty+0.12, 6.95, 3.6, [
    P([R("$ ", size=11.5, color=GREEN, mono=True, bold=True, glow=(GREEN, 3, 45)),
       R("git clone https://github.com/<org>/zodiac-ai.git", size=11.5, color=TEXT, mono=True)], sa=4),
    P([R("$ ", size=11.5, color=GREEN, mono=True, bold=True, glow=(GREEN, 3, 45)),
       R("cd zodiac-ai", size=11.5, color=TEXT, mono=True)], sa=9),
    P([R("# .env", size=10.5, color=FAINT, mono=True)], sa=4),
    P([R("GROQ_API_KEY", size=11, color=GREEN, mono=True, bold=True),
       R("=••••••••••", size=11, color=TEXT, mono=True),
       R("   # primary provider", size=9.5, color=FAINT, mono=True)], sa=4),
    P([R("GEMINI_API_KEY", size=11, color=BLUE, mono=True, bold=True),
       R("=••••••••••", size=11, color=TEXT, mono=True),
       R("   # optional fallback", size=9.5, color=FAINT, mono=True)], sa=9),
    P([R("$ ", size=11.5, color=GREEN, mono=True, bold=True, glow=(GREEN, 3, 45)),
       R("pip install -r requirements.txt", size=11.5, color=TEXT, mono=True),
       R("  # backend ✓", size=9.5, color=FAINT, mono=True)], sa=4),
    P([R("$ ", size=11.5, color=GREEN, mono=True, bold=True, glow=(GREEN, 3, 45)),
       R("npm install && npm run dev", size=11.5, color=TEXT, mono=True),
       R("         # frontend ✓", size=9.5, color=FAINT, mono=True)], sa=8),
    P([R("[", size=11, color=FAINT, mono=True), R("██████████░░", size=11, color=GREEN, mono=True, glow=(GREEN, 3, 40)),
       R("]", size=11, color=FAINT, mono=True), R(" LOADING...", size=10, color=MUTED, mono=True)], sa=8),
    P([R("▲ ", size=12, color=BLUE, mono=True, bold=True, glow=(BLUE, 3, 45)),
       R("open the local URL  →  start asking", size=12, color=BLUE, mono=True, bold=True)])])
py = panel(s, 8.4, 1.78, 4.38, 4.4, "RUN NOTES", accent=BLUE)
bullets(s, 8.7, py+0.15, 3.85, 3.5, [
    ("Prerequisites", "Python 3.10+ · Node 18+ · Git · ≥ 1 API key"),
    ("No DB server", "SQLite ships inside the repo"),
    ("One key runs it", "a second key activates the failover router"),
    ("< 5 minutes", "from git clone to first grounded answer"),
], size=11, marker="▸", mcolor=BLUE, gap=9)
bottom_bar(s, "CLONE IT DURING Q&A — THE DEMO IS THE PITCH", accent=GREEN, y=6.42)
notes(s, "A full terminal styled as a boot sequence with CRT scanlines: command lines with a typing cursor, executed lines flashing green with ✓, a progress-bar motif [████████░░] LOADING. The final command launches the app. Side panel: run notes. The .env keys are masked.",
      "The pitch of this slide is friction — or its absence. SQLite means no database server to stand up: the schema ships inside the repo. One API key is enough to run; the second is optional insurance that activates failover. From git clone to first grounded answer: under five minutes. Invite the audience to do it live during Q&A.")

# ==================================================== SLIDE 11 · CONCLUSION & Q&A
s = prs.slides.add_slide(BLANK)
s.shapes.add_picture("assets/hero-closing.jpg", 0, 0, Inches(SW), Inches(SH))
scrim = rect(s, 0, 0, 9.0, SH, fill=BG)
gradient_fill(scrim, [(0, BG, 95), (60, BG, 78), (100, BG, 0)], 0)
r1 = rect(s, 0, 0, SW*0.42, 0.04, fill=GREEN); add_glow(r1, GREEN, 5, 45)
r2 = rect(s, SW*0.42, 0, SW*0.58, 0.04, fill=BLUE);  add_glow(r2, BLUE, 5, 45)
text(s, 0.62, 0.5, 9.0, 0.3,
     [P([R("$ status: ", size=11, color=MUTED, mono=True),
         R("ALL SYSTEMS OPERATIONAL", size=11, color=GREEN, mono=True, bold=True, glow=(GREEN, 4, 50))])])
text(s, 0.62, 0.98, 9.6, 0.6,
     [P([R("The Database Is Now ", size=31, color=TEXT, bold=True, font=DISP),
         R("a Conversation", size=31, color=BLUE, bold=True, font=DISP, glow=(BLUE, 5, 35))])])
b1 = rect(s, 0.64, 1.68, 1.7, 0.034, fill=GREEN); add_glow(b1, GREEN, 4, 55)
b2 = rect(s, 2.34, 1.68, 0.55, 0.034, fill=BLUE);  add_glow(b2, BLUE, 4, 55)
bullets(s, 0.62, 2.0, 8.3, 3.2, [
    ("What we built", "exact SQL · interactive Plotly charts · live Mermaid diagrams · grounded explanations"),
    ("Why it matters", "analytics at conversational speed — without the hallucination tax"),
    ("How it holds up", "read-only guardrails · runtime schema awareness · multi-provider failover · 8-step reasoning"),
    ("What's next", "Postgres / MySQL connectors · richer chart grammar · conversational follow-ups · shared query memory"),
], size=12.5, gap=12, mglow=True)
cta = rect(s, 0.62, 5.42, 4.6, 0.54, fill=BG, line=GREEN, lw=1.2, rounded=True, adj=0.5)
set_fill_alpha(cta, 62); add_glow(cta, GREEN, 6, 45)
text(s, 0.62, 5.42, 4.6, 0.54,
     [P([R("▸ ", size=14, color=GREEN, mono=True, bold=True),
         R("CLONE IT. QUERY IT. BREAK IT.", size=13, color=GREEN, mono=True, bold=True, spc=100, glow=(GREEN, 3, 35))],
        align=PP_ALIGN.CENTER)], anchor=MSO_ANCHOR.MIDDLE)
display(s, 8.95, 4.35, 3.9, 1.7, "Q&A", 92)
text(s, 9.0, 6.12, 3.8, 0.35,
     [P([R("> ask anything", size=13, color=TEXT, mono=True),
         R("_", size=13, color=GREEN, mono=True, bold=True, glow=(GREEN, 4, 60))], align=PP_ALIGN.CENTER)])
text(s, 0.62, 6.95, 8.0, 0.3,
     [P([R("ZODIAC AI // THANK YOU", size=9, color=FAINT, mono=True, spc=280)])])
notes(s, "Wide cinematic shot of the Zodiac constellation — 12 stars connected by neon lines — reflected on the cyber-grid floor, closing the visual loop with Slide 1. Data streams flow from the constellation down into a small glowing database cylinder. Terminal reads '$ status: ALL SYSTEMS OPERATIONAL' with a blinking '> ask anything_'. Oversized Q&A as a neon sign with chromatic fringing.",
      "Close the loop with the opening line: databases used to be things you queried — now they're things you talk to. Recap the four pillars in one breath: exact SQL, live visuals, grounded numbers, resilient architecture. Open the floor — and if the room is quiet, seed the first question: 'Ask me what happens when Groq goes down mid-query.'")

# ---------------------------------------------------------------- save
prs.core_properties.title = "Zodiac AI — Pitch Deck & Technical Overview"
prs.core_properties.author = "Zodiac AI"
prs.core_properties.subject = "Chat with your Database — text-to-SQL agent"
OUT = "Zodiac_AI_Pitch_Deck.pptx"
prs.save(OUT)
print(f"saved {OUT} — {len(prs.slides._sldIdLst)} slides")
