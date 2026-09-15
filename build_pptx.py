#!/usr/bin/env python3
"""
ZODIAC AI — pitch deck builder v3 (MIDNIGHT AURORA)
Linear/Vercel-grade aesthetic: near-black canvas, violet→indigo aurora glows,
hairline rules, gradient display text, generous restraint.
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
BG      = "0A0B10";  PANEL  = "10131C";  PANEL_D = "0D0F16"
LINE    = "262B3D";  INK    = "1A1E2B";  HAIR    = "2E3348"
VIO     = "A78BFA";  VIO_D  = "1D1737";  VIO_L   = "3B3160"   # violet (primary)
IND     = "818CF8";  IND_D  = "151B36"                          # indigo (secondary)
LAV     = "C4B5FD"                                             # lavender (code)
RED     = "F87171";  RED_D  = "2A161C"
TEXT    = "F4F6FB";  MUTED  = "9AA3B8";  FAINT  = "5D6578"

BODY = "Segoe UI";  LIGHT = "Segoe UI Light";  MONO = "Consolas"

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

def add_glow(shape, color, rad_pt=4, alpha=30):
    spPr = shape._element.spPr
    for e in spPr.findall(qn('a:effectLst')): spPr.remove(e)
    eff = etree.Element(qn('a:effectLst'))
    g = _sub(eff, 'a:glow', rad=Pt(rad_pt))
    c = _sub(g, 'a:srgbClr', val=color); _sub(c, 'a:alpha', val=int(alpha*1000))
    spPr.append(eff)

def add_shadow(shape, blur=0.2, dist=0.05, alpha=45):
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
    spPr = shape._element.spPr
    for e in spPr.findall(qn('a:solidFill')): spPr.remove(e)
    grad = etree.Element(qn('a:gradFill')); grad.set('rotWithShape', '1')
    lst = _sub(grad, 'a:gsLst')
    for pos, hx, al in stops:
        gs = _sub(lst, 'a:gs', pos=int(pos*1000))
        c = _sub(gs, 'a:srgbClr', val=hx)
        if al is not None: _sub(c, 'a:alpha', val=int(al*1000))
    _sub(grad, 'a:lin', ang=int(angle_deg*60000))
    ln = spPr.find(qn('a:ln'))
    if ln is not None: ln.addprevious(grad)
    else: spPr.append(grad)

def run_alpha(run, pct):
    rPr = run._r.get_or_add_rPr()
    sf = rPr.find(qn('a:solidFill'))
    if sf is not None:
        c = sf.find(qn('a:srgbClr'))
        if c is not None: _sub(c, 'a:alpha', val=int(pct*1000))

def run_gradient(run, stops, angle_deg=0):
    rPr = run._r.get_or_add_rPr()
    for tag in ('a:noFill', 'a:solidFill', 'a:gradFill'):
        for e in rPr.findall(qn(tag)): rPr.remove(e)
    grad = etree.Element(qn('a:gradFill'))
    lst = _sub(grad, 'a:gsLst')
    for pos, hx, al in stops:
        gs = _sub(lst, 'a:gs', pos=int(pos*1000))
        c = _sub(gs, 'a:srgbClr', val=hx)
        if al is not None: _sub(c, 'a:alpha', val=int(al*1000))
    _sub(grad, 'a:lin', ang=int(angle_deg*60000))
    latin = rPr.find(qn('a:latin'))
    if latin is not None: latin.addprevious(grad)
    else: rPr.insert(0, grad)

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

def R(t, size=12, color=TEXT, bold=False, italic=False, mono=False, strike=False,
      spc=None, alpha=None, font=None, tgrad=None):
    return (t, dict(size=size, color=color, bold=bold, italic=italic, mono=mono,
                    strike=strike, spc=spc, alpha=alpha, font=font, tgrad=tgrad))

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
            if st.get('tgrad') is not None:
                stops, ang = st['tgrad']; run_gradient(r, stops, ang)
    return tb

def picture_fit(s, path, x, y, maxW, maxH):
    iw, ih = Image.open(path).size
    ar = iw / ih
    w = maxW; h = w / ar
    if h > maxH: h = maxH; w = h * ar
    px = x + (maxW - w) / 2; py = y + (maxH - h) / 2
    pic = s.shapes.add_picture(path, Inches(px), Inches(py), Inches(w), Inches(h))
    try: add_shadow(pic, 0.22, 0.05, 55)
    except Exception: pass
    return pic, px, py, w, h

# ---------------------------------------------------------------- components
GRAD_TEXT = ([(0, "FFFFFF", None), (55, "C4B5FD", None), (100, "818CF8", None)], 0)

def chrome(s, num, tag, bg):
    """slide chrome: aurora bg, ghost numeral, hairlines, footer, progress"""
    s.shapes.add_picture(bg, 0, 0, Inches(SW), Inches(SH))
    text(s, 10.15, 0.0, 2.75, 1.8,
         [P([R(f"{num:02d}", size=115, color="FFFFFF", font=LIGHT, alpha=5)], align=PP_ALIGN.RIGHT)])
    hl = rect(s, 0, 0, SW, 0.025, fill=VIO)
    gradient_fill(hl, [(0, VIO, 70), (45, IND, 50), (100, IND, 0)], 0)
    rect(s, 0.55, 7.05, SW-1.1, 0.011, fill=LINE)
    text(s, 0.55, 7.11, 6.6, 0.22,
         [P([R("ZODIAC AI ", size=8, color=VIO, mono=True, bold=True),
             R("// CHAT WITH YOUR DATABASE // ", size=8, color=FAINT, mono=True),
             R(tag, size=8, color=IND, mono=True)])])
    text(s, SW-1.6, 7.11, 1.05, 0.22,
         [P([R(f"{num:02d}", size=8, color=VIO, mono=True, bold=True),
             R(" / 11", size=8, color=FAINT, mono=True)], align=PP_ALIGN.RIGHT)])
    seg_w, gap = 0.30, 0.085
    x0 = (SW - (11*seg_w + 10*gap)) / 2
    for i in range(11):
        if i == num-1:
            sg = rect(s, x0+i*(seg_w+gap), 7.30, seg_w, 0.045, fill=VIO)
            add_glow(sg, VIO, 4, 30)
        else:
            rect(s, x0+i*(seg_w+gap), 7.30, seg_w, 0.045,
                 fill="2A2F49" if i < num-1 else INK)

def kicker_title(s, kicker, title_runs):
    k = rect(s, 0.57, 0.53, 0.09, 0.09, fill=VIO, rounded=True, adj=0.5)
    add_glow(k, VIO, 3, 35)
    text(s, 0.78, 0.45, 11.5, 0.3, [P([R(kicker, size=10, color=VIO, spc=230)])])
    text(s, 0.55, 0.76, 12.3, 0.62, [P(title_runs)])
    bar = rect(s, 0.57, 1.47, 2.3, 0.028, fill=VIO)
    gradient_fill(bar, [(0, VIO, 95), (70, IND, 65), (100, IND, 0)], 0)

def panel(s, x, y, w, h, title, accent=VIO, glass=82):
    p = rect(s, x, y, w, h, fill=PANEL, line=LINE, lw=1.0, rounded=True, adj=0.05)
    set_fill_alpha(p, glass); add_shadow(p, 0.2, 0.05, 45)
    d = rect(s, x+0.26, y+0.225, 0.09, 0.09, fill=accent, rounded=True, adj=0.5)
    add_glow(d, accent, 3, 35)
    text(s, x+0.48, y+0.17, w-0.75, 0.3, [P([R(title, size=10.5, color=accent, spc=150)])])
    return y + 0.52

def bullets(s, x, y, w, h, items, size=12.5, marker="▸", mcolor=VIO, gap=10):
    paras = []
    for it in items:
        if isinstance(it, tuple):
            lead, rest = it
            runs = [R(marker+"  ", size=size, color=mcolor, bold=True, mono=True),
                    R(lead, size=size, color=TEXT, bold=True)]
            if rest: runs.append(R(" — " + rest, size=size, color=MUTED))
        else:
            runs = [R(marker+"  ", size=size, color=mcolor, bold=True, mono=True),
                    R(it, size=size, color=TEXT)]
        paras.append(P(runs, sa=gap, ls=1.12, hang=0.26))
    text(s, x, y, w, h, paras)

def chip(s, x, y, w, h, label, color=MUTED, fill=None, size=9.5, border=HAIR):
    rect(s, x, y, w, h, fill=fill if fill else BG, line=border, lw=1.0, rounded=True, adj=0.5,
         alpha=None if fill else 55)
    text(s, x, y, w, h, [P([R(label, size=size, color=color, mono=True)], align=PP_ALIGN.CENTER)],
         anchor=MSO_ANCHOR.MIDDLE)

def bottom_bar(s, label, accent=VIO, y=6.55):
    b = rect(s, 0.55, y, SW-1.1, 0.46, fill=PANEL, line=LINE, lw=1.0, rounded=True, adj=0.5)
    set_fill_alpha(b, 88); add_shadow(b, 0.16, 0.04, 40)
    d = rect(s, 0.78, y+0.185, 0.09, 0.09, fill=accent, rounded=True, adj=0.5)
    add_glow(d, accent, 3, 40)
    text(s, 1.02, y, SW-2.2, 0.46, [P([R(label, size=10, color=TEXT, spc=70)])],
         anchor=MSO_ANCHOR.MIDDLE)

def terminal(s, x, y, w, h, label, accent=VIO):
    p = rect(s, x, y, w, h, fill=PANEL_D, line=LINE, lw=1.0, rounded=True, adj=0.045)
    set_fill_alpha(p, 93); add_shadow(p, 0.2, 0.05, 45)
    for i, dc in enumerate(["D0707F", "CFAF6B", "7CB88A"]):
        d = s.shapes.add_shape(MSO_SHAPE.OVAL, Inches(x+0.18+i*0.17), Inches(y+0.17), Inches(0.09), Inches(0.09))
        d.shadow.inherit = False; d.fill.solid(); d.fill.fore_color.rgb = C(dc); d.line.fill.background()
    text(s, x+0.82, y+0.13, w-1.05, 0.26,
         [P([R("zodiac:~ ", size=10, color=FAINT, mono=True),
             R(label, size=10, color=accent, mono=True, bold=True)])])
    return y + 0.56

def display(s, x, y, w, h, txt, size, align=PP_ALIGN.LEFT):
    """gradient display text: white → lavender → indigo"""
    text(s, x, y, w, h, [P([R(txt, size=size, bold=True, tgrad=GRAD_TEXT)], align=align)])

def notes(s, visual, speaker):
    s.notes_slide.notes_text_frame.text = "VISUAL DIRECTION — " + visual + "\n\nSPEAKER NOTES — " + speaker

def num_badge(s, x, y, n, size=0.42):
    b = rect(s, x, y, size, size, fill=VIO_D, line=VIO, lw=1.1, rounded=True, adj=0.24)
    add_glow(b, VIO, 4, 22)
    text(s, x, y, size, size, [P([R(str(n), size=13, color=VIO, mono=True, bold=True)], align=PP_ALIGN.CENTER)],
         anchor=MSO_ANCHOR.MIDDLE)

# ============================================================ SLIDE 1 · TITLE
s = prs.slides.add_slide(BLANK)
s.shapes.add_picture("assets/hero-title.jpg", 0, 0, Inches(SW), Inches(SH))
scrim = rect(s, 0, 0, 8.4, SH, fill=BG)
gradient_fill(scrim, [(0, BG, 96), (55, BG, 80), (100, BG, 0)], 0)
hl = rect(s, 0, 0, SW, 0.03, fill=VIO)
gradient_fill(hl, [(0, VIO, 75), (45, IND, 55), (100, IND, 0)], 0)
text(s, 0.62, 0.52, 9.0, 0.3,
     [P([R("ZODIAC OS v1.0 // TEXT-TO-SQL AGENT // STATUS: ", size=10.5, color=MUTED, mono=True),
         R("ONLINE", size=10.5, color=VIO, mono=True, bold=True),
         R("  ▮", size=10.5, color=VIO, mono=True)])])
display(s, 0.55, 1.38, 8.6, 1.7, "ZODIAC AI", 92)
text(s, 0.63, 3.2, 8.0, 0.55, [P([R("Chat with Your Database", size=26, color=TEXT, font=LIGHT)])])
bar = rect(s, 0.65, 3.86, 2.6, 0.028, fill=VIO)
gradient_fill(bar, [(0, VIO, 95), (70, IND, 65), (100, IND, 0)], 0)
text(s, 0.63, 4.1, 7.5, 1.0, [
    P([R("A conversational agent that talks directly to relational databases —", size=13, color=MUTED)], sa=3, ls=1.18),
    P([R("answering with ", size=13, color=MUTED), R("exact SQL", size=13, color=VIO, bold=True),
       R(", ", size=13, color=MUTED), R("interactive Plotly charts", size=13, color=IND, bold=True),
       R(", ", size=13, color=MUTED), R("live Mermaid diagrams", size=13, color=IND, bold=True),
       R(" and", size=13, color=MUTED)], sa=3, ls=1.18),
    P([R("mathematically grounded explanations", size=13, color=VIO, bold=True),
       R(".", size=13, color=MUTED)], ls=1.18)])
tp = rect(s, 0.63, 5.36, 5.15, 0.54, fill=BG, line=VIO_L, lw=1.1, rounded=True, adj=0.5, alpha=60)
text(s, 0.63, 5.36, 5.15, 0.54,
     [P([R("> ", size=14, color=IND, mono=True, bold=True),
         R("NOT GUESSES. GROUND TRUTH.", size=13.5, color=VIO, mono=True, bold=True, spc=150)],
        align=PP_ALIGN.CENTER)], anchor=MSO_ANCHOR.MIDDLE)
cx = 0.63
for lbl in ["TEXT-TO-SQL", "AGENTIC LOOP", "MULTI-PROVIDER", "8-STEP REASONING"]:
    wch = 0.18 + 0.082*len(lbl)
    chip(s, cx, 6.18, wch, 0.36, lbl, color=MUTED, size=9)
    cx += wch + 0.17
text(s, 0.63, 6.94, 8.0, 0.3, [P([R("PITCH DECK & TECHNICAL OVERVIEW", size=8.5, color=FAINT, spc=300)])])
notes(s, "Near-black canvas with a soft violet-indigo aurora bleeding in from the right. 'ZODIAC AI' set as large gradient display type (white → lavender → indigo) with a thin gradient rule beneath. Terminal status line on top; feature pills along the bottom.",
      "Open with the one-liner: 'Every company runs on data — almost nobody can talk to it directly.' Introduce Zodiac AI as a conversational agent between humans and the relational database: ask in plain English, get the exact SQL, the chart, the diagram, and the numbers — every figure traceable to an executed query. Set the hook fast; promise the architecture tour.")

# ==================================================== SLIDE 2 · EXECUTIVE VISION
s = prs.slides.add_slide(BLANK)
chrome(s, 2, "EXECUTIVE VISION", "assets/bg-a.jpg")
kicker_title(s, "02 — EXECUTIVE VISION",
             [R("The Wall Between ", size=30, color=TEXT, bold=True),
              R("You and Your Data", size=30, color=VIO, bold=True)])
seam = rect(s, 6.62, 1.85, 0.03, 4.3, fill=VIO)
gradient_fill(seam, [(0, BG, 0), (50, VIO, 85), (100, BG, 0)], 90)
py = panel(s, 0.55, 1.78, 5.85, 4.5, "THE OLD WAY — TRADITIONAL BI", accent=MUTED)
bullets(s, 0.88, py+0.22, 5.25, 3.6, [
    ("Every question becomes a ticket", "analysts queue, answers ship days later"),
    ("Static dashboards", "answer yesterday's questions, not today's"),
    ("SQL fluency is the bottleneck", "the data exists; direct access doesn't"),
    ("Naive LLM chatbots", "hallucinate numbers with total confidence"),
], size=13, marker="✕", mcolor=RED, gap=14)
py = panel(s, 6.98, 1.78, 5.85, 4.5, "THE ZODIAC WAY", accent=VIO)
bullets(s, 7.31, py+0.22, 5.25, 3.6, [
    ("Ask in plain English", "get the exact SQL, executed live"),
    ("Visual answers", "interactive Plotly charts + live Mermaid diagrams"),
    ("Grounded numbers", "every figure computed from real result sets"),
    ("Zero-latency analytics", "answers at the speed of conversation"),
], size=13, marker="▸", mcolor=VIO, gap=14)
bottom_bar(s, "ANALYTICS AT THE SPEED OF CONVERSATION — WITHOUT THE HALLUCINATION TAX", accent=VIO, y=6.5)
notes(s, "Two quiet glass panels divided by a thin vertical violet light. Left: the old world in neutral grey with soft-red ✕ markers. Right: the Zodiac world in violet with ▸ markers. Restrained, editorial, high-contrast typography.",
      "Frame it as economics: the real cost of BI is the latency between question and answer. When latency is days, people stop asking. Generic chatbots feel instant but fabricate — speed without trust. Zodiac collapses the loop to seconds while keeping every number traceable to an executed query. Position carefully: 'We're not replacing analysts — we're unblocking everyone else.'")

# ==================================================== SLIDE 3 · TECH STACK
s = prs.slides.add_slide(BLANK)
chrome(s, 3, "TECH STACK", "assets/bg-b.jpg")
kicker_title(s, "03 — TECH STACK",
             [R("The Stack — ", size=30, color=TEXT, bold=True),
              R("Engineered for Speed, Built to Swap", size=30, color=VIO, bold=True)])
layers = [
    ("FRONTEND", VIO, "React + TanStack Query", "Declarative UI · async server-state · smart caching & live refetch", "async · cached · live"),
    ("BACKEND",  IND, "FastAPI (Python)",       "Async endpoints · Pydantic-validated payloads · native streaming", "fast · typed · streams"),
    ("AI LAYER", VIO, "Groq + Gemini",          "Dual LLM providers behind a single router interface", "fallback-ready"),
    ("DATA",     IND, "SQLite",                 "Zero-config · file-based · instantly portable — swap for Postgres/MySQL", "ships in the repo"),
]
ly = 1.78
for name, acc, tech, desc, spec in layers:
    p = rect(s, 0.55, ly, 7.5, 1.02, fill=PANEL, line=LINE, lw=1.0, rounded=True, adj=0.12)
    set_fill_alpha(p, 88); add_shadow(p, 0.18, 0.045, 42)
    tag = rect(s, 0.55, ly, 2.3, 1.02, fill=VIO_D if acc == VIO else IND_D, rounded=True, adj=0.12)
    set_fill_alpha(tag, 92)
    rect(s, 2.83, ly+0.12, 0.024, 0.78, fill=acc)
    text(s, 0.55, ly, 2.3, 1.02, [P([R(name, size=13.5, color=acc, mono=True, bold=True, spc=110)], align=PP_ALIGN.CENTER)],
         anchor=MSO_ANCHOR.MIDDLE)
    text(s, 3.1, ly+0.15, 5.6, 0.38, [P([R(tech, size=14.5, color=TEXT, bold=True)])])
    text(s, 3.1, ly+0.54, 4.9, 0.42, [P([R(desc, size=9.5, color=MUTED)])])
    text(s, 8.9, ly, 0.62, 1.02, [P([R(spec, size=7.5, color=acc, mono=True)], align=PP_ALIGN.RIGHT)], anchor=MSO_ANCHOR.MIDDLE)
    ly += 1.17
py = panel(s, 8.35, 1.78, 4.43, 4.55, "SYSTEM CROSS-SECTION", accent=IND)
picture_fit(s, "assets/stack-visual.jpg", 8.55, py+0.08, 4.03, 3.15)
text(s, 8.55, py+3.4, 4.03, 0.3,
     [P([R("SWAP ANY LAYER — THE AGENT DOESN'T CARE", size=8, color=IND, mono=True, spc=60)], align=PP_ALIGN.CENTER)])
bottom_bar(s, "DESIGN PRINCIPLE — EVERY LAYER IS MODULAR. SWAP ANY COMPONENT WITHOUT REWRITING THE AGENT", accent=IND, y=6.55)
notes(s, "Four flat glass rows for the stack layers, violet and indigo alternating, with a soft 3D render of the four-layer frosted-glass stack beside them. Thin dividers, monospace spec tags on the right edge.",
      "Two reasons for these choices: iteration velocity and architectural portability. React + TanStack handles the chat-plus-live-chart pattern. FastAPI gives async Python with strict validation. SQLite makes the demo infinitely portable. Most important: the AI layer is deliberately abstracted — the agent never talks to a vendor, it talks to the router.")

# ==================================================== SLIDE 4 · MULTI-PROVIDER ARCHITECTURE
s = prs.slides.add_slide(BLANK)
chrome(s, 4, "AI ARCHITECTURE", "assets/bg-a.jpg")
kicker_title(s, "04 — MULTI-PROVIDER AI ARCHITECTURE",
             [R("Two Brains, One Loop — ", size=30, color=TEXT, bold=True),
              R("Never Offline", size=30, color=VIO, bold=True)])
py = panel(s, 0.55, 1.78, 5.35, 4.55, "PROVIDER FALLBACK ROUTER", accent=IND)
bullets(s, 0.86, py+0.14, 4.75, 2.1, [
    ("One unified LLM interface", "provider-agnostic by design"),
    ("Primary: Groq", "near-instant Llama-class inference"),
    ("Auto-failover to Gemini", "on rate limits, timeouts, outages"),
    ("Zero single points of failure", "run with one key or both"),
], size=11, marker="▸", mcolor=IND, gap=7)
picture_fit(s, "assets/router-visual.jpg", 0.8, py+2.32, 4.85, 1.62)
px = 6.1; pw = 6.68
py2 = panel(s, px, 1.78, pw, 4.55, "THE 8-STEP AGENT LOOP", accent=VIO)
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
     [P([R("▸ steps 3 → 6 repeat until the answer is complete and grounded", size=9.5, color=IND, mono=True)])])
bottom_bar(s, "ANSWERS ARE EARNED THROUGH ITERATION — NOT GUESSED ON THE FIRST PASS", accent=VIO, y=6.55)
notes(s, "Left panel: minimal 3D render — a glowing glass cube (the router) with two light paths: a violet path that dissolves halfway (Groq, mid-failover) and an indigo path carrying particles onward (Gemini). Right panel: the 8-step loop as a clean numbered grid with outlined violet badges.",
      "Two ideas. First, resilience: LLM providers are utilities, and utilities go down — the router treats Groq and Gemini as interchangeable, so when Groq rate-limits mid-session the user never notices. Second, the loop: this is not one-shot prompt → answer. The agent reasons, acts, observes, iterates — if the first query errors or returns partial data, it self-corrects. Answers are earned through iteration.")

# ==================================================== SLIDE 5 · CORE TOOLS 1
s = prs.slides.add_slide(BLANK)
chrome(s, 5, "CORE TOOLS 1/3", "assets/bg-b.jpg")
kicker_title(s, "05 — CORE TOOLS · 1 OF 3",
             [R("get_schema", size=29, color=VIO, bold=True, mono=True),
              R("  &  ", size=29, color=MUTED, bold=True),
              R("execute_query", size=29, color=VIO, bold=True, mono=True)])
ty = terminal(s, 0.55, 1.78, 5.85, 4.15, "get_schema", accent=VIO)
text(s, 0.9, ty+0.1, 5.2, 1.85, [
    P([R("sales_db", size=12.5, color=IND, mono=True, bold=True)], sa=3),
    P([R(" ├─ ", size=12.5, color=FAINT, mono=True), R("customers", size=12.5, color=TEXT, mono=True)], sa=3),
    P([R(" ├─ ", size=12.5, color=FAINT, mono=True), R("products", size=12.5, color=TEXT, mono=True)], sa=3),
    P([R(" ├─ ", size=12.5, color=FAINT, mono=True), R("orders", size=12.5, color=TEXT, mono=True)], sa=3),
    P([R(" └─ ", size=12.5, color=FAINT, mono=True), R("order_items", size=12.5, color=TEXT, mono=True)], sa=4),
    P([R("   tables · columns · types · relations", size=8.5, color=FAINT, mono=True)])])
bullets(s, 0.9, ty+2.25, 5.2, 1.0, [
    ("Live runtime introspection", "no hardcoded, stale schema assumptions"),
], size=11.5, gap=6)
ty = terminal(s, 6.98, 1.78, 5.85, 4.15, "execute_query", accent=VIO)
text(s, 7.33, ty+0.1, 5.2, 1.6, [
    P([R("SELECT c.name, SUM(oi.qty*oi.price)", size=10.5, color=LAV, mono=True)], sa=1),
    P([R("FROM order_items oi", size=10.5, color=LAV, mono=True)], sa=1),
    P([R("JOIN orders o ON o.id = oi.order_id", size=10.5, color=LAV, mono=True)], sa=1),
    P([R("JOIN customers c ON c.id = o.customer_id", size=10.5, color=LAV, mono=True)], sa=1),
    P([R("GROUP BY c.name;", size=10.5, color=LAV, mono=True)], sa=5),
    P([R("✓ 1,284 rows returned", size=10, color=IND, mono=True, bold=True),
       R("  [READ-ONLY ✓]", size=10, color=VIO, mono=True, bold=True)])])
bullets(s, 7.33, ty+2.0, 5.2, 0.7, [
    ("Exact SQL", "generated & executed on the real database"),
], size=11.5, gap=5)
blk = rect(s, 7.33, ty+2.72, 5.15, 0.52, fill=RED_D, line=RED, lw=1.0, rounded=True, adj=0.28)
text(s, 7.5, ty+2.72, 4.9, 0.52,
     [P([R("DROP TABLE orders;", size=10, color=RED, mono=True, strike=True),
         R("  ⚠ BLOCKED — WRITE GUARD", size=9.5, color=RED, mono=True, bold=True)])],
     anchor=MSO_ANCHOR.MIDDLE)
bottom_bar(s, "RADICAL TRANSPARENCY — EVERY ANSWER SHIPS WITH THE SQL THAT PRODUCED IT. COPY IT. AUDIT IT. RERUN IT.", accent=VIO, y=6.15)
notes(s, "Two quiet dark terminals. Left (get_schema): the schema tree in soft indigo and white monospace. Right (execute_query): the SQL in lavender, a violet READ-ONLY badge, and beneath it a soft-red rejected 'DROP TABLE orders;' stamped BLOCKED — WRITE GUARD.",
      "Trust is the product, and these two tools are its foundation. The agent reads the schema fresh at runtime — it cannot reference a column that doesn't exist — and it executes real SQL, so answers come from the database, not the model's memory. The read-only guard means the worst-case interaction is an empty result set. Users always see the exact SQL behind every number.")

# ==================================================== SLIDE 6 · CORE TOOLS 2
s = prs.slides.add_slide(BLANK)
chrome(s, 6, "CORE TOOLS 2/3", "assets/bg-a.jpg")
kicker_title(s, "06 — CORE TOOLS · 2 OF 3",
             [R("generate_chart", size=29, color=IND, bold=True, mono=True),
              R("  &  ", size=29, color=MUTED, bold=True),
              R("generate_flowchart", size=29, color=IND, bold=True, mono=True)])
py = panel(s, 0.55, 1.78, 7.8, 2.24, "generate_chart — PLOTLY", accent=IND)
bullets(s, 0.88, py+0.12, 7.15, 1.6, [
    ("Specs authored from real query results", "the model writes the definition, the browser renders it"),
    ("Fully interactive", "zoom · hover · pan · filter — exploration, not screenshots"),
    ("Deterministic rendering", "bar · line · scatter · pie, matched to the question — no hallucinated visuals"),
], size=11.5, marker="▸", mcolor=IND, gap=7)
py = panel(s, 0.55, 4.24, 7.8, 2.24, "generate_flowchart — MERMAID", accent=IND)
bullets(s, 0.88, py+0.12, 7.15, 1.6, [
    ("Mermaid syntax, written live by the agent", "diagrams render inside the chat"),
    ("Schema maps & ER diagrams", "entity relationships on demand"),
    ("Logic & pipeline flows", "the shape of the system, not just the numbers"),
], size=11.5, marker="▸", mcolor=IND, gap=7)
py = panel(s, 8.6, 1.78, 4.18, 4.7, "LIVE RENDER", accent=VIO)
picture_fit(s, "assets/viz-visual-square.jpg", 8.78, py+0.1, 3.82, 3.42)
text(s, 8.78, py+3.66, 3.82, 0.3,
     [P([R("PLOTLY × MERMAID — RENDERED LIVE", size=8, color=VIO, mono=True, spc=60)], align=PP_ALIGN.CENTER)])
bottom_bar(s, "ONE PROMPT → THE NUMBERS AND THE PICTURE OF THE NUMBERS", accent=IND, y=6.62)
notes(s, "Two stacked glass panels for the Plotly and Mermaid tools on the left; on the right, a soft hologram render — a violet line chart above an indigo node graph, light particles drifting between them.",
      "Text-only answers waste an LLM's real strength: structured generation. Zodiac treats visuals as first-class outputs — the model doesn't describe a chart, it writes the chart definition, and Plotly and Mermaid render it deterministically. Because the spec is generated from executed query results, the visual carries the same grounding as the SQL behind it.")

# ==================================================== SLIDE 7 · CORE TOOLS 3
s = prs.slides.add_slide(BLANK)
chrome(s, 7, "CORE TOOLS 3/3", "assets/bg-b.jpg")
kicker_title(s, "07 — CORE TOOLS · 3 OF 3",
             [R("explain_data", size=29, color=VIO, bold=True, mono=True),
              R(" — Grounded, Not Guessed", size=29, color=IND, bold=True)])
pipeline = [("1 · QUERY", VIO, "against the live database"),
            ("2 · COMPUTE", IND, "totals · means · min/max · shares"),
            ("3 · NARRATE", VIO, "prose around proven numbers")]
pxs = 0.85
for lbl, acc, d in pipeline:
    pc = rect(s, pxs, 1.8, 3.35, 1.0, fill=PANEL, line=LINE, lw=1.0, rounded=True, adj=0.14)
    set_fill_alpha(pc, 90); add_shadow(pc, 0.18, 0.045, 42)
    rect(s, pxs+0.14, 1.8, 3.07, 0.04, fill=acc)
    text(s, pxs, 1.97, 3.35, 0.4, [P([R(lbl, size=14.5, color=acc, mono=True, bold=True)], align=PP_ALIGN.CENTER)])
    text(s, pxs+0.1, 2.38, 3.15, 0.32, [P([R(d, size=8.5, color=MUTED, mono=True)], align=PP_ALIGN.CENTER)])
    if lbl != "3 · NARRATE":
        text(s, pxs+3.38, 1.8, 0.72, 1.0, [P([R("→", size=16, color=IND, mono=True, bold=True)], align=PP_ALIGN.CENTER)], anchor=MSO_ANCHOR.MIDDLE)
    pxs += 4.1
bullets(s, 0.88, 3.25, 7.0, 2.9, [
    ("Numeric summaries from executed result sets", "totals, averages, min/max, distributions, shares — computed, not recalled"),
    ("Every figure traces to a query result", "text is generated from computed numbers, never from model memory"),
    ("The math ships with the story", "growth rates, proportions, outliers — derived first, then narrated"),
    ("Anti-hallucination by architecture", "untraceable numbers don't make the answer"),
], size=12.5, gap=13)
py = panel(s, 8.35, 3.05, 4.43, 3.15, "GROUNDING LENS", accent=VIO)
picture_fit(s, "assets/lens-visual.jpg", 8.53, py+0.08, 4.07, 2.15)
text(s, 8.53, py+2.32, 4.07, 0.3,
     [P([R("EVERY NUMBER TRACED TO A ROW", size=8, color=VIO, mono=True, spc=60)], align=PP_ALIGN.CENTER)])
q = rect(s, 0.55, 6.5, SW-1.1, 0.5, fill=VIO_D, line=VIO_L, lw=1.0, rounded=True, adj=0.5)
set_fill_alpha(q, 85)
text(s, 0.9, 6.5, SW-1.8, 0.5,
     [P([R("> ", size=12, color=VIO, mono=True, bold=True),
         R("We don't ask the model to remember your revenue. We ask your database.", size=12, color=TEXT, mono=True)])],
     anchor=MSO_ANCHOR.MIDDLE)
notes(s, "A restrained pipeline motif — QUERY → COMPUTE → NARRATE — with a soft render of a violet-rimmed lens over a faint indigo data grid on the right, points of light lifting from the cells. The anti-hallucination quote sits in a quiet violet-tinted bar.",
      "This slide separates Zodiac from a wrapped chat API. Generic LLMs interpolate plausible-sounding statistics; Zodiac computes the statistics first and writes the narrative around them. The pipeline is strict: query → compute → narrate. If a number can't be traced to a result set, it doesn't make it into the answer. Say it plainly: 'We don't ask the model to remember your revenue. We ask your database.'")

# ==================================================== SLIDE 8 · DATABASE ARCHITECTURE
s = prs.slides.add_slide(BLANK)
chrome(s, 8, "DATA MODEL", "assets/bg-a.jpg")
kicker_title(s, "08 — DATABASE ARCHITECTURE",
             [R("Under the Hood — ", size=30, color=TEXT, bold=True),
              R("The Sample Sales Schema", size=30, color=VIO, bold=True)])
tables = [
    ("CUSTOMERS", VIO, "who is buying", "customer_id · name\ncity · signup_date"),
    ("ORDERS",    IND, "the transaction header", "order_id · customer_id\ndate · status"),
    ("ORDER_ITEMS", VIO, "line-level detail", "order_id · product_id\nqty · price"),
    ("PRODUCTS",  IND, "what is being sold", "product_id · name\ncategory · price"),
]
rels = [("1 — ∞", True), ("1 — ∞", True), ("∞ — 1", False)]
tx = 0.55; tw = 2.72; tg = (SW-1.1-4*tw)/3
for i, (name, acc, d, cols) in enumerate(tables):
    p = rect(s, tx, 1.88, tw, 1.66, fill=PANEL, line=LINE, lw=1.0, rounded=True, adj=0.09)
    set_fill_alpha(p, 90); add_shadow(p, 0.18, 0.045, 42)
    rect(s, tx+0.14, 1.88, tw-0.28, 0.04, fill=acc)
    text(s, tx, 2.04, tw, 0.32, [P([R(name, size=12.5, color=acc, mono=True, bold=True, spc=70)], align=PP_ALIGN.CENTER)])
    text(s, tx+0.1, 2.4, tw-0.2, 0.28, [P([R(d, size=9.5, color=TEXT)], align=PP_ALIGN.CENTER)])
    text(s, tx+0.08, 2.72, tw-0.16, 0.72,
         [P([R(cols.split("\n")[0], size=8, color=MUTED, mono=True)], align=PP_ALIGN.CENTER, sa=2),
          P([R(cols.split("\n")[1], size=8, color=MUTED, mono=True)], align=PP_ALIGN.CENTER)])
    if i < 3:
        lbl, fwd = rels[i]
        text(s, tx+tw-0.05, 2.24, tg+0.1, 0.3,
             [P([R("━━━→" if fwd else "←━━━", size=10, color=IND, mono=True)], align=PP_ALIGN.CENTER)])
        text(s, tx+tw-0.05, 2.5, tg+0.1, 0.3, [P([R(lbl, size=11, color=IND, mono=True, bold=True)], align=PP_ALIGN.CENTER)])
    tx += tw + tg
bullets(s, 0.88, 3.95, 11.6, 1.85, [
    ("Normalized e-commerce model — four tables, one star", "join-rich by design: every interesting question crosses at least one join"),
    ("The schema that stresses text-to-SQL edge cases", "if the agent is correct here, it generalizes"),
    ("Nothing is hardcoded", "the agent maps all relationships at runtime via get_schema"),
], size=12.5, gap=12)
bottom_bar(s, "SCHEMA DISCOVERED LIVE — NEVER BAKED INTO PROMPTS", accent=VIO, y=6.15)
notes(s, "Four quiet glass cards for the tables with thin violet/indigo top rules, connected by delicate cardinality labels (1 — ∞). Editorial and schematic rather than decorative.",
      "The demo dataset is deliberately canonical e-commerce: instantly understandable, and join-heavy — exactly where naive text-to-SQL breaks. Every interesting question (top customers, revenue per product, average order value) crosses at least one join. If the agent is correct on this schema, it generalizes. And the schema is never baked into prompts — it's discovered live.")

# ==================================================== SLIDE 9 · SAMPLE PROMPTS
s = prs.slides.add_slide(BLANK)
chrome(s, 9, "SAMPLE PROMPTS", "assets/bg-b.jpg")
kicker_title(s, "09 — VALID QUERY CATEGORIES & SAMPLE PROMPTS",
             [R("Ask It Anything ", size=30, color=TEXT, bold=True),
              R("(Structured)", size=30, color=VIO, bold=True)])
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
    set_fill_alpha(p, 90); add_shadow(p, 0.16, 0.04, 42)
    acc = VIO if i % 2 == 0 else IND
    rect(s, cx, cy+0.14, 0.05, 1.24, fill=acc)
    text(s, cx+0.26, cy+0.17, cw-0.45, 0.3,
         [P([R("// ", size=10, color=FAINT, mono=True), R(cat, size=10, color=acc, mono=True, bold=True, spc=130)])])
    l1, l2 = q.split("\n")
    text(s, cx+0.26, cy+0.52, cw-0.45, 0.85,
         [P([R(l1, size=12, color=TEXT, italic=True)], ls=1.15, sa=1),
          P([R(l2, size=12, color=TEXT, italic=True)], ls=1.15)])
bottom_bar(s, "EVERY ANSWER = EXACT SQL + INTERACTIVE CHART + GROUNDED EXPLANATION", accent=VIO, y=5.6)
ap = rect(s, 0.55, 6.28, 4.35, 0.5, fill=PANEL_D, line=LINE, lw=1.0, rounded=True, adj=0.5)
set_fill_alpha(ap, 90)
text(s, 0.85, 6.28, 4.0, 0.5,
     [P([R("> ", size=13, color=VIO, mono=True, bold=True),
         R("ask anything", size=13, color=TEXT, mono=True),
         R("_", size=13, color=VIO, mono=True, bold=True)])],
     anchor=MSO_ANCHOR.MIDDLE)
notes(s, "Six quiet glass cards, each with a thin accent edge, a small monospace category label, and an italic sample prompt in white. Beneath, the answer-bundle formula and a terminal input field: '> ask anything_'",
      "Walk one or two cards aloud, then let the audience read the rest. Point out the data-quality category last — it proves the agent can reason about the database, not just retrieve from it. Close on the bundle: a single prompt yields the SQL, a rendered chart, and a grounded explanation. That triple — query, visual, narrative — is the product experience.")

# ==================================================== SLIDE 10 · GETTING STARTED
s = prs.slides.add_slide(BLANK)
chrome(s, 10, "GETTING STARTED", "assets/bg-a.jpg")
kicker_title(s, "10 — GETTING STARTED",
             [R("Zero to Query in ", size=30, color=TEXT, bold=True),
              R("Five Minutes", size=30, color=VIO, bold=True)])
ty = terminal(s, 0.55, 1.78, 7.55, 4.4, "boot-sequence — zodiac-ai", accent=VIO)
text(s, 0.9, ty+0.12, 6.95, 3.6, [
    P([R("$ ", size=11.5, color=VIO, mono=True, bold=True),
       R("git clone https://github.com/<org>/zodiac-ai.git", size=11.5, color=TEXT, mono=True)], sa=4),
    P([R("$ ", size=11.5, color=VIO, mono=True, bold=True),
       R("cd zodiac-ai", size=11.5, color=TEXT, mono=True)], sa=9),
    P([R("# .env", size=10.5, color=FAINT, mono=True)], sa=4),
    P([R("GROQ_API_KEY", size=11, color=VIO, mono=True, bold=True),
       R("=••••••••••", size=11, color=TEXT, mono=True),
       R("   # primary provider", size=9.5, color=FAINT, mono=True)], sa=4),
    P([R("GEMINI_API_KEY", size=11, color=IND, mono=True, bold=True),
       R("=••••••••••", size=11, color=TEXT, mono=True),
       R("   # optional fallback", size=9.5, color=FAINT, mono=True)], sa=9),
    P([R("$ ", size=11.5, color=VIO, mono=True, bold=True),
       R("pip install -r requirements.txt", size=11.5, color=TEXT, mono=True),
       R("  # backend ✓", size=9.5, color=FAINT, mono=True)], sa=4),
    P([R("$ ", size=11.5, color=VIO, mono=True, bold=True),
       R("npm install && npm run dev", size=11.5, color=TEXT, mono=True),
       R("         # frontend ✓", size=9.5, color=FAINT, mono=True)], sa=8),
    P([R("[", size=11, color=FAINT, mono=True), R("██████████░░", size=11, color=VIO, mono=True),
       R("]", size=11, color=FAINT, mono=True), R(" LOADING...", size=10, color=MUTED, mono=True)], sa=8),
    P([R("▲ ", size=12, color=IND, mono=True, bold=True),
       R("open the local URL  →  start asking", size=12, color=IND, mono=True, bold=True)])])
py = panel(s, 8.4, 1.78, 4.38, 4.4, "RUN NOTES", accent=IND)
bullets(s, 8.7, py+0.15, 3.85, 3.5, [
    ("Prerequisites", "Python 3.10+ · Node 18+ · Git · ≥ 1 API key"),
    ("No DB server", "SQLite ships inside the repo"),
    ("One key runs it", "a second key activates the failover router"),
    ("< 5 minutes", "from git clone to first grounded answer"),
], size=11, marker="▸", mcolor=IND, gap=9)
bottom_bar(s, "CLONE IT DURING Q&A — THE DEMO IS THE PITCH", accent=VIO, y=6.42)
notes(s, "A clean dark terminal with the boot sequence — clone, .env with masked keys, installs, a slim progress bar, and the final 'open the local URL' line. Run Notes panel beside it.",
      "The pitch of this slide is friction — or its absence. SQLite means no database server to stand up: the schema ships inside the repo. One API key is enough to run; the second is optional insurance that activates failover. From git clone to first grounded answer: under five minutes. Invite the audience to do it live during Q&A.")

# ==================================================== SLIDE 11 · CONCLUSION & Q&A
s = prs.slides.add_slide(BLANK)
s.shapes.add_picture("assets/hero-closing.jpg", 0, 0, Inches(SW), Inches(SH))
scrim = rect(s, 0, 0, 9.0, SH, fill=BG)
gradient_fill(scrim, [(0, BG, 95), (60, BG, 78), (100, BG, 0)], 0)
hl = rect(s, 0, 0, SW, 0.03, fill=VIO)
gradient_fill(hl, [(0, VIO, 75), (45, IND, 55), (100, IND, 0)], 0)
text(s, 0.62, 0.5, 9.0, 0.3,
     [P([R("$ status: ", size=10.5, color=MUTED, mono=True),
         R("ALL SYSTEMS OPERATIONAL", size=10.5, color=VIO, mono=True, bold=True)])])
text(s, 0.62, 0.98, 9.6, 0.6,
     [P([R("The Database Is Now ", size=31, color=TEXT, bold=True),
         R("a Conversation", size=31, color=VIO, bold=True)])])
bar = rect(s, 0.64, 1.7, 2.3, 0.028, fill=VIO)
gradient_fill(bar, [(0, VIO, 95), (70, IND, 65), (100, IND, 0)], 0)
bullets(s, 0.62, 2.02, 8.3, 3.2, [
    ("What we built", "exact SQL · interactive Plotly charts · live Mermaid diagrams · grounded explanations"),
    ("Why it matters", "analytics at conversational speed — without the hallucination tax"),
    ("How it holds up", "read-only guardrails · runtime schema awareness · multi-provider failover · 8-step reasoning"),
    ("What's next", "Postgres / MySQL connectors · richer chart grammar · conversational follow-ups · shared query memory"),
], size=12.5, gap=12)
cta = rect(s, 0.62, 5.44, 4.5, 0.54, fill=BG, line=VIO_L, lw=1.1, rounded=True, adj=0.5, alpha=60)
text(s, 0.62, 5.44, 4.5, 0.54,
     [P([R("▸ ", size=13, color=VIO, mono=True, bold=True),
         R("CLONE IT. QUERY IT. BREAK IT.", size=12.5, color=VIO, mono=True, bold=True, spc=100)],
        align=PP_ALIGN.CENTER)], anchor=MSO_ANCHOR.MIDDLE)
display(s, 8.95, 4.35, 3.9, 1.7, "Q&A", 92)
text(s, 9.0, 6.14, 3.8, 0.35,
     [P([R("> ask anything", size=13, color=TEXT, mono=True),
         R("_", size=13, color=VIO, mono=True, bold=True)], align=PP_ALIGN.CENTER)])
text(s, 0.62, 6.96, 8.0, 0.3, [P([R("ZODIAC AI // THANK YOU", size=8.5, color=FAINT, spc=300)])])
notes(s, "Closing frame mirrors the title: near-black canvas, the minimal violet constellation with its indigo reflection on the right, clean dark space on the left for the recap. 'Q&A' set as large gradient display type (white → lavender → indigo).",
      "Close the loop with the opening line: databases used to be things you queried — now they're things you talk to. Recap the four pillars in one breath: exact SQL, live visuals, grounded numbers, resilient architecture. Open the floor — and if the room is quiet, seed the first question: 'Ask me what happens when Groq goes down mid-query.'")

# ---------------------------------------------------------------- save
prs.core_properties.title = "Zodiac AI — Pitch Deck & Technical Overview"
prs.core_properties.author = "Zodiac AI"
prs.core_properties.subject = "Chat with your Database — text-to-SQL agent"
OUT = "Zodiac_AI_Pitch_Deck.pptx"
prs.save(OUT)
print(f"saved {OUT} — {len(prs.slides._sldIdLst)} slides")
