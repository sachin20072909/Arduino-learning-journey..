#!/usr/bin/env python3
"""Build Zodiac AI pitch deck (16:9, cyberpunk 'Cyber-Grid' theme) as .pptx"""
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn
from lxml import etree

# ---------------------------------------------------------------- palette
BG      = "0A0E14"   # near-black canvas
PANEL   = "0E1622"   # panel fill
PANEL_H = "131E2E"   # panel header bar
LINE    = "223046"   # panel border
GREEN   = "00FF41"   # neon green  (ground truth / executed / safe)
BLUE    = "00D4FF"   # electric blue (in-flight / routed / rendered)
RED     = "FF3355"   # blocked / hallucinated
TEXT    = "E6F1FF"   # primary text
MUTED   = "8B9BB4"   # secondary text
FAINT   = "56637A"   # comments / tertiary
GRID    = "121A28"   # background grid lines
GREEN_D = "0D2A18"   # dark green tint fill
BLUE_D  = "0A2431"   # dark blue tint fill
RED_D   = "2A0F1A"   # dark red tint fill

MONO = "Consolas"
SANS = "Arial"

SW, SH = 13.333, 7.5

def C(h): return RGBColor(int(h[0:2],16), int(h[2:4],16), int(h[4:6],16))

prs = Presentation()
prs.slide_width  = Inches(SW)
prs.slide_height = Inches(SH)
BLANK = prs.slide_layouts[6]

# ---------------------------------------------------------------- helpers
def _set_alpha(clr_elem, pct):
    a = etree.SubElement(clr_elem, qn('a:alpha')); a.set('val', str(int(pct*1000)))

def set_fill_alpha(shape, pct):
    sf = shape._element.spPr.find(qn('a:solidFill'))
    if sf is not None: _set_alpha(sf.find(qn('a:srgbClr')), pct)

def slide_new():
    return prs.slides.add_slide(BLANK)

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

def text(s, x, y, w, h, paras, anchor=MSO_ANCHOR.TOP, wrap=True):
    """paras: list of dicts {runs:[(txt,style)], align, sb, sa, ls}
       style: {size, color, bold, italic, mono, strike, spc}"""
    tb = s.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = tb.text_frame
    tf.word_wrap = wrap
    tf.vertical_anchor = anchor
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    for i, p in enumerate(paras):
        par = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        par.alignment = p.get('align', PP_ALIGN.LEFT)
        if p.get('sb') is not None: par.space_before = Pt(p['sb'])
        if p.get('sa') is not None: par.space_after  = Pt(p['sa'])
        par.line_spacing = p.get('ls', 1.0)
        if p.get('hang'):
            pPr = par._p.get_or_add_pPr()
            pPr.set('marL', str(int(Inches(p['hang'])))); pPr.set('indent', str(-int(Inches(p['hang']))))
        for txt, st in p['runs']:
            r = par.add_run(); r.text = txt
            f = r.font
            f.size = Pt(st.get('size', 12))
            f.color.rgb = C(st.get('color', TEXT))
            f.bold = st.get('bold', False)
            f.italic = st.get('italic', False)
            f.name = MONO if st.get('mono') else SANS
            if st.get('strike'):
                r._r.get_or_add_rPr().set('strike', 'sngStrike')
            if st.get('spc'):
                r._r.get_or_add_rPr().set('spc', str(st['spc']))
    return tb

def R(t, **kw):  # run shortcut
    st = {'size': kw.get('size',12), 'color': kw.get('color',TEXT), 'bold': kw.get('bold',False),
          'italic': kw.get('italic',False), 'mono': kw.get('mono',False), 'strike': kw.get('strike',False),
          'spc': kw.get('spc')}
    return (t, st)

def P(runs, **kw):  # paragraph shortcut
    return {'runs': runs, 'align': kw.get('align'), 'sb': kw.get('sb'), 'sa': kw.get('sa'),
            'ls': kw.get('ls',1.0), 'hang': kw.get('hang')}

def bg_plain(s):
    rect(s, 0, 0, SW, SH, fill=BG)
    # subtle cyber grid
    gx = SW/12
    for i in range(1, 12):
        ln = rect(s, i*gx, 0, 0.012, SH, fill=GRID)
    for j in range(1, 7):
        rect(s, 0, j*(SH/7), SW, 0.012, fill=GRID)

def chrome(s, num, tag):
    """top accent bar, HUD corners, footer"""
    rect(s, 0, 0, SW*0.42, 0.045, fill=GREEN)
    rect(s, SW*0.42, 0, SW*0.58, 0.045, fill=BLUE)
    # HUD corner brackets
    rect(s, SW-0.62, 0.28, 0.34, 0.02, fill=BLUE); rect(s, SW-0.30, 0.28, 0.02, 0.34, fill=BLUE)
    rect(s, 0.28, SH-0.62, 0.34, 0.02, fill=BLUE); rect(s, 0.28, SH-0.60, 0.02, 0.34, fill=BLUE)
    # footer
    rect(s, 0.55, 7.08, SW-1.1, 0.014, fill=LINE)
    text(s, 0.55, 7.14, 6.5, 0.25, [P([R("ZODIAC AI ", size=8, color=GREEN, mono=True, bold=True),
                                       R("// CHAT WITH YOUR DATABASE // ", size=8, color=FAINT, mono=True),
                                       R(tag, size=8, color=BLUE, mono=True)])])
    text(s, SW-1.55, 7.14, 1.0, 0.25, [P([R(f"{num:02d}", size=8, color=GREEN, mono=True, bold=True),
                                          R(" / 11", size=8, color=FAINT, mono=True)], align=PP_ALIGN.RIGHT)])

def kicker_title(s, kicker, title_runs):
    text(s, 0.55, 0.42, 10.5, 0.3, [P([R("> ", size=12, color=GREEN, mono=True, bold=True),
                                       R(kicker, size=11, color=GREEN, mono=True, spc=140)])])
    text(s, 0.55, 0.72, 12.2, 0.62, [P(title_runs)])
    rect(s, 0.57, 1.42, 1.7, 0.035, fill=GREEN)
    rect(s, 2.27, 1.42, 0.55, 0.035, fill=BLUE)

def panel(s, x, y, w, h, title, accent=GREEN):
    rect(s, x, y, w, h, fill=PANEL, line=LINE, lw=1.0, rounded=True, adj=0.035)
    rect(s, x, y, w, 0.46, fill=PANEL_H, rounded=True, adj=0.16)
    rect(s, x, y, 0.07, 0.46, fill=accent)
    text(s, x+0.22, y+0.10, w-0.4, 0.3, [P([R(title, size=11.5, color=accent, mono=True, bold=True, spc=100)])])
    return y + 0.46

def bullets(s, x, y, w, h, items, size=12.5, marker="▸", mcolor=GREEN, tcolor=TEXT, gap=8):
    paras = []
    for it in items:
        if isinstance(it, tuple):
            lead, rest = it
            runs = [R(marker+"  ", size=size, color=mcolor, bold=True, mono=True),
                    R(lead, size=size, color=tcolor, bold=True)]
            if rest: runs.append(R(" — "+rest, size=size, color=tcolor))
        else:
            runs = [R(marker+"  ", size=size, color=mcolor, bold=True, mono=True),
                    R(it, size=size, color=tcolor)]
        paras.append(P(runs, sa=gap, ls=1.08, hang=0.24))
    text(s, x, y, w, h, paras)

def term_header(s, x, y, w, label, accent=GREEN):
    rect(s, x, y, w, 0.42, fill=PANEL_H, rounded=True, adj=0.18)
    for i, dc in enumerate(["FF5F56", "FFBD2E", "27C93F"]):
        d = s.shapes.add_shape(MSO_SHAPE.OVAL, Inches(x+0.16+i*0.17), Inches(y+0.15), Inches(0.11), Inches(0.11))
        d.shadow.inherit = False; d.fill.solid(); d.fill.fore_color.rgb = C(dc); d.line.fill.background()
    text(s, x+0.78, y+0.09, w-1.0, 0.26, [P([R("zodiac:~ ", size=10, color=FAINT, mono=True),
                                              R(label, size=10, color=accent, mono=True, bold=True)])])

def chip(s, x, y, w, h, label, color=BLUE, fill=None, size=10, bold=True):
    rect(s, x, y, w, h, fill=fill if fill else BG, line=color, lw=1.2, rounded=True, adj=0.5)
    text(s, x, y, w, h, [P([R(label, size=size, color=color, mono=True, bold=bold)], align=PP_ALIGN.CENTER)],
         anchor=MSO_ANCHOR.MIDDLE)

def bottom_bar(s, label, accent=GREEN, y=6.5):
    rect(s, 0.55, y, SW-1.1, 0.46, fill=PANEL, line=LINE, lw=1.0, rounded=True, adj=0.5)
    rect(s, 0.72, y+0.10, 0.055, 0.26, fill=accent)
    text(s, 0.95, y, SW-2.0, 0.46, [P([R(label, size=10.5, color=accent, mono=True, bold=True, spc=60)])],
         anchor=MSO_ANCHOR.MIDDLE)

def notes(s, visual, speaker):
    s.notes_slide.notes_text_frame.text = (
        "VISUAL DIRECTION — " + visual + "\n\nSPEAKER NOTES — " + speaker)

def chromatic(s, x, y, w, h, txt, size, align=PP_ALIGN.LEFT):
    """neon chromatic-aberration display text: blue ghost behind, green front"""
    text(s, x+0.045, y+0.03, w, h, [P([R(txt, size=size, color=BLUE, bold=True)], align=align)])
    text(s, x, y, w, h, [P([R(txt, size=size, color=GREEN, bold=True)], align=align)])

# ================================================================ SLIDE 1
s = slide_new()
s.shapes.add_picture("assets/title-bg.png", 0, 0, Inches(SW), Inches(SH))
ov = rect(s, 0, 0, 7.9, SH, fill=BG); set_fill_alpha(ov, 55)
rect(s, 0, 0, SW*0.42, 0.05, fill=GREEN); rect(s, SW*0.42, 0, SW*0.58, 0.05, fill=BLUE)
text(s, 0.62, 0.55, 8.0, 0.3, [P([R("v1.0 // TEXT-TO-SQL AGENT // STATUS: ", size=11, color=MUTED, mono=True),
                                  R("ONLINE", size=11, color=GREEN, mono=True, bold=True),
                                  R(" ▮", size=11, color=GREEN, mono=True)])])
chromatic(s, 0.55, 1.55, 8.2, 1.5, "ZODIAC AI", 84)
text(s, 0.62, 3.15, 7.6, 0.55, [P([R("Chat with Your Database", size=27, color=TEXT, bold=False)])])
rect(s, 0.64, 3.78, 2.4, 0.035, fill=GREEN); rect(s, 3.04, 3.78, 0.8, 0.035, fill=BLUE)
text(s, 0.62, 4.0, 7.4, 0.9, [
    P([R("A conversational agent that talks directly to relational databases —", size=13.5, color=MUTED)], sa=3, ls=1.15),
    P([R("answering with ", size=13.5, color=MUTED),
       R("exact SQL", size=13.5, color=GREEN, bold=True), R(", ", size=13.5, color=MUTED),
       R("interactive Plotly charts", size=13.5, color=BLUE, bold=True), R(", ", size=13.5, color=MUTED),
       R("live Mermaid diagrams", size=13.5, color=BLUE, bold=True), R(" and", size=13.5, color=MUTED)], sa=3, ls=1.15),
    P([R("mathematically grounded explanations", size=13.5, color=GREEN, bold=True),
       R(".", size=13.5, color=MUTED)], ls=1.15)])
text(s, 0.62, 5.15, 7.4, 0.4, [P([R("> ", size=15, color=BLUE, mono=True, bold=True),
                                  R("NOT GUESSES. GROUND TRUTH.", size=15, color=GREEN, mono=True, bold=True, spc=200)])])
cx = 0.62
for lbl in ["TEXT-TO-SQL", "AGENTIC LOOP", "MULTI-PROVIDER", "8-STEP REASONING"]:
    wch = 0.16 + 0.085*len(lbl)
    chip(s, cx, 5.75, wch, 0.34, lbl, color=BLUE, size=9.5)
    cx += wch + 0.18
text(s, 0.62, 6.85, 8.0, 0.3, [P([R("PITCH DECK & TECHNICAL OVERVIEW", size=9, color=FAINT, mono=True, spc=250)])])
notes(s, "Full-bleed cyber-grid scene. Holographic wireframe database cylinder (right) streaming neon-green particles into a floating chat terminal. Electric-blue perspective grid floor, faint 12-star Zodiac constellation overhead. 'ZODIAC' set with chromatic-aberration glitch; terminal status bar at top.",
      "Open with the one-liner: 'Every company runs on data — almost nobody can talk to it directly.' Introduce Zodiac AI as a conversational agent between humans and the relational database: ask in plain English, get the exact SQL, the chart, the diagram, and the numbers — every figure traceable to an executed query. Set the hook fast; promise the architecture tour.")

# ================================================================ SLIDE 2
s = slide_new(); bg_plain(s)
kicker_title(s, "02 · EXECUTIVE VISION", [R("The Wall Between ", size=30, color=TEXT, bold=True),
                                          R("You and Your Data", size=30, color=BLUE, bold=True)])
py = panel(s, 0.55, 1.72, 5.95, 4.55, "THE OLD WAY — TRADITIONAL BI", accent=RED)
bullets(s, 0.85, py+0.28, 5.35, 3.8, [
    ("Every question becomes a ticket", "analysts queue, answers ship days later"),
    ("Static dashboards", "answer yesterday's questions, not today's"),
    ("SQL fluency is the bottleneck", "the data exists; direct access doesn't"),
    ("Naive LLM chatbots", "hallucinate numbers with total confidence"),
], size=13, marker="✕", mcolor=RED, gap=13)
py = panel(s, 6.83, 1.72, 5.95, 4.55, "THE ZODIAC WAY", accent=GREEN)
bullets(s, 7.13, py+0.28, 5.35, 3.8, [
    ("Ask in plain English", "get the exact SQL, executed live"),
    ("Visual answers", "interactive Plotly charts + live Mermaid diagrams"),
    ("Grounded numbers", "every figure computed from real result sets"),
    ("Zero-latency analytics", "answers at the speed of conversation"),
], size=13, marker="▸", mcolor=GREEN, gap=13)
chrome(s, 2, "EXECUTIVE VISION")
notes(s, "Split-screen divided by a pulsing electric-blue lightning seam. Left: desaturated grey-blue mosaic of stale BI dashboards, CSV exports, ticket queues, slightly glitching. Right: vivid neon chat exchange — 'What drove revenue last quarter?' answered with a glowing SQL snippet, mini chart and bullet explanation.",
      "Frame it as economics: the real cost of BI is the latency between question and answer. When latency is days, people stop asking. Generic chatbots feel instant but fabricate — speed without trust. Zodiac collapses the loop to seconds while keeping every number traceable to an executed query. Position carefully: 'We're not replacing analysts — we're unblocking everyone else.'")

# ================================================================ SLIDE 3
s = slide_new(); bg_plain(s)
kicker_title(s, "03 · TECH STACK", [R("The Stack — ", size=30, color=TEXT, bold=True),
                                    R("Engineered for Speed, Built to Swap", size=30, color=BLUE, bold=True)])
layers = [
    ("FRONTEND", GREEN, "React + TanStack Query", "Declarative UI · async server-state · smart caching & live refetching", "async · cached · live"),
    ("BACKEND",  BLUE,  "FastAPI (Python)",        "Async endpoints · Pydantic-validated payloads · native streaming", "fast · typed · streams"),
    ("AI LAYER", GREEN, "Groq + Gemini",           "Dual LLM providers behind a single router interface", "fallback-ready"),
    ("DATA",     BLUE,  "SQLite",                  "Zero-config · file-based · instantly portable — swappable for Postgres / MySQL", "ships in the repo"),
]
ly = 1.75
for name, acc, tech, desc, spec in layers:
    rect(s, 0.55, ly, SW-1.1, 1.02, fill=PANEL, line=LINE, lw=1.0, rounded=True, adj=0.12)
    rect(s, 0.55, ly, 2.35, 1.02, fill=GREEN_D if acc==GREEN else BLUE_D, rounded=True, adj=0.12)
    rect(s, 2.86, ly, 0.03, 1.02, fill=acc)
    text(s, 0.55, ly, 2.35, 1.02, [P([R(name, size=15, color=acc, mono=True, bold=True, spc=120)], align=PP_ALIGN.CENTER)],
         anchor=MSO_ANCHOR.MIDDLE)
    text(s, 3.15, ly+0.14, 7.3, 0.4, [P([R(tech, size=15, color=TEXT, bold=True)])])
    text(s, 3.15, ly+0.52, 8.2, 0.4, [P([R(desc, size=10.5, color=MUTED)])])
    text(s, 11.0, ly, 1.9, 1.02, [P([R(spec, size=8.5, color=acc, mono=True)], align=PP_ALIGN.RIGHT)], anchor=MSO_ANCHOR.MIDDLE)
    ly += 1.17
bottom_bar(s, "DESIGN PRINCIPLE — every layer is modular: swap any component without rewriting the agent", accent=BLUE, y=6.5)
chrome(s, 3, "TECH STACK")
notes(s, "Vertical layered-architecture diagram rendered as a neon circuit-board cross-section (chip die-shot). Four glowing strata: FRONTEND, API, AI (two chips GROQ/GEMINI through a router switch), DATA (green SQLite cylinder). Data beams travel through vias between layers; monospace spec labels on the right edge.",
      "Two reasons for these choices: iteration velocity and architectural portability. React + TanStack handles the chat-plus-live-chart pattern. FastAPI gives async Python with strict validation. SQLite makes the demo infinitely portable. Most important: the AI layer is deliberately abstracted — the agent never talks to a vendor, it talks to the router.")

# ================================================================ SLIDE 4
s = slide_new(); bg_plain(s)
kicker_title(s, "04 · MULTI-PROVIDER AI ARCHITECTURE", [R("Two Brains, One Loop — ", size=30, color=TEXT, bold=True),
                                                        R("Never Offline", size=30, color=BLUE, bold=True)])
py = panel(s, 0.55, 1.72, 5.25, 4.62, "PROVIDER FALLBACK ROUTER", accent=BLUE)
bullets(s, 0.85, py+0.22, 4.65, 2.3, [
    ("One unified LLM interface", "provider-agnostic by design"),
    ("Primary: Groq", "near-instant Llama-class inference"),
    ("Auto-failover to Gemini", "on rate limits, timeouts, outages"),
    ("Zero single points of AI failure", "run with one key or both"),
], size=11.5, marker="▸", mcolor=BLUE, gap=9)
rect(s, 0.85, py+2.75, 4.65, 1.35, fill=BG, line=LINE, lw=1.0, rounded=True, adj=0.1)
chip(s, 1.15, py+2.95, 4.05, 0.44, "GROQ  //  PRIMARY", color=GREEN, fill=GREEN_D, size=11)
text(s, 1.15, py+3.44, 4.05, 0.26, [P([R("▼  auto-failover · seamless reroute", size=9, color=FAINT, mono=True)], align=PP_ALIGN.CENTER)])
chip(s, 1.15, py+3.72, 4.05, 0.44, "GEMINI  //  FALLBACK", color=BLUE, fill=BLUE_D, size=11)
px = 6.05; pw = 6.73
py2 = panel(s, px, 1.72, pw, 4.62, "THE 8-STEP AGENT LOOP", accent=GREEN)
steps = [
    ("1", "INGEST",     "parse prompt + context"),
    ("2", "INTROSPECT", "pull live schema"),
    ("3", "PLAN",       "decompose, choose tools"),
    ("4", "ACT",        "execute next tool call"),
    ("5", "OBSERVE",    "inspect results / errors"),
    ("6", "VALIDATE",   "complete & grounded?"),
    ("7", "ITERATE",    "loop 3 → 6 until yes"),
    ("8", "SYNTHESIZE", "SQL + chart + explanation"),
]
for i, (n, lbl, d) in enumerate(steps):
    col, row = divmod(i, 4)
    sx = px + 0.28 + col*3.22; sy = py2 + 0.22 + row*0.92
    rect(s, sx, sy, 0.4, 0.4, fill=GREEN, rounded=True, adj=0.22)
    text(s, sx, sy, 0.4, 0.4, [P([R(n, size=14, color=BG, mono=True, bold=True)], align=PP_ALIGN.CENTER)], anchor=MSO_ANCHOR.MIDDLE)
    text(s, sx+0.55, sy-0.03, 2.6, 0.3, [P([R(lbl, size=12, color=TEXT, bold=True, mono=True, spc=60)])])
    text(s, sx+0.55, sy+0.25, 2.6, 0.26, [P([R(d, size=8.5, color=MUTED, mono=True)])])
text(s, px+0.28, py2+3.95, pw-0.56, 0.3, [P([R("▸ steps 3 → 6 repeat until the answer is complete and grounded", size=9.5, color=BLUE, mono=True)])])
bottom_bar(s, "ANSWERS ARE EARNED THROUGH ITERATION — not guessed on the first pass", accent=GREEN, y=6.55)
chrome(s, 4, "AI ARCHITECTURE")
notes(s, "Two panels. Left: central PROVIDER ROUTER node; primary path burns neon green (GROQ — ultra-low latency) shown mid-failover with a red glitch-sever, traffic rerouting onto the electric-blue GEMINI path. Right: orbital ring of 8 nodes with a directional energy pulse around an AGENT CONTEXT core; segment 4→7 glows brighter, labeled ITERATE UNTIL GROUNDED.",
      "Two ideas. First, resilience: LLM providers are utilities, and utilities go down — the router treats Groq and Gemini as interchangeable, so when Groq rate-limits mid-session the user never notices. Second, the loop: this is not one-shot prompt → answer. The agent reasons, acts, observes, iterates — if the first query errors or returns partial data, it self-corrects. Answers are earned through iteration.")

# ================================================================ SLIDE 5
s = slide_new(); bg_plain(s)
kicker_title(s, "05 · CORE TOOLS — 1 OF 3", [R("get_schema", size=30, color=GREEN, bold=True, mono=True),
                                             R("  &  ", size=30, color=MUTED, bold=True),
                                             R("execute_query", size=30, color=GREEN, bold=True, mono=True)])
# left terminal
term_header(s, 0.55, 1.72, 5.95, "get_schema", accent=GREEN)
rect(s, 0.55, 2.14, 5.95, 3.55, fill=PANEL, line=LINE, lw=1.0)
text(s, 0.85, 2.38, 5.4, 1.75, [
    P([R("sales_db", size=12, color=BLUE, mono=True, bold=True)], sa=2),
    P([R(" ├─ ", size=12, color=FAINT, mono=True), R("customers", size=12, color=TEXT, mono=True)], sa=2),
    P([R(" ├─ ", size=12, color=FAINT, mono=True), R("products", size=12, color=TEXT, mono=True)], sa=2),
    P([R(" ├─ ", size=12, color=FAINT, mono=True), R("orders", size=12, color=TEXT, mono=True)], sa=2),
    P([R(" └─ ", size=12, color=FAINT, mono=True), R("order_items", size=12, color=TEXT, mono=True)], sa=2),
    P([R("   tables · columns · types · relations", size=8.5, color=FAINT, mono=True)])])
bullets(s, 0.85, 4.35, 5.4, 1.2, [
    ("Live runtime introspection", "no hardcoded, stale schema assumptions"),
], size=11.5, gap=8)
# right terminal
term_header(s, 6.83, 1.72, 5.95, "execute_query", accent=GREEN)
rect(s, 6.83, 2.14, 5.95, 3.55, fill=PANEL, line=LINE, lw=1.0)
text(s, 7.13, 2.38, 5.4, 1.15, [
    P([R("SELECT c.name, SUM(oi.qty*oi.price)", size=11, color=GREEN, mono=True)], sa=1),
    P([R("FROM order_items oi", size=11, color=GREEN, mono=True)], sa=1),
    P([R("JOIN orders o ON o.id = oi.order_id", size=11, color=GREEN, mono=True)], sa=1),
    P([R("JOIN customers c ON c.id = o.customer_id", size=11, color=GREEN, mono=True)], sa=1),
    P([R("GROUP BY c.name;", size=11, color=GREEN, mono=True)], sa=4),
    P([R("✓ 1,284 rows returned", size=10, color=BLUE, mono=True, bold=True),
       R("   [READ-ONLY ✓]", size=10, color=GREEN, mono=True, bold=True)])])
bullets(s, 7.13, 4.0, 5.4, 0.75, [
    ("Exact SQL", "generated & executed on the real database"),
], size=11.5, gap=6)
rect(s, 7.13, 4.72, 5.35, 0.55, fill=RED_D, line=RED, lw=1.0, rounded=True, adj=0.25)
text(s, 7.3, 4.72, 5.1, 0.55, [P([R("DROP TABLE orders;", size=10.5, color=RED, mono=True, strike=True),
                                  R("  ⚠ BLOCKED — WRITE GUARD", size=10, color=RED, mono=True, bold=True)])],
     anchor=MSO_ANCHOR.MIDDLE)
bottom_bar(s, "RADICAL TRANSPARENCY — every answer ships with the SQL that produced it. Copy it. Audit it. Rerun it.", accent=GREEN, y=6.05)
chrome(s, 5, "CORE TOOLS 1/3")
notes(s, "Two floating terminals. Left (get_schema): glowing schema tree — sales_db → customers/products/orders/order_items — column types cascading in green monospace. Right (execute_query): a SELECT+JOIN executing with a neon shield icon and green READ-ONLY badge; beneath it, a red-stamped rejected 'DROP TABLE orders;' marked BLOCKED — WRITE GUARD. Results stream below like a matrix cascade.",
      "Trust is the product, and these two tools are its foundation. The agent reads the schema fresh at runtime — it cannot reference a column that doesn't exist — and it executes real SQL, so answers come from the database, not the model's memory. The read-only guard means the worst-case interaction is an empty result set. Users always see the exact SQL behind every number.")

# ================================================================ SLIDE 6
s = slide_new(); bg_plain(s)
kicker_title(s, "06 · CORE TOOLS — 2 OF 3", [R("generate_chart", size=30, color=BLUE, bold=True, mono=True),
                                             R("  &  ", size=30, color=MUTED, bold=True),
                                             R("generate_flowchart", size=30, color=BLUE, bold=True, mono=True)])
py = panel(s, 0.55, 1.72, 5.95, 4.1, "generate_chart — PLOTLY", accent=BLUE)
bullets(s, 0.85, py+0.25, 5.35, 3.3, [
    ("Chart specs authored from real query results", "the model writes the definition, the browser renders it"),
    ("Fully interactive", "zoom · hover · pan · filter — exploration, not screenshots"),
    ("Type matched to the question", "bar · line · scatter · pie"),
    ("Deterministic rendering", "no hallucinated visuals"),
], size=12, marker="▸", mcolor=BLUE, gap=11)
py = panel(s, 6.83, 1.72, 5.95, 4.1, "generate_flowchart — MERMAID", accent=BLUE)
bullets(s, 7.13, py+0.25, 5.35, 3.3, [
    ("Mermaid syntax written by the agent", "diagrams render live inside the chat"),
    ("Schema maps & ER diagrams", "entity relationships on demand"),
    ("Logic & pipeline flows", "the shape of the system, not just the numbers"),
    ("Code-to-visual pipeline", "graph TD; style — generated, then rendered"),
], size=12, marker="▸", mcolor=BLUE, gap=11)
bottom_bar(s, "ONE PROMPT → the numbers AND the picture of the numbers", accent=BLUE, y=6.05)
chrome(s, 6, "CORE TOOLS 2/3")
notes(s, "Split dashboard inside a futuristic browser chrome. Left: interactive Plotly chart mid-render — glowing series, neon-gradient fills, hover tooltip caught on a data point. Right: Mermaid relationship diagram as a neon blueprint — nodes and arrows in electric blue with green highlights. Behind both, faint source code (fig = px.bar(...) / graph TD;) streams like a data waterfall and morphs into the visuals.",
      "Text-only answers waste an LLM's real strength: structured generation. Zodiac treats visuals as first-class outputs — the model doesn't describe a chart, it writes the chart definition, and Plotly and Mermaid render it deterministically. Because the spec is generated from executed query results, the visual carries the same grounding as the SQL behind it.")

# ================================================================ SLIDE 7
s = slide_new(); bg_plain(s)
kicker_title(s, "07 · CORE TOOLS — 3 OF 3", [R("explain_data", size=30, color=GREEN, bold=True, mono=True),
                                             R(" — Grounded, Not Guessed", size=30, color=BLUE, bold=True)])
# pipeline
pl = [("1 · QUERY", GREEN, "against the live database"),
      ("2 · COMPUTE", BLUE, "totals · means · min/max · shares"),
      ("3 · NARRATE", GREEN, "prose around proven numbers")]
pxs = 1.45
for lbl, acc, d in pl:
    rect(s, pxs, 1.85, 3.1, 1.05, fill=PANEL, line=acc, lw=1.4, rounded=True, adj=0.14)
    text(s, pxs, 1.98, 3.1, 0.4, [P([R(lbl, size=15, color=acc, mono=True, bold=True)], align=PP_ALIGN.CENTER)])
    text(s, pxs+0.1, 2.4, 2.9, 0.35, [P([R(d, size=8.5, color=MUTED, mono=True)], align=PP_ALIGN.CENTER)])
    if lbl != "3 · NARRATE":
        text(s, pxs+3.12, 1.85, 0.75, 1.05, [P([R("▶▶", size=16, color=BLUE, mono=True, bold=True)], align=PP_ALIGN.CENTER)], anchor=MSO_ANCHOR.MIDDLE)
    pxs += 3.87
bullets(s, 1.15, 3.35, 11.1, 2.4, [
    ("Numeric summaries from executed result sets", "totals, averages, min/max, distributions, shares — computed, not recalled"),
    ("Every figure traces back to a query result", "text is generated from computed numbers, never from model memory"),
    ("The math ships with the story", "growth rates, proportions, outliers — derived first, then narrated"),
    ("Anti-hallucination by architecture", "if a number can't be traced to a result set, it doesn't make the answer"),
], size=13, gap=12)
rect(s, 0.55, 6.05, SW-1.1, 0.5, fill=GREEN_D, line=GREEN, lw=1.0, rounded=True, adj=0.5)
text(s, 0.9, 6.05, SW-1.8, 0.5, [P([R("> ", size=12, color=GREEN, mono=True, bold=True),
                                    R("We don't ask the model to remember your revenue. We ask your database.", size=12, color=GREEN, mono=True, bold=True)])],
     anchor=MSO_ANCHOR.MIDDLE)
chrome(s, 7, "CORE TOOLS 3/3")
notes(s, "A neon magnifying lens hovers over a glowing result-set table; cells illuminate green and lift out, flowing upward into a paragraph where the same numbers appear highlighted — numbers traced to rows. Math glyphs (∑, μ, x̄, Δ%) orbit the lens. In the corner, a ghostly chatbot bubble with a made-up statistic is slashed in red — the anti-hallucination motif.",
      "This slide separates Zodiac from a wrapped chat API. Generic LLMs interpolate plausible-sounding statistics; Zodiac computes the statistics first and writes the narrative around them. The pipeline is strict: query → compute → narrate. If a number can't be traced to a result set, it doesn't make it into the answer. Say it plainly: 'We don't ask the model to remember your revenue. We ask your database.'")

# ================================================================ SLIDE 8
s = slide_new(); bg_plain(s)
kicker_title(s, "08 · DATABASE ARCHITECTURE", [R("Under the Hood — ", size=30, color=TEXT, bold=True),
                                               R("The Sample Sales Schema", size=30, color=BLUE, bold=True)])
tables = [
    ("CUSTOMERS", GREEN, "who is buying", "customer_id · name · city · signup"),
    ("ORDERS",    BLUE,  "the transaction header", "order_id · customer_id · date · status"),
    ("ORDER_ITEMS", GREEN, "line-level detail", "order_id · product_id · qty · price"),
    ("PRODUCTS",  BLUE,  "what is being sold", "product_id · name · category · price"),
]
rels = ["1 — ∞", "1 — ∞", "∞ — 1"]
tx = 0.55; tw = 2.72; tg = (SW-1.1-4*tw)/3
for i, (name, acc, d, cols) in enumerate(tables):
    rect(s, tx, 1.85, tw, 1.55, fill=PANEL, line=acc, lw=1.3, rounded=True, adj=0.1)
    rect(s, tx, 1.85, tw, 0.42, fill=GREEN_D if acc==GREEN else BLUE_D, rounded=True, adj=0.18)
    text(s, tx, 1.9, tw, 0.32, [P([R(name, size=12.5, color=acc, mono=True, bold=True, spc=60)], align=PP_ALIGN.CENTER)])
    text(s, tx+0.1, 2.36, tw-0.2, 0.3, [P([R(d, size=9.5, color=TEXT)], align=PP_ALIGN.CENTER)])
    text(s, tx+0.08, 2.68, tw-0.16, 0.65, [P([R(cols, size=8, color=MUTED, mono=True)], align=PP_ALIGN.CENTER, ls=1.15)])
    if i < 3:
        text(s, tx+tw-0.06, 2.28, tg+0.12, 0.35, [P([R(rels[i], size=11, color=BLUE, mono=True, bold=True)], align=PP_ALIGN.CENTER)])
        text(s, tx+tw-0.06, 2.05, tg+0.12, 0.3, [P([R("━━━▶" if i < 2 else "◀━━━", size=10, color=BLUE, mono=True)], align=PP_ALIGN.CENTER)])
    tx += tw + tg
bullets(s, 0.85, 3.85, 11.6, 1.9, [
    ("Normalized e-commerce model — four tables, one star", "join-rich by design: every interesting question crosses at least one join"),
    ("The schema that stresses text-to-SQL edge cases", "if the agent is correct here, it generalizes"),
    ("Nothing is hardcoded", "the agent maps all relationships at runtime via get_schema"),
], size=12.5, gap=11)
bottom_bar(s, "SCHEMA DISCOVERED LIVE — never baked into prompts", accent=GREEN, y=6.05)
chrome(s, 8, "DATA MODEL")
notes(s, "An ERD reimagined as a neon circuit schematic on a dark PCB. Four table-nodes as glowing chips — CUSTOMERS (green), PRODUCTS (blue), ORDERS (green), ORDER_ITEMS (blue) — connected by luminous traces with pulsing cardinality glyphs (1 — ∞) at each junction. Pin-labels expose key columns. Faint grid and ruler ticks frame it like engineering blueprints.",
      "The demo dataset is deliberately canonical e-commerce: instantly understandable, and join-heavy — exactly where naive text-to-SQL breaks. Every interesting question (top customers, revenue per product, average order value) crosses at least one join. If the agent is correct on this schema, it generalizes. And the schema is never baked into prompts — it's discovered live.")

# ================================================================ SLIDE 9
s = slide_new(); bg_plain(s)
kicker_title(s, "09 · VALID QUERY CATEGORIES & SAMPLE PROMPTS", [R("Ask It Anything ", size=30, color=TEXT, bold=True),
                                                                 R("(Structured)", size=30, color=BLUE, bold=True)])
cards = [
    ("AGGREGATION", '"What was total revenue per month this year?"'),
    ("RANKING",     '"Show me the top 10 customers by lifetime value."'),
    ("TREND",       '"Plot quarterly sales growth as a line chart."'),
    ("JOIN",        '"Which products are most frequently bought together?"'),
    ("SCHEMA",      '"Diagram how my tables relate to each other."'),
    ("QUALITY",     '"Are there any orders with no line items?"'),
]
cw = (SW-1.1-0.44)/3
for i, (cat, q) in enumerate(cards):
    col, row = divmod(i, 3)
    cx = 0.55 + col*(cw+0.22); cy = 1.8 + row*1.72
    rect(s, cx, cy, cw, 1.5, fill=PANEL, line=LINE, lw=1.0, rounded=True, adj=0.08)
    rect(s, cx, cy, 0.06, 1.5, fill=BLUE if i % 2 == 0 else GREEN)
    text(s, cx+0.25, cy+0.18, cw-0.45, 0.3, [P([R("// ", size=10, color=FAINT, mono=True),
                                                 R(cat, size=10.5, color=BLUE if i % 2 == 0 else GREEN, mono=True, bold=True, spc=120)])])
    text(s, cx+0.25, cy+0.52, cw-0.45, 0.85, [P([R(q, size=12.5, color=TEXT, italic=True)], ls=1.12)])
bottom_bar(s, "EVERY ANSWER = EXACT SQL + INTERACTIVE CHART + GROUNDED EXPLANATION", accent=GREEN, y=5.6)
text(s, 0.55, 6.25, SW-1.1, 0.35, [P([R("> ", size=13, color=GREEN, mono=True, bold=True),
                                      R("ask anything", size=13, color=TEXT, mono=True),
                                      R("_", size=13, color=GREEN, mono=True, bold=True)])])
chrome(s, 9, "SAMPLE PROMPTS")
notes(s, "Dark-mode chat UI with a vertical stack of user message bubbles in neon-green outline, each paired with an electric-blue category chip (AGGREGATION, RANKING, TREND, JOIN, SCHEMA, QUALITY). One exchange expanded to show the full multi-part response: SQL block, mini Plotly chart, explanation paragraph. An empty input field with a blinking cursor: '> ask anything_'",
      "Walk one or two bubbles aloud, then let the audience read the rest. Point out the data-quality category last — it proves the agent can reason about the database, not just retrieve from it. Close on the bundle: a single prompt yields the SQL, a rendered chart, and a grounded explanation. That triple — query, visual, narrative — is the product experience.")

# ================================================================ SLIDE 10
s = slide_new(); bg_plain(s)
kicker_title(s, "10 · GETTING STARTED", [R("Zero to Query in ", size=30, color=TEXT, bold=True),
                                         R("Five Minutes", size=30, color=GREEN, bold=True)])
term_header(s, 0.55, 1.72, 7.5, "boot-sequence — zodiac-ai", accent=GREEN)
rect(s, 0.55, 2.14, 7.5, 3.9, fill=PANEL, line=LINE, lw=1.0)
text(s, 0.85, 2.4, 6.95, 3.5, [
    P([R("$ ", size=12, color=GREEN, mono=True, bold=True),
       R("git clone https://github.com/<org>/zodiac-ai.git", size=12, color=TEXT, mono=True)], sa=4),
    P([R("$ ", size=12, color=GREEN, mono=True, bold=True),
       R("cd zodiac-ai", size=12, color=TEXT, mono=True)], sa=10),
    P([R("# .env", size=11, color=FAINT, mono=True)], sa=4),
    P([R("GROQ_API_KEY", size=11.5, color=GREEN, mono=True, bold=True),
       R("=••••••••••", size=11.5, color=TEXT, mono=True),
       R("   # primary provider", size=10, color=FAINT, mono=True)], sa=4),
    P([R("GEMINI_API_KEY", size=11.5, color=BLUE, mono=True, bold=True),
       R("=••••••••••", size=11.5, color=TEXT, mono=True),
       R("   # optional fallback", size=10, color=FAINT, mono=True)], sa=10),
    P([R("$ ", size=12, color=GREEN, mono=True, bold=True),
       R("pip install -r requirements.txt", size=12, color=TEXT, mono=True),
       R("  # backend ✓", size=10, color=FAINT, mono=True)], sa=4),
    P([R("$ ", size=12, color=GREEN, mono=True, bold=True),
       R("npm install && npm run dev", size=12, color=TEXT, mono=True),
       R("         # frontend ✓", size=10, color=FAINT, mono=True)], sa=10),
    P([R("▲ ", size=12, color=BLUE, mono=True, bold=True),
       R("open the local URL  →  start asking", size=12, color=BLUE, mono=True, bold=True)]),
])
px2 = 8.35
py3 = panel(s, px2, 1.72, 4.43, 4.32, "RUN NOTES", accent=BLUE)
bullets(s, px2+0.28, py3+0.25, 3.9, 3.5, [
    ("Prerequisites", "Python 3.10+ · Node 18+ · Git · ≥ 1 API key"),
    ("No DB server", "SQLite ships inside the repo"),
    ("One key runs it", "a second key activates the failover router"),
    ("< 5 minutes", "from git clone to first grounded answer"),
], size=11.5, marker="▸", mcolor=BLUE, gap=10)
bottom_bar(s, "CLONE IT DURING Q&A — the demo is the pitch", accent=GREEN, y=6.25)
chrome(s, 10, "GETTING STARTED")
notes(s, "A full terminal window styled as a boot sequence: command lines appearing with a typing cursor, executed lines flashing green with ✓, a progress-bar motif [████████░░] LOADING. The final command launches the app — a small browser thumbnail pops out showing the Zodiac chat UI. At right, a .env file rendered as a glowing document with masked keys.",
      "The pitch of this slide is friction — or its absence. SQLite means no database server to stand up: the schema ships inside the repo. One API key is enough to run; the second is optional insurance that activates failover. From git clone to first grounded answer: under five minutes. Invite the audience to do it live during Q&A.")

# ================================================================ SLIDE 11
s = slide_new()
s.shapes.add_picture("assets/closing-bg.png", 0, 0, Inches(SW), Inches(SH))
ov = rect(s, 0, 0, 8.6, SH, fill=BG); set_fill_alpha(ov, 52)
rect(s, 0, 0, SW*0.42, 0.05, fill=GREEN); rect(s, SW*0.42, 0, SW*0.58, 0.05, fill=BLUE)
text(s, 0.62, 0.5, 8.0, 0.3, [P([R("$ status: ", size=11, color=MUTED, mono=True),
                                 R("ALL SYSTEMS OPERATIONAL", size=11, color=GREEN, mono=True, bold=True)])])
text(s, 0.62, 0.95, 9.5, 0.6, [P([R("The Database Is Now ", size=31, color=TEXT, bold=True),
                                  R("a Conversation", size=31, color=BLUE, bold=True)])])
rect(s, 0.64, 1.62, 1.7, 0.035, fill=GREEN); rect(s, 2.34, 1.62, 0.55, 0.035, fill=BLUE)
bullets(s, 0.62, 1.95, 8.1, 3.3, [
    ("What we built", "exact SQL · interactive Plotly charts · live Mermaid diagrams · grounded explanations"),
    ("Why it matters", "analytics at conversational speed — without the hallucination tax"),
    ("How it holds up", "read-only guardrails · runtime schema awareness · multi-provider failover · 8-step reasoning"),
    ("What's next", "Postgres / MySQL connectors · richer chart grammar · conversational follow-ups · shared query memory"),
], size=12.5, gap=11)
text(s, 0.62, 5.35, 8.0, 0.4, [P([R("▸ ", size=15, color=GREEN, mono=True, bold=True),
                                  R("Clone it. Query it. Break it.", size=15, color=GREEN, mono=True, bold=True, spc=120)])])
chromatic(s, 9.15, 4.7, 3.6, 1.6, "Q&A", 88)
text(s, 9.2, 6.25, 3.5, 0.35, [P([R("> ask anything", size=13, color=TEXT, mono=True),
                                  R("_", size=13, color=GREEN, mono=True, bold=True)], align=PP_ALIGN.CENTER)])
text(s, 0.62, 6.85, 8.0, 0.3, [P([R("ZODIAC AI // THANK YOU", size=9, color=FAINT, mono=True, spc=250)])])
notes(s, "Wide cinematic shot of the Zodiac constellation — 12 stars connected by neon lines — reflected on the cyber-grid floor, closing the visual loop with Slide 1. Data streams flow from the constellation down into a small glowing database cylinder. Terminal reads '$ status: ALL SYSTEMS OPERATIONAL' with a blinking '> ask anything_'. Oversized Q&A in glitch typography.",
      "Close the loop with the opening line: databases used to be things you queried — now they're things you talk to. Recap the four pillars in one breath: exact SQL, live visuals, grounded numbers, resilient architecture. Open the floor — and if the room is quiet, seed the first question: 'Ask me what happens when Groq goes down mid-query.'")

# ---------------------------------------------------------------- save
prs.core_properties.title = "Zodiac AI — Pitch Deck & Technical Overview"
prs.core_properties.author = "Zodiac AI"
prs.core_properties.subject = "Chat with your Database — text-to-SQL agent"
OUT = "Zodiac_AI_Pitch_Deck.pptx"
prs.save(OUT)
print(f"saved {OUT} with {len(prs.slides._sldIdLst)} slides")
