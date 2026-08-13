"""
CityVerse AI — National Hackathon Pitch Deck
Builds a 12-slide premium .pptx with Morph transitions, cinematic animations
and full speaker notes.  Run:  python3 build_deck.py
"""
import math, os
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.shapes import MSO_SHAPE
from deck_lib import *
from deck_lib import _xfrag
from pptx.oxml.ns import qn

HERE = os.path.dirname(os.path.abspath(__file__))
ASSETS = os.path.join(HERE, "assets")
def A(name): return os.path.join(ASSETS, name)

prs = Presentation()
prs.slide_width  = SLIDE_W
prs.slide_height = SLIDE_H
BLANK = prs.slide_layouts[6]

def new_slide():
    return prs.slides.add_slide(BLANK)

# =====================================================================
# SLIDE 1 — OPENING IMPACT
# =====================================================================
def slide01():
    s = new_slide()
    an = Anim()
    p = pic(s, A("hero_city_twin.png"), 0, 0, 13.333, 7.5, name="hero")
    # cinematic scrims
    top = rect(s, 0, 0, 13.333, 1.5,
               grad=([(0, "02040C", 72), (100, "02040C", 0)], 90))
    bot = rect(s, 0, 3.1, 13.333, 4.4,
               grad=([(0, "02040C", 0), (62, "02040C", 55), (100, "02040C", 86)], 90))
    # HUD corner brackets
    m, L = 0.34, 0.52
    br = []
    for (x1, y1, x2, y2, x3, y3) in [
            (m, m + L, m, m, m + L, m),
            (13.333 - m - L, m, 13.333 - m, m, 13.333 - m, m + L),
            (m, 7.5 - m - L, m, 7.5 - m, m + L, 7.5 - m),
            (13.333 - m - L, 7.5 - m, 13.333 - m, 7.5 - m, 13.333 - m, 7.5 - m - L)]:
        br.append(seg(s, x1, y1, x2, y2, color=ICE, w=1.2, alpha=55))
        br.append(seg(s, x2, y2, x3, y3, color=ICE, w=1.2, alpha=55))
    k = txt(s, 2.667, 0.62, 8.0, 0.32,
            [[R("NATIONAL HACKATHON", 11, CYAN, True, spc=340),
              R("   ·   ", 11, DIM, True),
              R("SMART CITIES TRACK", 11, ICE, True, spc=340)]],
            align="c", name="kicker")
    t = txt(s, 1.667, 4.42, 10.0, 1.15,
            [[R("CityVerse ", 64, WHITE, True),
              R("AI", 64, WHITE, True,
                g=([(0, CYAN), (55, BLUE), (100, PURPLE)], 0))]],
            align="c", anchor="m", name="title")
    add_glow(t, CYAN, 16, alpha=45)
    sub = txt(s, 2.167, 5.62, 9.0, 0.42,
              [[R("A Living Digital Twin for Smarter, Safer & Sustainable Cities",
                  16.5, "DBE7FB", f=FONT_LT)]], align="c", name="subtitle")
    seg(s, 4.42, 6.42, 5.72, 6.42, color=CYAN, w=1.2, alpha=70, glow=CYAN, glow_r=4)
    seg(s, 7.62, 6.42, 8.92, 6.42, color=CYAN, w=1.2, alpha=70, glow=CYAN, glow_r=4)
    tag = txt(s, 1.167, 6.22, 11.0, 0.36,
              [[R("SEE THE CITY  ·  PREDICT THE FUTURE  ·  ACT BEFORE IT HAPPENS",
                  11.5, ICE, True, spc=280)]], align="c", name="tagline")
    # choreography
    an.add(p, "fade", dur=1500, start="click")
    an.add(top, "fade", dur=900, start="after", delay=200)
    an.add(bot, "fade", dur=900, start="with")
    an.add(k, "fade", dur=700, start="after", delay=500)
    an.add(t, "rise", dur=900, start="after", delay=250)
    an.add(sub, "fade", dur=800, start="after", delay=350)
    an.add(tag, "fade", dur=800, start="after", delay=250)
    for b in br:
        an.add(b, "fade", dur=500, start="with")
    add_timing(s, an)
    add_notes(s, "Good morning judges. Every minute, our cities generate millions of "
              "data points — yet the people running them still can't see the whole "
              "picture. This is CityVerse AI — a living digital twin of an entire "
              "city. It watches every signal in real time, predicts what happens "
              "next, and helps authorities act before problems occur. Over the next "
              "six minutes, I'll show you how a city learns to see itself. "
              "(Click once to build the scene.)")

# =====================================================================
# SLIDE 2 — THE PROBLEM
# =====================================================================
def slide02():
    s = new_slide()
    an = Anim()
    content_base(s, "The Problem", "Cities Have Data. But Not a Unified View.", 2)

    # ---- left card: the chaos
    lc = glass(s, 0.62, 1.95, 5.92, 4.12, edge=RED, edge_alpha=28, name="chaosCard")
    txt(s, 0.92, 2.18, 5.0, 0.3, [[R("WHAT CITIES FACE EVERY DAY", 10.5, RED,
                                     True, spc=280)]])
    issues = [("Traffic Congestion", "gridlock decides the morning"),
              ("Air Pollution", "AQI spikes go unnoticed"),
              ("Energy Waste", "consumption blind spots"),
              ("Aging Infrastructure", "failures found too late"),
              ("Emergency Delays", "when minutes decide lives")]
    rows = []
    for i, (t1, t2) in enumerate(issues):
        y = 2.62 + i * 0.665
        d = oval(s, 1.06, y + 0.19, 0.11, color=RED, glow=RED, glow_r=6)
        r1 = txt(s, 1.32, y, 4.9, 0.30, [[R(t1, 13.5, WHITE, True)]])
        r2 = txt(s, 1.32, y + 0.265, 4.9, 0.24, [[R(t2, 9.5, MUTED)]])
        if i < 4:
            seg(s, 1.32, y + 0.60, 6.22, y + 0.60, color=GRIDLN, w=0.75, alpha=60)
        rows += [d, r1, r2]

    # ---- right card: the silos
    rc = glass(s, 6.82, 1.95, 5.88, 4.12, edge_alpha=26, name="siloCard")
    txt(s, 7.12, 2.18, 5.0, 0.3, [[R("WHAT PLANNERS ACTUALLY HAVE", 10.5, ICE,
                                     True, spc=280)]])
    silos = [("Traffic Control", "own portal"), ("Pollution Board", "PDF reports"),
             ("Power Utility", "monthly CSV"), ("Disaster Mgmt", "phone lines")]
    cards = []
    for i, (t1, t2) in enumerate(silos):
        x = 7.12 + (i % 2) * 2.76
        y = 2.62 + (i // 2) * 1.24
        c = glass(s, x, y, 2.6, 1.08, radius=0.14, alpha=38, edge=DIM,
                  edge_alpha=45, name="silo%d" % i)
        c._element.spPr.find(qn("a:ln")).append(_xfrag('<a:prstDash val="dash"/>'))
        txt(s, x + 0.22, y + 0.18, 2.2, 0.3, [[R(t1, 12, "D7E1F2", True)]])
        txt(s, x + 0.22, y + 0.50, 2.2, 0.26, [[R(t2, 9, DIM)]])
        cross = seg(s, x + 2.24, y + 0.16, x + 2.40, y + 0.32, color=RED, w=1.6,
                    alpha=80)
        cross2 = seg(s, x + 2.40, y + 0.16, x + 2.24, y + 0.32, color=RED, w=1.6,
                     alpha=80)
        cards += [c, cross, cross2]
    warn = []
    for i in range(3):
        warn.append(oval(s, 12.28 + i * 0.16, 2.32, 0.075, color=RED,
                         glow=RED, glow_r=6))
    chip = rect(s, 7.12, 5.10, 5.32, 0.62, color=RED, alpha=13, radius=0.5,
                line=RED, line_alpha=55, line_w=1.1, name="alertChip")
    txt(s, 7.12, 5.10, 5.32, 0.62,
        [[R("⚠", 13, RED, True), R("  SIX SYSTEMS. ZERO CONVERSATION.", 11.5, RED,
                                   True, spc=160)]], align="c", anchor="m")

    # ---- bottom statement
    rect(s, 0.62, 6.36, 0.085, 0.62, grad=([(0, RED, None), (100, AMBER, None)], 90))
    q = txt(s, 0.92, 6.40, 11.4, 0.56,
            [[R("“City data exists in silos — making proactive planning nearly impossible.”",
                16.5, "E8EEF9", i=True)]], anchor="m", name="quote")

    tl = content = an
    an.add(lc, "wipe", direction="right", dur=650, start="click")
    for i, r in enumerate(rows):
        an.add(r, "fade", dur=400, start="after", delay=90 if i else 250)
    an.add(rc, "fade", dur=600, start="after", delay=250)
    for i, c in enumerate(cards):
        an.add(c, "zoom", dur=420, start="after", delay=60 if i else 120)
    an.add(chip, "fade", dur=500, start="after", delay=300)
    an.add(chip, "blink", dur=1400, start="with", loop=True)
    for w in warn:
        an.add(w, "blink", dur=900, start="with", loop=True)
    an.add(q, "wipe", direction="right", dur=700, start="after", delay=350)
    add_timing(s, an)
    add_notes(s, "Here's the problem. A city's traffic, pollution, energy, "
              "infrastructure and emergency data all live in different departments, "
              "different dashboards, different formats. Traffic control doesn't talk "
              "to the pollution board; the power utility doesn't talk to disaster "
              "management. So decisions are slow, reactive, and too late. City data "
              "exists in silos — and in a siloed city, proactive planning is "
              "nearly impossible. That's exactly the gap CityVerse attacks.")

# =====================================================================
# SLIDE 3 — OUR VISION
# =====================================================================
def slide03():
    s = new_slide()
    an = Anim()
    p = pic(s, A("city_transform.png"), 0, 0, 13.333, 7.5, name="vision")
    rect(s, 0, 0, 13.333, 7.5, color="040812", alpha=58)
    rect(s, 0, 4.9, 13.333, 2.6,
         grad=([(0, "02040C", 0), (100, "02040C", 80)], 90))
    k = txt(s, 3.667, 0.66, 6.0, 0.3,
            [[R("OUR VISION", 10.5, CYAN, True, spc=380)]], align="c")
    t = txt(s, 1.667, 1.02, 10.0, 0.7,
            [[R("What If a City Could See Itself?", 38, WHITE, True)]], align="c")

    stages = [("Physical City", "streets · signals · sensors"),
              ("Live Digital Twin", "a real-time virtual replica"),
              ("AI Predictions", "see the next 30 minutes"),
              ("Better Decisions", "act before it happens")]
    nodes, arrows = [], []
    xw, gap = 2.56, 0.58
    x0 = (13.333 - (4 * xw + 3 * gap)) / 2
    for i, (t1, t2) in enumerate(stages):
        x = x0 + i * (xw + gap)
        n = glass(s, x, 3.42, xw, 1.10, radius=0.18, alpha=52,
                  edge=CYAN if i else ICE, edge_alpha=45, glow=CYAN,
                  name="stage%d" % i)
        txt(s, x, 3.62, xw, 0.32, [[R(t1, 15, WHITE, True)]], align="c")
        txt(s, x, 3.96, xw, 0.28, [[R(t2, 9.5, MUTED)]], align="c")
        nodes.append(n)
        if i < 3:
            ax = x + xw + 0.075
            ar = seg(s, ax, 3.97, ax + gap - 0.15, 3.97, color=CYAN, w=2.2,
                     alpha=85, arrow=True, glow=CYAN, glow_r=5)
            arrows.append(ar)
    cap = txt(s, 2.667, 5.32, 8.0, 0.4,
              [[R("One living model. Every signal. Zero guesswork.", 14.5, ICE,
                  i=True, f=FONT_LT)]], align="c")
    dots = []
    for i in range(3):
        d = oval(s, x0 + xw + i * 0.2, 3.97, 0.055, color=CYAN, glow=CYAN,
                 glow_r=6)
        an.add(d, "flow", dur=3400, delay=i * 1100, start="with",
               path="M 0 0 L 0.72 0")
        dots.append(d)

    an.add(p, "fade", dur=1300, start="click")
    an.add(t, "rise", dur=800, start="after", delay=300)
    an.add(k, "fade", dur=600, start="with")
    for i, n in enumerate(nodes):
        an.add(n, "rise", dur=650, start="after", delay=160 if i else 300)
        if i < 3:
            an.add(arrows[i], "wipe", direction="right", dur=450, start="after",
                   delay=80)
    an.add(cap, "fade", dur=700, start="after", delay=300)
    add_timing(s, an)
    add_notes(s, "So we asked a simple question — what if a city could see itself? "
              "CityVerse AI mirrors the physical city into a live digital twin, "
              "layer by layer. IoT signals stream in, the twin updates continuously, "
              "AI looks ahead, and planners finally get decisions instead of just "
              "data. Physical city, to live twin, to predictions, to action — one "
              "continuous learning loop. That's our vision.")

# =====================================================================
# SLIDE 4 — SOLUTION OVERVIEW
# =====================================================================
def slide04():
    s = new_slide()
    an = Anim()
    content_base(s, "The Solution", "Meet CityVerse AI", 4)

    cx, cy, rx, ry = 6.667, 4.50, 3.02, 2.28
    angles = [270, 342, 54, 126, 198]
    names = [("01", "Live Digital City", "the city, mirrored in real time"),
             ("02", "Infrastructure Health", "roads & bridges, self-reporting"),
             ("03", "Traffic Intelligence", "flow simulated and eased"),
             ("04", "Energy & Pollution", "consumption & AQI analytics"),
             ("05", "Emergency Simulation", "rehearse any scenario")]
    pts = [(cx + rx * math.cos(math.radians(a)),
            cy + ry * math.sin(math.radians(a))) for a in angles]
    lines = []
    for (px, py) in pts:
        lines.append(seg(s, cx, cy, px, py, color=CYAN, w=1.0, alpha=30))

    ring = oval(s, cx, cy, 2.92, line=CYAN, line_w=1.1, line_alpha=45,
                glow=CYAN, glow_r=10, name="hubRing")
    hub = pic(s, A("hero_city_twin.png"), cx - 1.12, cy - 1.12, 2.24, 2.24,
              name="hub")
    # mask hub picture into a circle
    spPr = hub._element.spPr
    for el in spPr.findall(qn("a:prstGeom")):
        spPr.remove(el)
    spPr.append(_xfrag('<a:prstGeom prst="ellipse"><a:avLst/></a:prstGeom>'))
    ring2 = oval(s, cx, cy, 2.34, line=ICE, line_w=1.0, line_alpha=60)

    chips = []
    for (px, py), (num, t1, t2) in zip(pts, names):
        w, h = 2.58, 1.04
        x, y = px - w / 2, py - h / 2
        c = glass(s, x, y, w, h, radius=0.15, alpha=62, edge=CYAN,
                  edge_alpha=38, name="mod" + num)
        txt(s, x + 0.17, y + 0.27, 0.55, 0.5,
            [[R(num, 17, CYAN, True, f=FONT_SB)]])
        txt(s, x + 0.70, y + 0.115, w - 0.84, 0.56, [[R(t1, 11.5, WHITE, True)]],
            wrap=True, line_spacing=1.0)
        txt(s, x + 0.70, y + 0.68, w - 0.84, 0.28, [[R(t2, 8, MUTED)]])
        chips.append(c)
    ov = oval(s, cx, cy, 0.12, color=ICE, glow=CYAN, glow_r=8)

    an.add(ring, "zoom", dur=800, start="click")
    an.add(hub, "zoom", dur=800, start="with")
    an.add(ring2, "fade", dur=700, start="after", delay=150)
    for i, ln in enumerate(lines):
        an.add(ln, "fade", dur=380, start="after", delay=70 if i else 200)
    for i, c in enumerate(chips):
        an.add(c, "zoom", dur=480, start="after", delay=110)
    an.add(ov, "fade", dur=400, start="after", delay=150)
    an.add(ov, "pulse", dur=1600, start="with", loop=True)
    add_timing(s, an)
    add_notes(s, "Meet CityVerse AI — five modules around one living model. At the "
              "core, the Live Digital City: a real-time replica of every road, "
              "building and signal. Around it, Infrastructure Health monitoring, "
              "Traffic Intelligence, Energy and Pollution analytics, and Emergency "
              "Simulation. Every module reads from the same twin — so a change "
              "anywhere is visible everywhere. One city, one model, five "
              "superpowers. Now let me show you the engine underneath.")

# =====================================================================
# SLIDE 5 — HOW IT WORKS
# =====================================================================
def slide05():
    s = new_slide()
    an = Anim()
    content_base(s, "System Architecture", "From Sensors to Smart Decisions", 5)

    steps = [("01", "IoT Sensor Network", "10k+ live data points, every second"),
             ("02", "GIS + Real-Time Data", "every signal pinned to the map"),
             ("03", "AI Analytics Engine", "patterns · predictions · anomalies"),
             ("04", "Digital Twin Engine", "a live, queryable simulation"),
             ("05", "Governance Dashboard", "insight → one-click action")]
    NX, NW, NH = 3.72, 3.05, 0.74
    ys = [1.92 + i * (NH + 0.36) for i in range(5)]
    nodes = []
    for i, (num, t1, t2) in enumerate(steps):
        y = ys[i]
        n = glass(s, NX, y, NW, NH, radius=0.2, alpha=58, edge=CYAN,
                  edge_alpha=40, glow=CYAN, glow_r=5, name="pipe" + num)
        txt(s, NX + 0.20, y + 0.115, 0.55, 0.4, [[R(num, 15, CYAN, True)]])
        txt(s, NX + 0.62, y + 0.115, NW - 0.8, 0.3, [[R(t1, 13, WHITE, True)]])
        txt(s, NX + 0.62, y + 0.415, NW - 0.8, 0.26, [[R(t2, 8.5, MUTED)]])
        nodes.append(n)
    arrows = []
    for i in range(4):
        xa = NX + NW / 2
        arrows.append(seg(s, xa, ys[i] + NH + 0.035, xa, ys[i] + NH + 0.315,
                          color=CYAN, w=2.0, alpha=80, arrow=True, glow=CYAN,
                          glow_r=4))

    sensors = ["Traffic cameras", "Air quality sensors", "Smart meters",
               "Road vibration sensors", "Weather feeds"]
    chips, fan = [], []
    for i, name in enumerate(sensors):
        y = 1.92 + i * 0.60
        c = glass(s, 0.85, y, 2.15, 0.5, radius=0.5, alpha=42, edge=ICE,
                  edge_alpha=32, name="sensor%d" % i)
        oval(s, 1.07, y + 0.25, 0.075, color=CYAN, glow=CYAN, glow_r=5)
        txt(s, 1.22, y, 1.75, 0.5, [[R(name, 9.8, "C9DAF2")]], anchor="m")
        chips.append(c)
        fan.append(seg(s, 3.00, y + 0.25, NX - 0.06, ys[0] + NH / 2,
                       color=CYAN, w=0.75, alpha=22))

    rc = glass(s, 8.30, 1.92, 4.42, 4.35, edge_alpha=24, name="whyCard")
    txt(s, 8.60, 2.18, 3.8, 0.3, [[R("DESIGNED FOR THE REAL WORLD", 10, ICE,
                                     True, spc=240)]])
    txt(s, 8.60, 2.62, 3.9, 0.75,
        [[R("≈ 1 s", 34, WHITE, True,
            g=([(0, CYAN), (100, PURPLE)], 0))]])
    txt(s, 8.60, 3.36, 3.9, 0.3, [[R("sensor → decision latency", 10.5, MUTED)]])
    seg(s, 8.60, 3.86, 12.42, 3.86, color=GRIDLN, w=0.75, alpha=70)
    feats = [("Modular microservices", "swap any layer without downtime"),
             ("Open standards", "GeoJSON · REST · WebSockets everywhere"),
             ("Ward → megacity", "same codebase, linear scale-out")]
    for i, (t1, t2) in enumerate(feats):
        y = 4.12 + i * 0.68
        ok = oval(s, 8.78, y + 0.13, 0.24, color=CYAN, alpha=15, line=CYAN,
                  line_alpha=70, line_w=1.2)
        seg(s, 8.73, y + 0.135, 8.765, y + 0.175, color=CYAN, w=1.4, alpha=95)
        seg(s, 8.765, y + 0.175, 8.85, y + 0.075, color=CYAN, w=1.4, alpha=95)
        txt(s, 9.12, y, 3.4, 0.28, [[R(t1, 11.5, WHITE, True)]])
        txt(s, 9.12, y + 0.26, 3.4, 0.26, [[R(t2, 8.5, MUTED)]])

    parts = []
    for i in range(4):
        d = oval(s, NX + NW / 2, ys[0] + NH + 0.02, 0.06, color=CYAN, glow=CYAN,
                 glow_r=6)
        an.add(d, "flow", dur=3000, delay=1800 + i * 750, start="with",
               path="M 0 0 L 0 0.405")
        parts.append(d)

    an.add(chips[0], "wipe", direction="right", dur=420, start="click")
    for i in range(1, 5):
        an.add(chips[i], "wipe", direction="right", dur=400, start="after",
               delay=110)
    for i, f in enumerate(fan):
        an.add(f, "fade", dur=350, start="with")
    for i, n in enumerate(nodes):
        an.add(n, "wipe", direction="up", dur=480, start="after",
               delay=140 if i else 260)
        if i < 4:
            an.add(arrows[i], "fade", dur=300, start="with")
    an.add(rc, "fade", dur=700, start="after", delay=300)
    add_timing(s, an)
    add_notes(s, "The architecture is a five-stage pipeline. IoT sensors — traffic "
              "cameras, air-quality monitors, smart meters, road-vibration sensors "
              "and weather feeds — stream into a GIS layer that geotags everything "
              "in real time. The AI engine finds patterns and makes predictions. "
              "The digital twin engine turns those into a living simulation. And "
              "the governance dashboard turns insight into one-click decisions. "
              "It's fully modular — every layer is swappable and scalable.")

# =====================================================================
# SLIDE 6 — LIVE DIGITAL CITY
# =====================================================================
def slide06():
    s = new_slide()
    an = Anim()
    content_base(s, "Module 01 — Live Digital City", "A City That Updates in Real Time", 6)

    mp = pic(s, A("city_map_dark.png"), 0.62, 1.95, 7.55, 4.42, radius=0.04,
             line=CYAN, line_w=1.1, line_alpha=45, glow=CYAN, glow_r=10,
             name="map")
    txt(s, 0.90, 6.02, 4.5, 0.26,
        [[R("● ", 8, GREEN), R("SECTOR GRID 7 · ALL SYSTEMS NOMINAL", 8, MUTED,
                               spc=180)]])
    dots = []
    for (dx, dy, col) in [(1.35, 2.6, CYAN), (2.6, 3.4, CYAN), (3.5, 2.45, PURPLE),
                          (4.45, 3.1, CYAN), (5.6, 2.55, CYAN), (6.7, 3.5, PURPLE),
                          (7.5, 2.7, CYAN), (3.0, 5.3, CYAN), (5.2, 5.5, CYAN),
                          (6.5, 5.9, CYAN)]:
        d = oval(s, dx, dy, 0.075, color=col, glow=col, glow_r=7)
        halo = oval(s, dx, dy, 0.22, line=col, line_w=0.9, line_alpha=45)
        dots += [d, halo]
    scan = rect(s, 0.72, 1.95, 0.022, 4.42, color=CYAN, alpha=35, glow=CYAN,
                glow_r=6, name="scanline")

    stats = [("TRAFFIC", "Moderate", "avg 24 km/h downtown", AMBER, None),
             ("AIR QUALITY", "AQI 82", "moderate · trending down", GREEN, None),
             ("ENERGY LOAD", "74%", "peak expected 19:00", CYAN, 0.74),
             ("ROAD HEALTH", "Good", "3 minor flags open", GREEN, None)]
    cards, bars = [], []
    for i, (lab, val, sub, col, prog) in enumerate(stats):
        y = 1.95 + i * 1.145
        c = glass(s, 8.42, y, 4.28, 1.0, radius=0.15, alpha=55, edge_alpha=30,
                  name="stat%d" % i)
        d = oval(s, 8.72, y + 0.28, 0.10, color=col, glow=col, glow_r=6)
        txt(s, 8.95, y + 0.155, 1.9, 0.3, [[R(lab, 9, MUTED, True, spc=200)]])
        txt(s, 8.95, y + 0.42, 2.2, 0.32, [[R(sub, 9, DIM)]])
        txt(s, 10.35, y + 0.24, 2.15, 0.5, [[R(val, 21, WHITE, True)]],
            align="r")
        cards.append(c)
        if prog:
            rect(s, 8.95, y + 0.80, 3.45, 0.075, color=GRIDLN, radius=0.5)
            b = rect(s, 8.95, y + 0.80, 3.45 * prog, 0.075, color=CYAN,
                     radius=0.5, glow=CYAN, glow_r=5, name="energyBar")
            bars.append(b)
    an.add(mp, "wipe", direction="right", dur=850, start="click")
    for i, d in enumerate(dots):
        an.add(d, "fade", dur=350, start="after", delay=55 if i else 260)
    for i in range(0, len(dots), 2):
        an.add(dots[i], "pulse", dur=1900, start="with", loop=True)
    an.add(scan, "flow", dur=5200, start="with", path="M 0 0 L 0.535 0")
    for i, c in enumerate(cards):
        an.add(c, "rise", dur=520, start="after", delay=130 if i else 300)
    for b in bars:
        an.add(b, "wipe", direction="right", dur=700, start="after", delay=150)
    add_timing(s, an)
    add_notes(s, "This is the live digital city — the command-center view. Every "
              "road, building, hospital, school and signal, mapped on a real GIS "
              "canvas, with sensors streaming in real time. Check the vitals on "
              "the right: traffic is moderate, AQI at 82, energy load at "
              "74 percent, road health good. The moment anything changes in the "
              "real city, the twin updates — sub-second. Nobody waits for a "
              "monthly PDF ever again.")

# =====================================================================
# SLIDE 7 — AI INTELLIGENCE
# =====================================================================
def slide07():
    s = new_slide()
    an = Anim()
    content_base(s, "Module 03 — AI Engine", "Don't Just Monitor. Predict.", 7)

    br = pic(s, A("ai_core.png"), 0.62, 1.95, 5.3, 4.5, radius=0.045,
             line=CYAN, line_w=1.1, line_alpha=45, glow=PURPLE, glow_r=12,
             name="aibrain")
    chips = []
    for i, name in enumerate(["Forecasting ML", "Anomaly AI", "Vision AI"]):
        x = 0.92 + i * 1.62
        c = glass(s, x, 5.85, 1.5, 0.42, radius=0.5, alpha=72, edge=CYAN,
                  edge_alpha=45, name="model%d" % i)
        txt(s, x, 5.85, 1.5, 0.42, [[R(name, 8.5, ICE, True, spc=60)]],
            align="c", anchor="m")
        chips.append(c)

    preds = [("TF", "Traffic Foresight", "Congestion predicted 30 minutes before it forms", "94%", CYAN),
             ("IN", "Infrastructure Watch", "Road & bridge fatigue flagged before failure", "89%", BLUE),
             ("EN", "Environment Outlook", "Tomorrow's pollution hotspots, mapped today", "91%", PURPLE)]
    cards, links, fillbars = [], [], []
    for i, (mono, t1, t2, conf, col) in enumerate(preds):
        y = 1.95 + i * 1.54
        lk = seg(s, 5.95, y + 0.69, 6.42, y + 0.69, color=col, w=1.2, alpha=45,
                 dash="dash")
        ln = oval(s, 6.40, y + 0.69, 0.07, color=col, glow=col, glow_r=5)
        links += [lk, ln]
        c = glass(s, 6.55, y, 6.15, 1.38, radius=0.13, alpha=55, edge=col,
                  edge_alpha=40, name="pred%d" % i)
        chip = rect(s, 6.85, y + 0.36, 0.66, 0.66,
                    grad=([(0, col, 85), (100, PURPLE if col != PURPLE else BLUE, 85)], 135),
                    radius=0.22, glow=col, glow_r=6)
        txt(s, 6.85, y + 0.36, 0.66, 0.66, [[R(mono, 13, WHITE, True)]],
            align="c", anchor="m")
        txt(s, 7.70, y + 0.24, 3.4, 0.3, [[R(t1, 15, WHITE, True)]])
        txt(s, 7.70, y + 0.58, 3.6, 0.5, [[R(t2, 10, MUTED)]], wrap=True)
        txt(s, 11.35, y + 0.30, 1.15, 0.3, [[R(conf, 16, col, True)]], align="r")
        txt(s, 11.15, y + 0.62, 1.35, 0.24, [[R("CONFIDENCE", 7, DIM, True,
                                                spc=160)]], align="r")
        rect(s, 11.30, y + 0.95, 1.20, 0.07, color=GRIDLN, radius=0.5)
        fb = rect(s, 11.30, y + 0.95, 1.20 * (int(conf[:-1]) / 100.0), 0.07,
                  color=col, radius=0.5, glow=col, glow_r=4, name="conf%d" % i)
        fillbars.append(fb)
        cards.append(c)
    txt(s, 6.55, 6.62, 6.15, 0.28,
        [[R("Models retrain nightly on live city data — the twin gets smarter every day.",
            9.5, DIM, i=True)]])

    an.add(br, "zoom", dur=850, start="click")
    for i, c in enumerate(chips):
        an.add(c, "fade", dur=400, start="after", delay=300 if i == 0 else 110)
    for i, c in enumerate(cards):
        an.add(links[2 * i], "wipe", direction="right", dur=350, start="after",
               delay=200 if i == 0 else 60)
        an.add(links[2 * i + 1], "fade", dur=250, start="with")
        an.add(c, "zoom", dur=500, start="after", delay=130)
        an.add(fillbars[i], "wipe", direction="right", dur=600, start="after",
               delay=260)
    add_timing(s, an)
    add_notes(s, "But monitoring isn't enough. CityVerse doesn't just watch — it "
              "predicts. Traffic models forecast congestion thirty minutes before "
              "it forms. Infrastructure models detect stress signatures on roads "
              "and bridges before they fail. Environmental models flag tomorrow's "
              "pollution hotspots today. Every prediction carries a confidence "
              "score, so planners know what to trust. You fix the problem on "
              "Tuesday — instead of apologizing for it on Friday.")

# =====================================================================
# SLIDE 8 — EMERGENCY SIMULATION
# =====================================================================
def slide08():
    s = new_slide()
    an = Anim()
    content_base(s, "Module 05 — Simulation", "Test Emergencies Before They Happen", 8)

    # left vertical timeline
    steps = [("Incident", "simulated flood hits Sector 4", RED),
             ("AI Analyzes", "exposure, risk & impact in seconds", CYAN),
             ("Traffic Reroutes", "signals adapt, roads reopen", CYAN),
             ("Fastest Route", "units dispatched on optimal path", CYAN),
             ("Full Impact View", "authorities see the outcome", GREEN)]
    seg(s, 1.06, 2.42, 1.06, 6.05, color=GRIDLN, w=1.0, alpha=90)
    nodes = []
    for i, (t1, t2, col) in enumerate(steps):
        y = 2.25 + i * 0.88
        d = oval(s, 1.06, y + 0.12, 0.15, color=col, glow=col, glow_r=7)
        r1 = txt(s, 1.40, y - 0.05, 3.3, 0.30, [[R(t1, 13.5, WHITE, True)]])
        r2 = txt(s, 1.40, y + 0.235, 3.45, 0.26, [[R(t2, 9, MUTED)]])
        nodes += [d, r1, r2]

    mp = pic(s, A("emergency_map.png"), 4.98, 1.95, 7.72, 4.42, radius=0.04,
             line=CYAN, line_w=1.1, line_alpha=45, glow=RED, glow_r=10,
             name="emap")
    chip = rect(s, 5.28, 2.22, 3.52, 0.52, color=RED, alpha=16, radius=0.5,
                line=RED, line_alpha=60, line_w=1.2, name="simChip")
    txt(s, 5.28, 2.22, 3.52, 0.52,
        [[R("● LIVE SIM · FLOOD, SECTOR 4", 9, "FFB3C0", True, spc=100)]],
        align="c", anchor="m")
    ping = oval(s, 9.42, 4.32, 0.16, color=RED, glow=RED, glow_r=8)
    halo = oval(s, 9.42, 4.32, 0.42, line=RED, line_w=1.1, line_alpha=55)
    bn = glass(s, 0.62, 6.50, 12.08, 0.56, radius=0.3, alpha=46, edge=CYAN,
               edge_alpha=35, name="caption")
    txt(s, 0.62, 6.50, 12.08, 0.56,
        [[R("“Plan the response before the real event occurs.”  ", 14.5, ICE, i=True),
          R("— that's the point of a twin.", 10.5, MUTED, i=True)]],
        align="c", anchor="m")

    an.add(nodes[0], "zoom", dur=420, start="click")
    for i in range(1, len(nodes)):
        an.add(nodes[i], "fade", dur=380, start="after",
               delay=140 if i % 3 else 220)
    an.add(mp, "wipe", direction="right", dur=850, start="after", delay=300)
    an.add(chip, "fade", dur=450, start="after", delay=250)
    an.add(chip, "blink", dur=1500, start="with", loop=True)
    an.add(ping, "fade", dur=350, start="with")
    an.add(ping, "pulse", dur=1300, start="with", loop=True)
    an.add(halo, "pulse", dur=1900, start="with", loop=True)
    an.add(bn, "wipe", direction="right", dur=650, start="after", delay=350)
    add_timing(s, an)
    add_notes(s, "Here's where it gets powerful: cities can test emergencies before "
              "they happen. We simulate a flood in Sector 4 — the AI instantly "
              "analyzes exposure, reroutes traffic around the danger zone, computes "
              "the fastest path for every emergency vehicle, and shows authorities "
              "the complete impact — before a single drop of rain falls. You don't "
              "rehearse disaster response during the disaster. You plan it in the "
              "twin, safely, as many times as you need.")

# =====================================================================
# SLIDE 9 — GOVERNANCE DASHBOARD
# =====================================================================
def slide09():
    s = new_slide()
    an = Anim()
    content_base(s, "Module 05+ · Command Center", "One Screen. The Entire City.", 9)

    win = glass(s, 0.62, 1.86, 12.08, 4.88, radius=0.035, alpha=42,
                edge=ICE, edge_alpha=26, glow=CYAN, glow_r=8, name="window")
    for i, col in enumerate([RED, AMBER, GREEN]):
        oval(s, 0.92 + i * 0.155, 2.115, 0.095, color=col, alpha=75)
    txt(s, 2.5, 2.0, 8.33, 0.26, [[R("CITYVERSE COMMAND CENTER — LIVE FEED", 9,
                                     MUTED, True, spc=260)]], align="c")
    seg(s, 0.82, 2.40, 12.52, 2.40, color=GRIDLN, w=0.75, alpha=80)

    # ---------- card A: traffic heatmap
    cA = glass(s, 0.85, 2.56, 3.66, 2.52, radius=0.07, alpha=50, name="cardA")
    txt(s, 1.05, 2.72, 2.4, 0.26, [[R("TRAFFIC HEATMAP", 8.5, ICE, True, spc=200)]])
    oval(s, 4.22, 2.80, 0.075, color=GREEN, glow=GREEN, glow_r=5)
    cells = []
    heat = [[2, 3, 4, 4, 5, 3], [2, 4, 6, 7, 6, 4], [3, 5, 8, 9, 7, 4],
            [2, 4, 6, 7, 5, 3], [2, 3, 4, 4, 4, 2]]
    hcol = {2: (CYAN, 16), 3: (CYAN, 34), 4: (BLUE, 48), 5: (PURPLE, 55),
            6: (AMBER, 60), 7: (AMBER, 78), 8: (RED, 62), 9: (RED, 82)}
    for r in range(5):
        for cidx in range(6):
            col, al = hcol[heat[r][cidx]]
            cells.append(rect(s, 1.05 + cidx * 0.555, 3.06 + r * 0.385,
                              0.50, 0.335, color=col, alpha=al,
                              radius=0.22, name="hcell"))
    # ---------- card B: energy bars
    cB = glass(s, 4.66, 2.56, 3.66, 2.52, radius=0.07, alpha=50, name="cardB")
    txt(s, 4.86, 2.72, 2.6, 0.26, [[R("ENERGY · TODAY", 8.5, ICE, True, spc=200)]])
    txt(s, 6.9, 2.66, 1.2, 0.3, [[R("74%", 12, CYAN, True)]], align="r")
    heights = [0.55, 0.8, 0.62, 1.0, 1.28, 1.5, 1.15, 0.9, 0.7, 0.5]
    cols = [CYAN, CYAN, BLUE, BLUE, PURPLE, PURPLE, BLUE, BLUE, CYAN, CYAN]
    bars = []
    for i, h in enumerate(heights):
        b = rect(s, 4.90 + i * 0.335, 4.82 - h, 0.24, h,
                 grad=([(0, cols[i], 90), (100, "14335F", 45)], 90),
                 radius=0.28, name="ebar")
        bars.append(b)
    seg(s, 4.86, 4.84, 8.28, 4.84, color=GRIDLN, w=1.0, alpha=90)
    # ---------- card C: AQI line
    cC = glass(s, 8.47, 2.56, 3.99, 2.52, radius=0.07, alpha=50, name="cardC")
    txt(s, 8.67, 2.72, 2.6, 0.26, [[R("AIR QUALITY TREND", 8.5, ICE, True, spc=200)]])
    txt(s, 10.9, 2.66, 1.35, 0.3, [[R("AQI 82", 12, GREEN, True)]], align="r")
    pts = [(8.72, 4.30), (9.15, 4.05), (9.58, 4.22), (10.01, 3.80),
           (10.44, 3.95), (10.87, 3.55), (11.30, 3.72), (11.73, 3.35),
           (12.16, 3.48)]
    area = pts + [(12.16, 4.84), (8.72, 4.84)]
    freeform_polyline(s, area, closed=True,
                      fill_grad=([(0, CYAN, 26), (100, CYAN, 2)], 90))
    line = freeform_polyline(s, pts, color=CYAN, w=2.0, glow=CYAN)
    seg(s, 8.67, 4.86, 12.36, 4.86, color=GRIDLN, w=1.0, alpha=90)
    endd = oval(s, 12.16, 3.48, 0.09, color=CYAN, glow=CYAN, glow_r=7)

    # ---------- bottom row: AI recommendation
    bn = glass(s, 0.85, 5.24, 7.47, 1.30, radius=0.10, alpha=52, edge=CYAN,
               edge_alpha=50, glow=CYAN, glow_r=8, name="aiBanner")
    rect(s, 0.85, 5.24, 7.47, 0.045,
         grad=([(0, CYAN, 85), (60, BLUE, 85), (100, PURPLE, 85)], 0))
    star = slide_star = s.shapes.add_shape(MSO_SHAPE.STAR_4_POINT, Inches(1.12),
                                           Inches(5.60), Inches(0.42),
                                           Inches(0.42))
    no_shadow(star); fill_solid(star, CYAN); no_line(star); add_glow(star, CYAN, 10)
    txt(s, 1.72, 5.42, 4.4, 0.26, [[R("AI RECOMMENDATION", 8, CYAN, True, spc=240)]])
    txt(s, 1.72, 5.70, 5.0, 0.6,
        [[R("Reroute traffic from Sector 4 — ", 12.5, "EAF1FB"),
          R("reduce congestion by 18%.", 12.5, ICE, True)]], wrap=True)
    ap = rect(s, 6.85, 5.66, 1.18, 0.46, color=CYAN, alpha=16, radius=0.5,
              line=CYAN, line_alpha=75, line_w=1.1, name="applyBtn")
    txt(s, 6.85, 5.66, 1.18, 0.46, [[R("APPLY", 10, CYAN, True, spc=200)]],
        align="c", anchor="m")
    # infra ring
    gI = glass(s, 8.47, 5.24, 1.96, 1.30, radius=0.10, alpha=52, name="cardI")
    ring = s.shapes.add_shape(MSO_SHAPE.DONUT, Inches(8.66), Inches(5.42),
                              Inches(0.94), Inches(0.94))
    no_shadow(ring)
    try:
        ring.adjustments[0] = 0.14
    except Exception:
        pass
    grad_ring = ([(0, CYAN, 92), (100, PURPLE, 92)], 135)
    fill_gradient(ring, *grad_ring); no_line(ring); add_glow(ring, CYAN, 7)
    txt(s, 8.66, 5.42, 0.94, 0.94, [[R("82", 15, WHITE, True)]], align="c",
        anchor="m")
    txt(s, 9.70, 5.56, 0.75, 0.72, [[R("INFRA", 8, DIM, True, spc=140)],
                                     [R("HEALTH", 8, DIM, True, spc=140)],
                                     [R("B+", 13, GREEN, True)]], wrap=True)
    # alerts
    gL = glass(s, 10.55, 5.24, 1.91, 1.30, radius=0.10, alpha=52, edge=RED,
               edge_alpha=35, name="cardL")
    txt(s, 10.72, 5.38, 1.6, 0.24, [[R("ALERTS", 8, DIM, True, spc=200)]])
    oval(s, 10.80, 5.78, 0.085, color=RED, glow=RED, glow_r=5)
    txt(s, 10.94, 5.65, 1.45, 0.24, [[R("Flood sim · S4", 8.5, "FBC4CD")]])
    oval(s, 10.80, 6.14, 0.085, color=AMBER, glow=AMBER, glow_r=5)
    txt(s, 10.94, 6.01, 1.45, 0.24, [[R("AQI spike · East", 8.5, "F5DFB0")]])

    an.add(win, "zoom", dur=600, start="click")
    an.add(cA, "wipe", direction="up", dur=450, start="after", delay=250)
    for i, cell in enumerate(cells):
        an.add(cell, "zoom", dur=260, start="with" if i % 6 else "after",
               delay=(i % 6) * 45)
    an.add(cB, "wipe", direction="up", dur=450, start="after", delay=300)
    for i, b in enumerate(bars):
        an.add(b, "wipe", direction="up", dur=420, start="with" if i else "after",
               delay=i * 60)
    an.add(cC, "wipe", direction="up", dur=450, start="after", delay=200)
    an.add(line, "wipe", direction="right", dur=800, start="after", delay=200)
    an.add(endd, "fade", dur=250, start="after", delay=130)
    an.add(endd, "pulse", dur=1600, start="with", loop=True)
    an.add(bn, "wipe", direction="right", dur=600, start="after", delay=250)
    an.add(star, "pulse", dur=2000, start="with", loop=True)
    an.add(gI, "wipe", direction="up", dur=400, start="after", delay=150)
    an.add(ring, "wheel", dur=900, start="after", delay=120)
    an.add(gL, "wipe", direction="up", dur=400, start="after", delay=150)
    add_timing(s, an)
    add_notes(s, "Everything converges on one screen. Live traffic heatmap, "
              "energy consumption, air-quality trends, infrastructure health, and "
              "active alerts. But the star feature is the AI recommendation "
              "engine — it doesn't just show problems, it suggests actions: "
              "'reroute traffic from Sector 4, reduce congestion by eighteen "
              "percent.' One click on Apply, and the city adapts. That is what "
              "data-driven governance actually looks like.")

# =====================================================================
# SLIDE 10 — TECHNOLOGY STACK
# =====================================================================
def slide10():
    s = new_slide()
    an = Anim()
    content_base(s, "Under the Hood", "Built for Scale", 10)

    layers = [("LAYER 01", "Frontend", MSO_SHAPE.RECTANGLE,
               ["React + glassmorphic UI", "Mapbox GL / GIS canvas",
                "WebGL 3D twin view"]),
              ("LAYER 02", "Backend", MSO_SHAPE.ROUNDED_RECTANGLE,
               ["Node.js · Python services", "REST + WebSocket APIs",
                "Stream ingestion"]),
              ("LAYER 03", "AI Engine", MSO_SHAPE.STAR_4_POINT,
               ["Forecasting ML (LSTM)", "Anomaly detection",
                "Traffic vision models"]),
              ("LAYER 04", "Data", MSO_SHAPE.CAN,
               ["PostgreSQL + PostGIS", "Time-series sensor store",
                "GeoJSON pipelines"]),
              ("LAYER 05", "Cloud", MSO_SHAPE.CLOUD,
               ["Containerized deploy", "Autoscaling compute",
                "Edge-ready nodes"])]
    cards = []
    for i, (lay, name, icon, items) in enumerate(layers):
        x = 0.62 + i * 2.445
        c = glass(s, x, 2.02, 2.30, 3.55, radius=0.09, alpha=50,
                  edge_alpha=30, name="stack%d" % i)
        rect(s, x + 0.12, 2.02, 2.06, 0.055,
             grad=([(0, CYAN, 85), (100, PURPLE, 0)], 0))
        ic = s.shapes.add_shape(icon, Inches(x + 0.88), Inches(2.32),
                                Inches(0.55), Inches(0.55))
        no_shadow(ic); ic.fill.background()
        line_style(ic, CYAN, 1.4, 75); add_glow(ic, CYAN, 6)
        txt(s, x, 3.06, 2.30, 0.24, [[R(lay, 8, DIM, True, spc=260)]], align="c")
        txt(s, x, 3.30, 2.30, 0.34, [[R(name, 16.5, WHITE, True)]], align="c")
        seg(s, x + 0.55, 3.80, x + 1.75, 3.80, color=CYAN, w=1.0, alpha=45)
        for j, it in enumerate(items):
            oval(s, x + 0.28, 4.075 + j * 0.48, 0.055, color=CYAN, alpha=85)
            txt(s, x + 0.44, 3.96 + j * 0.48, 1.85, 0.42,
                [[R(it, 9, "B9C7DE")]], wrap=True)
        cards.append(c)
    strip = glass(s, 0.62, 5.86, 12.08, 0.88, radius=0.12, alpha=44,
                  edge_alpha=24, name="strip")
    msgs = [("Sub-second sync", "twin never lags reality"),
            ("Modular by design", "every layer swappable"),
            ("Hackathon-proven", "prototype → city scale")]
    for i, (t1, t2) in enumerate(msgs):
        xc = 0.62 + 12.08 / 3 * i
        if i:
            seg(s, xc, 6.06, xc, 6.58, color=GRIDLN, w=0.75, alpha=80)
        txt(s, xc, 6.02, 12.08 / 3, 0.3, [[R(t1, 13, ICE, True)]], align="c")
        txt(s, xc, 6.32, 12.08 / 3, 0.26, [[R(t2, 9, MUTED)]], align="c")

    an.add(cards[0], "wipe", direction="up", dur=480, start="click")
    for i in range(1, 5):
        an.add(cards[i], "wipe", direction="up", dur=480, start="after",
               delay=130)
    an.add(strip, "fade", dur=600, start="after", delay=350)
    add_timing(s, an)
    add_notes(s, "Under the hood — a stack realistic for a student team, yet "
              "genuinely scalable. React with Mapbox GL renders the twin in the "
              "browser. Node and Python services ingest streams over REST and "
              "WebSockets. The AI layer runs LSTM forecasting and anomaly "
              "detection. PostGIS stores the geography; a time-series store "
              "handles the sensor firehose. It all ships containerized to the "
              "cloud — and every layer is modular and swappable.")

# =====================================================================
# SLIDE 11 — IMPACT
# =====================================================================
def slide11():
    s = new_slide()
    an = Anim()
    content_base(s, "Why It Matters", "Better Cities. Better Lives.", 11)

    stats = [("20–30%", "Faster incident response", "when seconds decide outcomes"),
             ("15–25%", "Smoother traffic flow", "modeled congestion relief"),
             ("10–20%", "Energy optimization", "less waste · lower bills"),
             ("EARLY", "Infra risk detection", "fix it before it fails")]
    cards, nums = [], []
    for i, (num, t1, t2) in enumerate(stats):
        x = 0.62 + i * 3.075
        c = glass(s, x, 2.02, 2.87, 2.42, radius=0.11, alpha=52,
                  edge_alpha=30, glow=CYAN if i == 0 else None, name="impact%d" % i)
        tri = s.shapes.add_shape(MSO_SHAPE.ISOSCELES_TRIANGLE, Inches(x + 2.42),
                                 Inches(2.26), Inches(0.26), Inches(0.24))
        no_shadow(tri); fill_solid(tri, GREEN if i != 3 else CYAN, 85)
        no_line(tri); add_glow(tri, GREEN if i != 3 else CYAN, 7)
        n = txt(s, x + 0.24, 2.44, 2.4, 0.62,
                [[R(num, 30 if len(num) > 5 else 30, WHITE, True,
                   g=([(0, CYAN), (100, PURPLE)], 0))]])
        txt(s, x + 0.24, 3.22, 2.45, 0.3, [[R(t1, 12.5, WHITE, True)]])
        txt(s, x + 0.24, 3.52, 2.45, 0.28, [[R(t2, 9, MUTED)]])
        seg(s, x + 0.24, 3.05, x + 1.14, 3.05, color=CYAN, w=1.4, alpha=60,
            glow=CYAN, glow_r=4)
        cards, nums = cards + [c], nums + [n]
    txt(s, 3.667, 4.86, 6.0, 0.28, [[R("WHO BENEFITS", 9.5, DIM, True, spc=400)]],
        align="c")
    who = [("Citizens", "safer, cleaner commutes", MSO_SHAPE.OVAL),
           ("Urban Planners", "simulate before building", MSO_SHAPE.DIAMOND),
           ("Emergency Services", "minutes that save lives", MSO_SHAPE.ISOSCELES_TRIANGLE),
           ("City Government", "budgets spent wisely", MSO_SHAPE.STAR_4_POINT)]
    chips = []
    for i, (t1, t2, icn) in enumerate(who):
        x = 0.777 + i * 3.02
        c = glass(s, x, 5.28, 2.72, 0.92, radius=0.13, alpha=46, edge_alpha=26,
                  name="who%d" % i)
        ring = oval(s, x + 0.42, 5.74, 0.5, line=CYAN, line_w=1.2, line_alpha=70)
        ic = s.shapes.add_shape(icn, Inches(x + 0.31), Inches(5.63),
                                Inches(0.22), Inches(0.22))
        no_shadow(ic); fill_solid(ic, CYAN, 90); no_line(ic)
        txt(s, x + 0.82, 5.47, 1.85, 0.28, [[R(t1, 11.5, WHITE, True)]])
        txt(s, x + 0.82, 5.75, 1.85, 0.26, [[R(t2, 8.5, MUTED)]])
        chips.append(c)

    an.add(cards[0], "fade", dur=500, start="click")
    for i in range(4):
        an.add(nums[i], "wipe", direction="up", dur=520, start="after",
               delay=160 if i else 200)
        if i:
            an.add(cards[i], "fade", dur=420, start="with")
    for i, c in enumerate(chips):
        an.add(c, "zoom", dur=460, start="after", delay=90 if i else 300)
    add_timing(s, an)
    add_notes(s, "What does this mean in numbers? Published smart-city studies show "
              "platforms like this deliver 20 to 30 percent faster incident "
              "response, 15 to 25 percent better traffic flow, and 10 to 20 "
              "percent energy savings — plus early detection of infrastructure "
              "risks before they become failures. Citizens get safer, cleaner "
              "commutes. Planners simulate before they build. Emergency services "
              "save the minutes that save lives. And governments spend every rupee "
              "smarter.")

# =====================================================================
# SLIDE 12 — CLOSING
# =====================================================================
def slide12():
    s = new_slide()
    an = Anim()
    rect(s, 0, 0, 13.333, 7.5, color="000000", name="bg")
    p = pic(s, A("city_dawn.png"), 0, 0, 13.333, 7.5, name="dawn")
    gradT = rect(s, 0, 0, 13.333, 2.0,
                 grad=([(0, "02040C", 78), (100, "02040C", 0)], 90))
    gradB = rect(s, 0, 4.2, 13.333, 3.3,
                 grad=([(0, "02040C", 0), (55, "02040C", 55), (100, "02040C", 90)], 90))
    scrim = blob(s, 6.667, 3.1, 12.5, "02040C", 46)
    qm = txt(s, 3.167, 1.30, 7.0, 0.9, [[R("“", 60, CYAN, True)]], align="c")
    q = txt(s, 1.417, 2.06, 10.5, 1.5,
            [[R("The smartest city isn't the one with the most sensors.", 25,
               WHITE, f=FONT_LT, i=True)],
             [R("It's the one that understands what they're saying.", 25, WHITE,
               True, i=True)]], align="c", line_spacing=1.25)
    div = seg(s, 5.867, 4.02, 7.467, 4.02, color=CYAN, w=1.3, alpha=75,
              glow=CYAN, glow_r=5)
    brand = txt(s, 3.667, 4.30, 6.0, 0.55,
                [[R("CityVerse ", 27, WHITE, True),
                  R("AI", 27, WHITE, True,
                    g=([(0, CYAN), (55, BLUE), (100, PURPLE)], 0))]], align="c")
    tag = txt(s, 1.167, 5.02, 11.0, 0.32,
              [[R("SEE THE CITY  ·  PREDICT THE FUTURE  ·  ACT BEFORE IT HAPPENS",
                  11, ICE, True, spc=260)]], align="c")
    thx = txt(s, 5.167, 6.55, 3.0, 0.3,
              [[R("THANK YOU", 10.5, DIM, True, spc=500)]], align="c")

    an.add(p, "fade", dur=1600, start="click")
    an.add(qm, "fade", dur=800, start="after", delay=400)
    an.add(q, "rise", dur=900, start="after", delay=250)
    an.add(brand, "fade", dur=800, start="after", delay=500)
    an.add(tag, "fade", dur=700, start="after", delay=300)
    an.add(thx, "fade", dur=700, start="after", delay=550)
    # cinematic fade-out: every layer exits together, revealing black
    outs = [p, gradT, gradB, scrim, qm, q, div, brand, tag, thx]
    an.add(outs[0], "fadeout", dur=1600, start="after", delay=2400)
    for sh in outs[1:]:
        an.add(sh, "fadeout", dur=1600, start="with", delay=0)
    add_timing(s, an)
    add_notes(s, "We'll leave you with this. The smartest city isn't the one with "
              "the most sensors — it's the one that understands what they're "
              "saying. CityVerse AI gives cities that understanding: see the "
              "city, predict the future, act before it happens. We're a student "
              "team with a working blueprint and a scalable prototype — and we'd "
              "love to build this for real. Thank you. We welcome your questions.")

# =====================================================================
for fn in [slide01, slide02, slide03, slide04, slide05, slide06, slide07,
           slide08, slide09, slide10, slide11, slide12]:
    fn()

for i, sl in enumerate(prs.slides):
    add_transition(sl, 1100 if i not in (0, 11) else 1250)

out = os.path.join(HERE, "CityVerse_AI_Pitch_Deck.pptx")
prs.save(out)
print("saved:", out, "%.1f MB" % (os.path.getsize(out) / 1e6))
