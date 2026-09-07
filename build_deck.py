#!/usr/bin/env python3
"""
AI EMERGENCY COMMANDER  --  Hackathon presentation generator
Builds a 10-slide, futuristic "emergency command center" deck with python-pptx.
Simulated prototype / human decision-support framing; no invented results.
"""
import os
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn

# ----------------------------------------------------------------------------
# Palette / fonts
# ----------------------------------------------------------------------------
CYAN      = RGBColor(0x00, 0xE5, 0xFF)
CYAN_DIM  = RGBColor(0x0F, 0x6E, 0x8C)
BLUE      = RGBColor(0x2E, 0x8B, 0xFF)
DEEP      = RGBColor(0x06, 0x0B, 0x14)
PANEL     = RGBColor(0x0B, 0x15, 0x24)
PANEL_2   = RGBColor(0x10, 0x1D, 0x31)
TEXT      = RGBColor(0xE8, 0xF2, 0xFF)
MUTED     = RGBColor(0x8A, 0xA4, 0xC4)
AMBER     = RGBColor(0xFF, 0xB3, 0x00)
RED       = RGBColor(0xFF, 0x3B, 0x30)
GREEN     = RGBColor(0x2E, 0xD1, 0x8F)

FONT   = "Arial"
MONO   = "Consolas"

ASSETS = "/home/user/assets"

# ----------------------------------------------------------------------------
# Helpers
# ----------------------------------------------------------------------------
def set_fill(shape, rgb, opacity=100, line_rgb=None, line_pt=1.0, line_opacity=100):
    shape.fill.solid()
    shape.fill.fore_color.rgb = rgb
    if opacity < 100:
        spPr = shape._element.spPr
        srgbClr = spPr.find('.//' + qn('a:srgbClr'))
        if srgbClr is not None:
            a = srgbClr.makeelement(qn('a:alpha'), {'val': str(int(opacity * 1000))})
            srgbClr.append(a)
    if line_rgb is not None:
        shape.line.color.rgb = line_rgb
        shape.line.width = Pt(line_pt)
    else:
        shape.line.fill.background()
    return shape


def add_bg(slide, image):
    if image:
        slide.shapes.add_picture(image, 0, 0, Inches(13.333), Inches(7.5))


def add_shape(slide, x, y, w, h, stype=MSO_SHAPE.RECTANGLE, fill=PANEL, opacity=100,
              line=None, line_pt=1.0, line_opacity=100):
    sp = slide.shapes.add_shape(stype, Inches(x), Inches(y), Inches(w), Inches(h))
    sp.shadow.inherit = False
    if fill is None:
        sp.fill.background()
    else:
        set_fill(sp, fill, opacity, line, line_pt, line_opacity)
    return sp


def add_text(slide, x, y, w, h, text, size=14, color=TEXT, bold=False, align=PP_ALIGN.LEFT,
             font=FONT, anchor=MSO_ANCHOR.TOP, italic=False, spacing=None, line_spacing=None):
    tb = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    p = tf.paragraphs[0]
    p.alignment = align
    if spacing:
        p.space_after = Pt(spacing)
    if line_spacing:
        p.line_spacing = line_spacing
    r = p.add_run()
    r.text = text
    r.font.size = Pt(size)
    r.font.bold = bold
    r.font.italic = italic
    r.font.name = font
    r.font.color.rgb = color
    return tb


def add_para(tf, text, size=13, color=TEXT, bold=False, align=PP_ALIGN.LEFT,
             font=FONT, space_after=6, line_spacing=1.0, first=False):
    p = tf.paragraphs[0] if first else tf.add_paragraph()
    p.alignment = align
    p.space_after = Pt(space_after)
    p.line_spacing = line_spacing
    r = p.add_run()
    r.text = text
    r.font.size = Pt(size)
    r.font.bold = bold
    r.font.name = font
    r.font.color.rgb = color
    return p


def add_rich_box(slide, x, y, w, h, items, anchor=MSO_ANCHOR.TOP):
    """items: list of (text, size, color, bold, align, space_after, font)"""
    tb = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    for i, it in enumerate(items):
        text, size, color, bold, align, space_after = it[0], it[1], it[2], it[3], it[4], it[5]
        font = it[6] if len(it) > 6 else FONT
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        p.space_after = Pt(space_after)
        r = p.add_run()
        r.text = text
        r.font.size = Pt(size)
        r.font.bold = bold
        r.font.name = font
        r.font.color.rgb = color
    return tb


def header(slide, kicker, title):
    add_text(slide, 0.85, 0.45, 8.0, 0.35, f"// {kicker}", 12, CYAN, True, font=MONO)
    add_text(slide, 0.85, 0.72, 11.6, 0.75, title, 30, TEXT, True)
    add_shape(slide, 0.85, 1.58, 1.4, 0.045, fill=CYAN)
    add_shape(slide, 2.35, 1.63, 10.13, 0.012, fill=CYAN_DIM, opacity=70)


def footer(slide, num):
    add_text(slide, 0.85, 7.13, 9.0, 0.25,
             "AI EMERGENCY COMMANDER   ·   SIMULATED PROTOTYPE   ·   HUMAN DECISION-SUPPORT SYSTEM",
             9, MUTED, False, font=MONO)
    add_text(slide, 12.2, 7.13, 0.9, 0.25, f"{num:02d}", 11, CYAN, True, align=PP_ALIGN.RIGHT, font=MONO)


def dot(slide, x, y, d, color):
    sp = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(x), Inches(y), Inches(d), Inches(d))
    sp.shadow.inherit = False
    set_fill(sp, color, 100)
    return sp


def arrow_down(slide, cx, y, h, color=CYAN):
    add_shape(slide, cx - 0.18, y, 0.36, h, MSO_SHAPE.DOWN_ARROW, fill=color, opacity=90)


def arrow_right(slide, x, cy, w, color=CYAN):
    add_shape(slide, x, cy - 0.18, w, 0.36, MSO_SHAPE.RIGHT_ARROW, fill=color, opacity=90)


# ----------------------------------------------------------------------------
# Build presentation
# ----------------------------------------------------------------------------
prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)
blank = prs.slide_layouts[6]

BG = os.path.join(ASSETS, "bg_commandcenter.png")
HERO = os.path.join(ASSETS, "hero_title.png")

# ============================================================================
# SLIDE 1 -- TITLE
# ============================================================================
s = prs.slides.add_slide(blank)
add_bg(s, HERO)
# dark overlay band for title legibility
add_shape(s, 0, 2.75, 13.333, 4.75, fill=DEEP, opacity=72)
add_shape(s, 0, 2.75, 13.333, 0.03, fill=CYAN, opacity=60)
# top-left prototype tag
add_shape(s, 0.85, 0.55, 0.05, 0.34, fill=CYAN)
add_text(s, 1.05, 0.57, 9.0, 0.34, "HACKATHON PROTOTYPE  ·  ELECTRICITY GRID OPERATIONS", 13, CYAN, True, font=MONO)

# Title
add_text(s, 1.0, 3.15, 11.6, 1.1, "AI EMERGENCY COMMANDER", 44, TEXT, True, align=PP_ALIGN.LEFT)
# Tagline
add_text(s, 1.02, 4.18, 11.6, 0.55, "PREDICT.  SIMULATE.  RECOMMEND.", 22, CYAN, True, font=MONO)
# subtitle
add_text(s, 1.02, 4.92, 10.5, 0.55,
         "AI-powered emergency decision support for critical infrastructure.",
         17, MUTED, False)

# three mini capability chips
chips = ["MONITOR", "PREDICT", "SIMULATE", "RECOMMEND"]
cx = 1.02
for ch in chips:
    w = 0.34 + 0.14 * len(ch)
    add_shape(s, cx, 5.55, w, 0.42, fill=PANEL, opacity=85, line=CYAN_DIM, line_pt=1.0)
    add_text(s, cx, 5.55, w, 0.42, ch, 12, CYAN, True, align=PP_ALIGN.CENTER, font=MONO, anchor=MSO_ANCHOR.MIDDLE)
    cx += w + 0.18

# footer disclaimer
add_text(s, 1.02, 6.55, 11.3, 0.4,
         "A simulated decision-support system for operators — not connected to any real power grid.",
         12, MUTED, False, font=MONO)

# ============================================================================
# SLIDE 2 -- PROBLEM
# ============================================================================
s = prs.slides.add_slide(blank)
add_bg(s, BG)
header(s, "02 · PROBLEM", "Failures don't stop at one component")
footer(s, 2)

add_text(s, 0.85, 1.95, 6.0, 1.0,
         "A normal monitoring system shows an alarm.  An operator still needs to know what comes next.",
         17, TEXT, False, line_spacing=1.15)

# three question cards
qs = [
    ("WHAT MAY FAIL NEXT?", "The next weak point in the network", RED),
    ("HOW WILL IT SPREAD?", "The cascade path through the grid", AMBER),
    ("WHAT ACTION FIRST?", "The safest response to take", CYAN),
]
qx = 0.85
for i, (q, d, col) in enumerate(qs):
    w = 3.7
    add_shape(s, qx, 3.2, w, 3.05, fill=PANEL, opacity=90, line=col, line_pt=1.25)
    add_shape(s, qx, 3.2, w, 0.055, fill=col)
    dot(s, qx + 0.32, 3.62, 0.16, col)
    add_text(s, qx + 0.62, 3.52, w - 0.9, 0.5, q, 15, col, True, font=MONO)
    add_text(s, qx + 0.32, 4.15, w - 0.64, 1.4, d, 14, TEXT, False, line_spacing=1.2)
    add_text(s, qx + 0.32, 5.4, w - 0.64, 0.6, f"0{i+1}", 26, col, True, font=MONO, opacity_hint=None) if False else None
    add_text(s, qx + 0.32, 5.55, 1.2, 0.5, f"0{i+1}", 22, col, True, font=MONO)
    qx += w + 0.26

# ============================================================================
# SLIDE 3 -- OUR SOLUTION
# ============================================================================
s = prs.slides.add_slide(blank)
add_bg(s, BG)
header(s, "03 · OUR SOLUTION", "The AI Emergency Commander")

add_text(s, 0.85, 1.92, 11.6, 0.5,
         "An AI-based decision-support copilot that watches the grid and thinks ahead for the operator.",
         17, TEXT, False)

caps = [
    ("MONITORS", "the power network", CYAN),
    ("DETECTS", "abnormal conditions", AMBER),
    ("PREDICTS", "failure risk", RED),
    ("SIMULATES", "cascading failures", AMBER),
    ("TESTS", "possible actions", CYAN),
    ("RECOMMENDS", "the safest action", GREEN),
]
gw = 3.7
gy = 2.75
gx = 0.85
for i, (verb, obj, col) in enumerate(caps):
    col_i = i % 3
    row = i // 3
    x = 0.85 + col_i * (gw + 0.26)
    y = gy + row * (1.7)
    add_shape(s, x, y, gw, 1.45, fill=PANEL, opacity=90, line=CYAN_DIM, line_pt=1.0)
    add_shape(s, x, y, 0.07, 1.45, fill=col)
    add_text(s, x + 0.28, y + 0.22, gw - 0.5, 0.45, verb, 17, col, True, font=MONO)
    add_text(s, x + 0.28, y + 0.72, gw - 0.5, 0.5, obj, 14, TEXT, False)

# ============================================================================
# SLIDE 4 -- HOW IT WORKS
# ============================================================================
s = prs.slides.add_slide(blank)
add_bg(s, BG)
header(s, "04 · HOW IT WORKS", "From grid data to a recommended action")
footer(s, 4)

flow = [
    ("Power Grid Data", CYAN),
    ("Danger Detection", AMBER),
    ("Failure Prediction", RED),
    ("Cascade Simulation", AMBER),
    ("What-If Analysis", BLUE),
    ("AI Recommendation", GREEN),
]
top = 2.0
node_h = 0.62
gap = 0.22
node_w = 6.6
node_x = 0.85
for i, (label, col) in enumerate(flow):
    y = top + i * (node_h + gap)
    add_shape(s, node_x, y, node_w, node_h, fill=PANEL, opacity=90, line=col, line_pt=1.25)
    add_shape(s, node_x, y, 0.06, node_h, fill=col)
    dot(s, node_x + 0.26, y + node_h/2 - 0.09, 0.18, col)
    add_text(s, node_x + 0.62, y, 5.0, node_h, label, 15, TEXT, True, anchor=MSO_ANCHOR.MIDDLE)
    if i < len(flow) - 1:
        arrow_down(s, node_x + 0.35, y + node_h + 0.01, 0.2, CYAN_DIM)

# right side: live-loop annotation
add_shape(s, 8.05, 2.0, 4.4, node_h * 6 + gap * 5, fill=PANEL, opacity=65, line=CYAN_DIM, line_pt=1.0, line_opacity=60)
add_text(s, 8.4, 2.3, 3.7, 0.5, "CLOSED LOOP", 14, CYAN, True, font=MONO)
add_text(s, 8.4, 2.85, 3.7, 3.8,
         "Each step feeds the next, and the recommendation loops back into the next what-if run.",
         14, MUTED, False, line_spacing=1.25)
for j in range(4):
    add_text(s, 8.4, 4.55 + j*0.55, 3.7, 0.4, f"ITERATION {j+1}", 12, CYAN_DIM, True, font=MONO)

# ============================================================================
# SLIDE 5 -- EXAMPLE EMERGENCY (cascade)
# ============================================================================
s = prs.slides.add_slide(blank)
add_bg(s, BG)
header(s, "05 · EXAMPLE EMERGENCY", "A failure spreads through the grid")
footer(s, 5)

# risk meter
add_text(s, 0.85, 1.9, 4.0, 0.4, "THREAT ESCALATION", 12, MUTED, True, font=MONO)
add_shape(s, 0.85, 2.32, 11.6, 0.12, fill=PANEL, line=CYAN_DIM, line_pt=0.75)

cascade = [
    ("Transformer A", "Overheated + overloaded", AMBER),
    ("Transformer A", "May fail", RED),
    ("Power shifts", "to Transformer B", CYAN),
    ("Transformer B", "Becomes overloaded", AMBER),
    ("Hospital power", "Supply at risk", RED),
]
cw = 2.16
cx = 0.85
cy = 2.95
ch = 1.95
prev_risk = 0.0
for i, (name, sub, col) in enumerate(cascade):
    # severity bar grows
    sev = 0.18 + i * 0.17
    risk_col = [CYAN, AMBER, AMBER, RED, RED][i]
    add_shape(s, cx + cw/2 - sev/2, 2.3, sev, 0.12, fill=risk_col)
    # card
    add_shape(s, cx, cy, cw, ch, fill=PANEL, opacity=92, line=col, line_pt=1.5)
    add_shape(s, cx, cy, cw, 0.07, fill=col)
    # step chip
    add_text(s, cx + 0.14, cy + 0.2, cw - 0.28, 0.35, f"STEP {i+1}", 11, col, True, font=MONO)
    # icon circle
    ico = MSO_SHAPE.LIGHTNING_BOLT if i in (0, 1, 3) else MSO_SHAPE.RIGHT_ARROW
    dot(s, cx + cw/2 - 0.28, cy + 0.72, 0.56, col)
    lbl = "⚡" if i in (0, 1, 3) else "▸"
    add_text(s, cx + cw/2 - 0.28, cy + 0.72, 0.56, 0.56, lbl, 22, DEEP, True, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    add_text(s, cx + 0.14, cy + 1.36, cw - 0.28, 0.3, name, 12.5, TEXT, True, align=PP_ALIGN.CENTER)
    add_text(s, cx + 0.14, cy + 1.63, cw - 0.28, 0.3, sub, 11, col, False, align=PP_ALIGN.CENTER)
    # intermediate pulse dot for spread
    if i == 1:
        dot(s, cx + cw + 0.06, cy + ch/2 - 0.04, 0.12, RED)
    cx += cw + 0.26

# bottom outcome strip
add_shape(s, 0.85, 5.45, 11.6, 0.95, fill=PANEL, opacity=92, line=RED, line_pt=1.0)
dot(s, 1.15, 5.82, 0.16, RED)
add_text(s, 1.45, 5.65, 10.8, 0.55,
         "A single overloaded component can pull down critical services — the operator must see the whole cascade, not one alarm.",
         14, TEXT, False)

# ============================================================================
# SLIDE 6 -- WHAT-IF SIMULATOR  (MAIN HIGHLIGHT)
# ============================================================================
s = prs.slides.add_slide(blank)
add_bg(s, BG)
header(s, "06 · WHAT-IF SIMULATOR", "Two futures. The AI picks the safer one.")
footer(s, 6)

# --- LEFT PANEL: DO NOTHING ---
lx, lw = 0.85, 5.6
ly, lh = 2.15, 4.35
add_shape(s, lx, ly, lw, lh, fill=PANEL, opacity=92, line=RED, line_pt=1.25)
add_shape(s, lx, ly, lw, 0.62, fill=RED, opacity=85)
add_text(s, lx + 0.3, ly + 0.14, lw - 0.6, 0.4, "OPTION 1  —  DO NOTHING", 15, DEEP, True, font=MONO)
add_text(s, lx + 0.3, ly + 0.78, lw - 0.6, 0.4, "Failure spreads", 16, RED, True)

out_l = [
    ("Failure spreads through the network", RED),
    ("Higher risk across the grid", RED),
    ("Hospital power may be affected", RED),
]
oy = ly + 1.4
for label, col in out_l:
    add_shape(s, lx + 0.3, oy, lw - 0.6, 0.62, fill=PANEL_2, line=col, line_pt=0.9)
    add_shape(s, lx + 0.3, oy, 0.32, 0.62, fill=col)
    add_text(s, lx + 0.78, oy, lw - 1.0, 0.62, label, 13, TEXT, True, anchor=MSO_ANCHOR.MIDDLE)
    oy += 0.78
# risk gauge (high)
add_text(s, lx + 0.3, ly + 3.85, 2.0, 0.3, "RISK", 12, MUTED, True, font=MONO)
add_shape(s, lx + 0.3, ly + 4.15, lw - 0.6, 0.16, fill=PANEL_2, line=RED, line_pt=0.75)
add_shape(s, lx + 0.3, ly + 4.15, (lw - 0.6) * 0.9, 0.16, fill=RED)

# --- RIGHT PANEL: REDUCE LOAD (highlight) ---
rx, rw = 6.9, 5.6
add_shape(s, rx - 0.12, ly - 0.12, rw + 0.24, lh + 0.24, fill=None, line=CYAN, line_pt=2.0, line_opacity=100)
add_shape(s, rx, ly, rw, lh, fill=RGBColor(0x0B, 0x1C, 0x2A), opacity=95, line=CYAN, line_pt=1.25)
add_shape(s, rx, ly, rw, 0.62, fill=GREEN, opacity=90)
add_text(s, rx + 0.3, ly + 0.14, rw - 0.6, 0.4, "OPTION 2  —  REDUCE NON-CRITICAL LOAD", 15, DEEP, True, font=MONO)
add_text(s, rx + 0.3, ly + 0.78, rw - 0.6, 0.4, "Cascade prevented", 16, GREEN, True)

out_r = [
    ("Transformer stabilizes", GREEN),
    ("Cascade is prevented", GREEN),
    ("Hospital stays protected", GREEN),
]
oy = ly + 1.4
for label, col in out_r:
    add_shape(s, rx + 0.3, oy, rw - 0.6, 0.62, fill=PANEL_2, line=col, line_pt=0.9)
    add_shape(s, rx + 0.3, oy, 0.32, 0.62, fill=col)
    add_text(s, rx + 0.78, oy, rw - 1.0, 0.62, label, 13, TEXT, True, anchor=MSO_ANCHOR.MIDDLE)
    oy += 0.78
add_text(s, rx + 0.3, ly + 3.85, 2.0, 0.3, "RISK", 12, MUTED, True, font=MONO)
add_shape(s, rx + 0.3, ly + 4.15, rw - 0.6, 0.16, fill=PANEL_2, line=GREEN, line_pt=0.75)
add_shape(s, rx + 0.3, ly + 4.15, (rw - 0.6) * 0.15, 0.16, fill=GREEN)

# --- VS badge ---
add_shape(s, 6.29, 3.85, 0.75, 0.75, MSO_SHAPE.OVAL, fill=CYAN, line=DEEP, line_pt=1.0)
add_text(s, 6.29, 3.85, 0.75, 0.75, "VS", 15, DEEP, True, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, font=MONO)

# --- bottom callout ---
add_shape(s, 0.85, 6.55, 11.6, 0.5, fill=PANEL, opacity=85, line=CYAN_DIM, line_pt=1.0)
add_text(s, 1.15, 6.6, 11.0, 0.4,
         "The simulator tests both actions on a digital model and rates the safest outcome.",
         13, MUTED, False, font=MONO, anchor=MSO_ANCHOR.MIDDLE)

# ============================================================================
# SLIDE 7 -- AI RECOMMENDATION
# ============================================================================
s = prs.slides.add_slide(blank)
add_bg(s, BG)
header(s, "07 · AI RECOMMENDATION", "The safest action, stated clearly")
footer(s, 7)

# main action card
add_shape(s, 0.85, 1.95, 11.6, 2.5, fill=PANEL, opacity=93, line=GREEN, line_pt=1.5)
add_shape(s, 0.85, 1.95, 0.1, 2.5, fill=GREEN)
add_text(s, 1.25, 2.2, 5.4, 0.4, "RECOMMENDED ACTION", 13, GREEN, True, font=MONO)
add_text(s, 1.25, 2.6, 5.5, 0.95,
         "Reduce non-critical\nZone B load",
         27, TEXT, True, line_spacing=1.05)
add_text(s, 1.25, 3.55, 5.5, 0.7,
         "by 15%", 30, GREEN, True)

# gauge illustration on right
gz = 7.15
add_text(s, gz, 2.2, 4.9, 0.4, "ZONE B LOAD", 12, MUTED, True, font=MONO)
add_shape(s, gz, 2.85, 4.9, 0.22, fill=PANEL_2, line=CYAN_DIM, line_pt=0.75)
add_shape(s, gz, 2.85, 4.9, 0.22, fill=AMBER)
add_shape(s, gz, 3.2, 4.9 * 0.85, 0.22, fill=GREEN)
add_text(s, gz, 3.5, 4.9, 0.35, "100%  →  85%", 13, GREEN, True, font=MONO, align=PP_ALIGN.CENTER)

# reason block
add_shape(s, 0.85, 4.75, 11.6, 1.55, fill=PANEL, opacity=90, line=CYAN, line_pt=1.0)
add_shape(s, 0.85, 4.75, 0.1, 1.55, fill=CYAN)
add_text(s, 1.25, 4.95, 4.0, 0.35, "WHY", 13, CYAN, True, font=MONO)
add_text(s, 1.25, 5.32, 10.7, 0.9,
         "Reduces transformer overload and lowers the risk of a cascade — protecting the critical hospital feed.",
         15, TEXT, False, line_spacing=1.2)

# ============================================================================
# SLIDE 8 -- SYSTEM ARCHITECTURE
# ============================================================================
s = prs.slides.add_slide(blank)
add_bg(s, BG)
header(s, "08 · SYSTEM ARCHITECTURE", "A digital twin that thinks ahead")
footer(s, 8)

arch = [
    ("Simulated Sensors", "grid telemetry", CYAN),
    ("Monitoring Engine", "anomaly detection", CYAN),
    ("Prediction Engine", "failure risk", AMBER),
    ("Digital Power-Grid Model", "the digital twin", BLUE),
    ("What-If Simulator", "action testing", CYAN),
    ("AI Decision Engine", "safest recommendation", GREEN),
    ("Emergency Dashboard", "operator interface", GREEN),
]
# centered vertical stack with downward flow (top = sensors, bottom = dashboard)
top_y = 1.95
layer_h = 0.6
gap = 0.1
bar_w = 8.6
block_x = 2.5   # left inset of the block
arrow_cx = block_x + 0.38
for i, (name, sub, col) in enumerate(arch):
    y = top_y + i * (layer_h + gap)
    add_shape(s, block_x, y, bar_w, layer_h, fill=PANEL, opacity=90, line=col, line_pt=1.1)
    add_shape(s, block_x, y, 0.06, layer_h, fill=col)
    dot(s, block_x + 0.28, y + layer_h/2 - 0.09, 0.18, col)
    add_text(s, block_x + 0.62, y, 5.6, layer_h, name, 13.5, TEXT, True, anchor=MSO_ANCHOR.MIDDLE)
    add_text(s, block_x + 6.2, y, 2.2, layer_h, sub, 11.5, MUTED, False, align=PP_ALIGN.RIGHT, anchor=MSO_ANCHOR.MIDDLE, font=MONO)
# downward arrows between layers
for i in range(len(arch) - 1):
    y0 = top_y + i * (layer_h + gap) + layer_h
    arrow_down(s, arrow_cx, y0 + 0.005, gap - 0.005, CYAN_DIM)
# left flow axis (downward DATA direction)
add_shape(s, 0.62, 2.35, 0.26, 4.2, MSO_SHAPE.DOWN_ARROW, fill=CYAN, opacity=85)
add_text(s, 0.35, 2.0, 0.8, 0.3, "DATA", 10, CYAN, True, align=PP_ALIGN.CENTER, font=MONO)

# ============================================================================
# SLIDE 9 -- INNOVATION & IMPACT
# ============================================================================
s = prs.slides.add_slide(blank)
add_bg(s, BG)
header(s, "09 · INNOVATION & IMPACT", "Beyond detection — to safer decisions")
footer(s, 9)

# innovation panel
add_shape(s, 0.85, 1.95, 5.6, 4.6, fill=PANEL, opacity=92, line=CYAN, line_pt=1.25)
add_shape(s, 0.85, 1.95, 5.6, 0.06, fill=CYAN)
add_text(s, 1.15, 2.25, 5.0, 0.4, "INNOVATION", 13, CYAN, True, font=MONO)
add_text(s, 1.15, 2.7, 5.0, 1.4,
         "Not just detection.",
         24, TEXT, True, line_spacing=1.1)
add_text(s, 1.15, 3.4, 5.0, 2.6,
         "The system predicts consequences, compares possible actions, and recommends the safer response.",
         17, MUTED, False, line_spacing=1.25)
add_shape(s, 1.15, 5.5, 4.6, 0.012, fill=CYAN_DIM)
add_text(s, 1.15, 5.65, 4.8, 0.7,
         "A human remains in command.",
         13, TEXT, True, font=MONO)

# impact panel
impacts = [
    ("Faster emergency response", CYAN),
    ("Reduced infrastructure downtime", CYAN),
    ("Protection of critical services", GREEN),
    ("Better decisions for operators", GREEN),
    ("Scalable to other infrastructure", BLUE),
]
add_shape(s, 6.85, 1.95, 5.6, 4.6, fill=PANEL, opacity=92, line=GREEN, line_pt=1.25)
add_shape(s, 6.85, 1.95, 5.6, 0.06, fill=GREEN)
add_text(s, 7.15, 2.25, 5.0, 0.4, "IMPACT", 13, GREEN, True, font=MONO)
iy = 2.95
for label, col in impacts:
    dot(s, 7.2, iy + 0.13, 0.16, col)
    add_text(s, 7.5, iy, 4.6, 0.45, label, 15, TEXT, True)
    iy += 0.68

# ============================================================================
# SLIDE 10 -- FINAL
# ============================================================================
s = prs.slides.add_slide(blank)
add_bg(s, HERO)
add_shape(s, 0, 2.35, 13.333, 5.15, fill=DEEP, opacity=78)
add_shape(s, 0, 2.35, 13.333, 0.03, fill=CYAN, opacity=60)

add_text(s, 1.0, 2.75, 11.3, 0.7, "DON'T JUST DETECT THE FAILURE.", 24, MUTED, True)
add_text(s, 1.0, 3.55, 11.3, 1.2, "PREDICT IT.", 46, TEXT, True)
add_text(s, 1.0, 4.55, 11.3, 1.2, "SIMULATE IT.", 46, CYAN, True)
add_text(s, 1.0, 5.55, 11.3, 1.2, "PREVENT THE CASCADE.", 46, GREEN, True)

add_shape(s, 1.0, 6.75, 2.0, 0.03, fill=CYAN)
add_text(s, 1.0, 6.9, 11.3, 0.4,
         "AI EMERGENCY COMMANDER  ·  SIMULATED PROTOTYPE  ·  HUMAN DECISION-SUPPORT SYSTEM",
         12, MUTED, False, font=MONO)

# ----------------------------------------------------------------------------
out = "/home/user/Arduino-learning-journey../AI_Emergency_Commander.pptx"
prs.save(out)
print("Saved:", out, "| slides:", len(prs.slides.__iter__.__self__._sldIdLst))
