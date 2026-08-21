from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.shapes import MSO_SHAPE, MSO_CONNECTOR
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.dml.color import RGBColor
from pptx.enum.dml import MSO_THEME_COLOR
from pptx.enum.shapes import PP_PLACEHOLDER
from pathlib import Path
from PIL import Image
import math

# -----------------------------------------------------------------------------
# Autonomous Edge-AI Leak Intelligence — editable, vector-first presentation
# -----------------------------------------------------------------------------

OUT = Path(__file__).resolve().parents[1] / "deliverables" / "autonomous_edge_ai_leak_intelligence.pptx"
OUT.parent.mkdir(parents=True, exist_ok=True)

prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)
BLANK = prs.slide_layouts[6]

# Palette: midnight industrial with cyan sensing and amber action accents.
BG = "0A0F15"
PANEL = "111A24"
PANEL_2 = "172330"
PANEL_3 = "1C2B38"
WHITE = "F4F7FA"
MUTED = "9AA9B8"
MUTED_2 = "6F8292"
GRID = "20313D"
CYAN = "48E0D2"
CYAN_DARK = "1E8F91"
BLUE = "5EA7FF"
AMBER = "F4B860"
ORANGE = "F28D52"
RED = "FF6B6B"
GREEN = "7FE0A4"
PURPLE = "B796FF"

FONT = "Aptos"
FONT_DISPLAY = "Aptos Display"
ASSET_DIR = Path(__file__).resolve().parent / "assets"
HERO_IMAGE = ASSET_DIR / "hero_robot_pipeline.png"
LEAK_IMAGE = ASSET_DIR / "pipeline_leak_acoustic.png"
ROBOT_IMAGE = ASSET_DIR / "inspection_robot_closeup.png"


def rgb(hexstr):
    return RGBColor.from_string(hexstr)


def set_fill(shape, color, transparency=0):
    shape.fill.solid()
    shape.fill.fore_color.rgb = rgb(color)
    if transparency:
        shape.fill.transparency = transparency


def set_line(shape, color, width=1.0, transparency=0):
    shape.line.color.rgb = rgb(color)
    shape.line.width = Pt(width)
    if transparency:
        shape.line.transparency = transparency


def no_line(shape):
    shape.line.fill.background()


def rect(slide, x, y, w, h, fill=PANEL, radius=False, line=None, lw=1.0, transparency=0):
    shp = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE if radius else MSO_SHAPE.RECTANGLE,
                                 Inches(x), Inches(y), Inches(w), Inches(h))
    set_fill(shp, fill, transparency)
    if line:
        set_line(shp, line, lw)
    else:
        no_line(shp)
    return shp


def ellipse(slide, x, y, w, h, fill=PANEL, line=None, lw=1.0, transparency=0):
    shp = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(x), Inches(y), Inches(w), Inches(h))
    set_fill(shp, fill, transparency)
    if line:
        set_line(shp, line, lw)
    else:
        no_line(shp)
    return shp


def line(slide, x1, y1, x2, y2, color=GRID, width=1.0, dash=None):
    shp = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(x1), Inches(y1), Inches(x2), Inches(y2))
    set_line(shp, color, width)
    if dash:
        try:
            shp.line.dash_style = dash
        except Exception:
            pass
    return shp


def text(slide, x, y, w, h, value, size=12, color=WHITE, bold=False,
         font=FONT, align=PP_ALIGN.LEFT, valign=MSO_ANCHOR.TOP,
         margin=0.04, italic=False, tracking=None):
    box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = box.text_frame
    tf.clear()
    tf.word_wrap = True
    tf.margin_left = Inches(margin)
    tf.margin_right = Inches(margin)
    tf.margin_top = Inches(margin)
    tf.margin_bottom = Inches(margin)
    tf.vertical_anchor = valign
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = value
    run.font.name = font
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.italic = italic
    run.font.color.rgb = rgb(color)
    if tracking is not None:
        try:
            run.font.kerning = Pt(tracking)
        except Exception:
            pass
    return box


def rich_text(slide, x, y, w, h, parts, size=12, align=PP_ALIGN.LEFT, valign=MSO_ANCHOR.TOP,
              margin=0.04, font=FONT):
    box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = box.text_frame
    tf.clear(); tf.word_wrap = True
    tf.margin_left = Inches(margin); tf.margin_right = Inches(margin)
    tf.margin_top = Inches(margin); tf.margin_bottom = Inches(margin)
    tf.vertical_anchor = valign
    p = tf.paragraphs[0]; p.alignment = align
    for part in parts:
        run = p.add_run(); run.text = part.get('text', '')
        run.font.name = part.get('font', font)
        run.font.size = Pt(part.get('size', size))
        run.font.bold = part.get('bold', False)
        run.font.italic = part.get('italic', False)
        run.font.color.rgb = rgb(part.get('color', WHITE))
    return box


def pill(slide, x, y, w, label, fill=PANEL_2, color=MUTED, line_color=None, size=7.5):
    p = rect(slide, x, y, w, 0.27, fill, radius=True, line=line_color, lw=0.7)
    text(slide, x, y+0.005, w, 0.25, label.upper(), size=size, color=color, bold=True,
         align=PP_ALIGN.CENTER, valign=MSO_ANCHOR.MIDDLE, margin=0.01)
    return p


def dot(slide, x, y, d=0.08, fill=CYAN, line_color=None):
    return ellipse(slide, x, y, d, d, fill, line=line_color or fill)


def chevron(slide, x, y, w=0.22, h=0.22, color=CYAN):
    s = slide.shapes.add_shape(MSO_SHAPE.CHEVRON, Inches(x), Inches(y), Inches(w), Inches(h))
    set_fill(s, color); no_line(s); return s


def label(slide, x, y, value, color=MUTED_2, size=7.5, w=2.5):
    return text(slide, x, y, w, 0.18, value.upper(), size=size, color=color, bold=True, tracking=0.2)


def picture_cover(slide, path, x, y, w, h):
    """Place a picture in a frame without distorting its aspect ratio."""
    with Image.open(path) as im:
        iw, ih = im.size
    target = w / h
    aspect = iw / ih
    pic = slide.shapes.add_picture(str(path), Inches(x), Inches(y), Inches(w), Inches(h))
    if aspect > target:
        crop = (1.0 - target / aspect) / 2.0
        pic.crop_left = crop
        pic.crop_right = crop
    elif aspect < target:
        crop = (1.0 - aspect / target) / 2.0
        pic.crop_top = crop
        pic.crop_bottom = crop
    return pic


def slide_bg(slide, kicker, title, subtitle=None, number=None, accent=CYAN):
    bg = slide.background.fill
    bg.solid(); bg.fore_color.rgb = rgb(BG)
    # quiet structural grid
    for xx in [0.45, 3.0, 6.0, 9.0, 12.55]:
        line(slide, xx, 0.25, xx, 7.24, GRID, 0.45)
    for yy in [0.35, 1.26, 6.94, 7.22]:
        line(slide, 0.35, yy, 12.98, yy, GRID, 0.45)
    # header
    label(slide, 0.64, 0.36, kicker, accent, 7.5, 6.6)
    if number is not None:
        text(slide, 11.75, 0.34, 0.95, 0.2, f"{number:02d} / 12", size=8, color=MUTED_2,
             bold=True, align=PP_ALIGN.RIGHT, margin=0.0)
    text(slide, 0.60, 0.63, 11.0, 0.48, title, size=25, color=WHITE, bold=True,
         font=FONT_DISPLAY, margin=0.0)
    if subtitle:
        text(slide, 0.64, 1.11, 11.7, 0.25, subtitle, size=9.5, color=MUTED, margin=0.0)
    # footer
    text(slide, 0.64, 7.24, 6.0, 0.15, "AUTONOMOUS LEAK INTELLIGENCE  /  HACKATHON MVP", size=6.5,
         color=MUTED_2, bold=True, margin=0.0)
    text(slide, 10.15, 7.24, 2.55, 0.15, "PERIODIC PATROL · LOCAL INTELLIGENCE", size=6.5,
         color=MUTED_2, bold=True, align=PP_ALIGN.RIGHT, margin=0.0)


def add_notes(slide, value):
    slide.notes_slide.notes_text_frame.text = value


def icon_robot(slide, x, y, s=1.0, color=CYAN, accent=AMBER):
    # compact line robot, designed to survive small sizes
    rect(slide, x+0.16*s, y+0.14*s, 0.68*s, 0.48*s, PANEL_3, radius=True, line=color, lw=1.3)
    rect(slide, x+0.27*s, y+0.24*s, 0.16*s, 0.13*s, color, radius=True)
    rect(slide, x+0.57*s, y+0.24*s, 0.16*s, 0.13*s, color, radius=True)
    line(slide, x+0.50*s, y+0.02*s, x+0.50*s, y+0.14*s, color, 1.2)
    dot(slide, x+0.46*s, y-0.03*s, 0.08*s, accent)
    line(slide, x+0.10*s, y+0.70*s, x+0.90*s, y+0.70*s, color, 1.6)
    ellipse(slide, x+0.19*s, y+0.58*s, 0.18*s, 0.18*s, BG, line=color, lw=1.2)
    ellipse(slide, x+0.65*s, y+0.58*s, 0.18*s, 0.18*s, BG, line=color, lw=1.2)
    line(slide, x+0.50*s, y+0.62*s, x+0.50*s, y+0.84*s, accent, 1.0)


def icon_wave(slide, x, y, w=1.0, h=0.5, color=CYAN, thin=1.2):
    pts = [(0.0, 0.55), (0.12, 0.55), (0.22, 0.17), (0.31, 0.85), (0.42, 0.42), (0.52, 0.60), (0.63, 0.28), (0.74, 0.70), (0.84, 0.45), (1.0, 0.45)]
    for (a, b), (c, d) in zip(pts[:-1], pts[1:]):
        line(slide, x+a*w, y+b*h, x+c*w, y+d*h, color, thin)


def icon_signal(slide, x, y, s=1.0, color=CYAN):
    line(slide, x+0.50*s, y+0.54*s, x+0.50*s, y+0.20*s, color, 1.4)
    line(slide, x+0.50*s, y+0.20*s, x+0.22*s, y+0.06*s, color, 1.4)
    line(slide, x+0.50*s, y+0.20*s, x+0.78*s, y+0.06*s, color, 1.4)
    line(slide, x+0.50*s, y+0.20*s, x+0.50*s, y+0.06*s, color, 1.4)
    for r, a in [(0.19, 0.0), (0.34, 0.05), (0.49, 0.12)]:
        # small arc approximation
        line(slide, x+(0.50-r)*s, y+(0.62-r*0.25)*s, x+(0.50-r*0.55)*s, y+(0.72-r*0.1)*s, color, 1.0)
        line(slide, x+(0.50+r*0.55)*s, y+(0.72-r*0.1)*s, x+(0.50+r)*s, y+(0.62-r*0.25)*s, color, 1.0)


def icon_pin(slide, x, y, s=1.0, color=AMBER, filled=False):
    # pin body with circle head and stem
    ellipse(slide, x+0.18*s, y+0.02*s, 0.32*s, 0.32*s, color if filled else BG, line=color, lw=1.1)
    line(slide, x+0.34*s, y+0.34*s, x+0.34*s, y+0.64*s, color, 1.2)
    line(slide, x+0.34*s, y+0.64*s, x+0.20*s, y+0.46*s, color, 1.2)
    line(slide, x+0.34*s, y+0.64*s, x+0.48*s, y+0.46*s, color, 1.2)


def icon_battery(slide, x, y, w=0.78, h=0.42, color=GREEN):
    rect(slide, x, y, w, h, BG, radius=True, line=color, lw=1.1)
    rect(slide, x+w, y+h*0.29, 0.07, h*0.42, color, radius=True)
    rect(slide, x+0.08, y+0.08, w*0.35, h-0.16, color, radius=True)
    rect(slide, x+0.48, y+0.08, w*0.20, h-0.16, color, radius=True)


def icon_battery_small(slide, x, y, s=1.0, color=GREEN):
    icon_battery(slide, x, y, 0.68*s, 0.38*s, color)


def icon_brain(slide, x, y, s=1.0, color=PURPLE):
    ellipse(slide, x+0.18*s, y+0.10*s, 0.32*s, 0.38*s, BG, line=color, lw=1.0)
    ellipse(slide, x+0.43*s, y+0.10*s, 0.32*s, 0.38*s, BG, line=color, lw=1.0)
    line(slide, x+0.50*s, y+0.12*s, x+0.50*s, y+0.48*s, color, 1.0)
    line(slide, x+0.27*s, y+0.25*s, x+0.42*s, y+0.25*s, color, 0.9)
    line(slide, x+0.58*s, y+0.34*s, x+0.73*s, y+0.34*s, color, 0.9)
    dot(slide, x+0.30*s, y+0.14*s, 0.07*s, color)
    dot(slide, x+0.61*s, y+0.42*s, 0.07*s, color)


def icon_wave_small(slide, x, y, s=1.0, color=CYAN):
    icon_wave(slide, x, y, 0.76*s, 0.42*s, color, 1.0)


def icon_filter(slide, x, y, s=1.0, color=BLUE):
    line(slide, x+0.12*s, y+0.12*s, x+0.82*s, y+0.12*s, color, 1.4)
    line(slide, x+0.25*s, y+0.35*s, x+0.69*s, y+0.35*s, color, 1.4)
    line(slide, x+0.39*s, y+0.58*s, x+0.55*s, y+0.58*s, color, 1.4)


def icon_clock(slide, x, y, s=1.0, color=AMBER):
    ellipse(slide, x+0.08*s, y+0.08*s, 0.56*s, 0.56*s, BG, line=color, lw=1.1)
    line(slide, x+0.36*s, y+0.36*s, x+0.36*s, y+0.19*s, color, 1.2)
    line(slide, x+0.36*s, y+0.36*s, x+0.49*s, y+0.43*s, color, 1.2)
    dot(slide, x+0.32*s, y+0.32*s, 0.08*s, color)


def icon_check(slide, x, y, s=1.0, color=GREEN):
    line(slide, x+0.10*s, y+0.38*s, x+0.29*s, y+0.57*s, color, 1.6)
    line(slide, x+0.29*s, y+0.57*s, x+0.70*s, y+0.14*s, color, 1.6)


def icon_alert(slide, x, y, s=1.0, color=RED):
    # triangle-ish warning mark
    tri = slide.shapes.add_shape(MSO_SHAPE.ISOSCELES_TRIANGLE, Inches(x), Inches(y), Inches(0.72*s), Inches(0.64*s))
    set_fill(tri, BG); set_line(tri, color, 1.3)
    text(slide, x+0.22*s, y+0.18*s, 0.28*s, 0.30*s, "!", size=13*s, color=color, bold=True, align=PP_ALIGN.CENTER, margin=0.0)


def icon_route(slide, x, y, w=1.0, h=0.55, color=CYAN):
    # angular route with nodes
    pts = [(0.05,0.75),(0.24,0.25),(0.48,0.56),(0.68,0.12),(0.94,0.42)]
    for (a,b),(c,d) in zip(pts[:-1],pts[1:]):
        line(slide, x+a*w, y+b*h, x+c*w, y+d*h, color, 1.25)
    for a,b in pts:
        dot(slide, x+a*w-0.04, y+b*h-0.04, 0.08, color)


def icon_route_small(slide, x, y, s=1.0, color=CYAN):
    icon_route(slide, x, y, 0.76*s, 0.52*s, color)


def icon_wrench(slide, x, y, s=1.0, color=AMBER):
    # simple diagonal wrench
    line(slide, x+0.22*s, y+0.58*s, x+0.62*s, y+0.18*s, color, 2.4)
    ellipse(slide, x+0.52*s, y+0.04*s, 0.22*s, 0.22*s, BG, line=color, lw=1.5)
    line(slide, x+0.55*s, y+0.04*s, x+0.72*s, y+0.21*s, BG, 3.0)
    ellipse(slide, x+0.10*s, y+0.48*s, 0.23*s, 0.23*s, BG, line=color, lw=1.4)


def icon_cloud(slide, x, y, s=1.0, color=MUTED):
    ellipse(slide, x+0.12*s, y+0.25*s, 0.36*s, 0.27*s, BG, line=color, lw=1.0)
    ellipse(slide, x+0.35*s, y+0.12*s, 0.43*s, 0.40*s, BG, line=color, lw=1.0)
    ellipse(slide, x+0.67*s, y+0.25*s, 0.30*s, 0.27*s, BG, line=color, lw=1.0)
    line(slide, x+0.23*s, y+0.51*s, x+0.83*s, y+0.51*s, color, 1.0)


def icon_server(slide, x, y, w=0.72, h=0.62, color=BLUE):
    for i in range(3):
        rect(slide, x, y+i*0.22, w, 0.16, BG, radius=True, line=color, lw=0.9)
        dot(slide, x+0.08, y+i*0.22+0.055, 0.05, color)
        line(slide, x+0.24, y+i*0.22+0.08, x+w-0.08, y+i*0.22+0.08, color, 0.7)


def icon_server_small(slide, x, y, s=1.0, color=BLUE):
    icon_server(slide, x, y, 0.62*s, 0.54*s, color)


def small_card(slide, x, y, w, h, heading, body, accent=CYAN, icon=None, body_color=MUTED):
    rect(slide, x, y, w, h, PANEL, radius=True, line=GRID, lw=0.8)
    rect(slide, x, y, 0.05, h, accent, radius=True)
    if icon:
        icon(slide, x+0.20, y+0.20, 0.62, accent)
        text(slide, x+0.96, y+0.16, w-1.14, 0.25, heading, size=10.5, color=WHITE, bold=True, margin=0.0)
        text(slide, x+0.96, y+0.49, w-1.14, h-0.58, body, size=8.6, color=body_color, margin=0.0)
    else:
        text(slide, x+0.20, y+0.16, w-0.36, 0.25, heading, size=10.5, color=WHITE, bold=True, margin=0.0)
        text(slide, x+0.20, y+0.49, w-0.36, h-0.58, body, size=8.6, color=body_color, margin=0.0)


def workflow_node(slide, x, y, w, h, n, heading, sub, accent=CYAN, icon_fn=None):
    rect(slide, x, y, w, h, PANEL, radius=True, line=GRID, lw=0.8)
    ellipse(slide, x+0.17, y+0.17, 0.31, 0.31, accent, line=accent)
    text(slide, x+0.17, y+0.173, 0.31, 0.30, str(n), size=8.5, color=BG, bold=True, align=PP_ALIGN.CENTER, valign=MSO_ANCHOR.MIDDLE, margin=0.0)
    if icon_fn:
        icon_fn(slide, x+w-0.75, y+0.14, 0.58, accent)
    text(slide, x+0.17, y+0.62, w-0.28, 0.26, heading, size=9.2, color=WHITE, bold=True, margin=0.0)
    text(slide, x+0.17, y+0.91, w-0.28, h-0.97, sub, size=7.4, color=MUTED, margin=0.0)


def add_arrow(slide, x, y, w=0.20, color=CYAN):
    chev = slide.shapes.add_shape(MSO_SHAPE.CHEVRON, Inches(x), Inches(y), Inches(w), Inches(0.26))
    set_fill(chev, color); no_line(chev); return chev


def add_status_dot(slide, x, y, color, label_text, w=1.1):
    dot(slide, x, y+0.045, 0.08, color)
    text(slide, x+0.14, y, w, 0.16, label_text, size=7.2, color=MUTED, bold=True, margin=0.0)


def draw_pipe(slide, x, y, w, h=0.30, leak_at=0.68, pipe_color=BLUE, leak_color=AMBER):
    # horizontal pipeline and segment joints
    rect(slide, x, y, w, h, PANEL_3, radius=True, line=pipe_color, lw=1.0)
    for ratio in [0.20, 0.42, 0.64, 0.86]:
        line(slide, x+ratio*w, y+0.02, x+ratio*w, y+h-0.02, pipe_color, 0.8)
    lx = x + leak_at*w
    ellipse(slide, lx-0.05, y+h*0.30, 0.10, 0.10, leak_color, line=leak_color)
    # air jets
    for off, length in [(-0.13,0.28),(0.0,0.42),(0.13,0.25)]:
        line(slide, lx, y+h*0.5, lx+off, y-0.08-length*0.12, leak_color, 1.2)
    return lx


def draw_route(slide, x, y, w, h, active=0.66, color=CYAN, muted=GRID):
    pts = [(0.02,0.74),(0.16,0.22),(0.34,0.46),(0.52,0.16),(0.70,0.62),(0.88,0.32),(0.98,0.54)]
    for (a,b),(c,d) in zip(pts[:-1], pts[1:]):
        line(slide, x+a*w, y+b*h, x+c*w, y+d*h, muted, 2.4)
    # active trace until active ratio along points
    segs = []
    for i in range(len(pts)-1):
        mid = (i+0.5)/(len(pts)-1)
        if mid <= active:
            segs.append((pts[i],pts[i+1]))
    for (a,b),(c,d) in segs:
        line(slide, x+a*w, y+b*h, x+c*w, y+d*h, color, 3.0)
    for i,(a,b) in enumerate(pts):
        node_color = color if i/len(pts) <= active else muted
        dot(slide, x+a*w-0.055, y+b*h-0.055, 0.11, node_color)
    return [(x+a*w,y+b*h) for a,b in pts]


def draw_signal_panel(slide, x, y, w, h, event=True):
    rect(slide, x, y, w, h, PANEL, radius=True, line=GRID, lw=0.8)
    label(slide, x+0.20, y+0.16, "RAW ACOUSTIC WINDOW", MUTED_2, 7.0, w-0.4)
    # grid
    for i in range(1,5): line(slide, x+0.22, y+0.45+i*(h-0.64)/5, x+w-0.20, y+0.45+i*(h-0.64)/5, GRID, 0.5)
    icon_wave(slide, x+0.22, y+0.55, w-0.44, h-0.86, color=MUTED_2, thin=0.8)
    if event:
        pts = [(0.0,0.58),(0.08,0.58),(0.14,0.28),(0.19,0.84),(0.27,0.53),(0.33,0.58),(0.40,0.18),(0.46,0.86),(0.53,0.46),(0.63,0.57),(0.71,0.22),(0.76,0.74),(0.85,0.50),(0.93,0.58),(1.0,0.58)]
        for (a,b),(c,d) in zip(pts[:-1],pts[1:]):
            line(slide, x+0.22+a*(w-0.44), y+0.55+b*(h-0.86), x+0.22+c*(w-0.44), y+0.55+d*(h-0.86), CYAN, 1.5)
        rect(slide, x+w-1.26, y+0.19, 1.0, 0.22, CYAN_DARK, radius=True)
        text(slide, x+w-1.24, y+0.20, 0.96, 0.18, "EVENT WINDOW", size=6.7, color=CYAN, bold=True, align=PP_ALIGN.CENTER, margin=0.0)


def draw_robot_diagram(slide, x, y, scale=1.0):
    # floor line / route
    line(slide, x-0.05*scale, y+2.65*scale, x+3.18*scale, y+2.65*scale, CYAN_DARK, 2.0)
    line(slide, x+0.12*scale, y+2.71*scale, x+0.58*scale, y+2.71*scale, CYAN, 1.0)
    line(slide, x+0.78*scale, y+2.71*scale, x+1.26*scale, y+2.71*scale, CYAN, 1.0)
    line(slide, x+1.47*scale, y+2.71*scale, x+1.95*scale, y+2.71*scale, CYAN, 1.0)
    line(slide, x+2.16*scale, y+2.71*scale, x+2.64*scale, y+2.71*scale, CYAN, 1.0)
    # body
    rect(slide, x+0.62*scale, y+1.15*scale, 1.95*scale, 1.05*scale, PANEL_2, radius=True, line=CYAN, lw=1.3)
    rect(slide, x+0.85*scale, y+1.34*scale, 0.62*scale, 0.48*scale, BG, radius=True, line=BLUE, lw=1.0)
    line(slide, x+1.00*scale, y+1.48*scale, x+1.30*scale, y+1.48*scale, BLUE, 1.2)
    dot(slide, x+1.06*scale, y+1.57*scale, 0.07*scale, CYAN)
    dot(slide, x+1.22*scale, y+1.57*scale, 0.07*scale, AMBER)
    # battery
    icon_battery(slide, x+1.77*scale, y+1.48*scale, 0.52*scale, 0.27*scale, GREEN)
    # wheels
    ellipse(slide, x+0.82*scale, y+2.02*scale, 0.49*scale, 0.49*scale, BG, line=CYAN, lw=1.4)
    ellipse(slide, x+2.00*scale, y+2.02*scale, 0.49*scale, 0.49*scale, BG, line=CYAN, lw=1.4)
    dot(slide, x+1.00*scale, y+2.20*scale, 0.12*scale, MUTED_2)
    dot(slide, x+2.18*scale, y+2.20*scale, 0.12*scale, MUTED_2)
    # top sensor mast
    line(slide, x+1.58*scale, y+1.14*scale, x+1.58*scale, y+0.62*scale, CYAN, 1.4)
    rect(slide, x+1.30*scale, y+0.26*scale, 0.58*scale, 0.40*scale, PANEL_3, radius=True, line=AMBER, lw=1.2)
    icon_signal(slide, x+1.34*scale, y+0.31*scale, 0.50*scale, AMBER)
    text(slide, x+1.17*scale, y-0.02*scale, 1.1*scale, 0.18*scale, "ACOUSTIC SENSOR", size=7.0, color=AMBER, bold=True, align=PP_ALIGN.CENTER, margin=0.0)
    # labels and callouts
    line(slide, x+0.85*scale, y+1.34*scale, x+0.12*scale, y+0.84*scale, MUTED_2, 0.8)
    text(slide, x-0.10*scale, y+0.62*scale, 1.25*scale, 0.24*scale, "EDGE AI PROCESSOR", size=7.0, color=MUTED, bold=True, margin=0.0)
    line(slide, x+2.45*scale, y+1.53*scale, x+3.02*scale, y+1.07*scale, MUTED_2, 0.8)
    text(slide, x+2.72*scale, y+0.85*scale, 1.15*scale, 0.24*scale, "COMMS LINK", size=7.0, color=MUTED, bold=True, margin=0.0)
    # acoustic halo
    for d, tr in [(0.80, 75), (1.10, 85), (1.42, 92)]:
        ellipse(slide, x+1.58*scale-d*0.5*scale, y+0.46*scale-d*0.5*scale, d*scale, d*scale, BG, line=AMBER, lw=0.7, transparency=tr)


def draw_sparkline(slide, x, y, w, h, color=CYAN, points=None):
    pts = points or [0.55,0.60,0.42,0.48,0.29,0.34,0.18,0.24,0.12,0.19,0.11]
    for i in range(len(pts)-1):
        line(slide, x+i*w/(len(pts)-1), y+pts[i]*h, x+(i+1)*w/(len(pts)-1), y+pts[i+1]*h, color, 1.5)
    for i,p in enumerate(pts):
        dot(slide, x+i*w/(len(pts)-1)-0.025, y+p*h-0.025, 0.05, color)


def add_slide_number_badge(slide, n, x=0.64, y=6.84):
    ellipse(slide, x, y, 0.25, 0.25, CYAN, line=CYAN)
    text(slide, x, y+0.01, 0.25, 0.22, str(n), size=7.5, color=BG, bold=True, align=PP_ALIGN.CENTER, valign=MSO_ANCHOR.MIDDLE, margin=0.0)

# -----------------------------------------------------------------------------
# SLIDE 1 — COVER
# -----------------------------------------------------------------------------
slide = prs.slides.add_slide(BLANK)
slide_bg(slide, "INDUSTRIAL LEAK INTELLIGENCE  /  HACKATHON MVP", "", number=None)
# Clear the standard title area for the premium cover composition.
text(slide, 0.66, 0.58, 5.7, 0.22, "AUTONOMOUS INSPECTION SYSTEM", size=8.0, color=CYAN, bold=True, margin=0.0)
text(slide, 0.66, 1.06, 6.1, 1.82, "Autonomous Edge-AI\nRobotic System for\nCompressed-Air Leak Detection", size=28, color=WHITE, bold=True, font=FONT_DISPLAY, margin=0.0)
text(slide, 0.68, 3.16, 5.72, 0.52, "Automated Inspection  •  Intelligent Detection  •  Leak Localization  •  Maintenance Prioritization", size=11.2, color=MUTED, bold=False, margin=0.0)
pill(slide, 0.68, 4.18, 1.65, "PERIODIC PATROL", PANEL_2, CYAN, line_color=CYAN_DARK, size=7.3)
pill(slide, 2.47, 4.18, 1.55, "EDGE AI", PANEL_2, PURPLE, line_color=PURPLE, size=7.3)
pill(slide, 4.16, 4.18, 1.92, "ACTIONABLE ALERTS", PANEL_2, AMBER, line_color=AMBER, size=7.3)
text(slide, 0.68, 5.28, 5.5, 0.48, "A practical, scalable way to turn robot patrol data into maintenance decisions.", size=13, color=WHITE, bold=True, margin=0.0)
text(slide, 0.68, 6.55, 4.0, 0.24, "PROBLEM → GAP → SOLUTION → PROOF → SCALE", size=8.0, color=MUTED_2, bold=True, margin=0.0)
# cover visual right: pipeline + robot + signal halo
rect(slide, 6.70, 1.02, 5.92, 5.55, PANEL, radius=True, line=GRID, lw=0.9)
# generated hero image: the robot and leak relationship is immediately legible
picture_cover(slide, HERO_IMAGE, 6.70, 1.02, 5.92, 5.55)
# restrained overlays keep the image premium while preserving the visual story
rect(slide, 6.70, 5.76, 5.92, 0.81, BG, radius=False, transparency=22)
pill(slide, 7.02, 5.94, 1.28, "PATROL ONLINE", "14313A", CYAN, line_color=CYAN_DARK, size=6.6)
pill(slide, 11.02, 5.94, 1.28, "ACOUSTIC EVENT", "352B1F", AMBER, line_color=AMBER, size=6.4)
text(slide, 7.03, 6.37, 5.25, 0.18, "A route-aware robot for automated periodic inspection", size=8.2, color=WHITE, align=PP_ALIGN.CENTER, margin=0.0)
label(slide, 7.03, 1.34, "LIVE VISUAL CONCEPT", CYAN, 7.0, 2.3)
add_notes(slide, "Open by reframing the project. This is not a line-following robot looking for a clever demo; it is an automated industrial leak-intelligence workflow. The robot patrols a predefined route, senses acoustic signatures locally, and returns a maintenance-ready record with approximate location and priority. We will show a realistic MVP boundary first, then the path to scale. The promise is simple: turn manual leak hunting into repeatable, actionable inspection data.")

# -----------------------------------------------------------------------------
# SLIDE 2 — PROBLEM
# -----------------------------------------------------------------------------
slide = prs.slides.add_slide(BLANK)
slide_bg(slide, "01  /  THE PROBLEM", "Compressed air disappears before it creates value.", "A small acoustic event can represent wasted energy, unplanned effort, and a maintenance decision waiting to be made.", 2, AMBER)
# left narrative
pill(slide, 0.66, 1.62, 1.08, "WHY NOW", "352B1F", AMBER, line_color=AMBER, size=7.0)
text(slide, 0.66, 2.05, 3.38, 0.98, "Leaks are\nquietly expensive.", size=25, color=WHITE, bold=True, font=FONT_DISPLAY, margin=0.0)
text(slide, 0.68, 3.20, 3.20, 0.72, "Industrial teams spend time finding a problem that is often hard to hear, hard to locate, and hard to prioritize.", size=11.2, color=MUTED, margin=0.0)
# three vertical consequence tags
for i,(head,body,accent) in enumerate([
    ("ENERGY WASTED", "Compressed air escapes before it does useful work.", AMBER),
    ("FINANCIAL LOSS", "Wasted energy becomes recurring operating cost.", BLUE),
    ("MANUAL EFFORT", "Finding it still requires worker time.", RED),
]):
    yy=4.32+i*0.66
    dot(slide,0.70,yy+0.06,0.09,accent)
    text(slide,0.90,yy,1.50,0.18,head,size=7.2,color=accent,bold=True,margin=0.0)
    text(slide,2.02,yy-0.02,2.05,0.34,body,size=8.3,color=MUTED,margin=0.0)
# right visual panel
rect(slide, 4.35, 1.62, 8.25, 4.95, PANEL, radius=True, line=GRID, lw=0.9)
label(slide, 4.68, 1.91, "THE COST OF NOT KNOWING", MUTED_2, 7.4, 3.0)
# compressor and pipeline
rect(slide, 4.77, 3.25, 1.13, 1.20, PANEL_2, radius=True, line=BLUE, lw=1.0)
ellipse(slide, 4.98, 3.52, 0.58, 0.58, BG, line=BLUE, lw=1.1)
line(slide, 5.27,3.66,5.27,3.96,BLUE,1.1); line(slide,5.12,3.81,5.42,3.81,BLUE,1.1)
text(slide,4.90,4.58,0.85,0.18,"COMPRESSOR",size=6.8,color=MUTED,bold=True,align=PP_ALIGN.CENTER,margin=0.0)
line(slide,5.90,3.82,11.55,3.82,BLUE,3.4)
line(slide,5.90,3.98,11.55,3.98,BLUE,1.0)
for xx in [6.66,7.53,8.40,9.27,10.14,11.01]: line(slide,xx,3.69,xx,4.11,BLUE,0.8)
# air flow arrows
for xx in [6.15,7.35,8.55,9.75,10.95]:
    chevron(slide,xx,3.65,0.19,0.24,BLUE)
# leak point and wasted energy plume
lx=9.10
dot(slide,lx-0.06,3.72,0.12,AMBER)
for dx,dy,dx2 in [(-0.22,0.65,-0.34),(0.0,0.87,0.0),(0.23,0.58,0.34)]:
    line(slide,lx,3.88,lx+dx,3.88+dy,AMBER,1.4)
    line(slide,lx+dx,3.88+dy,lx+dx+dx2*0.22,3.88+dy+0.14,AMBER,1.1)
text(slide,8.55,5.05,1.45,0.18,"WASTED ENERGY",size=7.0,color=AMBER,bold=True,align=PP_ALIGN.CENTER,margin=0.0)
# noisy background bands
for i,lab in enumerate(["MOTORS","VALVES","FANS","PEOPLE"]):
    xx=4.79+i*1.72
    rect(slide,xx,2.38,1.30,0.42,PANEL_2,radius=True,line=GRID,lw=0.6)
    dot(slide,xx+0.12,2.55,0.06,MUTED_2)
    text(slide,xx+0.24,2.48,0.92,0.18,lab,size=6.8,color=MUTED,bold=True,margin=0.0)
# noise annotation
line(slide,10.10,2.80,10.10,2.10,RED,0.8)
text(slide,10.26,2.02,1.90,0.34,"INDUSTRIAL NOISE\ncan mask the signature",size=8.0,color=RED,bold=True,margin=0.0)
# photographic anchor: a leak event makes the problem tangible
picture_cover(slide, LEAK_IMAGE, 4.43, 2.26, 7.98, 2.92)
rect(slide, 4.43, 2.26, 7.98, 0.46, BG, radius=False, transparency=28)
pill(slide, 4.70, 2.37, 1.20, "LEAK EVENT", "352B1F", AMBER, line_color=AMBER, size=6.7)
text(slide, 10.00, 2.39, 2.02, 0.16, "ULTRASONIC SIGNATURE", size=6.8, color=AMBER, bold=True, align=PP_ALIGN.RIGHT, margin=0.0)
# bottom three impact cards
for x,head,body,accent in [(4.78,"HARD TO DETECT","Subtle leaks compete with plant noise.",RED),(7.34,"HARD TO REPEAT","Inspection depends on people and availability.",BLUE),(9.90,"HARD TO PRIORITIZE","A finding without context slows action.",AMBER)]:
    rect(slide,x,5.58,2.32,0.67,PANEL_2,radius=True,line=GRID,lw=0.6)
    text(slide,x+0.12,5.70,2.05,0.16,head,size=7.0,color=accent,bold=True,margin=0.0)
    text(slide,x+0.12,5.91,2.05,0.22,body,size=7.4,color=MUTED,margin=0.0)
add_notes(slide, "Start with the operational reality, not a technical feature. Compressed air is valuable only when it reaches the process. A leak turns that value into wasted energy and another maintenance task. The detection challenge is not just hearing a sound: industrial motors, valves, fans, and people create a noisy acoustic environment. The second issue is repeatability. If inspection is only a person walking a route with a handheld device, coverage depends on time, labor, and memory. That is the gap this project addresses.")

# -----------------------------------------------------------------------------
# SLIDE 3 — EXISTING APPROACH
# -----------------------------------------------------------------------------
slide = prs.slides.add_slide(BLANK)
slide_bg(slide, "02  /  THE GAP", "Manual inspection is useful — but not scalable.", "Traditional handheld ultrasonic detection finds leaks; it does not create an automated inspection memory.", 3, BLUE)
# left today panel
rect(slide, 0.66, 1.66, 5.12, 4.96, PANEL, radius=True, line=GRID, lw=0.8)
pill(slide, 0.92, 1.92, 0.76, "TODAY", "203042", BLUE, line_color=BLUE, size=7.0)
text(slide,0.92,2.35,3.2,0.30,"Handheld ultrasonic rounds",size=15.5,color=WHITE,bold=True,font=FONT_DISPLAY,margin=0.0)
text(slide,0.92,2.77,4.16,0.44,"A trained worker inspects pipeline sections one by one, often recording findings after the fact.",size=9.7,color=MUTED,margin=0.0)
# human + detector line art
ellipse(slide,1.20,3.56,0.34,0.34,PANEL_2,line=BLUE,lw=1.0)
line(slide,1.37,3.90,1.37,4.78,BLUE,1.4)
line(slide,1.37,4.15,1.00,4.42,BLUE,1.2); line(slide,1.37,4.15,1.76,4.42,BLUE,1.2)
line(slide,1.37,4.78,1.06,5.28,BLUE,1.2); line(slide,1.37,4.78,1.70,5.28,BLUE,1.2)
# handheld instrument
rect(slide,1.74,4.22,0.42,0.60,PANEL_3,radius=True,line=AMBER,lw=1.0)
line(slide,1.95,4.22,2.22,3.84,AMBER,1.0)
# pipe route behind
line(slide,2.65,4.12,5.32,4.12,BLUE,2.2); line(slide,2.65,4.22,5.32,4.22,BLUE,0.7)
for xx in [3.1,3.65,4.2,4.75]: line(slide,xx,4.02,xx,4.32,BLUE,0.8)
for xx in [2.85,3.60,4.35,5.10]:
    ellipse(slide,xx-0.05,4.02,0.10,0.10,BG,line=BLUE,lw=0.9)
    text(slide,xx-0.15,4.43,0.30,0.16,chr(65+int((xx-2.85)/0.75)),size=7.2,color=MUTED_2,bold=True,align=PP_ALIGN.CENTER,margin=0.0)
# limitations row
lims=[("HUMAN-DEPENDENT", "coverage follows\nworker availability", BLUE),("TIME-INTENSIVE", "walk section by\nsection", AMBER),("SNAPSHOT-BASED", "finding may lack\nroute context", RED)]
for i,(head,body,accent) in enumerate(lims):
    xx=0.92+i*1.52
    rect(slide,xx,5.48,1.32,0.72,PANEL_2,radius=True,line=GRID,lw=0.6)
    dot(slide,xx+0.12,5.62,0.07,accent)
    text(slide,xx+0.24,5.57,0.94,0.20,head,size=6.4,color=accent,bold=True,margin=0.0)
    text(slide,xx+0.12,5.82,1.08,0.26,body,size=7.0,color=MUTED,margin=0.0)
# right gap statement
rect(slide, 6.12, 1.66, 6.48, 4.96, PANEL_2, radius=True, line=CYAN_DARK, lw=1.0)
pill(slide, 6.43, 1.92, 0.74, "GAP", "14313A", CYAN, line_color=CYAN_DARK, size=7.0)
text(slide,6.43,2.35,5.28,0.78,"Industries need a flexible system capable of automated periodic inspection across predefined pipeline routes.",size=19.5,color=WHITE,bold=True,font=FONT_DISPLAY,margin=0.0)
line(slide,6.44,3.36,12.22,3.36,CYAN_DARK,1.0)
# desired capability ladder
text(slide,6.45,3.66,2.4,0.20,"WHAT CHANGES",size=7.2,color=CYAN,bold=True,margin=0.0)
for i,(head,body,accent) in enumerate([
    ("PATROL", "repeatable route coverage", CYAN),
    ("LISTEN", "continuous sensing while moving", BLUE),
    ("REMEMBER", "location + confidence + time", AMBER),
    ("ACT", "priority for maintenance", RED),
]):
    yy=4.03+i*0.49
    ellipse(slide,6.48,yy+0.02,0.22,0.22,accent,line=accent)
    text(slide,6.53,yy+0.035,0.12,0.13,str(i+1),size=6.8,color=BG,bold=True,align=PP_ALIGN.CENTER,margin=0.0)
    text(slide,6.84,yy,1.25,0.18,head,size=8.0,color=accent,bold=True,margin=0.0)
    text(slide,8.18,yy,3.75,0.18,body,size=8.6,color=MUTED,margin=0.0)
add_notes(slide, "We are not saying handheld ultrasonic tools are ineffective. They are valuable tools, but the workflow is human-dependent and difficult to scale across routes. A worker still has to be present, the inspection is a snapshot, and the record may not carry enough route context for the next maintenance decision. Our design goal is therefore not to replace expertise; it is to make the first-pass inspection more repeatable and structured. The gap is automated periodic inspection across predefined pipeline routes.")

# -----------------------------------------------------------------------------
# SLIDE 4 — SOLUTION WORKFLOW
# -----------------------------------------------------------------------------
slide = prs.slides.add_slide(BLANK)
slide_bg(slide, "03  /  OUR SOLUTION", "Move from leak hunting to leak intelligence.", "One patrol produces a connected chain from acoustic evidence to a maintenance-ready alert.", 4, CYAN)
# key statement band
rect(slide,0.66,1.60,11.98,0.62,"14313A",radius=True,line=CYAN_DARK,lw=0.8)
icon_robot(slide,0.90,1.74,0.62,CYAN,AMBER)
text(slide,1.72,1.72,9.35,0.22,"Turn manual leak inspection into an automated inspection and intelligence workflow.",size=13.5,color=WHITE,bold=True,margin=0.0)
pill(slide,11.24,1.78,1.04,"MVP FLOW",PANEL_2,CYAN,line_color=CYAN_DARK,size=6.8)
# workflow row 1
nodes=[
    ("Robot patrol","predefined route",CYAN,icon_robot),
    ("Acoustic sense","capture while moving",AMBER,icon_signal),
    ("Preprocess","filter + window",BLUE,icon_filter),
    ("Edge AI","classify locally",PURPLE,icon_brain),
]
startx=0.68; y=2.72; w=2.53; h=1.38; gap=0.40
for i,(head,sub,accent,ic) in enumerate(nodes):
    workflow_node(slide,startx+i*(w+gap),y,w,h,i+1,head,sub,accent,ic)
    if i<len(nodes)-1: add_arrow(slide,startx+i*(w+gap)+w+0.10,y+0.56,0.18,MUTED_2)
# workflow row 2
nodes2=[
    ("Leak detected","leak / no leak",RED,icon_alert),
    ("Locate","route + distance",CYAN,icon_pin),
    ("Prioritize","severity / action",AMBER,icon_wrench),
    ("Alert","dashboard record",GREEN,icon_server_small),
]
y2=4.62
for i,(head,sub,accent,ic) in enumerate(nodes2):
    workflow_node(slide,startx+i*(w+gap),y2,w,h,i+5,head,sub,accent,ic)
    if i<len(nodes2)-1: add_arrow(slide,startx+i*(w+gap)+w+0.10,y2+0.56,0.18,MUTED_2)
# connector from row one to row two
line(slide,9.33,4.10,9.33,4.46,GRID,1.0)
chevron(slide,9.22,4.27,0.22,0.18,CYAN)
# bottom proof bar
rect(slide,0.68,6.33,11.96,0.37,PANEL,radius=True,line=GRID,lw=0.6)
text(slide,0.92,6.43,1.20,0.16,"THE OUTPUT",size=7.0,color=CYAN,bold=True,margin=0.0)
text(slide,2.15,6.40,9.92,0.18,"An approximate, route-aware leak record — not just a raw sensor reading.",size=9.6,color=WHITE,bold=True,margin=0.0)
add_notes(slide, "This is the core narrative slide. The value comes from the connected workflow. The robot patrols a known route, senses acoustics as it moves, applies signal preprocessing, and runs an Edge AI classifier locally. When the event is likely a leak, the system attaches route and distance context, estimates a basic priority, and sends an alert to the dashboard. The output is not just a sound sample; it is an actionable inspection record.")

# -----------------------------------------------------------------------------
# SLIDE 5 — SYSTEM
# -----------------------------------------------------------------------------
slide = prs.slides.add_slide(BLANK)
slide_bg(slide, "04  /  SYSTEM DESIGN", "A compact robot closes the loop.", "Sensing, movement, inference, position, and communication are designed as one inspection system.", 5, CYAN)
# left robot diagram panel
rect(slide,0.66,1.62,5.36,4.94,PANEL,radius=True,line=GRID,lw=0.8)
label(slide,0.94,1.90,"MOBILE INSPECTION NODE",MUTED_2,7.4,3.0)
draw_robot_diagram(slide,1.02,2.23,1.18)
# small photographic inset grounds the system diagram in a believable form factor
rect(slide,4.44,4.78,1.22,1.34,PANEL_2,radius=True,line=GRID,lw=0.7)
picture_cover(slide, ROBOT_IMAGE, 4.50, 4.84, 1.10, 1.10)
label(slide,4.52,6.00,"FORM FACTOR",AMBER,5.8,1.04)
# right components
text(slide,6.39,1.80,4.0,0.24,"System blocks",size=13.5,color=WHITE,bold=True,font=FONT_DISPLAY,margin=0.0)
text(slide,6.40,2.15,5.5,0.38,"Each block has a clear role in the MVP — no hidden magic.",size=9.4,color=MUTED,margin=0.0)
components=[
    ("01", "Line following", "Autonomous movement on a controlled route.", CYAN, icon_route_small),
    ("02", "Ultrasonic / acoustic sensor", "Captures leak-related signal windows.", AMBER, icon_signal),
    ("03", "Embedded / Edge AI processor", "Runs local classification and scoring.", PURPLE, icon_brain),
    ("04", "Position + distance", "Maps detections to section and route distance.", BLUE, icon_pin),
    ("05", "Communication module", "Sends compact records to the dashboard.", GREEN, icon_server_small),
    ("06", "Battery / power", "Supports the patrol cycle and safe return.", ORANGE, icon_battery_small),
]
for i,(num,head,body,accent,ic) in enumerate(components):
    row=i//2; col=i%2; xx=6.40+col*3.08; yy=2.79+row*0.89
    rect(slide,xx,yy,2.78,0.68,PANEL,radius=True,line=GRID,lw=0.65)
    dot(slide,xx+0.16,yy+0.18,0.08,accent)
    text(slide,xx+0.33,yy+0.11,0.27,0.18,num,size=6.8,color=accent,bold=True,margin=0.0)
    text(slide,xx+0.66,yy+0.10,1.95,0.20,head,size=8.0,color=WHITE,bold=True,margin=0.0)
    text(slide,xx+0.66,yy+0.34,1.94,0.22,body,size=7.0,color=MUTED,margin=0.0)
    ic(slide,xx+2.28,yy+0.15,0.38,accent)
# system loop footer
rect(slide,6.40,5.63,6.08,0.72,"14313A",radius=True,line=CYAN_DARK,lw=0.8)
text(slide,6.65,5.77,0.9,0.16,"LOOP",size=7.0,color=CYAN,bold=True,margin=0.0)
text(slide,7.56,5.73,4.63,0.25,"Move → listen → decide → log → continue",size=12.0,color=WHITE,bold=True,margin=0.0)
add_notes(slide, "Walk through the architecture from left to right. The robot is deliberately compact: a line-following base provides repeatable movement on a controlled route; the acoustic sensor captures the signal; an embedded processor makes the local decision; position and distance create context; communication publishes a compact record; and the battery supports the patrol. The system is modular, so the hackathon can demonstrate the smallest viable loop without pretending to have a fully industrialized platform.")

# -----------------------------------------------------------------------------
# SLIDE 6 — EDGE AI
# -----------------------------------------------------------------------------
slide = prs.slides.add_slide(BLANK)
slide_bg(slide, "05  /  EDGE AI", "Edge AI makes the sensor useful in the factory.", "The differentiator is not a microphone alone — it is local classification that can separate signal from surrounding noise.", 6, PURPLE)
# left signal journey
rect(slide,0.66,1.63,5.06,4.95,PANEL,radius=True,line=GRID,lw=0.8)
label(slide,0.94,1.91,"FROM SOUND TO DECISION",MUTED_2,7.4,3.0)
draw_signal_panel(slide,0.94,2.26,4.50,1.52,True)
# mini arrows and stages
stages=[("FILTER","remove irrelevant\nfrequency / noise",BLUE,icon_filter),("FEATURES","describe the\nsignal window",CYAN,icon_wave_small),("CLASSIFY","compare with\nlearned patterns",PURPLE,icon_brain)]
for i,(head,body,accent,ic) in enumerate(stages):
    xx=0.94+i*1.48
    if i>0: add_arrow(slide,xx-0.27,4.29,0.17,MUTED_2)
    rect(slide,xx,4.10,1.22,0.91,PANEL_2,radius=True,line=GRID,lw=0.6)
    ic(slide,xx+0.12,4.23,0.35,accent)
    text(slide,xx+0.12,4.60,0.99,0.16,head,size=6.7,color=accent,bold=True,margin=0.0)
    text(slide,xx+0.12,4.79,0.96,0.20,body,size=6.8,color=MUTED,margin=0.0)
# output
rect(slide,0.94,5.33,4.50,0.74,"1A2F38",radius=True,line=CYAN_DARK,lw=0.8)
text(slide,1.15,5.51,1.14,0.18,"OUTPUT",size=7.0,color=CYAN,bold=True,margin=0.0)
pill(slide,2.36,5.46,0.93,"LEAK", "352B1F", AMBER, line_color=AMBER, size=6.9)
text(slide,3.42,5.52,1.64,0.16,"confidence score",size=8.4,color=WHITE,bold=True,align=PP_ALIGN.RIGHT,margin=0.0)
# right edge AI rationale
text(slide,6.20,1.84,3.2,0.28,"Why process locally?",size=14.5,color=WHITE,bold=True,font=FONT_DISPLAY,margin=0.0)
benefits=[
    ("LOW LATENCY", "A detection decision does not wait for a round trip.", CYAN),
    ("LESS CLOUD DEPENDENCY", "The patrol can keep sensing with limited connectivity.", BLUE),
    ("FASTER RESPONSE", "A likely event can be surfaced close to the moment it occurs.", AMBER),
    ("RELIABLE + PRIVATE", "Raw acoustic windows can stay on-device unless shared deliberately.", GREEN),
]
for i,(head,body,accent) in enumerate(benefits):
    yy=2.42+i*0.67
    rect(slide,6.20,yy,6.30,0.50,PANEL,radius=True,line=GRID,lw=0.6)
    dot(slide,6.41,yy+0.19,0.09,accent)
    text(slide,6.63,yy+0.10,1.80,0.18,head,size=7.3,color=accent,bold=True,margin=0.0)
    text(slide,8.53,yy+0.10,3.70,0.25,body,size=8.0,color=MUTED,margin=0.0)
# model guardrail
rect(slide,6.20,5.32,6.30,0.87,"211B2A",radius=True,line=PURPLE,lw=0.8)
text(slide,6.45,5.48,1.50,0.18,"MODEL GUARDRAIL",size=7.0,color=PURPLE,bold=True,margin=0.0)
text(slide,6.45,5.70,5.65,0.28,"Edge AI classifier — model to be selected based on dataset and validation.",size=8.3,color=WHITE,bold=True,margin=0.0)
add_notes(slide, "This is the technical heart of the concept. The acoustic sensor sees a window that may contain a leak signature plus background noise. We filter and window the signal, extract features, and pass them into a local classifier that outputs leak or no leak with a confidence score. Edge processing keeps the first decision close to the robot, reduces dependence on connectivity, and supports faster response. We intentionally do not name a model or claim accuracy before the dataset and validation plan are complete.")

# -----------------------------------------------------------------------------
# SLIDE 7 — LOCALIZATION & PRIORITIZATION
# -----------------------------------------------------------------------------
slide = prs.slides.add_slide(BLANK)
slide_bg(slide, "06  /  LOCATION + PRIORITY", "Every alert arrives with context.", "For the MVP, approximate location comes from a known route, distance tracking, section ID, and event time.", 7, AMBER)
# map panel
rect(slide,0.66,1.64,7.02,4.94,PANEL,radius=True,line=GRID,lw=0.8)
label(slide,0.94,1.90,"MVP ROUTE MAP  /  P-03",MUTED_2,7.4,3.0)
# section headers
for xx,sec in [(1.16,"A"),(2.82,"B"),(4.58,"C"),(6.08,"D")]:
    pill(slide,xx,2.28,0.36,sec,PANEL_2,MUTED,size=7.0)
# route
pts=draw_route(slide,1.03,3.00,5.95,1.70,active=0.58,color=CYAN,muted=GRID)
# section divider ticks
for xx in [2.53,4.20,5.80]: line(slide,xx,2.84,xx,5.34,GRID,0.8)
# robot marker at approximate B
rx,ry=pts[4]
ellipse(slide,rx-0.20,ry-0.20,0.40,0.40,BG,line=AMBER,lw=1.3)
dot(slide,rx-0.06,ry-0.06,0.12,AMBER)
for r,tr in [(0.42,65),(0.78,84),(1.10,92)]: ellipse(slide,rx-r/2,ry-r/2,r,r,BG,line=AMBER,lw=0.8,transparency=tr)
text(slide,rx-0.58,ry+0.38,1.20,0.20,"18.6 m  /  SECTION B",size=7.6,color=AMBER,bold=True,align=PP_ALIGN.CENTER,margin=0.0)
# distance scale
line(slide,1.02,5.45,6.98,5.45,MUTED_2,0.8)
for i in range(8):
    xx=1.02+i*(5.96/7)
    line(slide,xx,5.39,xx,5.52,MUTED_2,0.8)
    text(slide,xx-0.17,5.60,0.34,0.16,str(i*5),size=6.4,color=MUTED_2,align=PP_ALIGN.CENTER,margin=0.0)
text(slide,6.41,5.59,0.55,0.17,"m",size=6.8,color=MUTED_2,bold=True,margin=0.0)
# record panel
rect(slide,8.02,1.64,4.58,3.56,PANEL_2,radius=True,line=AMBER,lw=1.0)
pill(slide,8.34,1.93,1.44,"EXAMPLE RECORD", "352B1F", AMBER, line_color=AMBER, size=6.7)
text(slide,8.34,2.38,2.5,0.28,"Maintenance record",size=14.5,color=WHITE,bold=True,font=FONT_DISPLAY,margin=0.0)
records=[("Pipeline", "P-03"),("Section", "B"),("Distance", "18.6 m"),("Leak probability", "94%"),("Priority", "HIGH"),("Status", "Maintenance Required")]
for i,(k,v) in enumerate(records):
    yy=2.86+i*0.31
    line(slide,8.34,yy+0.25,12.24,yy+0.25,GRID,0.5)
    text(slide,8.34,yy,1.68,0.17,k,size=7.7,color=MUTED,margin=0.0)
    col=RED if k=="Priority" else (AMBER if k=="Leak probability" else WHITE)
    text(slide,10.10,yy,2.14,0.17,v,size=8.0,color=col,bold=(k in ["Priority","Leak probability"]),align=PP_ALIGN.RIGHT,margin=0.0)
# separation band
rect(slide,8.02,5.44,4.58,1.14,PANEL,radius=True,line=GRID,lw=0.7)
pill(slide,8.32,5.65,0.73,"MVP", "14313A", CYAN, line_color=CYAN_DARK, size=6.8)
text(slide,9.18,5.62,3.05,0.28,"route + distance + section + timestamp",size=8.5,color=WHITE,bold=True,margin=0.0)
pill(slide,8.32,6.06,0.89,"FUTURE", "211B2A", PURPLE, line_color=PURPLE, size=6.8)
text(slide,9.35,6.03,2.85,0.28,"SLAM / UWB / multi-sensor refinement",size=8.2,color=MUTED,margin=0.0)
add_notes(slide, "The MVP does not need to claim pinpoint geolocation. It knows the route, tracks distance, and maps the event to a pipeline and section. That is enough to produce a useful approximate location for a maintenance team. The record also carries confidence, timestamp, and a basic priority. The example values are a dashboard record format, not a deployment result. More advanced localization such as UWB, SLAM, or multi-sensor triangulation is clearly a future extension.")

# -----------------------------------------------------------------------------
# SLIDE 8 — DASHBOARD
# -----------------------------------------------------------------------------
slide = prs.slides.add_slide(BLANK)
slide_bg(slide, "07  /  DASHBOARD", "The dashboard answers: what needs attention first?", "A maintenance team should see the route, the evidence, and the next action in one glance.", 8, RED)
# dashboard shell
rect(slide,0.66,1.63,11.96,4.98,"0D141D",radius=True,line=GRID,lw=0.9)
# sidebar
rect(slide,0.66,1.63,1.45,4.98,"0B1118",radius=True,line=GRID,lw=0.6)
text(slide,0.88,1.93,0.94,0.22,"AIR//SENSE",size=10.5,color=WHITE,bold=True,font=FONT_DISPLAY,margin=0.0)
pill(slide,0.88,2.35,0.83,"LIVE", "14313A", CYAN, line_color=CYAN_DARK, size=6.8)
for i,(head,accent) in enumerate([("Overview",CYAN),("Patrols",MUTED),("Alerts",RED),("History",MUTED),("Settings",MUTED)]):
    yy=2.97+i*0.50
    if i==0: rect(slide,0.82,yy-0.08,1.08,0.31,PANEL_2,radius=True,line=CYAN_DARK,lw=0.6)
    dot(slide,0.96,yy+0.01,0.07,accent)
    text(slide,1.11,yy-0.02,0.68,0.18,head,size=7.6,color=WHITE if i==0 else MUTED,bold=(i==0),margin=0.0)
# main heading
text(slide,2.42,1.93,3.3,0.26,"Patrol overview",size=13.5,color=WHITE,bold=True,font=FONT_DISPLAY,margin=0.0)
text(slide,2.43,2.23,3.6,0.18,"P-03  /  Section B  /  08:42 UTC",size=7.8,color=MUTED,margin=0.0)
pill(slide,9.36,1.92,1.18,"DEMO DATA", "211B2A", AMBER, line_color=AMBER, size=6.7)
pill(slide,10.80,1.92,1.34,"ROBOT ONLINE", "14313A", CYAN, line_color=CYAN_DARK, size=6.7)
# status cards
status=[("ROBOT STATUS","Patrolling","08:42","CYAN"),("ROUTE PROGRESS","58%","P-03 / B","BLUE"),("OPEN ALERTS","01","High priority","RED")]
colors={"CYAN":CYAN,"BLUE":BLUE,"RED":RED}
for i,(head,big,sub,cname) in enumerate(status):
    xx=2.42+i*1.72
    rect(slide,xx,2.61,1.52,0.76,PANEL,radius=True,line=GRID,lw=0.6)
    text(slide,xx+0.12,2.73,1.25,0.15,head,size=6.5,color=MUTED_2,bold=True,margin=0.0)
    text(slide,xx+0.12,2.91,1.22,0.24,big,size=13.0,color=colors[cname],bold=True,font=FONT_DISPLAY,margin=0.0)
    text(slide,xx+0.12,3.19,1.23,0.13,sub,size=6.5,color=MUTED,margin=0.0)
# route map center
rect(slide,2.42,3.63,4.08,2.52,PANEL,radius=True,line=GRID,lw=0.6)
text(slide,2.64,3.84,1.8,0.17,"PIPELINE ROUTE",size=7.0,color=MUTED_2,bold=True,margin=0.0)
text(slide,5.24,3.84,0.98,0.17,"P-03 / 32 m",size=7.0,color=MUTED,align=PP_ALIGN.RIGHT,margin=0.0)
pts=draw_route(slide,2.72,4.21,3.46,1.22,active=0.58,color=CYAN,muted=GRID)
rx,ry=pts[4]
dot(slide,rx-0.07,ry-0.07,0.14,AMBER)
for r,tr in [(0.30,68),(0.56,86)]: ellipse(slide,rx-r/2,ry-r/2,r,r,BG,line=AMBER,lw=0.7,transparency=tr)
icon_robot(slide,3.04,5.18,0.52,CYAN,AMBER)
text(slide,3.76,5.25,2.1,0.18,"ROBOT  /  18.6 m",size=7.5,color=AMBER,bold=True,margin=0.0)
# alerts table right
rect(slide,6.72,2.61,5.55,3.54,PANEL,radius=True,line=GRID,lw=0.6)
text(slide,6.96,2.84,1.8,0.18,"DETECTIONS",size=7.0,color=MUTED_2,bold=True,margin=0.0)
text(slide,11.18,2.84,0.82,0.18,"1 open",size=7.0,color=RED,bold=True,align=PP_ALIGN.RIGHT,margin=0.0)
# table headings
for x,lab,w in [(6.96,"LOCATION",1.32),(8.30,"CONF.",0.72),(9.05,"PRIORITY",0.85),(10.02,"TIME",0.75),(10.88,"STATUS",1.18)]:
    text(slide,x,3.20,w,0.16,lab,size=6.3,color=MUTED_2,bold=True,margin=0.0)
line(slide,6.94,3.43,12.02,3.43,GRID,0.7)
rows=[("P-03 / B", "94%", "HIGH", "08:42", "REQUIRED", RED),
      ("P-03 / A", "—", "CLEAR", "08:27", "CLOSED", GREEN),
      ("P-02 / C", "—", "CLEAR", "YEST.", "CLOSED", GREEN)]
for i,(loc,conf,prio,tm,st,accent) in enumerate(rows):
    yy=3.67+i*0.67
    line(slide,6.94,yy+0.38,12.02,yy+0.38,GRID,0.5)
    dot(slide,7.00,yy+0.10,0.08,accent)
    text(slide,7.16,yy+0.05,1.12,0.18,loc,size=7.2,color=WHITE,bold=(i==0),margin=0.0)
    text(slide,8.30,yy+0.05,0.70,0.18,conf,size=7.2,color=AMBER if i==0 else MUTED, bold=(i==0),margin=0.0)
    text(slide,9.05,yy+0.05,0.85,0.18,prio,size=7.0,color=accent,bold=True,margin=0.0)
    text(slide,10.02,yy+0.05,0.75,0.18,tm,size=7.0,color=MUTED,margin=0.0)
    text(slide,10.88,yy+0.05,1.18,0.18,st,size=6.6,color=accent if i==0 else MUTED,bold=(i==0),margin=0.0)
# bottom history and priority ladder inside dashboard
rect(slide,6.72,6.26,2.65,0.30,PANEL_2,radius=True,line=GRID,lw=0.5)
text(slide,6.90,6.34,0.82,0.15,"HISTORY",size=6.5,color=MUTED_2,bold=True,margin=0.0)
draw_sparkline(slide,7.80,6.33,1.31,0.15,RED,[0.88,0.72,0.78,0.43,0.58,0.30,0.35,0.20])
rect(slide,9.58,6.26,2.69,0.30,PANEL_2,radius=True,line=GRID,lw=0.5)
text(slide,9.77,6.34,0.92,0.15,"PRIORITY",size=6.5,color=MUTED_2,bold=True,margin=0.0)
for i,(lab,col) in enumerate([("CRITICAL",RED),("HIGH",AMBER),("MEDIUM",ORANGE),("LOW",GREEN)]):
    dot(slide,10.78+i*0.32,6.35,0.07,col)
# callout below shell
text(slide,0.68,6.82,11.9,0.16,"The interface is deliberately maintenance-first: location, confidence, urgency, timestamp, and status — in that order.",size=8.5,color=WHITE,bold=True,align=PP_ALIGN.CENTER,margin=0.0)
add_notes(slide, "The dashboard is designed for a maintenance decision, not a robotics demo. The operator can see whether the robot is online, where it is on the route, which alerts are open, and how those alerts are prioritized. The example record shows the fields that matter: location, confidence, priority, time, and maintenance status. A history view creates a record of repeat detections without claiming a predictive model yet.")

# -----------------------------------------------------------------------------
# SLIDE 9 — DIFFERENTIATION
# -----------------------------------------------------------------------------
slide = prs.slides.add_slide(BLANK)
slide_bg(slide, "08  /  DIFFERENTIATION", "The step-change is the workflow — not just the detector.", "The proposed system combines mobility, local intelligence, and maintenance context in a single inspection loop.", 9, CYAN)
# left framing
rect(slide,0.66,1.64,2.25,4.95,PANEL,radius=True,line=GRID,lw=0.8)
icon_robot(slide,1.30,2.07,0.94,CYAN,AMBER)
text(slide,0.93,3.22,1.72,0.67,"Not just a\nleak detector.",size=18,color=WHITE,bold=True,font=FONT_DISPLAY,align=PP_ALIGN.CENTER,margin=0.0)
line(slide,0.95,4.18,2.60,4.18,CYAN_DARK,1.0)
text(slide,0.94,4.46,1.70,0.80,"A route-aware inspection and maintenance-intelligence platform.",size=10.0,color=MUTED,align=PP_ALIGN.CENTER,margin=0.0)
pill(slide,1.03,5.70,1.52,"WORKFLOW VALUE", "14313A", CYAN, line_color=CYAN_DARK, size=7.0)
# table
x0=3.18; y0=1.64; totalw=9.42
colw=[2.35,2.10,2.10,2.87]
# column headers
headers=["CRITERIA","MANUAL HANDHELD","FIXED SENSOR","PROPOSED AUTONOMOUS\nEDGE-AI ROBOT"]
for i,(head,wc) in enumerate(zip(headers,colw)):
    xx=x0+sum(colw[:i])
    fill="14313A" if i==3 else PANEL
    rect(slide,xx,y0,wc,0.70,fill,radius=True,line=CYAN_DARK if i==3 else GRID,lw=0.8)
    text(slide,xx+0.14,y0+0.16,wc-0.28,0.35,head,size=8.0,color=CYAN if i==3 else MUTED,bold=True,align=PP_ALIGN.CENTER,margin=0.0)
rows=[
    ("Inspection method", "Worker + handheld", "Installed nodes", "Mobile route patrol"),
    ("Automation", "Human-triggered", "Always-on at node", "Automated periodic inspection"),
    ("Mobility", "Worker moves", "Fixed location", "Robot moves on route"),
    ("Noise filtering", "Operator judgment", "Local thresholding", "Edge AI classifier"),
    ("Localization", "Manual notes", "Known node", "Route + distance + section"),
    ("Sensing model", "Point snapshot", "Point sensing", "Continuous sensing during patrol"),
    ("Maintenance priority", "After review", "Alert at node", "Confidence + basic priority"),
    ("Scalability", "Labor-limited", "Wiring / install-limited", "Routes + fleet over time"),
]
rowh=0.48
for r,(crit,manual,fixed,proposed) in enumerate(rows):
    yy=y0+0.78+r*rowh
    vals=[crit,manual,fixed,proposed]
    for i,(val,wc) in enumerate(zip(vals,colw)):
        xx=x0+sum(colw[:i])
        fill="101E28" if i==3 else PANEL_2
        rect(slide,xx,yy,wc,rowh-0.05,fill,radius=False,line=GRID,lw=0.45)
        col=CYAN if i==3 else (WHITE if i==0 else MUTED)
        text(slide,xx+0.14,yy+0.13,wc-0.28,rowh-0.15,val,size=7.6,color=col,bold=(i in [0,3]),margin=0.0)
        if i==3: line(slide,xx+0.02,yy+0.05,xx+0.02,yy+rowh-0.11,CYAN,2.0)
# note
rect(slide,3.18,6.32,9.42,0.27,"211B2A",radius=True,line=PURPLE,lw=0.6)
text(slide,3.36,6.38,9.04,0.14,"Boundary: periodic autonomous patrol — not continuous monitoring of every pipeline point.",size=7.4,color=PURPLE,bold=True,align=PP_ALIGN.CENTER,margin=0.0)
add_notes(slide, "The differentiation is the combination. A manual handheld tool has a human in the loop. A fixed sensor can be useful but only where it is installed. Our proposed robot is mobile, route-aware, and can continuously sense during its patrol while filtering signals locally. The important wording is periodic autonomous inspection, not continuous monitoring of every point. That honest boundary makes the concept more credible while preserving its value: repeatable coverage plus maintenance context.")

# -----------------------------------------------------------------------------
# SLIDE 10 — MVP FEASIBILITY
# -----------------------------------------------------------------------------
slide = prs.slides.add_slide(BLANK)
slide_bg(slide, "09  /  MVP + FEASIBILITY", "Smallest working system. Strongest proof.", "The hackathon demo proves the core loop on a controlled route — and makes the future boundary explicit.", 10, GREEN)
# scope column
rect(slide,0.66,1.64,6.15,4.96,PANEL,radius=True,line=GRID,lw=0.8)
pill(slide,0.94,1.91,1.05,"HACKATHON MVP", "14313A", GREEN, line_color=CYAN_DARK, size=6.9)
text(slide,0.94,2.34,4.5,0.30,"What we will demonstrate",size=15.0,color=WHITE,bold=True,font=FONT_DISPLAY,margin=0.0)
mvp=[
    ("01", "Line-following robot + controlled pipeline route", CYAN),
    ("02", "Ultrasonic / acoustic sensing while moving", AMBER),
    ("03", "Leak / no-leak Edge AI classification", PURPLE),
    ("04", "Position estimate, confidence score, basic severity / priority", BLUE),
    ("05", "Dashboard record + alert generation", GREEN),
]
for i,(num,item,accent) in enumerate(mvp):
    yy=2.92+i*0.53
    ellipse(slide,0.96,yy,0.30,0.30,accent,line=accent)
    text(slide,0.96,yy+0.03,0.30,0.23,num,size=7.0,color=BG,bold=True,align=PP_ALIGN.CENTER,margin=0.0)
    text(slide,1.44,yy+0.03,4.75,0.23,item,size=9.3,color=WHITE,bold=True,margin=0.0)
    line(slide,1.44,yy+0.39,6.35,yy+0.39,GRID,0.5)
# demo success bar
rect(slide,0.94,5.80,5.55,0.48,"14313A",radius=True,line=CYAN_DARK,lw=0.7)
text(slide,1.14,5.96,1.15,0.15,"PROOF POINT",size=6.8,color=CYAN,bold=True,margin=0.0)
text(slide,2.46,5.91,3.70,0.21,"detect → locate → prioritize → alert",size=9.2,color=WHITE,bold=True,margin=0.0)
# future column
rect(slide,7.14,1.64,5.46,4.96,PANEL_2,radius=True,line=PURPLE,lw=0.9)
pill(slide,7.44,1.91,1.02,"FUTURE", "211B2A", PURPLE, line_color=PURPLE, size=7.0)
text(slide,7.44,2.34,4.2,0.30,"Enhancements, not MVP claims",size=14.5,color=WHITE,bold=True,font=FONT_DISPLAY,margin=0.0)
future=[("SLAM / UWB", "refined localization", PURPLE),
        ("ADVANCED SEVERITY", "validated severity estimation", AMBER),
        ("MULTI-ROBOT", "fleet coordination", BLUE),
        ("INDUSTRIAL INTEGRATION", "maintenance-system connectors", GREEN),
        ("LARGE-SCALE ROUTES", "deployment across plant assets", CYAN)]
for i,(head,body,accent) in enumerate(future):
    yy=2.92+i*0.54
    rect(slide,7.44,yy,4.86,0.38,PANEL,radius=True,line=GRID,lw=0.55)
    dot(slide,7.63,yy+0.14,0.08,accent)
    text(slide,7.84,yy+0.09,1.90,0.17,head,size=6.9,color=accent,bold=True,margin=0.0)
    text(slide,9.84,yy+0.09,2.22,0.17,body,size=7.2,color=MUTED,margin=0.0)
# feasibility strap
text(slide,7.44,5.92,4.80,0.20,"Build the loop first. Validate before scaling.",size=9.5,color=WHITE,bold=True,align=PP_ALIGN.CENTER,margin=0.0)
add_notes(slide, "Judges should leave this slide knowing exactly what will be built. The MVP is a controlled route, a line-following robot, acoustic sensing, a leak or no-leak classifier, a position estimate, a confidence score, basic priority, and a dashboard alert. That is enough to prove the core value. SLAM, UWB, advanced severity estimation, multi-robot fleets, and industrial system integration are future enhancements. We are intentionally choosing a small, demonstrable system rather than overpromising a plant-ready deployment.")

# -----------------------------------------------------------------------------
# SLIDE 11 — IMPACT + SCALE
# -----------------------------------------------------------------------------
slide = prs.slides.add_slide(BLANK)
slide_bg(slide, "10  /  IMPACT + SCALE", "Create value before you scale complexity.", "Start with better inspection records and faster maintenance decisions; scale only what validation supports.", 11, GREEN)
# impact cards
text(slide,0.66,1.65,3.2,0.26,"Industrial value",size=14.5,color=WHITE,bold=True,font=FONT_DISPLAY,margin=0.0)
impact=[("REDUCE EFFORT","Less manual first-pass inspection effort.",CYAN,icon_robot),
        ("IDENTIFY EARLIER","Surface likely events during a patrol.",AMBER,icon_alert),
        ("PRIORITIZE BETTER","Give teams confidence, context, and urgency.",RED,icon_wrench),
        ("EFFICIENCY + RECORDS","Support energy-efficiency work with an automated inspection history.",BLUE,icon_clock)]
for i,(head,body,accent,ic) in enumerate(impact):
    xx=0.66+(i%2)*2.82; yy=2.08+(i//2)*1.18
    small_card(slide,xx,yy,2.55,0.94,head,body,accent,ic)
# measurement bar
rect(slide,0.66,4.75,5.37,1.53,"14313A",radius=True,line=CYAN_DARK,lw=0.8)
label(slide,0.92,5.02,"MEASURE BEFORE CLAIMING ROI",CYAN,7.0,3.4)
text(slide,0.92,5.35,4.60,0.50,"Track detections, route coverage, response time, and [validated savings].",size=12.0,color=WHITE,bold=True,font=FONT_DISPLAY,margin=0.0)
# scale ladder right
text(slide,6.66,1.65,3.2,0.26,"Path to scale",size=14.5,color=WHITE,bold=True,font=FONT_DISPLAY,margin=0.0)
scale_steps=[("01","ONE CONTROLLED ROUTE","prove the loop",CYAN),
             ("02","MULTIPLE ROUTES","schedule periodic patrols",BLUE),
             ("03","MULTI-ROBOT FLEET","parallelize coverage",PURPLE),
             ("04","MAINTENANCE ECOSYSTEM","history → predictive workflows",AMBER)]
for i,(num,head,body,accent) in enumerate(scale_steps):
    yy=2.10+i*0.93
    # connector
    if i<3: line(slide,7.03,yy+0.55,7.03,yy+0.93,GRID,1.4)
    ellipse(slide,6.75,yy,0.56,0.56,accent,line=accent)
    text(slide,6.75,yy+0.16,0.56,0.20,num,size=8.2,color=BG,bold=True,align=PP_ALIGN.CENTER,margin=0.0)
    text(slide,7.60,yy+0.05,3.10,0.18,head,size=8.5,color=accent,bold=True,margin=0.0)
    text(slide,7.60,yy+0.29,4.35,0.18,body,size=8.7,color=MUTED,margin=0.0)
# scale boundary note
rect(slide,6.66,5.94,5.94,0.35,PANEL,radius=True,line=GRID,lw=0.6)
text(slide,6.87,6.03,5.50,0.16,"Scale the routes, the records, and the learning — not the claims.",size=7.5,color=WHITE,bold=True,align=PP_ALIGN.CENTER,margin=0.0)
add_notes(slide, "The impact story is intentionally measurable without invented numbers. The system can reduce manual first-pass effort, help surface likely leaks earlier in a patrol, improve maintenance prioritization, and create automated inspection records. We would measure route coverage, detections, response time, and validated savings before making ROI claims. Scaling then follows a controlled path: one route, multiple scheduled routes, a fleet, and eventually integration into a predictive maintenance ecosystem.")

# -----------------------------------------------------------------------------
# SLIDE 12 — FINAL PITCH
# -----------------------------------------------------------------------------
slide = prs.slides.add_slide(BLANK)
slide_bg(slide, "11  /  FINAL PITCH", "From manual leak hunting to autonomous leak intelligence.", "A realistic MVP today. A scalable operating model tomorrow.", 12, CYAN)
# full-width journey
rect(slide,0.66,1.70,11.98,1.42,PANEL,radius=True,line=GRID,lw=0.8)
journey=[("SINGLE\nROBOT",CYAN,icon_robot),("MULTIPLE\nROUTES",BLUE,icon_route_small),("MULTI-ROBOT\nFLEET",PURPLE,icon_robot),("SCHEDULED\nPATROLS",AMBER,icon_clock),("HISTORICAL\nANALYTICS",ORANGE,icon_server_small),("PREDICTIVE\nMAINTENANCE",GREEN,icon_wrench)]
for i,(head,accent,ic) in enumerate(journey):
    xx=0.98+i*1.91
    if i>0: add_arrow(slide,xx-0.32,2.24,0.20,MUTED_2)
    ellipse(slide,xx,2.00,0.53,0.53,accent,line=accent)
    ic(slide,xx+0.08,2.07,0.38,accent)
    text(slide,xx-0.14,2.66,0.82,0.28,head,size=6.4,color=accent,bold=True,align=PP_ALIGN.CENTER,margin=0.0)
# final statement
text(slide,0.70,3.77,6.15,0.52,"The robot does not replace the maintenance team.\nIt gives the team better evidence, earlier.",size=17.2,color=WHITE,bold=True,font=FONT_DISPLAY,margin=0.0)
# three final points right
finals=[("DETECT EARLIER",CYAN),("LOCATE FASTER",AMBER),("PRIORITIZE SMARTER",GREEN)]
for i,(lab,accent) in enumerate(finals):
    yy=3.82+i*0.57
    rect(slide,7.72,yy,4.82,0.42,PANEL_2,radius=True,line=GRID,lw=0.6)
    dot(slide,7.96,yy+0.16,0.09,accent)
    text(slide,8.22,yy+0.11,4.05,0.18,lab,size=10.2,color=accent,bold=True,margin=0.0)
# closing band
rect(slide,0.66,5.75,11.98,0.63,"14313A",radius=True,line=CYAN_DARK,lw=0.9)
text(slide,0.92,5.94,4.65,0.20,"PERIODIC PATROL  •  LOCAL INTELLIGENCE  •  ACTIONABLE RECORDS",size=8.2,color=CYAN,bold=True,margin=0.0)
text(slide,7.05,5.90,5.17,0.24,"A strong MVP starts with a route — and ends with a decision.",size=10.3,color=WHITE,bold=True,align=PP_ALIGN.RIGHT,margin=0.0)
# tiny final footer
text(slide,0.70,6.80,11.72,0.18,"AUTONOMOUS EDGE-AI ROBOTIC SYSTEM FOR COMPRESSED-AIR LEAK DETECTION",size=7.3,color=MUTED_2,bold=True,align=PP_ALIGN.CENTER,margin=0.0)
add_notes(slide, "Close on the operating model, not on a feature list. The first step is one controlled route and one robot. From there, the same inspection record can support multiple routes, scheduled patrols, fleets, historical analytics, and eventually predictive maintenance workflows. The MVP is intentionally honest: it demonstrates periodic autonomous inspection with local intelligence. The final pitch is the value loop judges should remember: detect earlier, locate faster, prioritize smarter.")

# Metadata and save.
prs.core_properties.title = "Autonomous Edge-AI Robotic System for Compressed-Air Leak Detection"
prs.core_properties.subject = "Automated industrial leak intelligence and maintenance prioritization"
prs.core_properties.author = "Arena.ai Agent Mode"
prs.core_properties.keywords = "compressed air, leak detection, edge AI, robotics, predictive maintenance, hackathon"
prs.core_properties.comments = "Editable vector-first hackathon presentation. Claims are intentionally bounded to the MVP and proposed functionality."
prs.save(OUT)
print(f"Wrote {OUT}")
