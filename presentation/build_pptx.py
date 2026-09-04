#!/usr/bin/env python3
"""Build the PhishGuard 16:9 hackathon pitch deck as an editable PPTX."""

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import nsmap, qn
from pptx.util import Emu, Inches, Pt
from lxml import etree

# --- palette ---
BG = RGBColor(6, 9, 16)
BG_ELEV = RGBColor(12, 18, 32)
CARD = RGBColor(16, 24, 40)
CARD_ALT = RGBColor(15, 23, 42)
LINE = RGBColor(30, 58, 74)
CYAN = RGBColor(34, 211, 238)
CYAN_SOFT = RGBColor(103, 232, 249)
BLUE = RGBColor(96, 165, 250)
TEXT = RGBColor(241, 245, 249)
MUTED = RGBColor(148, 163, 184)
DIM = RGBColor(100, 116, 139)
GREEN = RGBColor(52, 211, 153)
YELLOW = RGBColor(251, 191, 36)
RED = RGBColor(251, 113, 133)
WHITE = RGBColor(255, 255, 255)
BAD_LINE = RGBColor(80, 40, 52)
GOOD_LINE = RGBColor(24, 80, 64)

W = Inches(13.333)
H = Inches(7.5)
FONT = "Calibri"


def _set_run(run, size, color, bold=False, italic=False):
    run.font.size = Pt(size)
    run.font.color.rgb = color
    run.font.bold = bold
    run.font.italic = italic
    run.font.name = FONT
    rPr = run._r.get_or_add_rPr()
    # Force latin/cs/ea font so Google Slides / Office keep Calibri
    for tag in ("a:latin", "a:cs", "a:ea"):
        el = rPr.find(qn(tag))
        if el is None:
            el = etree.SubElement(rPr, qn(tag))
        el.set("typeface", FONT)


def add_text(
    slide,
    l,
    t,
    w,
    h,
    text,
    size=14,
    color=TEXT,
    bold=False,
    italic=False,
    align=PP_ALIGN.LEFT,
    anchor=MSO_ANCHOR.TOP,
    margin=0.08,
):
    box = slide.shapes.add_textbox(Inches(l), Inches(t), Inches(w), Inches(h))
    tf = box.text_frame
    tf.word_wrap = True
    tf.auto_size = None
    tf.margin_left = Inches(margin)
    tf.margin_right = Inches(margin)
    tf.margin_top = Inches(0.04)
    tf.margin_bottom = Inches(0.02)
    try:
        tf._txBody.bodyPr.set("anchor", {MSO_ANCHOR.TOP: "t", MSO_ANCHOR.MIDDLE: "ctr", MSO_ANCHOR.BOTTOM: "b"}[anchor])
    except Exception:
        pass
    lines = text.split("\n")
    for i, line in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        p.space_before = Pt(0)
        p.space_after = Pt(2)
        p.line_spacing = 1.08
        run = p.add_run()
        run.text = line
        _set_run(run, size, color, bold, italic)
    return box


def shape_fill_line(shape, fill, line=None, weight=1.0):
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill
    if line is None:
        shape.line.fill.background()
    else:
        shape.line.color.rgb = line
        shape.line.width = Pt(weight)


def rect(slide, l, t, w, h, fill=CARD, line=LINE, radius=True, weight=1.15):
    kind = MSO_SHAPE.ROUNDED_RECTANGLE if radius else MSO_SHAPE.RECTANGLE
    s = slide.shapes.add_shape(kind, Inches(l), Inches(t), Inches(w), Inches(h))
    shape_fill_line(s, fill, line, weight)
    if radius:
        try:
            s.adjustments[0] = 0.08
        except Exception:
            pass
    s.shadow.inherit = False
    return s


def oval(slide, l, t, w, h, fill, line=None):
    s = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(l), Inches(t), Inches(w), Inches(h))
    shape_fill_line(s, fill, line)
    return s


def set_bg(slide):
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = BG
    # Full-bleed rectangle so Google Slides keeps the dark theme on import
    bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), W, H)
    shape_fill_line(bg, BG, None)
    bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), W, Inches(0.055))
    shape_fill_line(bar, CYAN, None)


def chrome(slide, n, total=12):
    add_text(slide, 0.48, 0.14, 6.5, 0.32, "PHISHGUARD  ·  ON-DEVICE ML", 11, CYAN, True)
    add_text(slide, 11.2, 0.14, 1.7, 0.32, f"{n:02d}  /  {total:02d}", 11, DIM, False, align=PP_ALIGN.RIGHT)


def kicker(slide, text, y=0.48):
    add_text(slide, 0.5, y, 12.2, 0.32, text.upper(), 12, CYAN, True)


def title(slide, text, y=0.78, h=1.15, size=32):
    add_text(slide, 0.5, y, 12.3, h, text, size, TEXT, True)


def lede(slide, text, y, w=11.5):
    add_text(slide, 0.5, y, w, 0.55, text, 16, MUTED, False)


def pill(slide, l, t, w, h, text, fg=TEXT, bg=CARD_ALT, ln=LINE):
    rect(slide, l, t, w, h, bg, ln, True, 1.0)
    add_text(slide, l, t + 0.04, w, h - 0.04, text, 12, fg, True, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, margin=0.04)


def node(slide, l, t, w, h, text, sub="", accent=False, fg=None):
    ln = CYAN if accent else LINE
    fill = RGBColor(10, 32, 42) if accent else CARD
    rect(slide, l, t, w, h, fill, ln, True, 1.25 if accent else 1.0)
    color = fg or (CYAN_SOFT if accent else TEXT)
    if sub:
        add_text(slide, l, t + 0.08, w, 0.28, text, 12, color, True, align=PP_ALIGN.CENTER, margin=0.04)
        add_text(slide, l, t + 0.32, w, 0.28, sub, 10, DIM, False, align=PP_ALIGN.CENTER, margin=0.04)
    else:
        add_text(slide, l, t, w, h, text, 13, color, True, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, margin=0.04)


def arrow(slide, l, t, text="→"):
    add_text(slide, l, t, 0.32, 0.4, text, 18, CYAN, True, align=PP_ALIGN.CENTER)


def card_block(slide, l, t, w, h, heading, body, icon=None, ln=LINE):
    rect(slide, l, t, w, h, CARD, ln)
    y = t + 0.12
    if icon:
        add_text(slide, l + 0.08, y, w - 0.16, 0.32, icon, 14, CYAN, True)
        y += 0.32
    add_text(slide, l + 0.06, y, w - 0.16, 0.36, heading, 15, TEXT, True)
    add_text(slide, l + 0.06, y + 0.34, w - 0.16, h - (y - t) - 0.42, body, 13, MUTED, False)


def bullets(slide, l, t, w, h, items, size=14, color=MUTED, dot=CYAN):
    # one textbox; cyan-ish bullets via "▸ "
    box = slide.shapes.add_textbox(Inches(l), Inches(t), Inches(w), Inches(h))
    tf = box.text_frame
    tf.word_wrap = True
    tf.margin_left = Inches(0.04)
    tf.margin_right = Inches(0.06)
    for i, item in enumerate(items):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.space_before = Pt(3)
        p.space_after = Pt(4)
        p.line_spacing = 1.08
        run = p.add_run()
        run.text = "▸  " + item
        _set_run(run, size, color, False)
    return box


def notes(slide, text):
    slide.notes_slide.notes_text_frame.text = text


def build():
    prs = Presentation()
    prs.slide_width = W
    prs.slide_height = H
    blank = prs.slide_layouts[6]

    # ========== 1 TITLE ==========
    s = prs.slides.add_slide(blank)
    set_bg(s)
    chrome(s, 1)
    kicker(s, "Hackathon pitch  ·  Privacy-first web security")
    add_text(s, 0.5, 0.95, 7.4, 1.9, "Browser-Based\nPhishing Detector", 36, TEXT, True)
    add_text(
        s, 0.5, 2.95, 7.3, 0.9,
        "On-device machine learning for real-time web security.\nDetect phishing websites without sending user data to the cloud.",
        16, MUTED, False,
    )
    pill(s, 0.5, 4.0, 1.7, 0.38, "On-device ML", CYAN_SOFT, RGBColor(10, 32, 42), CYAN)
    pill(s, 2.32, 4.0, 1.45, 0.38, "Real-time", TEXT, CARD, LINE)
    pill(s, 3.88, 4.0, 1.7, 0.38, "Privacy-first", TEXT, CARD, LINE)
    pill(s, 5.68, 4.0, 1.7, 0.38, "Lightweight", TEXT, CARD, LINE)

    add_text(s, 0.5, 4.7, 2.3, 0.22, "TEAM", 10, DIM, True)
    add_text(s, 0.5, 4.92, 2.3, 0.35, "[Team Name]", 14, TEXT, False)
    add_text(s, 2.9, 4.7, 4.4, 0.22, "MEMBERS", 10, DIM, True)
    add_text(s, 2.9, 4.92, 4.4, 0.35, "[Member 1]  ·  [Member 2]  ·  [Member 3]", 14, TEXT, False)
    add_text(s, 0.5, 5.4, 7.0, 0.22, "EVENT", 10, DIM, True)
    add_text(s, 0.5, 5.62, 7.0, 0.35, "[Hackathon Name]", 14, TEXT, False)

    # browser mockup
    rect(s, 8.15, 1.15, 4.7, 5.15, RGBColor(11, 18, 32), CYAN, True, 1.4)
    rect(s, 8.15, 1.15, 4.7, 0.55, RGBColor(17, 24, 39), LINE, False)
    oval(s, 8.35, 1.32, 0.18, 0.18, RED)
    oval(s, 8.58, 1.32, 0.18, 0.18, YELLOW)
    oval(s, 8.81, 1.32, 0.18, 0.18, GREEN)
    rect(s, 9.15, 1.28, 3.5, 0.32, BG, LINE, True)
    add_text(s, 9.2, 1.28, 3.4, 0.32, "⚠  http://paypa1-secure-login.xyz/auth", 10, MUTED, False, anchor=MSO_ANCHOR.MIDDLE)
    add_text(s, 8.4, 1.9, 3.0, 0.4, "Paypa1 Secure Login", 16, TEXT, True)
    add_text(s, 11.55, 1.85, 1.1, 0.5, "🛡", 22, CYAN, True, align=PP_ALIGN.RIGHT)
    rect(s, 8.45, 2.5, 4.1, 0.42, RGBColor(17, 24, 39), LINE, True)
    rect(s, 8.45, 3.05, 4.1, 0.42, RGBColor(17, 24, 39), LINE, True)
    rect(s, 8.45, 3.6, 4.1, 0.42, RGBColor(17, 24, 39), LINE, True)
    rect(s, 8.45, 4.25, 4.1, 1.55, RGBColor(40, 18, 28), RGBColor(120, 50, 64), True, 1.2)
    add_text(s, 8.55, 4.4, 3.9, 1.25, "AI verdict\nLikely phishing\nAnalyzed locally — data never left this device", 14, RGBColor(254, 205, 211), True)
    notes(s, "Open on the privacy hook: we detect phishing in the browser without sending user data to the cloud. Introduce team. ~30s.")

    # ========== 2 PROBLEM ==========
    s = prs.slides.add_slide(blank)
    set_bg(s)
    chrome(s, 2)
    kicker(s, "The problem")
    title(s, "One Click Can Expose Everything.", y=0.78, h=0.7, size=32)
    cards = [
        ("Convincing fakes", "Phishing sites imitate banks, email, and cloud logins people already trust."),
        ("High-value theft", "One form submit can leak passwords, cards, and personal details."),
        ("Cloud trade-offs", "Many detectors send URLs or page data to remote servers."),
        ("No time to wait", "A warning must appear before the user hits “Sign in”."),
    ]
    positions = [(0.5, 1.7), (4.05, 1.7), (0.5, 3.55), (4.05, 3.55)]
    for (l, t), (hd, bd) in zip(positions, cards):
        card_block(s, l, t, 3.4, 1.7, hd, bd)
    rect(s, 7.65, 1.7, 5.2, 3.55, CARD, LINE)
    add_text(s, 7.8, 1.82, 4.9, 0.28, "ATTACK PATH", 11, DIM, True)
    steps = [
        (2.18, "User clicks a link", LINE, TEXT),
        (2.88, "Looks like a trusted site", LINE, TEXT),
        (3.58, "Password / banking details entered", LINE, TEXT),
        (4.28, "Credentials stolen", RGBColor(120, 50, 64), RGBColor(254, 205, 211)),
    ]
    for y, label, ln, fg in steps:
        node(s, 7.95, y, 4.6, 0.55, label, accent=(ln != LINE), fg=fg)
    add_text(s, 7.95, 2.7, 4.6, 0.22, "↓", 14, CYAN, True, align=PP_ALIGN.CENTER)
    add_text(s, 7.95, 3.4, 4.6, 0.22, "↓", 14, CYAN, True, align=PP_ALIGN.CENTER)
    add_text(s, 7.95, 4.1, 4.6, 0.22, "↓", 14, CYAN, True, align=PP_ALIGN.CENTER)
    rect(s, 0.5, 5.45, 12.35, 0.7, RGBColor(10, 32, 42), CYAN, True, 1.3)
    add_text(
        s, 0.65, 5.5, 12.05, 0.6,
        "Key question   ·   How can we detect phishing in real time while keeping user data on the device?",
        16, CYAN_SOFT, True, anchor=MSO_ANCHOR.MIDDLE,
    )
    notes(s, "Make the privacy tension clear. End on the key question. ~35s.")

    # ========== 3 SOLUTION ==========
    s = prs.slides.add_slide(blank)
    set_bg(s)
    chrome(s, 3)
    kicker(s, "Our solution")
    title(s, "Phishing Detection, Directly Inside the Browser.", y=0.78, h=0.85, size=30)
    lede(s, "A Chrome extension that scores the current page locally — then shows a clear, human verdict.", 1.6)
    nodes = [
        (0.5, "Website", "Current tab", False),
        (3.05, "Feature extraction", "URL + page signals", False),
        (5.6, "On-device ML", "No cloud inference", True),
        (8.15, "Risk score", "0–100%", False),
        (10.7, "Verdict", "3 simple states", False),
    ]
    for l, a, b, acc in nodes:
        node(s, l, 2.25, 2.35, 0.85, a, b, acc)
    for l in (2.78, 5.33, 7.88, 10.43):
        arrow(s, l, 2.45)
    pill(s, 0.5, 3.3, 1.55, 0.38, "●  Safe", GREEN, CARD, LINE)
    pill(s, 2.2, 3.3, 2.15, 0.38, "●  Suspicious", YELLOW, CARD, LINE)
    pill(s, 4.5, 3.3, 2.45, 0.38, "●  Likely Phishing", RED, CARD, LINE)
    feats = [
        ("Local", "Runs on the user’s device."),
        ("Private", "No page data required in the cloud."),
        ("Real-time", "Scores the page as it loads."),
        ("Lightweight", "Compact model for the browser."),
        ("Transparent", "Shows the signals behind the call."),
    ]
    for i, (hd, bd) in enumerate(feats):
        card_block(s, 0.5 + i * 2.5, 3.9, 2.38, 2.15, hd, bd)
    notes(s, "Product slide. Hit the four selling points: on-device, private, real-time, lightweight. ~40s.")

    # ========== 4 ARCHITECTURE ==========
    s = prs.slides.add_slide(blank)
    set_bg(s)
    chrome(s, 4)
    kicker(s, "Architecture")
    title(s, "From URL to Verdict in Milliseconds", y=0.78, h=0.65, size=30)
    lede(s, "Every step — extraction, scoring, and the final warning — happens inside the browser. Inference does not require a cloud round-trip.", 1.45)
    rect(s, 0.5, 2.1, 12.35, 4.55, CARD, LINE)
    node(s, 1.1, 2.35, 3.3, 0.7, "Chrome browser", accent=False)
    arrow(s, 4.5, 2.48)
    node(s, 4.85, 2.35, 3.3, 0.7, "Browser extension", accent=False)
    arrow(s, 8.25, 2.48)
    node(s, 8.6, 2.35, 3.55, 0.7, "Feature extraction", accent=True)
    node(s, 1.3, 3.3, 5.15, 0.95, "URL features", "Length, host, TLD, lookalikes, IP, special chars", False)
    node(s, 6.9, 3.3, 5.15, 0.95, "Page content features", "Forms, HTTPS, iframes, brand mismatch", False)
    node(s, 0.9, 4.5, 3.7, 0.7, "30–40 engineered features")
    arrow(s, 4.7, 4.62)
    node(s, 5.05, 4.5, 3.5, 0.7, "On-device ML model", accent=True)
    arrow(s, 8.65, 4.62)
    node(s, 9.0, 4.5, 3.2, 0.7, "Phishing probability")
    node(s, 1.1, 5.5, 3.5, 0.7, "Safe", fg=GREEN)
    node(s, 4.95, 5.5, 3.5, 0.7, "Suspicious", fg=YELLOW)
    node(s, 8.8, 5.5, 3.5, 0.7, "Likely phishing", fg=RED)
    notes(s, "Walk the pipeline once, slowly. Emphasize local inference. ~40s.")

    # ========== 5 SIGNALS ==========
    s = prs.slides.add_slide(blank)
    set_bg(s)
    chrome(s, 5)
    kicker(s, "Signals")
    title(s, "30–40 Signals Behind Every Decision", y=0.78, h=0.65, size=30)
    rect(s, 0.5, 1.7, 6.05, 4.9, CARD, LINE)
    add_text(s, 0.7, 1.85, 5.7, 0.25, "URL-BASED FEATURES", 11, CYAN, True)
    add_text(s, 0.7, 2.15, 5.7, 0.4, "What the address reveals", 18, TEXT, True)
    bullets(s, 0.7, 2.65, 5.6, 3.7, [
        "URL length and nested path structure",
        "Number of subdomains and suspicious TLDs",
        "Special characters and encoding tricks",
        "IP address used instead of a domain",
        "Typosquatting / lookalike brand spellings",
        "Character similarity to known brands",
        "Suspicious URL structures overall",
    ], 15, MUTED)
    rect(s, 6.8, 1.7, 6.05, 4.9, CARD, LINE)
    add_text(s, 7.0, 1.85, 5.7, 0.25, "PAGE CONTENT FEATURES", 11, CYAN, True)
    add_text(s, 7.0, 2.15, 5.7, 0.4, "What the page is asking for", 18, TEXT, True)
    bullets(s, 7.0, 2.65, 5.6, 3.7, [
        "Login forms and password fields",
        "Credential-harvesting indicators",
        "HTTP / non-secure connections",
        "Embedded iframes",
        "Brand vs. content mismatch",
        "Suspicious form destinations",
        "Odd post-load form behavior",
    ], 15, MUTED)
    notes(s, "Don’t list every feature. Pick 2 URL + 2 page examples a judge can picture. ~35s.")

    # ========== 6 ML ==========
    s = prs.slides.add_slide(blank)
    set_bg(s)
    chrome(s, 6)
    kicker(s, "Machine learning  ·  model-selection stage")
    title(s, "Small Model. Fast Decision.", y=0.78, h=0.6, size=32)
    lede(s, "Trained offline, then shipped with the extension. Inference stays on-device. Final architecture is still being selected — we are comparing two compact options.", 1.4, w=12.2)
    rect(s, 0.5, 2.1, 6.05, 2.55, CARD, LINE)
    add_text(s, 0.7, 2.22, 5.7, 0.25, "CANDIDATE A", 11, CYAN, True)
    add_text(s, 0.7, 2.5, 5.7, 0.4, "Gradient-boosted decision tree", 18, TEXT, True)
    bullets(s, 0.7, 3.0, 5.6, 1.5, [
        "Strong on tabular URL / page features",
        "Stays compact after training",
        "Easier to inspect which signals matter",
    ], 14)
    rect(s, 6.8, 2.1, 6.05, 2.55, CARD, LINE)
    add_text(s, 7.0, 2.22, 5.7, 0.25, "CANDIDATE B", 11, CYAN, True)
    add_text(s, 7.0, 2.5, 5.7, 0.4, "Compact neural network", 18, TEXT, True)
    bullets(s, 7.0, 3.0, 5.6, 1.5, [
        "Small enough for browser runtimes",
        "Fast forward-pass on CPU",
        "Fits ONNX / TF.js-style export",
    ], 14)
    node(s, 0.5, 4.9, 2.55, 0.7, "30–40 features")
    arrow(s, 3.1, 5.02)
    node(s, 3.4, 4.9, 2.9, 0.7, "Selected compact model", accent=True)
    arrow(s, 6.35, 5.02)
    node(s, 6.7, 4.9, 2.55, 0.7, "Probability score")
    arrow(s, 9.3, 5.02)
    node(s, 9.65, 4.9, 3.2, 0.7, "Thresholds → verdict")
    pill(s, 0.5, 5.8, 1.85, 0.38, "Trained offline", TEXT, CARD, LINE)
    pill(s, 2.5, 5.8, 2.2, 0.38, "No cloud inference", TEXT, CARD, LINE)
    pill(s, 4.85, 5.8, 2.15, 0.38, "Low-latency design", TEXT, CARD, LINE)
    pill(s, 7.15, 5.8, 2.55, 0.38, "Feature-based = explainable", TEXT, CARD, LINE)
    notes(s, "Be honest: model not locked. Sell compactness, locality, interpretability. ~35s.")

    # ========== 7 PRIVACY ==========
    s = prs.slides.add_slide(blank)
    set_bg(s)
    chrome(s, 7)
    kicker(s, "Privacy by design")
    title(s, "Your Data Stays on Your Device.", y=0.78, h=0.65, size=32)
    rect(s, 0.5, 1.7, 6.05, 4.9, CARD, BAD_LINE, True, 1.4)
    add_text(s, 0.7, 1.88, 5.7, 0.25, "TRADITIONAL CLOUD DETECTION", 11, RED, True)
    add_text(s, 0.7, 2.2, 5.7, 0.45, "Page data leaves the browser", 20, TEXT, True)
    steps_bad = ["User", "Website information", "Cloud server", "Detection", "Result"]
    for i, lab in enumerate(steps_bad):
        node(s, 0.85, 2.8 + i * 0.68, 5.35, 0.55, lab, fg=MUTED)
    rect(s, 6.8, 1.7, 6.05, 4.9, CARD, GOOD_LINE, True, 1.4)
    add_text(s, 7.0, 1.88, 5.7, 0.25, "OUR APPROACH", 11, GREEN, True)
    add_text(s, 7.0, 2.2, 5.7, 0.45, "Inference never needs the cloud", 20, TEXT, True)
    steps_good = ["User", "Website information", "Local ML model", "Result"]
    for i, lab in enumerate(steps_good):
        acc = lab == "Local ML model"
        node(s, 7.15, 2.8 + i * 0.55, 5.35, 0.48, lab, accent=acc, fg=CYAN_SOFT if acc else TEXT)
    bullets(s, 7.15, 5.05, 5.4, 1.4, [
        "No URL required for inference",
        "No webpage content uploaded",
        "Works with weaker connectivity · faster local decisions",
    ], 13, MUTED)
    notes(s, "This is the differentiator. Contrast cloud upload vs local model. ~40s.")

    # ========== 8 UX ==========
    s = prs.slides.add_slide(blank)
    set_bg(s)
    chrome(s, 8)
    kicker(s, "User experience  ·  example UI  ·  not a benchmark")
    title(s, "What the User Sees", y=0.78, h=0.55, size=32)
    lede(s, "Simple language. One glance. Built for a non-technical user.", 1.35)
    # popup
    rect(s, 0.5, 1.95, 5.5, 4.7, RGBColor(16, 24, 38), CYAN, True, 1.4)
    add_text(s, 0.7, 2.08, 3.4, 0.3, "🛡  PhishGuard", 14, TEXT, True)
    pill(s, 4.15, 2.1, 1.6, 0.32, "LIVE EXAMPLE", RED, RGBColor(40, 18, 28), RED)
    add_text(s, 0.7, 2.5, 5.1, 0.28, "http://paypa1-secure-login.xyz", 12, MUTED, False)
    add_text(s, 0.7, 2.85, 5.1, 0.4, "⚠  LIKELY PHISHING", 20, RED, True)
    add_text(s, 0.7, 3.3, 2.2, 0.25, "RISK SCORE", 11, DIM, True)
    add_text(s, 3.7, 3.2, 2.0, 0.5, "94%", 28, RED, True, align=PP_ALIGN.RIGHT)
    rect(s, 0.7, 3.75, 5.05, 0.14, RGBColor(30, 41, 59), None, True)
    rect(s, 0.7, 3.75, 4.75, 0.14, RGBColor(251, 113, 133), None, True)
    add_text(s, 0.7, 4.05, 5.1, 0.25, "WHY THIS PAGE", 11, DIM, True)
    bullets(s, 0.7, 4.3, 5.1, 2.1, [
        "Suspicious URL structure",
        "Brand similarity detected",
        "Credential form found",
        "Non-secure connection",
    ], 14, MUTED)
    # statuses
    statuses = [
        (GREEN, "SAFE — Low Risk", "Looks consistent with a legitimate site."),
        (YELLOW, "SUSPICIOUS — Review Carefully", "Some signals don’t add up. Pause before you sign in."),
        (RED, "LIKELY PHISHING — Avoid This Site", "Do not enter passwords or payment details."),
    ]
    for i, (col, hd, bd) in enumerate(statuses):
        y = 1.95 + i * 1.55
        rect(s, 6.3, y, 6.55, 1.4, CARD, LINE)
        oval(s, 6.5, y + 0.55, 0.22, 0.22, col)
        add_text(s, 6.9, y + 0.22, 5.7, 0.4, hd, 16, col, True)
        add_text(s, 6.9, y + 0.65, 5.7, 0.55, bd, 14, MUTED, False)
    notes(s, "Show the mock UI. 94% is an example verdict for the popup, not a measured accuracy. ~35s.")

    # ========== 9 EVALUATION ==========
    s = prs.slides.add_slide(blank)
    set_bg(s)
    chrome(s, 9)
    kicker(s, "Evaluation  ·  targets, not claimed results")
    title(s, "We Measure What Actually Matters.", y=0.78, h=0.6, size=30)
    metrics = [
        ("PRECISION", "Target", "Of flagged sites, how many were truly phishing?"),
        ("RECALL", "Priority", "Of real phishing sites, how many did we catch?"),
        ("F1 SCORE", "Balance", "Single view of precision and recall together."),
        ("LATENCY", "ms-class", "Time from page load to on-device verdict."),
        ("MODEL SIZE", "Compact", "Small enough to ship inside the extension."),
    ]
    for i, (k, n, b) in enumerate(metrics):
        x = 0.5 + i * 2.5
        rect(s, x, 1.6, 2.38, 2.55, CARD, LINE)
        add_text(s, x + 0.12, 1.75, 2.14, 0.25, k, 11, DIM, True)
        add_text(s, x + 0.12, 2.1, 2.14, 0.45, n, 20, CYAN_SOFT, True)
        add_text(s, x + 0.12, 2.6, 2.14, 1.3, b, 13, MUTED, False)
    rect(s, 0.5, 4.35, 12.35, 1.15, RGBColor(10, 32, 42), CYAN, True, 1.2)
    add_text(
        s, 0.7, 4.45, 12.0, 0.95,
        "Recall is the safety metric.  A missed phishing page (false negative) can expose a password.\nWe will tune thresholds to catch attacks first, then reduce false alarms.",
        16, TEXT, False, anchor=MSO_ANCHOR.MIDDLE,
    )
    pill(s, 0.5, 5.75, 3.15, 0.4, "Target detection quality: 95%+", CYAN_SOFT, RGBColor(10, 32, 42), CYAN)
    pill(s, 3.8, 5.75, 3.05, 0.4, "Not yet experimentally verified", MUTED, CARD, LINE)
    pill(s, 7.0, 5.75, 2.55, 0.4, "Held-out test set planned", MUTED, CARD, LINE)
    notes(s, "Do not invent numbers. Say “target 95%+” clearly. Stress recall. ~30s.")

    # ========== 10 DATA ==========
    s = prs.slides.add_slide(blank)
    set_bg(s)
    chrome(s, 10)
    kicker(s, "Data  ·  planned public sources")
    title(s, "Built and Tested Against Real-World Data", y=0.78, h=0.6, size=30)
    sources = [
        ("PHISHING", "PhishTank", "Community-verified phishing URLs as malicious examples."),
        ("PHISHING", "OpenPhish", "Additional live phishing feeds for coverage and freshness."),
        ("BENIGN", "Tranco", "Popular legitimate websites as the safe class."),
    ]
    for i, (k, hd, bd) in enumerate(sources):
        x = 0.5 + i * 4.2
        rect(s, x, 1.55, 4.0, 1.85, CARD, LINE)
        add_text(s, x + 0.18, 1.68, 3.65, 0.25, k, 11, CYAN, True)
        add_text(s, x + 0.18, 1.98, 3.65, 0.4, hd, 20, TEXT, True)
        add_text(s, x + 0.18, 2.45, 3.65, 0.75, bd, 14, MUTED, False)
    pipe = [
        (0.5, "Dataset"),
        (3.05, "Feature extraction"),
        (5.6, "Train / test split"),
        (8.15, "Model training"),
        (10.7, "Held-out eval"),
    ]
    for i, (x, lab) in enumerate(pipe):
        node(s, x, 3.65, 2.35, 0.7, lab, accent=(lab == "Held-out eval"))
        if i < 4:
            arrow(s, x + 2.28, 3.78)
    card_block(s, 0.5, 4.6, 6.05, 1.55, "Consistent features", "The same 30–40 signals are computed for every URL and page snapshot.")
    card_block(s, 6.8, 4.6, 6.05, 1.55, "Generalization first", "We score the held-out set — not training accuracy — so results reflect new sites.")
    notes(s, "Name the three sources. Stress held-out evaluation, not training score. ~30s.")

    # ========== 11 LIMITS ==========
    s = prs.slides.add_slide(blank)
    set_bg(s)
    chrome(s, 11)
    kicker(s, "Honesty & roadmap")
    title(s, "Honest About the Gaps. Ready to Improve.", y=0.78, h=0.65, size=28)
    rect(s, 0.5, 1.6, 6.05, 4.0, CARD, LINE)
    add_text(s, 0.7, 1.75, 5.7, 0.3, "CURRENT LIMITATIONS", 11, YELLOW, True)
    bullets(s, 0.7, 2.15, 5.6, 3.25, [
        "Cannot reliably check live domain age without an external service",
        "Focused on web-based phishing, not every scam channel",
        "Sophisticated attacks may still evade feature-based detection",
        "Dynamic pages can change after the first analysis",
        "False positives are possible",
    ], 14, MUTED)
    rect(s, 6.8, 1.6, 6.05, 4.0, CARD, LINE)
    add_text(s, 7.0, 1.75, 5.7, 0.3, "FUTURE IMPROVEMENTS", 11, CYAN, True)
    bullets(s, 7.0, 2.15, 5.6, 3.25, [
        "Deeper behavioral analysis after load",
        "Stronger typosquatting detection",
        "Optional reputation lists (user-controlled)",
        "Continuous model improvement",
        "Richer social-engineering patterns and more browsers",
    ], 14, MUTED)
    rect(s, 0.5, 5.8, 12.35, 0.7, RGBColor(10, 32, 42), CYAN, True, 1.2)
    add_text(
        s, 0.65, 5.85, 12.05, 0.6,
        "Hackathon-ready core: extract signals locally → score with a compact model → warn before submit. Extras stay optional so privacy remains the default.",
        14, CYAN_SOFT, False, anchor=MSO_ANCHOR.MIDDLE,
    )
    notes(s, "This slide shows maturity. Name 2 limits, 2 next steps, then the feasible core. ~30s.")

    # ========== 12 CLOSE ==========
    s = prs.slides.add_slide(blank)
    set_bg(s)
    chrome(s, 12)
    kicker(s, "Impact")
    add_text(s, 0.5, 0.85, 12.3, 0.7, "Security Without Sacrificing Privacy.", 32, TEXT, True, align=PP_ALIGN.CENTER)
    trio = [
        ("DETECT", "See it in time", "Identify suspicious phishing patterns in real time, in the tab the user is already on."),
        ("PROTECT", "Warn before submit", "Give a clear Safe / Suspicious / Phishing call before sensitive data is entered."),
        ("PRESERVE", "Keep it local", "Run detection and ML inference on the device — not on someone else’s server."),
    ]
    for i, (k, hd, bd) in enumerate(trio):
        x = 0.5 + i * 4.2
        rect(s, x, 1.7, 4.0, 2.15, CARD, LINE)
        add_text(s, x + 0.2, 1.85, 3.6, 0.25, k, 12, CYAN, True)
        add_text(s, x + 0.2, 2.15, 3.6, 0.4, hd, 18, TEXT, True)
        add_text(s, x + 0.2, 2.6, 3.6, 1.0, bd, 14, MUTED, False)
    rect(s, 1.4, 4.1, 10.5, 0.7, RGBColor(10, 32, 42), CYAN, True, 1.3)
    add_text(
        s, 1.55, 4.15, 10.2, 0.6,
        "Effective phishing detection doesn't have to mean sending user data to the cloud.",
        16, CYAN_SOFT, True, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE,
    )
    add_text(s, 0.5, 5.0, 12.3, 0.55, "Thank You", 32, TEXT, True, align=PP_ALIGN.CENTER)
    add_text(s, 0.5, 5.5, 12.3, 0.35, "Questions?", 18, CYAN, True, align=PP_ALIGN.CENTER)
    add_text(
        s, 0.5, 5.95, 12.3, 0.4,
        "[Team Name]   ·   github.com/[your-team]/phishguard   ·   [contact@email]",
        14, MUTED, False, align=PP_ALIGN.CENTER,
    )
    notes(s, "Close on Detect / Protect / Preserve. Invite questions. ~25s.")

    out = "/home/user/Arduino-learning-journey../presentation/PhishGuard-Hackathon-Pitch.pptx"
    prs.save(out)
    print(out)


if __name__ == "__main__":
    build()
