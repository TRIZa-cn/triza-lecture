# -*- coding: utf-8 -*-
"""Build the TRIZa V8 15-minute conference deck (English, 20pt floor)."""
import sys
from pathlib import Path

from pptx import Presentation
from pptx.util import Inches as In, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
from notes_v8 import NOTES  # noqa: E402


def C(h):
    return RGBColor.from_string(h)


NAVY = C("0E2A47"); INK = C("16283C"); BLUE = C("1E6BA8"); SKY = C("4E93C4")
AMBER = C("D98A26"); CORAL = C("BE5138"); GREEN = C("2B7A63"); MUTED = C("6C7C8F")
LINE = C("D9E2EB"); SOFT = C("F4F8FC"); SOFT2 = C("E7EFF7")
WHITE = C("FFFFFF"); GREY = C("9AA8B8")
FONT = "Segoe UI"

LM = 0.72
CW = 11.893
FOOT_Y = 6.98

prs = Presentation()
prs.slide_width = In(13.3333)
prs.slide_height = In(7.5)
BLANK = prs.slide_layouts[6]


def new_slide():
    s = prs.slides.add_slide(BLANK)
    f = s.background.fill
    f.solid()
    f.fore_color.rgb = WHITE
    return s


def _fill_para(para, p, align):
    para.alignment = p.get("al", align)
    if p.get("ls"):
        para.line_spacing = p["ls"]
    if p.get("sp"):
        para.space_before = Pt(p["sp"])
    r = para.add_run()
    r.text = p["t"]
    f = r.font
    f.name = FONT
    f.size = Pt(p.get("sz", 20))
    f.bold = bool(p.get("b", False))
    f.color.rgb = p.get("c", INK)
    if p.get("spc"):
        r._r.get_or_add_rPr().set("spc", str(int(p["spc"] * 100)))
    return r


def text(s, x, y, w, h, paras, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP, name=None):
    tb = s.shapes.add_textbox(In(x), In(y), In(w), In(h))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = In(0.02)
    tf.margin_top = tf.margin_bottom = 0
    tf.vertical_anchor = anchor
    for i, p in enumerate(paras):
        para = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        _fill_para(para, p, align)
    if name:
        tb.name = name
    return tb


def shape(s, x, y, w, h, fill, line=None, kind=MSO_SHAPE.RECTANGLE, radius=None, lw=1.0, name=None):
    sh = s.shapes.add_shape(kind, In(x), In(y), In(w), In(h))
    if fill is None:
        sh.fill.background()
    else:
        sh.fill.solid()
        sh.fill.fore_color.rgb = fill
    if line is None:
        sh.line.fill.background()
    else:
        sh.line.color.rgb = line
        sh.line.width = Pt(lw)
    if radius is not None and kind == MSO_SHAPE.ROUNDED_RECTANGLE:
        try:
            sh.adjustments[0] = radius
        except Exception:
            pass
    sh.shadow.inherit = False
    if name:
        sh.name = name
    return sh


def card(s, x, y, w, h, fill=WHITE, line=LINE, radius=0.05, lw=1.0):
    return shape(s, x, y, w, h, fill, line, MSO_SHAPE.ROUNDED_RECTANGLE, radius, lw)


def oval(s, x, y, w, h, fill, line=None, lw=1.0):
    return shape(s, x, y, w, h, fill, line, MSO_SHAPE.OVAL, None, lw)


def label(sp, paras, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE,
          margins=(0.10, 0.10, 0.02, 0.02)):
    tf = sp.text_frame
    tf.word_wrap = True
    tf.margin_left = In(margins[0])
    tf.margin_right = In(margins[1])
    tf.margin_top = In(margins[2])
    tf.margin_bottom = In(margins[3])
    tf.vertical_anchor = anchor
    for i, p in enumerate(paras):
        para = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        _fill_para(para, p, align)
    return sp


def chip(s, x, y, w, h, txt, fill=NAVY, color=WHITE, sz=20, b=True):
    sp = card(s, x, y, w, h, fill=fill, line=None, radius=0.22)
    label(sp, [{"t": txt, "sz": sz, "b": b, "c": color}])
    return sp


def head(s, kicker, title, sub=None):
    shape(s, LM, 0.40, 0.055, 0.40, AMBER)
    text(s, LM + 0.20, 0.37, CW - 0.2, 0.44,
         [{"t": kicker, "sz": 20, "c": MUTED, "spc": 1.0}], name="kicker")
    text(s, LM, 0.80, CW, 0.66,
         [{"t": title, "sz": 30, "b": True, "c": NAVY, "ls": 1.06}], name="title")
    if sub:
        text(s, LM, 1.44, CW, 0.46,
             [{"t": sub, "sz": 20, "c": MUTED, "ls": 1.1}], name="sub")


def foot(s, section, page):
    shape(s, LM, FOOT_Y, CW, 0.012, LINE)
    text(s, LM, FOOT_Y + 0.07, 9.0, 0.42, [{"t": section, "sz": 20, "c": GREY}])
    text(s, LM + CW - 2.6, FOOT_Y + 0.07, 2.6, 0.42,
         [{"t": "%02d / %d" % (page, 20), "sz": 20, "c": GREY}], align=PP_ALIGN.RIGHT)


def band(s, y, txt, color=AMBER, fill=SOFT, h=0.72, sz=22, b=True):
    card(s, LM, y, CW, h, fill=fill, line=None, radius=0.10)
    shape(s, LM, y, 0.075, h, color)
    text(s, LM + 0.28, y, CW - 0.5, h,
         [{"t": txt, "sz": sz, "b": b, "c": INK}], anchor=MSO_ANCHOR.MIDDLE)


def bullets(s, x, y, w, h, items, sz=20, gap=6, color=INK, dot=BLUE):
    paras = []
    for i, it in enumerate(items):
        paras.append({"t": it, "sz": sz, "c": color, "ls": 1.14, "sp": 0 if i == 0 else gap})
    tb = text(s, x, y, w, h, paras)
    for para in tb.text_frame.paragraphs:
        pPr = para._p.get_or_add_pPr()
        pPr.set("marL", str(int(In(0.26))))
        pPr.set("indent", str(int(-In(0.26))))
    return tb

# ---------------------------------------------------------------- 01 title
def p01(s):
    shape(s, 0, 0, 13.3333, 0.16, NAVY)
    shape(s, LM, 1.06, 0.06, 0.98, AMBER)
    text(s, LM + 0.26, 1.02, CW - 0.3, 0.46,
         [{"t": "AI4Science 2026  \u00b7  Shenzhen, China  \u00b7  October 23-25", "sz": 20, "c": MUTED, "spc": 0.6}])
    text(s, LM, 1.58, 11.3, 1.50, [
        {"t": "TRIZ^a: Guiding Agent Evolution", "sz": 36, "b": True, "c": NAVY, "ls": 1.06},
        {"t": "from Pattern Recognition to Solution Invention", "sz": 36, "b": True, "c": NAVY, "ls": 1.06},
    ])
    shape(s, LM, 3.22, 2.20, 0.05, AMBER)
    text(s, LM, 3.42, 10.6, 0.90, [
        {"t": "Wenyin Liu (GDUT)   \u00b7   Yiheng Huang (BUPT)   \u00b7   Kai Wang (Beijing Denglu Technology)",
         "sz": 20, "c": INK, "ls": 1.2},
    ])
    text(s, LM, 4.02, 10.6, 0.90, [
        {"t": "Evaluation feedback says whether a candidate worked.", "sz": 22, "c": BLUE, "ls": 1.2},
        {"t": "TRIZ says which edit to try next.", "sz": 22, "c": BLUE, "ls": 1.2},
    ])
    card(s, LM, 5.40, CW, 1.24, fill=SOFT, line=None, radius=0.08)
    text(s, LM + 0.30, 5.40, CW - 0.6, 1.24, [
        {"t": "Paper + supplementary appendix A-H + contracts, code, measurements and decision ledgers:",
         "sz": 20, "c": MUTED, "ls": 1.15},
        {"t": "triza.cn/triza-paper-exp-data2026", "sz": 20, "b": True, "c": NAVY, "ls": 1.15},
    ], anchor=MSO_ANCHOR.MIDDLE)
    foot(s, "TRIZa V8  \u00b7  AI4Science 2026", 1)

# ---------------------------------------------------------------- 02 the gap
def p02(s):
    head(s, "02 · THE GAP", "A score says it failed. Not what to change.")
    card(s, LM, 1.90, 5.55, 3.35)
    text(s, LM + 0.28, 2.02, 5.0, 0.42,
         [{"t": "The loop you already have", "sz": 20, "b": True, "c": MUTED, "spc": 0.6}])
    steps = ["propose a candidate", "execute it", "evaluate it", "propose again"]
    for i, t in enumerate(steps):
        y = 2.46 + i * 0.56
        sp = oval(s, LM + 0.30, y + 0.06, 0.42, 0.42, SOFT2)
        label(sp, [{"t": str(i + 1), "sz": 20, "b": True, "c": BLUE}], margins=(0, 0, 0, 0))
        text(s, LM + 0.86, y, 4.3, 0.54, [{"t": t, "sz": 20, "c": INK}],
             anchor=MSO_ANCHOR.MIDDLE)
    text(s, LM + 0.30, 4.74, 5.0, 0.42,
         [{"t": "and each step is a guess", "sz": 20, "b": True, "c": CORAL}],
         anchor=MSO_ANCHOR.MIDDLE)
    card(s, 6.55, 1.90, 6.06, 3.35)
    text(s, 6.83, 2.02, 5.5, 0.42,
         [{"t": "What the score cannot tell you", "sz": 20, "b": True, "c": MUTED, "spc": 0.6}])
    bullets(s, 6.83, 2.52, 5.5, 2.5, [
        "Which property to trade against which.",
        "Whether the gain came from leakage, compute, or changed conditions.",
        "Whether a better number is permission to ship.",
    ], gap=10)
    band(s, 5.45, "Evaluation localises the weakness. It does not name the operation.")
    foot(s, "02 · The gap", 2)


# ---------------------------------------------------------------- 03 commitments
def p03(s):
    head(s, "03 \u00b7 THIS PAPER", "Three commitments, one workflow")
    items = [
        ("01", "A named direction", "Contradiction, principle and typed transformation, fixed before the candidate."),
        ("02", "A measurement of harm", "Functional information and Ideality, against a fixed reference population."),
        ("03", "An authorisation boundary", "A better metric is not permission to replace the running system."),
    ]
    for i, (n, h1, b1) in enumerate(items):
        x = LM + i * 4.07
        card(s, x, 1.90, 3.75, 3.58)
        chip(s, x + 0.28, 2.06, 0.78, 0.58, n, fill=NAVY)
        text(s, x + 0.28, 2.74, 3.20, 0.82,
             [{"t": h1, "sz": 22, "b": True, "c": NAVY, "ls": 1.1}])
        text(s, x + 0.28, 3.58, 3.22, 1.86,
             [{"t": b1, "sz": 20, "c": INK, "ls": 1.18}])
    band(s, 5.68, "Everything after this slide is evidence about how those three behave.",
         color=BLUE, fill=SOFT2)
    foot(s, "03 \u00b7 This paper", 3)

# ---------------------------------------------------------------- 04 TRIZ
def p04(s):
    head(s, "04 \u00b7 BACKGROUND", "TRIZ gives names, not recipes")
    cols = [
        ("Contradiction", "Improving A costs you B. The pair is stated, not assumed."),
        ("40 principles", "Named directions: segmentation, local quality, prior action, parameter change."),
        ("Ideality", "Useful function over harm and cost. A ratio, not a score."),
    ]
    for i, (h1, b1) in enumerate(cols):
        x = LM + i * 4.18
        card(s, x, 2.00, 3.30, 2.38, fill=SOFT if i == 1 else WHITE)
        text(s, x + 0.24, 2.14, 2.85, 0.46, [{"t": h1, "sz": 20, "b": True, "c": NAVY}])
        text(s, x + 0.24, 2.68, 2.86, 1.64,
             [{"t": b1, "sz": 20, "c": INK, "ls": 1.16}])
    text(s, LM + 3.34, 3.00, 0.80, 0.44, [{"t": ">", "sz": 24, "b": True, "c": AMBER}],
         align=PP_ALIGN.CENTER)
    text(s, LM + 7.52, 3.00, 0.80, 0.44, [{"t": ">", "sz": 24, "b": True, "c": AMBER}],
         align=PP_ALIGN.CENTER)
    band(s, 4.70, "A principle does not specify an estimator or a value. The LLM picks it; the frozen protocol tests it.")
    text(s, LM, 5.70, CW, 0.90, [
        {"t": "Our contribution is not the theory. It is making the theory executable inside a loop that can run code and measure results.",
         "sz": 20, "c": MUTED, "ls": 1.2},
    ])
    foot(s, "04 \u00b7 Background", 4)

# ---------------------------------------------------------------- 05 architecture
def p05(s):
    head(s, "05 · FRAMEWORK", "Four stages, three layers")
    w = 10.20
    h = w * 870.0 / 1769.0
    x = LM + (CW - w) / 2.0
    s.shapes.add_picture(str(HERE / "assets" / "arch.png"), In(x), In(1.50), In(w), In(h))
    text(s, LM, 1.50 + h + 0.06, CW, 0.40, [
        {"t": "Layer 1 formulate   ·   Layer 2 transform   ·   Layer 3 control",
         "sz": 20, "c": MUTED, "spc": 0.6},
    ], align=PP_ALIGN.CENTER)
    foot(s, "05 · Framework", 5)


# ---------------------------------------------------------------- 06 stage one
def p06(s):
    head(s, "06 \u00b7 STAGE ONE", "Fix the question before the answer")
    items = [
        ("HIL brief", "Target function and constraints that cannot move."),
        ("Contradiction", "Improving and worsening property, with testable criteria."),
        ("Frozen protocol", "Split, metric and procedure fixed before generation."),
        ("AKBank", "Source-tagged mappings, plus prior outcomes."),
    ]
    for i, (h1, b1) in enumerate(items):
        x = LM + (i % 2) * 5.83
        y = 1.88 + (i // 2) * 1.70
        card(s, x, y, 5.55, 1.56)
        chip(s, x + 0.24, y + 0.22, 0.58, 0.46, str(i + 1), fill=AMBER, sz=20)
        text(s, x + 0.96, y + 0.16, 4.40, 0.42,
             [{"t": h1, "sz": 20, "b": True, "c": NAVY}])
        text(s, x + 0.96, y + 0.62, 4.42, 0.88,
             [{"t": b1, "sz": 20, "c": INK, "ls": 1.14}])
    band(s, 5.40, "If the measurement is decided after the result, no later claim is safe.")
    foot(s, "06 \u00b7 Stage one", 6)

# ---------------------------------------------------------------- 07 typed proposal
def p07(s):
    head(s, "07 · THE OBJECT", "The candidate is a typed proposal")
    card(s, LM, 1.86, CW, 1.56, fill=SOFT, line=LINE, radius=0.06)
    text(s, LM + 0.30, 1.98, CW - 0.6, 1.34, [
        {"t": "\u03c0 = ( id(s), c, r, \u03c4_r, \u0394, \u03ba )", "sz": 26, "b": True, "c": NAVY, "ls": 1.14},
        {"t": "s\u2032 = \u03c4_r ( s ; \u0394 )", "sz": 26, "b": True, "c": BLUE, "ls": 1.14},
    ])
    fields = ["input / output types", "permitted operations", "protected conditions", "required checks"]
    for i, t in enumerate(fields):
        x = LM + i * 3.05
        sp = card(s, x, 3.66, 2.75, 0.66, fill=WHITE, line=SKY, radius=0.24)
        label(sp, [{"t": t, "sz": 20, "b": True, "c": NAVY}])
    text(s, LM, 4.42, CW, 0.40,
         [{"t": "What the contract has to state before generation:", "sz": 20, "c": MUTED}])
    band(s, 5.16, "A favourable result cannot retrospectively make an untyped edit TRIZ-guided.",
         color=CORAL, fill=C("FBEEEA"), h=0.80)
    text(s, LM, 6.12, CW, 0.72, [
        {"t": "id(s) parent  \u00b7  c contradiction  \u00b7  r principle  \u00b7  \u03c4_r declared operator  \u00b7  \u0394 arguments  \u00b7  \u03ba frozen protocol",
         "sz": 20, "c": MUTED, "ls": 1.14},
    ])
    foot(s, "07 · The object", 7)


# ---------------------------------------------------------------- 08 four directions
def p08(s):
    head(s, "08 · DIRECTIONS", "Four directions, made executable")
    text(s, LM, 1.96, 3.0, 0.42, [{"t": "Principle", "sz": 20, "b": True, "c": MUTED}])
    text(s, 3.98, 1.96, 8.6, 0.42, [{"t": "Allowed transformation class", "sz": 20, "b": True, "c": MUTED}])
    shape(s, LM, 2.44, CW, 0.014, LINE)
    rows = [
        ("Segmentation (1)", "Partition components or decisions behind a stable interface."),
        ("Local Quality (3)", "Assign context-specific behaviour under a declared context."),
        ("Prior Action (10)", "Prepare reusable state before the critical path."),
        ("Parameter Change (35)", "Adjust a type, threshold, loss, weight or operating point."),
    ]
    for i, (a, b1) in enumerate(rows):
        y = 2.52 + i * 0.90
        text(s, LM, y, 3.05, 0.80, [{"t": a, "sz": 20, "b": True, "c": NAVY, "ls": 1.14}],
             anchor=MSO_ANCHOR.MIDDLE)
        text(s, 3.98, y, 8.62, 0.80, [{"t": b1, "sz": 20, "c": INK, "ls": 1.14}],
             anchor=MSO_ANCHOR.MIDDLE)
        if i < 3:
            shape(s, LM, y + 0.90, CW, 0.010, LINE)
    band(s, 6.14, "Full 40-class map: Appendix B.  The experiments evaluate selected instances, not all 40.",
         color=AMBER, fill=SOFT, h=0.66, sz=20)
    foot(s, "08 · Directions", 8)


# ---------------------------------------------------------------- 09 four steps
def p09(s):
    head(s, "09 · PROCEDURE", "Selecting an applicable instance")
    steps = [
        ("1", "State the useful property, the property that may worsen, and what is protected."),
        ("2", "Retrieve principles with sources. A matrix hit records the pair; a fallback says so."),
        ("3", "State how the operation addresses benefit and harm, with its inputs and prohibited changes."),
        ("4", "Record the principle and the arguments before evaluation. A missing rationale cannot be added later."),
    ]
    for i, (n, t) in enumerate(steps):
        y = 1.86 + i * 1.24
        text(s, LM, y - 0.06, 0.80, 0.70, [{"t": n, "sz": 40, "b": True, "c": AMBER}])
        text(s, LM + 0.98, y, 10.85, 0.90, [{"t": t, "sz": 20, "c": INK, "ls": 1.16}],
             anchor=MSO_ANCHOR.MIDDLE)
        if i < 3:
            shape(s, LM + 0.98, y + 1.10, 10.85, 0.010, LINE)
    foot(s, "09 · Procedure", 9)


# ---------------------------------------------------------------- 10 functional information
def p10(s):
    head(s, "10 \u00b7 MEASUREMENT", "Functional information: rarity, not quality")
    cards = [
        ("Useful FI", "I_U = -log2[ (1 + h_U) / (1 + N) ]",
         "h_U counts references at or above the candidate's useful level."),
        ("Harm FI", "I_H = -log2[ (1 + h_H) / (1 + N) ]",
         "h_H counts the same over false positives, recall loss, leakage, resources."),
    ]
    for i, (h1, f1, b1) in enumerate(cards):
        x = LM + i * 6.14
        card(s, x, 1.86, 5.75, 2.10)
        text(s, x + 0.26, 1.98, 5.2, 0.40, [{"t": h1, "sz": 20, "b": True, "c": MUTED}])
        text(s, x + 0.26, 2.40, 5.2, 0.46, [{"t": f1, "sz": 20, "b": True, "c": NAVY}])
        text(s, x + 0.26, 2.94, 5.24, 0.94, [{"t": b1, "sz": 20, "c": INK, "ls": 1.14}])
    band(s, 4.12, "FI is relative to a declared reference population and protocol.", color=AMBER)
    text(s, LM, 4.98, CW, 1.72, [
        {"t": "Change the population, the qualification rule or the protocol, and the comparison changes with it.",
         "sz": 20, "c": INK, "ls": 1.2},
        {"t": "The add-one correction keeps the estimate finite when nothing matches. It is not a confidence bound.",
         "sz": 20, "c": INK, "ls": 1.2, "sp": 10},
    ])
    foot(s, "10 \u00b7 Measurement", 10)

# ---------------------------------------------------------------- 11 ideality
def p11(s):
    head(s, "11 \u00b7 MEASUREMENT", "Ideality, and a separate certificate")
    card(s, LM, 1.84, CW, 1.10, fill=SOFT, radius=0.06)
    text(s, LM + 0.30, 1.84, CW - 0.6, 1.10, [
        {"t": "J_l = I_U^LB  /  max( \u03b5 , I_H^UB + \u03bb_C \u00b7 C_path )",
         "sz": 24, "b": True, "c": NAVY},
    ], anchor=MSO_ANCHOR.MIDDLE)
    card(s, LM, 3.14, 5.75, 1.72)
    text(s, LM + 0.26, 3.24, 5.2, 0.40, [{"t": "Certification level", "sz": 20, "b": True, "c": MUTED}])
    text(s, LM + 0.26, 3.66, 5.26, 1.12, [
        {"t": "L_cert = the highest level whose validation and hard gates were passed; -1 if none.",
         "sz": 20, "c": INK, "ls": 1.14},
    ])
    card(s, 6.86, 3.14, 5.75, 1.72)
    text(s, 7.12, 3.24, 5.2, 0.40, [{"t": "What a ratio cannot buy", "sz": 20, "b": True, "c": MUTED}])
    text(s, 7.12, 3.66, 5.26, 1.12, [
        {"t": "Integrity, safety, replication and resource gates are mandatory. A larger J does not compensate.",
         "sz": 20, "c": INK, "ls": 1.14},
    ])
    band(s, 5.04, "A method specification, not a measured quantity.", color=CORAL, fill=C("FBEEEA"), h=0.66, sz=20)
    text(s, LM, 5.86, CW, 0.90, [
        {"t": "Historical runs use harm point estimates; the campaign uses construction steps over six as a cost proxy, not measured runtime or energy.",
         "sz": 20, "c": MUTED, "ls": 1.18},
    ])
    foot(s, "11 \u00b7 Measurement", 11)

# ---------------------------------------------------------------- 12 traceability
def p12(s):
    head(s, "12 · TRACEABILITY", "A decision you can reconstruct")
    card(s, LM, 1.84, 6.30, 3.34)
    text(s, LM + 0.26, 1.96, 5.8, 0.40,
         [{"t": "What one record links", "sz": 20, "b": True, "c": MUTED}])
    rows = ["parent and candidate identity", "contradiction and principle",
            "operator arguments and the edit", "protocol fingerprint",
            "metrics and gate outcomes", "the decision itself"]
    for i, t in enumerate(rows):
        y = 2.44 + i * 0.44
        oval(s, LM + 0.30, y + 0.10, 0.16, 0.16, SKY)
        text(s, LM + 0.60, y, 5.5, 0.40, [{"t": t, "sz": 20, "c": INK}])
    card(s, 7.22, 1.84, 5.39, 3.34)
    text(s, 7.48, 1.96, 4.9, 0.40,
         [{"t": "Two rules", "sz": 20, "b": True, "c": MUTED}])
    text(s, 7.48, 2.44, 4.88, 2.60, [
        {"t": "A matrix hit records the matched pair. A fallback is recorded as a fallback.", "sz": 20, "c": INK, "ls": 1.16},
        {"t": "Stage labels stay at their own stage. They are not merged into one principle afterwards.", "sz": 20, "c": INK, "ls": 1.16, "sp": 12},
    ])
    band(s, 5.40, "2,040 retained decisions replayed from a 108-file manifest, with no training calls.",
         color=BLUE, fill=SOFT2, h=0.76)
    foot(s, "12 · Traceability", 12)


# ---------------------------------------------------------------- 13 evidence ladder
def p13(s):
    head(s, "13 \u00b7 EVIDENCE", "Four groups, four different questions")
    cols = [(LM, 2.55, "Task"), (3.27, 2.95, "Comparison"), (6.22, 3.35, "Observed"), (9.57, 3.04, "Status")]
    for x, w, t in cols:
        text(s, x, 1.90, w, 0.42, [{"t": t, "sz": 20, "b": True, "c": MUTED}])
    shape(s, LM, 2.36, CW, 0.014, LINE)
    rows = [
        ("PowerDuck GOOSE", "paired weighted-F1", "0.96943 \u2192 0.99817   +2.87 pp",
         "complete; no principle attribution", GREEN),
        ("CICIoT2023 legacy", "macro-F1, 9 families", "+4.23 pp",
         "complete; no run-level identity chain", GREEN),
        ("CIC-DDoS2019", "weighted-F1, 10 folds", "+0.15 pp",
         "complete; full registered scope", GREEN),
        ("Formal V2", "macro-F1, 45 pairs", "0.83256 \u2192 0.84839   +0.01584",
         "complete; NOT promoted", CORAL),
        ("V17 / V18", "multi-metric sequence", "V17 passes a separate champion rule",
         "neither meets strict V2", CORAL),
    ]
    for i, (a, b1, c1, d1, col) in enumerate(rows):
        y = 2.42 + i * 0.86
        text(s, LM, y, 2.55, 0.82, [{"t": a, "sz": 20, "b": True, "c": NAVY, "ls": 1.12}],
             anchor=MSO_ANCHOR.MIDDLE)
        text(s, 3.27, y, 2.95, 0.82, [{"t": b1, "sz": 20, "c": INK, "ls": 1.12}],
             anchor=MSO_ANCHOR.MIDDLE)
        text(s, 6.22, y, 3.35, 0.82, [{"t": c1, "sz": 20, "c": INK, "ls": 1.12}],
             anchor=MSO_ANCHOR.MIDDLE)
        text(s, 9.57, y, 3.04, 0.82, [{"t": d1, "sz": 20, "b": True, "c": col, "ls": 1.12}],
             anchor=MSO_ANCHOR.MIDDLE)
        if i < 4:
            shape(s, LM, y + 0.86, CW, 0.010, LINE)
    band(s, 6.14, "\u201ccomplete\u201d means the stated evaluation finished, not that the candidate was promoted.",
         color=AMBER, fill=SOFT, h=0.72, sz=20)
    foot(s, "13 \u00b7 Evidence", 13)

# ---------------------------------------------------------------- 14 legacy pairs
def p14(s):
    head(s, "14 · RESULTS", "Three legacy pairs, three protocols")
    items = [
        ("PowerDuck GOOSE", "weighted-F1, 15 classes", "+2.87", "pp", "0.96943 \u2192 0.99817"),
        ("CICIoT2023", "macro-F1, families x seeds", "+4.23", "pp", "legacy retained endpoint"),
        ("CIC-DDoS2019", "weighted-F1, 10 folds", "+0.15", "pp", "full registered scope"),
    ]
    for i, (n, m, big, unit, sub) in enumerate(items):
        x = LM + i * 4.07
        card(s, x, 1.84, 3.75, 3.06)
        text(s, x + 0.26, 1.98, 3.25, 0.42, [{"t": n, "sz": 20, "b": True, "c": NAVY}])
        text(s, x + 0.26, 2.40, 3.25, 0.42, [{"t": m, "sz": 20, "c": MUTED}])
        text(s, x + 0.26, 2.86, 3.25, 0.92, [{"t": big, "sz": 46, "b": True, "c": AMBER}])
        text(s, x + 0.26, 3.76, 3.25, 0.40, [{"t": unit, "sz": 20, "b": True, "c": AMBER}])
        text(s, x + 0.26, 4.18, 3.25, 0.62, [{"t": sub, "sz": 20, "c": INK, "ls": 1.14}])
    band(s, 5.08, "Read each pair inside its own protocol. The sizes are not a cross-task ranking.",
         color=BLUE, fill=SOFT2, h=0.70)
    text(s, LM, 5.94, CW, 0.86, [
        {"t": "PowerDuck's +2.87 pp is against its paired baseline. Promotion was decided against a prior incumbent at 0.99600, where the gain is +0.22 pp.",
         "sz": 20, "c": INK, "ls": 1.16},
    ])
    foot(s, "14 · Results", 14)


# ---------------------------------------------------------------- 15 formal V2
def p15(s):
    head(s, "15 · THE HONEST RESULT", "It improved. It was not promoted.")
    tiles = [
        ("macro-F1, 45 pairs", "0.83256 \u2192 0.84839", "+0.01584   (1.90%)"),
        ("recall", "+0.01802", "over the same pairs"),
        ("false-positive rate", "-0.00168", "lower is better"),
    ]
    for i, (a, b1, c1) in enumerate(tiles):
        x = LM + i * 4.07
        card(s, x, 1.84, 3.75, 1.54, fill=SOFT, line=LINE)
        text(s, x + 0.26, 1.94, 3.25, 0.38, [{"t": a, "sz": 20, "c": MUTED}])
        text(s, x + 0.26, 2.34, 3.30, 0.46, [{"t": b1, "sz": 22, "b": True, "c": NAVY}])
        text(s, x + 0.26, 2.84, 3.25, 0.42, [{"t": c1, "sz": 20, "b": True, "c": GREEN}])
    card(s, LM, 3.58, 6.30, 1.70)
    text(s, LM + 0.26, 3.68, 5.8, 0.40, [{"t": "Gates: both required", "sz": 20, "b": True, "c": MUTED}])
    text(s, LM + 0.26, 4.10, 5.85, 1.10, [
        {"t": "relative F1 gain \u2265 3.2277%      actual 1.90%", "sz": 20, "c": INK, "ls": 1.2},
        {"t": "FPR reduction \u2265 0.005          actual 0.00168", "sz": 20, "c": INK, "ls": 1.2, "sp": 8},
    ])
    card(s, 7.22, 3.58, 5.39, 1.70)
    text(s, 7.48, 3.68, 4.9, 0.40, [{"t": "Consistency", "sz": 20, "b": True, "c": MUTED}])
    text(s, 7.48, 4.10, 4.9, 1.10, [
        {"t": "43 of 45 pairs positive; all five seed blocks moved the same way.",
         "sz": 20, "c": INK, "ls": 1.18},
    ])
    band(s, 5.48, "Complete evidence, insufficient authorisation.", color=CORAL, fill=C("FBEEEA"), h=0.78)
    text(s, LM, 6.36, CW, 0.50, [
        {"t": "Deltas are computed before rounding: 0.84839357180 - 0.83255586481 = 0.01583770699.",
         "sz": 20, "c": MUTED},
    ])
    foot(s, "15 · The honest result", 15)


# ---------------------------------------------------------------- 16 controller audit
def p16(s):
    head(s, "16 \u00b7 CONTROL", "Policy is a trade-off, not a winner",
         "510 overlapping subsets of nine held-out families, 2,040 decisions, four candidates each")
    px0, px1, py0, py1 = 2.30, 11.30, 1.98, 5.66
    xmin, xmax, ymin, ymax = 0.86, 1.005, 0.30, 1.00

    def X(v):
        return px0 + (v - xmin) / (xmax - xmin) * (px1 - px0)

    def Y(v):
        return py1 - (v - ymin) / (ymax - ymin) * (py1 - py0)

    for v in (0.4, 0.6, 0.8, 1.0):
        shape(s, px0, Y(v), px1 - px0, 0.008, LINE)
        text(s, 1.45, Y(v) - 0.19, 0.72, 0.38, [{"t": "%.1f" % v, "sz": 20, "c": GREY}],
             align=PP_ALIGN.RIGHT)
    for v in (0.88, 0.92, 0.96, 1.00):
        text(s, X(v) - 0.55, py1 + 0.06, 1.10, 0.42,
             [{"t": "%.2f" % v, "sz": 20, "c": GREY}], align=PP_ALIGN.CENTER)
    shape(s, px0, py0, 0.014, py1 - py0, GREY)
    shape(s, px0, py1, px1 - px0, 0.014, GREY)
    pts = [(0.963, 0.508, "1"), (0.998, 0.409, "2"), (0.989, 0.811, "3"), (0.883, 0.955, "4")]
    for sx, sy, n in pts:
        cx, cy = X(sx), Y(sy)
        sp = oval(s, cx - 0.22, cy - 0.22, 0.44, 0.44, NAVY)
        label(sp, [{"t": n, "sz": 20, "b": True, "c": WHITE}], margins=(0, 0, 0, 0))
    text(s, LM, 6.20, CW, 0.72, [
        {"t": "1 fixed 5%   \u00b7   2 safety-first   \u00b7   3 balanced V2   \u00b7   4 calibrated V1        x: specificity     y: qualification F1",
         "sz": 20, "c": MUTED},
    ])
    foot(s, "16 \u00b7 Control", 16)

# ---------------------------------------------------------------- 17 downstream
def p17(s):
    head(s, "17 · DOWNSTREAM", "Where the loop said no")
    card(s, LM, 1.84, 5.75, 3.30)
    text(s, LM + 0.26, 1.96, 5.2, 0.42, [{"t": "V17", "sz": 22, "b": True, "c": NAVY}])
    text(s, LM + 0.26, 2.48, 5.2, 2.50, [
        {"t": "macro-F1 0.87530. Recall 2.43 pp below the RF100 baseline. Detector FPR rises.",
         "sz": 20, "c": INK, "ls": 1.16},
        {"t": "Passes its own champion rule. Does not pass strict balanced V2.",
         "sz": 20, "b": True, "c": CORAL, "ls": 1.16, "sp": 12},
    ])
    card(s, 6.86, 1.84, 5.75, 3.30)
    text(s, 7.12, 1.96, 5.2, 0.42, [{"t": "V18", "sz": 22, "b": True, "c": NAVY}])
    text(s, 7.12, 2.48, 5.2, 2.50, [
        {"t": "Rejected under its retained contract.", "sz": 20, "b": True, "c": CORAL, "ls": 1.16},
        {"t": "MCC is higher than the same-fit V17 core and lower than the older cached V17. The reference changed, not the value.",
         "sz": 20, "c": INK, "ls": 1.16, "sp": 12},
    ])
    band(s, 5.34, "Neither variant establishes strict V2 acceptance.", color=CORAL, fill=C("FBEEEA"), h=0.70)
    text(s, LM, 6.20, CW, 0.72, [
        {"t": "Detector FPR and controller specificity are different decisions: one marks benign traffic as attack, the other rejects unqualified candidates.",
         "sz": 20, "c": MUTED, "ls": 1.16},
    ])
    foot(s, "17 · Downstream", 17)


# ---------------------------------------------------------------- 18 scope
def p18(s):
    head(s, "18 \u00b7 SCOPE", "What this paper does not show")
    items = [
        ("No component-level causation", "TRIZ, FI and the controller are not separated by an ablation."),
        ("No matched-budget search", "Related candidates and reused caches are not independent."),
        ("No population sensitivity", "We do not vary the reference set or its qualification rule."),
        ("Incomplete provenance", "Sample-level predictions and one specification original are missing."),
    ]
    for i, (h1, b1) in enumerate(items):
        x = LM + (i % 2) * 5.83
        y = 1.82 + (i // 2) * 1.54
        card(s, x, y, 5.55, 1.40)
        text(s, x + 0.26, y + 0.12, 5.0, 0.40,
             [{"t": h1, "sz": 20, "b": True, "c": NAVY}])
        text(s, x + 0.26, y + 0.56, 5.06, 0.82,
             [{"t": b1, "sz": 20, "c": INK, "ls": 1.14}])
    band(s, 4.92, "These limits are stated in the paper, not discovered in review.",
         color=BLUE, fill=SOFT2, h=0.70)
    text(s, LM, 5.82, CW, 0.80, [
        {"t": "Reproduction rests on released records, a 108-file manifest and a no-training replay of 2,040 retained decisions.",
         "sz": 20, "c": MUTED, "ls": 1.16},
    ])
    foot(s, "18 \u00b7 Scope", 18)

# ---------------------------------------------------------------- 19 take-aways
def p19(s):
    head(s, "19 \u00b7 TAKE-AWAYS", "Three things to take away")
    rows = [
        ("1", "Type the direction, not only the result.",
         "A contradiction, a principle and a contract make the search auditable before the number exists."),
        ("2", "Functional information is an instrument, not a score.",
         "It is defined against a fixed reference population and a frozen protocol."),
        ("3", "Promotion is a separate decision from improvement.",
         "Our gates ask for more than a better metric, and they refuse candidates that improve one."),
    ]
    for i, (n, h1, b1) in enumerate(rows):
        y = 1.84 + i * 1.42
        text(s, LM, y - 0.04, 0.80, 0.72, [{"t": n, "sz": 40, "b": True, "c": AMBER}])
        text(s, LM + 0.96, y, 10.90, 0.48, [{"t": h1, "sz": 22, "b": True, "c": NAVY}])
        text(s, LM + 0.96, y + 0.50, 10.90, 0.78, [{"t": b1, "sz": 20, "c": INK, "ls": 1.14}])
    band(s, 6.06, "The most interesting result here is a candidate that improved three measures and was still not promoted.",
         color=CORAL, fill=C("FBEEEA"), h=0.66, sz=20)
    foot(s, "19 \u00b7 Take-aways", 19)

# ---------------------------------------------------------------- 20 thanks
def p20(s):
    shape(s, 0, 0, 13.3333, 0.16, NAVY)
    shape(s, LM, 1.62, 0.06, 1.06, AMBER)
    text(s, LM + 0.26, 1.56, 9.0, 0.50,
         [{"t": "AI4Science 2026  \u00b7  Shenzhen", "sz": 20, "c": MUTED, "spc": 0.6}])
    text(s, LM, 2.02, 10.0, 0.86, [{"t": "Thank you", "sz": 46, "b": True, "c": NAVY}])
    text(s, LM, 3.16, CW, 1.30, [
        {"t": "Paper: 8 pages  \u00b7  Supplementary appendix: 20 pages, A-H", "sz": 22, "c": INK, "ls": 1.2},
        {"t": "Contracts, code, measurements, decision ledgers, no-training replay:", "sz": 22, "c": INK, "ls": 1.2},
        {"t": "triza.cn/triza-paper-exp-data2026", "sz": 22, "b": True, "c": BLUE, "ls": 1.2},
    ])
    card(s, LM, 4.86, CW, 1.06, fill=SOFT, line=None, radius=0.08)
    text(s, LM + 0.30, 4.86, CW - 0.6, 1.06, [
        {"t": "Wenyin Liu  \u00b7  liuwy@gdut.edu.cn  \u00b7  Guangdong University of Technology, Guangzhou",
         "sz": 20, "c": INK},
    ], anchor=MSO_ANCHOR.MIDDLE)
    text(s, LM, 6.16, CW, 0.50, [
        {"t": "Yiheng Huang (BUPT)   \u00b7   Kai Wang (Beijing Denglu Technology)",
         "sz": 20, "c": MUTED},
    ])
    foot(s, "TRIZa V8  \u00b7  AI4Science 2026", 20)


def main():
    for fn in (p01, p02, p03, p04, p05, p06, p07, p08, p09, p10,
               p11, p12, p13, p14, p15, p16, p17, p18, p19, p20):
        fn(new_slide())
    for sl, note in zip(prs.slides, NOTES):
        sl.notes_slide.notes_text_frame.text = note
    out = HERE / "TRIZa_V8_AI4Science_15min.pptx"
    prs.save(str(out))
    print("SAVED", len(prs.slides._sldIdLst), out.stat().st_size)


if __name__ == "__main__":
    main()
