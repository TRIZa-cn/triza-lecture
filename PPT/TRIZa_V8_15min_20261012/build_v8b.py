# -*- coding: utf-8 -*-
"""Build the 14-slide TRIZa conference deck (English, 20pt floor)."""
from pathlib import Path
import sys
from PIL import Image
from pptx import Presentation
from pptx.util import Inches as In, Pt
from pptx.dml.color import RGBColor as C
from pptx.enum.text import PP_ALIGN as AL, MSO_ANCHOR as AN
from pptx.enum.shapes import MSO_SHAPE, MSO_CONNECTOR
from pptx.oxml import parse_xml
from pptx.oxml.ns import nsdecls

D = Path(__file__).resolve().parent
sys.path.insert(0, str(D))
from notes_v8b import NOTES

NAVY  = C(0x0E, 0x2A, 0x47)
INK   = C(0x16, 0x28, 0x3C)
BLUE  = C(0x1E, 0x6B, 0xA8)
SKY   = C(0x4E, 0x93, 0xC4)
AMBER = C(0xD9, 0x8A, 0x26)
CORAL = C(0xBE, 0x51, 0x38)
GREEN = C(0x2B, 0x7A, 0x63)
MUTED = C(0x6C, 0x7C, 0x8F)
LINE  = C(0xD9, 0xE2, 0xEB)
SOFT  = C(0xF4, 0xF8, 0xFC)
SOFT2 = C(0xE7, 0xEF, 0xF7)
WHITE = C(0xFF, 0xFF, 0xFF)
GREY  = C(0x9A, 0xA8, 0xB8)
FONT  = "Segoe UI"

LM, CW, FOOT_Y = 0.72, 11.893, 6.98


# ---------------------------------------------------------------- text helpers
def text(s, x, y, w, h, paras, align=AL.LEFT, anchor=AN.TOP, name=None):
    tb = s.shapes.add_textbox(In(x), In(y), In(w), In(h))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = In(0)
    tf.margin_top = tf.margin_bottom = In(0)
    tf.vertical_anchor = anchor
    first = True
    for p in paras:
        para = tf.paragraphs[0] if first else tf.add_paragraph()
        first = False
        para.alignment = p.get("align", align)
        if "sb" in p:
            para.space_before = Pt(p["sb"])
        if "sa" in p:
            para.space_after = Pt(p["sa"])
        if "ls" in p:
            para.line_spacing = p["ls"]
        runs = p.get("runs") or [(p.get("t", ""), {})]
        for t, o in runs:
            r = para.add_run()
            r.text = t
            f = r.font
            f.name = o.get("font", FONT)
            f.size = Pt(o.get("sz", p.get("sz", 20)))
            f.bold = o.get("b", p.get("b", False))
            f.italic = o.get("i", p.get("i", False))
            f.color.rgb = o.get("c", p.get("c", INK))
    if name:
        tb.name = name
    return tb


def shape(s, x, y, w, h, fill=None, line=None, lw=0.75, kind=MSO_SHAPE.RECTANGLE,
          radius=None, name=None):
    sp = s.shapes.add_shape(kind, In(x), In(y), In(w), In(h))
    if fill is None:
        sp.fill.background()
    else:
        sp.fill.solid()
        sp.fill.fore_color.rgb = fill
    if line is None:
        sp.line.fill.background()
    else:
        sp.line.color.rgb = line
        sp.line.width = Pt(lw)
    sp.shadow.inherit = False
    if radius is not None:
        try:
            sp.adjustments[0] = radius
        except Exception:
            pass
    sp.text_frame.word_wrap = True
    if name:
        sp.name = name
    return sp


def card(s, x, y, w, h, fill=SOFT, edge=LINE, accent=None):
    shape(s, x, y, w, h, fill=fill, line=edge, kind=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.06)
    if accent:
        shape(s, x, y + 0.14, 0.055, h - 0.28, fill=accent)


def head(s, kicker, title, sub=None):
    text(s, LM, 0.46, CW, 0.30, [{"t": kicker, "sz": 20, "b": True, "c": SKY}])
    text(s, LM, 0.82, CW, 0.62, [{"t": title, "sz": 30, "b": True, "c": NAVY}])
    if sub:
        text(s, LM, 1.46, CW, 0.32, [{"t": sub, "sz": 20, "c": MUTED}])


def foot(s, section, page):
    ln = s.shapes.add_connector(MSO_CONNECTOR.STRAIGHT,
                                In(LM), In(FOOT_Y - 0.10), In(LM + CW), In(FOOT_Y - 0.10))
    ln.line.color.rgb = LINE
    ln.line.width = Pt(0.75)
    text(s, LM, FOOT_Y, 8.2, 0.32, [{"t": section, "sz": 20, "c": GREY}])
    text(s, LM + CW - 1.2, FOOT_Y, 1.2, 0.32, [{"t": page, "sz": 20, "c": GREY}], align=AL.RIGHT)


def band(s, y, runs, h=0.62, fill=SOFT2, accent=AMBER):
    shape(s, LM, y, CW, h, fill=fill, line=None, kind=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.05)
    shape(s, LM, y, 0.07, h, fill=accent)
    text(s, LM + 0.26, y, CW - 0.42, h, [{"runs": runs, "sz": 20}], anchor=AN.MIDDLE)


def bullets(s, x, y, w, h, items, marker="\u00b7", mc=SKY, sz=20, gap=8):
    paras = []
    for i, it in enumerate(items):
        if isinstance(it, str):
            it = [(it, {})]
        paras.append({"runs": [("\u00a0" + marker + "\u00a0\u00a0", {"c": mc, "b": True, "sz": sz})] + it,
                      "sz": sz, "sb": 0 if i == 0 else gap, "ls": 1.06})
    return text(s, x, y, w, h, paras)


def pic(s, path, x, y, w=None, h=None):
    iw, ih = Image.open(str(path)).size
    if w is not None:
        h = w * ih / float(iw)
    else:
        w = h * iw / float(ih)
    s.shapes.add_picture(str(path), In(x), In(y), In(w), In(h))
    return w, h


# ---------------------------------------------------------------- OMML helpers
NSM = 'xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"'
NSA14 = 'xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main"'


def esc(t):
    return t.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def mr(t, sz=20, italic=True, plain=False):
    sty = '<m:rPr><m:sty m:val="p"/></m:rPr>' if plain else ""
    it = "1" if italic else "0"
    return ('<m:r>%s<a:rPr lang="en-US" sz="%d" b="0" i="%s">'
            '<a:latin typeface="Cambria Math"/></a:rPr>'
            '<m:t xml:space="preserve">%s</m:t></m:r>' % (sty, int(sz * 100), it, esc(t)))


def mv(t, sz=20):
    return mr(t, sz, italic=True)


def mo(t, sz=20):
    return mr(t, sz, italic=False, plain=True)


def msub(b, sb, sz=20):
    return '<m:sSub><m:e>%s</m:e><m:sub>%s</m:sub></m:sSub>' % (b, sb)


def msup(b, sp, sz=20):
    return '<m:sSup><m:e>%s</m:e><m:sup>%s</m:sup></m:sSup>' % (b, sp)


def msubsup(b, sb, sp, sz=20):
    return ('<m:sSubSup><m:e>%s</m:e><m:sub>%s</m:sub><m:sup>%s</m:sup></m:sSubSup>'
            % (b, sb, sp))


def mfrac(num, den, sz=20):
    return ('<m:f><m:fPr><m:type m:val="bar"/></m:fPr><m:num>%s</m:num><m:den>%s</m:den></m:f>'
            % (num, den))


def math(s, x, y, w, h, body, anchor=AN.MIDDLE):
    tb = s.shapes.add_textbox(In(x), In(y), In(w), In(h))
    tf = tb.text_frame
    tf.word_wrap = False
    tf.margin_left = tf.margin_right = In(0.02)
    tf.margin_top = tf.margin_bottom = In(0)
    tf.vertical_anchor = anchor
    p = tf.paragraphs[0]._p
    for ch in list(p):
        p.remove(ch)
    xml = ('<a:p %s %s %s><a14:m><m:oMathPara><m:oMath>%s</m:oMath>'
           '</m:oMathPara></a14:m></a:p>' % (nsdecls('a'), NSM, NSA14, body))
    p.getparent().replace(p, parse_xml(xml))
    tb.name = "math"
    return tb


def new_slide(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    f = s.background.fill
    f.solid()
    f.fore_color.rgb = WHITE
    return s

# ---------------------------------------------------------------- slides 1-7
def p01(prs):
    s = new_slide(prs)
    shape(s, 0, 0, 13.3333, 0.34, fill=NAVY)
    text(s, LM, 0.70, CW, 0.32,
         [{"t": "AI4SCIENCE 2026   \u00b7   SHENZHEN   \u00b7   OCTOBER 23-25",
           "sz": 20, "b": True, "c": SKY}])
    text(s, LM, 1.16, 11.9, 1.40, [
        {"t": "TRIZa: Guiding Agent Evolution", "sz": 32, "b": True, "c": NAVY},
        {"t": "from Pattern Recognition to Solution Invention", "sz": 32, "b": True,
         "c": NAVY, "sb": 2}])

    ops = ["", "\u00d7", "="]
    lines = [("TRIZ parameterisation  (reasoning \u00b7 direction)", False),
             ("LLM\u00b7Agent automatic solving  (execution)", False),
             ("a self-evolving loop  (RSI)", True)]
    for i, ((txt, emph), op) in enumerate(zip(lines, ops)):
        y = 3.02 + i * 0.56
        if op:
            text(s, 1.55, y, 0.55, 0.50, [{"t": op, "sz": 26, "b": True, "c": AMBER}],
                 align=AL.LEFT, anchor=AN.MIDDLE)
        text(s, 2.22, y, 9.6, 0.50,
             [{"t": txt, "sz": 24 if emph else 22, "b": emph,
               "c": NAVY if emph else INK}], anchor=AN.MIDDLE)

    shape(s, 1.55, 4.86, 9.9, 0.014, fill=LINE)
    card(s, 1.55, 5.02, 9.9, 1.42, fill=SOFT)
    text(s, 1.85, 5.20, 9.4, 1.10, [
        {"t": "Wenyin Liu (GDUT)   \u00b7   Yiheng Huang (BUPT)   \u00b7   "
              "Kai Wang (Beijing Denglu Technology)", "sz": 20, "c": INK},
        {"t": "liuwy@gdut.edu.cn", "sz": 20, "c": MUTED, "sb": 6}])
    foot(s, "AI4Science 2026, Shenzhen", "01")
    return s



def p02(prs):
    s = new_slide(prs)
    head(s, "01 \u00b7 THE GAP", "Prompting yields code. Principles yield evidence.")
    left = [
        "The raw problem goes straight into the prompt.",
        "A few rounds may still produce runnable code.",
        "Nothing forces the model to name the trade-off.",
        "Different wording next run: not reproducible, no ablation.",
    ]
    right = [
        "A contradiction is written down before generation.",
        "Each candidate carries one principle and one parameter pair.",
        "Ablatable: remove the operator and the effect should vanish.",
        "Comparable: the protocol is frozen first, so runs line up.",
    ]
    for x, title, items, col in ((LM, "PROMPT ONLY", left, CORAL),
                                 (LM + 6.09, "THROUGH TRIZ", right, GREEN)):
        card(s, x, 1.78, 5.80, 4.14, fill=SOFT, accent=col)
        text(s, x + 0.26, 1.94, 5.30, 0.62,
             [{"t": title, "sz": 20, "b": True, "c": col}])
        bullets(s, x + 0.26, 2.62, 5.30, 3.30, items, mc=col, gap=8)
    band(s, 6.06, [("Ablatable, comparable, reproducible ", {"c": INK, "b": True}),
                   ("\u2014 or none of the three.", {"c": INK})], h=0.58)
    foot(s, "The gap", "02")
    return s

def p03(prs):
    s = new_slide(prs)
    head(s, "02 \u00b7 BACKGROUND", "TRIZ in one minute",
         "Forty named directions, one entry point, one criterion.")
    rows = [
        ("Contradiction",
         "You want more of A, and that costs you B. TRIZ asks for the pair, not for the wish. "
         "In our setting the pair is concrete: raise attack recall without raising false positives."),
        ("Forty principles",
         "Named directions rather than recipes: segmentation, local quality, prior action, "
         "parameter change, feedback, intermediary. Forty in total, tabulated in Appendix B."),
        ("Ideality",
         "Judge a result by useful function over harm and cost. The ideal final result delivers "
         "the function and carries none of the cost."),
    ]
    y = 1.96
    for i, (name, body) in enumerate(rows):
        text(s, LM, y, 2.55, 0.36, [{"t": name, "sz": 22, "b": True, "c": NAVY}])
        text(s, 3.46, y - 0.02, 9.15, 1.02, [{"t": body, "sz": 20, "c": INK, "ls": 1.06}])
        if i < 2:
            shape(s, LM, y + 1.12, CW, 0.012, fill=LINE)
        y += 1.24
    band(s, 5.78, [("TRIZ never picks your estimator. ", {"c": INK}),
                   ("It picks the class of change to try, and why that class should help.",
                    {"c": INK, "b": True})])
    foot(s, "Background", "03")
    return s


def p04(prs):
    s = new_slide(prs)
    head(s, "03 \u00b7 THE LOOP", "Four stages, three layers",
         "Human authority sits at both ends; the middle is where the agent works.")
    w, h = pic(s, D / "assets" / "arch.png", 1.867, 1.90, w=9.6)
    foot(s, "Architecture", "04")
    return s


def p05(prs):
    s = new_slide(prs)
    head(s, "04 \u00b7 STAGE ONE", "Fix the question before the answer",
         "Nothing is generated until the protocol is frozen.")
    items = [
        ("Human brief", "Target function, operating constraints, and the actions that "
                        "need a human authority to approve."),
        ("Contradiction", "The property to improve, the property that may worsen, and a "
                          "testable criterion for each side."),
        ("Frozen protocol", "Split, metric implementation and evaluation procedure, fixed "
                            "before any candidate exists."),
        ("Knowledge bank", "Source-tagged mappings and prior outcomes, so a recommendation "
                           "can say where it came from."),
    ]
    for i, (name, body) in enumerate(items):
        x = LM + (i % 2) * 6.09
        y = 1.84 + (i // 2) * 1.94
        card(s, x, y, 5.80, 1.72, fill=SOFT, accent=SKY)
        text(s, x + 0.26, y + 0.16, 5.28, 0.34,
             [{"t": name, "sz": 22, "b": True, "c": NAVY}])
        text(s, x + 0.26, y + 0.60, 5.28, 1.06,
             [{"t": body, "sz": 20, "c": INK, "ls": 1.05}])
    band(s, 5.80, [("If the measurement is chosen after the result, ", {"c": INK}),
                   ("no later claim is safe.", {"c": CORAL, "b": True})])
    foot(s, "Stage one", "05")
    return s


def p06(prs):
    s = new_slide(prs)
    head(s, "05 \u00b7 THE OBJECT", "A candidate is a typed operator, not an edit",
         "The contract is written before the run; that is what makes ablation possible.")
    pi = (mv("\u03c0") + mo(" = ( ") + mo("id(s), ") + mv("c") + mo(", ") + mv("r")
          + mo(", ") + msub(mv("\u03c4"), mv("r")) + mo(", ") + mv("\u0394") + mo(", ")
          + mv("\u03ba") + mo(" )"))
    math(s, LM, 1.90, 8.4, 0.66, pi)
    text(s, LM, 2.58, 9.0, 0.32,
         [{"t": "parent artifact, contradiction, principle, operator, arguments, protocol",
           "sz": 20, "c": MUTED}])
    sp = (msup(mv("s"), mo("\u2032")) + mo(" = ") + msub(mv("\u03c4"), mv("r"))
          + mo("( ") + mv("s") + mo(" ; ") + mv("\u0394") + mo(" )"))
    math(s, LM, 3.02, 5.2, 0.66, sp)
    text(s, LM + 5.45, 3.10, 6.2, 0.70,
         [{"t": "the candidate is that operator applied to the parent", "sz": 20, "c": MUTED}])

    labels = ["Input and output types", "Permitted operations",
              "Protected conditions", "Required checks"]
    for i, lb in enumerate(labels):
        x = LM + i * 3.00
        card(s, x, 4.02, 2.84, 1.16, fill=SOFT2, accent=BLUE)
        text(s, x + 0.20, 4.02, 2.48, 1.16,
             [{"runs": [("%d  " % (i + 1), {"c": AMBER, "b": True, "sz": 20}),
                        (lb, {"c": NAVY, "b": True, "sz": 20})], "sz": 20}],
             anchor=AN.MIDDLE)
    band(s, 5.42, [("A favourable result cannot retrospectively make ", {"c": INK}),
                   ("an untyped edit TRIZ-guided.", {"c": CORAL, "b": True})])
    foot(s, "The object", "06")
    return s


def p07(prs):
    s = new_slide(prs)
    head(s, "06 \u00b7 DIRECTIONS", "Four principles, one four-step procedure")
    rows = [("Segmentation", "partition decisions behind a stable interface"),
            ("Local Quality", "context-specific behaviour, declared in advance"),
            ("Prior Action", "reusable state prepared before the critical path"),
            ("Parameter Change", "a threshold or a weight is moved")]
    y = 1.82
    for name, gloss in rows:
        text(s, LM, y, 6.0, 0.34, [{"t": name, "sz": 22, "b": True, "c": NAVY}])
        text(s, LM, y + 0.40, 6.0, 0.34, [{"t": gloss, "sz": 20, "c": MUTED}])
        y += 1.02
    steps = [
        "Name the property to improve, the one that may worsen, and what is protected.",
        "Retrieve principles with their source; label a fallback as a fallback.",
        "State how the operation helps and how it harms.",
        "Record the principle and the arguments before evaluation.",
    ]
    yy = 1.82
    for i, st in enumerate(steps):
        text(s, 7.10, yy, 0.45, 0.34, [{"t": str(i + 1), "sz": 22, "b": True, "c": AMBER}])
        text(s, 7.60, yy, 4.95, 0.90, [{"t": st, "sz": 20, "c": INK, "ls": 1.05}])
        yy += 0.94
    band(s, 5.86, [("No rationale, no claim: ", {"c": CORAL, "b": True}),
                   ("the record has to exist before the number does.", {"c": INK})])
    foot(s, "Directions", "07")
    return s

# ---------------------------------------------------------------- slides 8-14
def p08(prs):
    s = new_slide(prs)
    head(s, "07 \u00b7 MEASUREMENT", "Two quantities, one order",
         "Both bounds come from the same reference population, and the protocol id travels with them.")

    iu = (msub(mv("I"), mo("U")) + mo("(s) = \u2212") + msub(mo("log"), mo("2")) + mo(" [ ")
          + mfrac(mo("(1 + ") + msub(mv("h"), mo("U")) + mo("(s))"), mo("(1 + N)")) + mo(" ]"))
    ih = (msub(mv("I"), mo("H")) + mo("(s) = \u2212") + msub(mo("log"), mo("2")) + mo(" [ ")
          + mfrac(mo("(1 + ") + msub(mv("h"), mo("H")) + mo("(s))"), mo("(1 + N)")) + mo(" ]"))
    jl = (msub(mv("J"), mv("l")) + mo(" = ")
          + mfrac(msubsup(mv("I"), mo("U"), mo("LB")),
                  mo("max(\u03b5, ") + msubsup(mv("I"), mo("H"), mo("UB"))
                  + mo(" + ") + msub(mv("\u03bb"), mo("C")) + mo("\u00b7")
                  + msub(mv("C"), mo("path")) + mo(")")))

    math(s, LM, 1.86, 7.00, 0.80, iu)
    math(s, LM, 2.84, 7.00, 0.80, ih)
    math(s, LM, 3.86, 7.00, 1.16, jl)

    text(s, 8.05, 1.90, 4.55, 1.02,
         [{"t": "Useful FI: how rare the candidate's level is among references at that level.",
           "sz": 20, "c": INK, "ls": 1.04}])
    text(s, 8.05, 2.88, 4.55, 1.02,
         [{"t": "Harm FI: the same construction over false positives, recall loss, leakage, cost.",
           "sz": 20, "c": INK, "ls": 1.04}])
    text(s, 8.05, 3.92, 4.55, 1.06,
         [{"t": "Ideality: a lower bound on useful FI over an upper bound on harm FI plus path cost.",
           "sz": 20, "c": INK, "ls": 1.04}])
    band(s, 5.34, [("Certification is separate: ", {"c": INK}),
                   ("the highest level whose gates were passed, or \u22121.",
                    {"c": NAVY, "b": True})])
    text(s, LM, 6.14, CW, 0.34,
         [{"t": "The add-one term keeps the estimate finite when nothing matches. "
                "It is not a confidence bound.", "sz": 20, "c": MUTED}])
    foot(s, "Measurement", "08")
    return s


def p09(prs):
    s = new_slide(prs)
    head(s, "08 \u00b7 EVIDENCE", "Paired outcomes, each inside its own protocol")
    w, h = pic(s, D / "assets" / "figA.png", LM, 1.76, w=5.60)
    items = [
        [("PowerDuck GOOSE (W): ", {"b": True}), ("0.96943 to 0.99817, +2.87 pp.", {})],
        [("CICIoT legacy (M): ", {"b": True}), ("+4.23 pp.", {})],
        [("CIC-DDoS (W): ", {"b": True}), ("+0.15 pp across all ten folds.", {})],
        [("Formal candidate (M): ", {"b": True}), ("+1.58 pp, and still not promoted.", {})],
    ]
    bullets(s, 6.72, 1.84, 5.90, 2.60, items, gap=8)
    band(s, 4.62, [("Two cautions. ", {"c": CORAL, "b": True}),
                   ("The sizes are not a ranking across tasks, because the metrics differ; "
                    "and PowerDuck's margin against its incumbent was +0.22 pp, not +2.87.",
                    {"c": INK})], h=1.10)
    foot(s, "Paired evidence", "09")
    return s



def p10(prs):
    s = new_slide(prs)
    head(s, "09 \u00b7 THE DECISION", "Improved on three measures. Not promoted.")
    text(s, LM, 1.74, 6.5, 1.02,
         [{"runs": [("0.83256", {"c": MUTED, "b": True, "sz": 32}),
                    ("  \u2192  ", {"c": AMBER, "b": True, "sz": 32}),
                    ("0.84839", {"c": NAVY, "b": True, "sz": 32})]}])
    text(s, LM, 2.80, 6.5, 0.62, [{"t": "macro-F1, exact endpoints", "sz": 20, "c": MUTED}])
    text(s, LM, 3.46, 6.5, 0.72,
         [{"runs": [("\u0394 = +0.01584", {"c": INK, "b": True, "sz": 24}),
                    ("     1.90% relative", {"c": MUTED, "sz": 20})]}])
    stats = [("43 / 45", "pairs positive"),
             ("5 / 5", "seed blocks positive"),
             ("+0.01802", "recall"),
             ("\u22120.00168", "false positive rate")]
    paras = []
    for i, (val, lab) in enumerate(stats):
        paras.append({"runs": [(val + "   ", {"c": GREEN, "b": True, "sz": 22}),
                               (lab, {"c": INK, "sz": 20})],
                      "sz": 20, "sb": 0 if i == 0 else 8})
    text(s, 7.30, 1.76, 5.30, 1.94, paras)
    card(s, LM, 4.22, CW, 1.30, fill=SOFT2, accent=CORAL)
    text(s, LM + 0.26, 4.30, CW - 0.52, 1.14, [
        {"runs": [("Our own gates: ", {"c": INK, "b": True, "sz": 20}),
                  ("relative F1 gain \u2265 3.2277%   and   FPR reduction \u2265 0.005",
                   {"c": INK, "sz": 20})], "sz": 20},
        {"runs": [("Both unmet. ", {"c": CORAL, "b": True, "sz": 20}),
                  ("A completed evaluation is not an authorisation.", {"c": INK, "sz": 20})],
         "sz": 20, "sb": 3},
        {"t": "Deltas are computed before rounding; the rounded pair alone would read +0.01583.",
         "sz": 20, "c": MUTED, "sb": 3}])
    band(s, 5.64, [("Complete evidence, insufficient authorisation.",
                    {"c": NAVY, "b": True})], h=0.66)
    foot(s, "The decision", "10")
    return s

def p11(prs):
    s = new_slide(prs)
    head(s, "10 \u00b7 THE CONTROLLER", "Qualification is a policy choice, not a winner",
         "Specificity is the rejection rate for unqualified candidates.")
    w, h = pic(s, D / "assets" / "figB.png", LM, 1.72, w=5.55)
    items = [
        [("Safety-first: ", {"b": True}), ("specificity 0.998, qualification F1 0.409.", {})],
        [("Balanced V2: ", {"b": True}), ("0.989 and 0.811.", {})],
        [("Original V1, fixed 5%: ", {"b": True}), ("0.963 and 0.508.", {})],
        [("Calibrated V1: ", {"b": True}), ("0.883 and 0.955.", {})],
    ]
    bullets(s, 6.68, 1.80, 5.95, 2.55, items, gap=8)
    band(s, 4.72, [("2,040 decisions over 510 overlapping subsets ", {"c": INK, "b": True}),
                   ("\u2014 reported as a trade-off with stated consequences, "
                    "not as independent replications.", {"c": INK})], h=0.90)
    foot(s, "The controller", "11")
    return s


def p12(prs):
    s = new_slide(prs)
    head(s, "11 \u00b7 DOWNSTREAM", "Where the loop said no")
    w, h = pic(s, D / "assets" / "fig2.png", LM, 1.76, w=4.50)
    items = [
        [("V14: ", {"b": True}), ("macro-F1 0.83256 \u2192 0.87490.", {})],
        [("V17: ", {"b": True}), ("0.87530, but detector FPR 0.05040 \u2192 0.02319 and recall "
                                  "2.43 pp below RF100.", {})],
        [("V17 passes ", {}), ("its own champion rule, not strict balanced V2.", {"b": True})],
        [("V18 ", {"b": True}), ("was rejected under its retained contract.", {})],
        [("MCC: ", {"b": True}), ("above the same-fit V17 core, below the older cached V17.", {})],
    ]
    bullets(s, 5.62, 1.80, 7.00, 3.60, items, gap=10)
    band(s, 5.78, [("The number did not change. ", {"c": INK}),
                   ("The reference did.", {"c": CORAL, "b": True})])
    foot(s, "Downstream", "12")
    return s



def p13(prs):
    s = new_slide(prs)
    head(s, "12 \u00b7 SCOPE", "What this does not show")
    items = [
        ("Not causal", "No isolation of TRIZ, FI or the controller: a workflow, not three "
                       "separable effects."),
        ("No matched search", "No matched-budget independent search, and no sensitivity study "
                              "over reference populations."),
        ("Not independent", "Related candidates and reused caches are not independent "
                            "confirmations."),
        ("Reproduction limited", "Sample-level predictions missing, one specification original "
                                 "missing."),
    ]
    for i, (name, body) in enumerate(items):
        x = LM + (i % 2) * 6.09
        y = 1.82 + (i // 2) * 1.94
        card(s, x, y, 5.80, 1.72, fill=SOFT, accent=CORAL)
        text(s, x + 0.26, y + 0.16, 5.28, 0.34,
             [{"t": name, "sz": 22, "b": True, "c": NAVY}])
        text(s, x + 0.26, y + 0.60, 5.28, 1.04,
             [{"t": body, "sz": 20, "c": INK, "ls": 1.04}])
    band(s, 5.78, [("Stated in the paper, ", {"c": INK, "b": True}),
                   ("not discovered in review.", {"c": INK})])
    foot(s, "Scope", "13")
    return s


def p14(prs):
    s = new_slide(prs)
    head(s, "13 \u00b7 TAKE-AWAYS", "Three things worth keeping")
    items = [
        "Type the direction, not only the result: contradiction, principle and contract "
        "make the search auditable before the number exists.",
        "Treat functional information as an instrument tied to a fixed reference population, "
        "not as a universal score.",
        "Keep promotion separate from improvement: our most interesting candidate improved "
        "three measures and was still not promoted.",
    ]
    y = 1.84
    for i, it in enumerate(items):
        text(s, LM, y, 0.55, 0.72, [{"t": str(i + 1), "sz": 26, "b": True, "c": AMBER}])
        text(s, LM + 0.62, y + 0.02, 11.2, 0.92, [{"t": it, "sz": 22, "c": INK, "ls": 1.05}])
        y += 1.06
    shape(s, LM, 5.14, CW, 0.014, fill=LINE)
    text(s, LM, 5.26, 6.0, 0.72, [{"t": "Thank you.", "sz": 34, "b": True, "c": NAVY}])
    text(s, 6.90, 5.30, 5.70, 1.14, [
        {"t": "TRIZa: Guiding Agent Evolution", "sz": 20, "b": True, "c": INK},
        {"t": "Paper, Appendix A\u2013H and the decision records travel with the proceedings.",
         "sz": 20, "c": MUTED, "ls": 1.04, "sb": 4}])
    foot(s, "Take-aways", "14")
    return s

def crop_panels():
    src = Image.open(str(D / "assets" / "figAB.png")).convert("RGB")
    w, h = src.size
    for i, name in enumerate(("figA.png", "figB.png")):
        im = src.crop((i * (w // 2), 0, (i + 1) * (w // 2), h))
        bb = Image.eval(im, lambda v: 255 - v).getbbox()
        if bb:
            pad = 6
            bb = (max(0, bb[0] - pad), max(0, bb[1] - pad),
                  min(im.width, bb[2] + pad), min(im.height, bb[3] + pad))
            im = im.crop(bb)
        im.save(str(D / "assets" / name))
        print("  panel", name, im.size)


def main():
    crop_panels()
    prs = Presentation()
    prs.slide_width = In(13.3333)
    prs.slide_height = In(7.5)
    for fn in (p01, p02, p03, p04, p05, p06, p07, p08, p09, p10, p11, p12, p13, p14):
        fn(prs)
    assert len(prs.slides) == len(NOTES), (len(prs.slides), len(NOTES))
    for s, note in zip(prs.slides, NOTES):
        s.notes_slide.notes_text_frame.text = note
    out = D / "TRIZa_V8_AI4Science_15min.pptx"
    prs.save(str(out))
    print("SAVED", len(prs.slides), out.stat().st_size)


if __name__ == "__main__":
    main()
