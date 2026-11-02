# -*- coding: utf-8 -*-
"""原生图元绘制：功能图 / 因果链。所有图元命名为 fig-*，供审计脚本豁免 20pt 下限。"""
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE, MSO_CONNECTOR
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.oxml.ns import qn

INK = RGBColor(0x1A, 0x1D, 0x23)
INK2 = RGBColor(0x56, 0x5C, 0x66)
INK3 = RGBColor(0x8A, 0x90, 0x9A)
LINE = RGBColor(0xDD, 0xD8, 0xCE)
Z1 = RGBColor(0xFC, 0xFB, 0xF8)
Z0 = RGBColor(0xF2, 0xEF, 0xE8)
RED = RGBColor(0xB0, 0x3A, 0x2B)
AMBER = RGBColor(0xA9, 0x6E, 0x24)
GREEN = RGBColor(0x35, 0x69, 0x4A)
FONT = "微软雅黑"


def _ea(run, name):
    rPr = run._r.get_or_add_rPr()
    for tag in ("a:latin", "a:ea", "a:cs"):
        e = rPr.find(qn(tag))
        if e is None:
            e = rPr.makeelement(qn(tag), {})
            rPr.append(e)
        e.set("typeface", name)


def _nm(sh, n):
    try:
        sh.name = n
    except Exception:
        pass
    return sh


def _box(s, x, y, w, h, text, size, fill, lc, bold=False, color=INK, name="fig-node"):
    sh = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x), Inches(y), Inches(w), Inches(h))
    try:
        sh.adjustments[0] = 0.14
    except Exception:
        pass
    sh.fill.solid()
    sh.fill.fore_color.rgb = fill
    if lc is None:
        sh.line.fill.background()
    else:
        sh.line.color.rgb = lc
        sh.line.width = Pt(1.25)
    try:
        sh.shadow.inherit = False
    except Exception:
        pass
    tf = sh.text_frame
    tf.word_wrap = True
    tf.margin_left = Inches(0.04)
    tf.margin_right = Inches(0.04)
    tf.margin_top = Inches(0.01)
    tf.margin_bottom = Inches(0.01)
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    r = p.add_run()
    r.text = text
    r.font.size = Pt(size)
    r.font.bold = bold
    r.font.color.rgb = color
    r.font.name = FONT
    _ea(r, FONT)
    return _nm(sh, name)


def _label(s, x, y, w, h, text, size, col=INK2, name="fig-label", bold=False, align=PP_ALIGN.CENTER):
    tb = s.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = tb.text_frame
    tf.word_wrap = False
    for m in ("margin_left", "margin_right", "margin_top", "margin_bottom"):
        setattr(tf, m, 0)
    p = tf.paragraphs[0]
    p.alignment = align
    r = p.add_run()
    r.text = text
    r.font.size = Pt(size)
    r.font.bold = bold
    r.font.color.rgb = col
    r.font.name = FONT
    _ea(r, FONT)
    return _nm(tb, name)


def _arrow(s, x1, y1, x2, y2, col=INK3, width=1.5, dash=None, name="fig-arrow"):
    c = s.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(x1), Inches(y1), Inches(x2), Inches(y2))
    c.line.color.rgb = col
    c.line.width = Pt(width)
    ln = c.line._get_or_add_ln()
    if dash:
        d = ln.makeelement(qn("a:prstDash"), {"val": dash})
        ln.append(d)
    te = ln.makeelement(qn("a:tailEnd"), {"type": "triangle", "w": "med", "len": "med"})
    ln.append(te)
    return _nm(c, name)


_STATUS = {
    "ok":   (INK2, 1.5, None),
    "weak": (INK3, 1.0, None),
    "harm": (RED,  2.25, None),
    "miss": (INK3, 1.25, "dash"),
}


def _rows(s, spec, x, y, w, h, acc=RED):
    rows = spec["rows"]
    n = len(rows)
    rh, gap = 0.70, 0.26
    total = n * rh + (n - 1) * gap
    y0 = y + max(0.0, (h - total) / 2.0)
    lw, rw = 1.55, 1.80
    lx = x + 0.04
    rx = x + w - rw - 0.04
    for k, (src, verb, tgt, st) in enumerate(rows):
        yy = y0 + k * (rh + gap)
        col, lwid, dash = _STATUS.get(st, _STATUS["ok"])
        _box(s, lx, yy, lw, rh, src, 14, Z1, LINE, name="fig-node-l%d" % (k + 1))
        _box(s, rx, yy, rw, rh, tgt, 14, Z1, col, bold=True, name="fig-node-r%d" % (k + 1))
        _arrow(s, lx + lw + 0.02, yy + rh / 2.0, rx - 0.02, yy + rh / 2.0, col, lwid, dash,
               name="fig-arrow-%d" % (k + 1))
        _label(s, lx + lw + 0.04, yy + rh / 2.0 - 0.34, rx - lx - lw - 0.08, 0.28, verb, 14,
               INK3, name="fig-verb-%d" % (k + 1))


def _chain(s, spec, x, y, w, h, acc=RED):
    nodes = spec["nodes"]
    root = spec.get("root", len(nodes) - 1)
    n = len(nodes)
    bh, gap = 0.62, 0.26
    total = n * bh + (n - 1) * gap
    y0 = y + max(0.0, (h - total) / 2.0)
    bw = w - 0.30
    bx = x + 0.15
    for k, t in enumerate(nodes):
        yy = y0 + k * (bh + gap)
        if k == root:
            _box(s, bx, yy, bw, bh, t, 18, acc, acc, bold=True, color=RGBColor(0xFF, 0xFF, 0xFF),
                 name="fig-node-%d" % (k + 1))
        else:
            _box(s, bx, yy, bw, bh, t, 18, Z1, LINE, name="fig-node-%d" % (k + 1))
        if k < n - 1:
            _arrow(s, bx + bw / 2.0, yy + bh + 0.01, bx + bw / 2.0, yy + bh + gap - 0.01,
                   INK3, 1.75, name="fig-arrow-%d" % (k + 1))


def caption(s, x, y, w, h, text, size=20, acc=RED):
    return _label(s, x, y, w, h, text, size, INK2, name="fig-cap", align=PP_ALIGN.LEFT)


def draw_fig(s, spec, x, y, w, h, acc=RED):
    t = spec.get("type")
    if t == "rows":
        _rows(s, spec, x, y, w, h, acc)
    elif t == "chain":
        _chain(s, spec, x, y, w, h, acc)
    else:
        raise ValueError("unknown fig type: %r" % (t,))