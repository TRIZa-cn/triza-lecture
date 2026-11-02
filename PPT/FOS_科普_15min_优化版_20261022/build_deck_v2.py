# -*- coding: utf-8 -*-
"""
重建 FOS 科普 15 分钟课件（15min 版）

设计目标
1. 页面上只留听众需要看的东西：没有「建议用时 / 讲者要点 / 来源 / 本页做法」。
2. notes 是可直接逐字照读的口语稿：连贯、多连接词与转折词、不含任何提醒。
3. 版式统一：章节色带 + 标题区 + 卡片网格 + 页脚页码；图表用原生形状绘制，可编辑。
"""
import sys
from pathlib import Path

from lxml import etree
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, MSO_AUTO_SIZE, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Inches, Pt

OUT_DIR = Path(r"D:\doc\ZGCA\Lectures\TRIZ\FOS_科普_15min_优化版_20261022")
PPTX_PATH = OUT_DIR / "FOS_科普_15min_优化版.pptx"

# ---------------------------------------------------------------- 设计令牌
NAVY = "0E2A47"
INK = "16283C"
BLUE = "1E6BA8"
SKY = "4E93C4"
AMBER = "D98A26"
CORAL = "BE5138"
GREEN = "2B7A63"
MUTED = "6C7C8F"
LINE = "D9E2EB"
SOFT = "F4F8FC"
SOFT2 = "E7EFF7"
WHITE = "FFFFFF"
FONT = "Microsoft YaHei"

SW, SH = 13.3333, 7.5
LM = 0.72
CW = SW - 2 * LM
FOOT_Y = 6.96
BODY_TOP = 1.98
BODY_TOP_N = 1.62


def C(h):
    return RGBColor.from_string(h)


def _apply_font(run, sz, bold, color, font=FONT, spacing=None):
    f = run.font
    f.name = font
    f.size = Pt(sz)
    f.bold = bold
    f.color.rgb = C(color)
    rPr = run._r.get_or_add_rPr()
    for tag in ("a:ea", "a:cs"):
        el = rPr.find(qn(tag))
        if el is None:
            el = etree.SubElement(rPr, qn(tag))
        el.set("typeface", font)
    if spacing is not None:
        rPr.set("spc", str(int(spacing * 100)))


def text(slide, x, y, w, h, paras, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP,
         line=1.32, margins=(0, 0, 0, 0), name=None):
    tb = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    if name:
        tb.name = name
    tf = tb.text_frame
    tf.word_wrap = True
    tf.auto_size = MSO_AUTO_SIZE.NONE
    tf.vertical_anchor = anchor
    tf.margin_left = Inches(margins[0])
    tf.margin_right = Inches(margins[1])
    tf.margin_top = Inches(margins[2])
    tf.margin_bottom = Inches(margins[3])
    for i, p in enumerate(paras):
        para = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        para.alignment = p.get("align", align)
        para.line_spacing = p.get("ls", line)
        if p.get("sp"):
            para.space_before = Pt(p["sp"])
        r = para.add_run()
        r.text = p["t"]
        _apply_font(r, p.get("sz", 14), p.get("b", False), p.get("c", INK),
                    p.get("font", FONT), p.get("spc"))
    return tb


def box(slide, x, y, w, h, fill=None, line=None, radius=0.05,
        shape=MSO_SHAPE.ROUNDED_RECTANGLE, lw=1.0, name=None):
    sp = slide.shapes.add_shape(shape, Inches(x), Inches(y), Inches(w), Inches(h))
    if name:
        sp.name = name
    if shape == MSO_SHAPE.ROUNDED_RECTANGLE:
        try:
            sp.adjustments[0] = radius
        except (IndexError, ValueError):
            pass
    if fill is None:
        sp.fill.background()
    else:
        sp.fill.solid()
        sp.fill.fore_color.rgb = C(fill)
    if line is None:
        sp.line.fill.background()
    else:
        sp.line.color.rgb = C(line)
        sp.line.width = Pt(lw)
    sp.shadow.inherit = False
    sp.text_frame.word_wrap = True
    return sp


def label(sp, paras, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE,
          margins=(0.06, 0.06, 0.02, 0.02), line=1.2):
    tf = sp.text_frame
    tf.word_wrap = True
    tf.auto_size = MSO_AUTO_SIZE.NONE
    tf.vertical_anchor = anchor
    tf.margin_left = Inches(margins[0])
    tf.margin_right = Inches(margins[1])
    tf.margin_top = Inches(margins[2])
    tf.margin_bottom = Inches(margins[3])
    for i, p in enumerate(paras):
        para = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        para.alignment = p.get("align", align)
        para.line_spacing = p.get("ls", line)
        if p.get("sp"):
            para.space_before = Pt(p["sp"])
        r = para.add_run()
        r.text = p["t"]
        _apply_font(r, p.get("sz", 14), p.get("b", False), p.get("c", INK),
                    p.get("font", FONT), p.get("spc"))
    return sp


def rect(slide, x, y, w, h, fill, name=None):
    return box(slide, x, y, w, h, fill=fill, line=None,
               shape=MSO_SHAPE.RECTANGLE, name=name)


# ---------------------------------------------------------------- 页面骨架
def new_slide(prs):
    return prs.slides.add_slide(prs.slide_layouts[6])


def footer(slide, page, total, section):
    rect(slide, LM, FOOT_Y, CW, 0.008, LINE, name="footer-line")
    text(slide, LM, FOOT_Y + 0.07, 8.2, 0.3,
         [{"t": section, "sz": 9.5, "c": MUTED, "spc": 0.6}], name="footer-left")
    text(slide, LM + CW - 2.2, FOOT_Y + 0.07, 2.2, 0.3,
         [{"t": "%02d / %02d" % (page, total), "sz": 9.5, "c": MUTED, "spc": 0.6}],
         align=PP_ALIGN.RIGHT, name="footer-right")


def header(slide, kicker, title, sub=None, accent=AMBER):
    rect(slide, LM, 0.44, 0.055, 0.30, accent, name="kicker-bar")
    text(slide, LM + 0.19, 0.42, CW - 0.19, 0.30,
         [{"t": kicker, "sz": 10.5, "c": MUTED, "spc": 1.2}], name="kicker")
    text(slide, LM, 0.74, CW, 0.66,
         [{"t": title, "sz": 25, "b": True, "c": NAVY, "ls": 1.08}], name="title")
    if sub:
        text(slide, LM, 1.44, CW, 0.40,
             [{"t": sub, "sz": 13.5, "c": MUTED, "ls": 1.15}], name="subtitle")


def section_tag(slide, x, y, w, t, fill=SOFT2, color=BLUE, sz=10):
    sp = box(slide, x, y, w, 0.30, fill=fill, line=None, radius=0.5)
    label(sp, [{"t": t, "sz": sz, "b": True, "c": color, "spc": 0.6}])
    return sp


def card(slide, x, y, w, h, fill=WHITE, line=LINE, radius=0.045, lw=1.0):
    return box(slide, x, y, w, h, fill=fill, line=line, radius=radius, lw=lw)


def numdot(slide, x, y, d, t, fill=NAVY, color=WHITE, sz=12):
    sp = box(slide, x, y, d, d, fill=fill, line=None, shape=MSO_SHAPE.OVAL)
    label(sp, [{"t": t, "sz": sz, "b": True, "c": color}])
    return sp


def arrow_right(slide, x, y, w, h, color=SKY):
    return box(slide, x, y, w, h, fill=color, line=None, shape=MSO_SHAPE.RIGHT_ARROW)


def arrow_down(slide, x, y, w, h, color=SKY):
    return box(slide, x, y, w, h, fill=color, line=None, shape=MSO_SHAPE.DOWN_ARROW)


# ---------------------------------------------------------------- 复合组件
def step_row(slide, x, y, w, h, items, accent=BLUE, gap=0.18, numbered=True,
             title_sz=13, body_sz=11.5):
    n = len(items)
    cw = (w - gap * (n - 1)) / n
    for i, (t, body) in enumerate(items):
        cx = x + i * (cw + gap)
        card(slide, cx, y, cw, h)
        rect(slide, cx, y + 0.16, 0.045, h - 0.32, accent)
        top = y + 0.20
        if numbered:
            text(slide, cx + 0.24, top, 0.6, 0.32,
                 [{"t": "%02d" % (i + 1), "sz": 11, "b": True, "c": accent, "spc": 0.5}])
            top += 0.36
        text(slide, cx + 0.24, top, cw - 0.44, 0.34,
             [{"t": t, "sz": title_sz, "b": True, "c": NAVY, "ls": 1.12}])
        if body:
            text(slide, cx + 0.24, top + 0.42, cw - 0.44, h - (top - y) - 0.56,
                 [{"t": body, "sz": body_sz, "c": MUTED, "ls": 1.30}])


def divider(prs, num, title, lead, section, page, total):
    s = new_slide(prs)
    rect(s, 0, 0, SW, SH, NAVY, name="divider-bg")
    rect(s, 0, 0, 0.30, SH, AMBER, name="divider-edge")
    text(s, 1.05, 2.02, 2.4, 1.5,
         [{"t": num, "sz": 72, "b": True, "c": "27476B", "ls": 0.95}], name="divider-num")
    text(s, 1.05, 3.52, 9.6, 0.95,
         [{"t": title, "sz": 34, "b": True, "c": WHITE, "ls": 1.10}], name="divider-title")
    rect(s, 1.10, 4.62, 1.05, 0.055, AMBER, name="divider-rule")
    text(s, 1.05, 4.88, 9.4, 0.9,
         [{"t": lead, "sz": 14, "c": "A9C2D8", "ls": 1.35}], name="divider-lead")
    text(s, LM + CW - 2.2, FOOT_Y + 0.07, 2.2, 0.3,
         [{"t": "%02d / %02d" % (page, total), "sz": 9.5, "c": "6E8CA6", "spc": 0.6}],
         align=PP_ALIGN.RIGHT, name="footer-right")
    return s


# ---------------------------------------------------------------- 40 条发明原理
PRINCIPLES = [
    "分割", "抽取", "局部质量", "非对称", "合并",
    "通用性", "嵌套", "重量补偿", "预先反作用", "预先作用",
    "预先防范", "等势", "反向作用", "曲面化", "动态化",
    "不足或过分的动作", "维数变化", "机械振动", "周期性作用", "有效作用的连续性",
    "快速通过", "变害为利", "反馈", "中介物", "自服务",
    "复制", "一次性用品", "机械系统替代", "气压或液压结构", "柔性壳体或薄膜",
    "多孔材料", "颜色变化", "同质性", "抛弃与修复", "参数变化",
    "相变", "热膨胀", "加速氧化", "惰性环境", "复合材料",
]
HIGHLIGHT = {1, 3, 4, 10, 24, 31, 35}


def principle_grid(slide, start, y=2.02):
    cols, rws = 5, 4
    gx, gy = 0.16, 0.15
    cw = (CW - gx * (cols - 1)) / cols
    ch = 1.00
    for i in range(20):
        n = start + i
        r, c = divmod(i, cols)
        x = LM + c * (cw + gx)
        yy = y + r * (ch + gy)
        hot = n in HIGHLIGHT
        card(slide, x, yy, cw, ch, fill="FDF3E4" if hot else WHITE,
             line=AMBER if hot else LINE, lw=1.4 if hot else 1.0)
        text(slide, x + 0.16, yy + 0.12, 0.44, 0.26,
             [{"t": "%02d" % n, "sz": 10, "b": True,
               "c": AMBER if hot else "A8B6C4"}])
        text(slide, x + 0.16, yy + 0.44, cw - 0.30, 0.46,
             [{"t": PRINCIPLES[n - 1], "sz": 12.5, "b": hot,
               "c": NAVY if hot else INK, "ls": 1.05}])

# ================================================================ 内容
def build():
    prs = Presentation()
    prs.slide_width = Inches(SW)
    prs.slide_height = Inches(SH)
    total = 30
    p = [0]

    def page():
        p[0] += 1
        return p[0]

    # ---------------- 01 封面
    s = new_slide(prs)
    rect(s, 0, 0, SW, SH, NAVY, name="cover-bg")
    rect(s, 0, 0, SW, 0.10, AMBER, name="cover-top")
    for x, w, a in [(9.55, 3.78, "1B3D5E"), (10.90, 2.43, "22496E"), (11.95, 1.38, "2A567F")]:
        box(s, x, 0, w, SH, fill=a, line=None, shape=MSO_SHAPE.RECTANGLE)
    text(s, LM + 0.02, 1.42, 8.4, 0.34,
         [{"t": "TRIZ 创新方法 · 十五分钟科普", "sz": 12, "c": "9FBBD4", "spc": 2.2}])
    text(s, LM, 1.92, 9.2, 1.30,
         [{"t": "功能导向检索", "sz": 50, "b": True, "c": WHITE, "ls": 1.05}])
    text(s, LM, 3.24, 9.0, 0.52,
         [{"t": "Function-Oriented Search   ·   FOS", "sz": 15, "c": AMBER, "spc": 1.4}])
    rect(s, LM + 0.04, 4.00, 1.10, 0.05, AMBER, name="cover-rule")
    text(s, LM, 4.26, 8.6, 0.95,
         [{"t": "离开本领域去找已经长出来的结构，", "sz": 16, "c": "D7E4EF", "ls": 1.42},
          {"t": "再把功能迁回自己的问题。", "sz": 16, "c": "D7E4EF", "ls": 1.42}])
    text(s, LM, 5.72, 8.4, 0.34,
         [{"t": "刘文印   ·   广东工业大学", "sz": 12.5, "c": "8FAEC8", "spc": 0.8}])
    text(s, LM, 6.14, 8.4, 0.34,
         [{"t": "全程 15 分钟 · 一个工具 · 一个完整例子", "sz": 12.5, "c": "8FAEC8"}])
    page()

    # ---------------- 02 三个问题
    s = new_slide(prs)
    header(s, "开场", "今天只回答三个问题",
           "十五分钟不够讲全部，刚好够把一件事说透。")
    qs = [("为什么跨行业借鉴", "常常变成生搬硬套？", "功能抽得太粗，粗到什么都能类比。"),
          ("有没有一套可重复的动作", "而不是靠灵感？", "顺序固定，换个人做也走得通。"),
          ("怎么判断一次迁移", "是真有用，还是自我安慰？", "过三关：结构、前提、可观测效果。")]
    cw = (CW - 0.36 * 2) / 3
    for i, (l1, l2, note) in enumerate(qs):
        x = LM + i * (cw + 0.36)
        card(s, x, 2.10, cw, 3.55)
        rect(s, x, 2.10, cw, 0.055, AMBER if i == 0 else BLUE)
        numdot(s, x + 0.34, 2.44, 0.52, str(i + 1),
               fill=NAVY if i == 0 else "2C5C86", sz=14)
        text(s, x + 0.34, 3.20, cw - 0.68, 1.55,
             [{"t": l1, "sz": 17, "b": True, "c": NAVY, "ls": 1.24},
              {"t": l2, "sz": 17, "b": True, "c": NAVY, "ls": 1.24, "sp": 2}])
        rect(s, x + 0.34, 4.84, cw - 0.68, 0.012, LINE)
        text(s, x + 0.34, 4.98, cw - 0.68, 0.52,
             [{"t": note, "sz": 11.5, "c": MUTED, "ls": 1.28}])
    footer(s, page(), total, "开场")

    # ---------------- 03 TRIZ 是什么
    s = new_slide(prs)
    header(s, "第一部分 · 认识 TRIZ", "TRIZ：给发明性问题的一套求解框架")
    card(s, LM, BODY_TOP_N, 5.55, 4.35, fill=SOFT)
    rect(s, LM, BODY_TOP_N, 0.055, 4.35, BLUE)
    text(s, LM + 0.36, BODY_TOP_N + 0.28, 4.85, 0.30,
         [{"t": "定义", "sz": 10.5, "b": True, "c": BLUE, "spc": 1.2}])
    text(s, LM + 0.36, BODY_TOP_N + 0.70, 4.85, 1.60,
         [{"t": "一套基于归纳的、面向发明性问题的", "sz": 15.5, "c": INK, "ls": 1.40},
          {"t": "创造性求解理论框架。", "sz": 15.5, "c": INK, "ls": 1.40}])
    text(s, LM + 0.36, BODY_TOP_N + 2.46, 4.85, 1.55,
         [{"t": "它把大量专利里的解法，归纳成可以复用的原理，", "sz": 12.5, "c": MUTED, "ls": 1.38},
          {"t": "所以同一类矛盾，换个人做也能走到相似的解。", "sz": 12.5, "c": MUTED, "ls": 1.38}])
    rx = LM + 5.95
    rw = CW - 5.95
    for i, (t, b) in enumerate([("发明性问题", "没有现成方案，必须发明一个新方案才能解决。"),
                                ("核心动作", "把一句模糊的问题，收成一个清楚的矛盾。"),
                                ("可复用", "原理只认名字，不认行业，所以能跨领域。")]):
        yy = BODY_TOP_N + i * 1.47
        card(s, rx, yy, rw, 1.30)
        numdot(s, rx + 0.26, yy + 0.44, 0.42, str(i + 1), fill="2C5C86", sz=11)
        text(s, rx + 0.86, yy + 0.20, rw - 1.10, 0.34,
             [{"t": t, "sz": 14, "b": True, "c": NAVY}])
        text(s, rx + 0.86, yy + 0.60, rw - 1.10, 0.58,
             [{"t": b, "sz": 11.5, "c": MUTED, "ls": 1.26}])
    text(s, LM, 6.14, CW, 0.40,
         [{"t": "它追求的不是灵感，是可重复。", "sz": 13, "b": True, "c": AMBER}])
    footer(s, page(), total, "认识 TRIZ")

    # ---------------- 04 / 05 四十条原理
    for idx, start in enumerate([1, 21]):
        s = new_slide(prs)
        header(s, "第一部分 · 认识 TRIZ",
               "四十条发明原理（一）" if idx == 0 else "四十条发明原理（二）",
               "四十条不是四十个公式，是四十个提问的角度。" if idx == 0
               else "同一条原理在机械、化工、软件里都成立，所以 TRIZ 天生鼓励你往外看。")
        principle_grid(s, start, y=2.02)
        footer(s, page(), total, "认识 TRIZ")

    # ---------------- 06 收成矛盾
    s = new_slide(prs)
    header(s, "第一部分 · 认识 TRIZ", "第一步：把问题收成一个矛盾")
    card(s, LM, BODY_TOP_N, 5.45, 3.05, fill=SOFT)
    section_tag(s, LM + 0.32, BODY_TOP_N + 0.24, 1.72, "初始问题")
    text(s, LM + 0.32, BODY_TOP_N + 0.74, 4.80, 0.95,
         [{"t": "怎么绕过网络审查？", "sz": 21, "b": True, "c": NAVY, "ls": 1.20}])
    text(s, LM + 0.32, BODY_TOP_N + 1.88, 4.80, 0.90,
         [{"t": "这里只讨论方法，不讨论任何具体做法。", "sz": 12, "c": MUTED, "ls": 1.30}])
    arrow_right(s, 6.28, BODY_TOP_N + 1.26, 0.62, 0.44)
    card(s, 7.05, BODY_TOP_N, CW - 6.33, 3.05, fill="FDF3E4", line=AMBER, lw=1.4)
    section_tag(s, 7.35, BODY_TOP_N + 0.24, 2.30, "收成物理矛盾", fill="F8E4C6", color="A5661A")
    text(s, 7.35, BODY_TOP_N + 0.74, CW - 6.93, 1.75,
         [{"t": "同一段网络流量，", "sz": 14.5, "c": INK, "ls": 1.34},
          {"t": "必须同时是标准的 HTTPS 网页流量，", "sz": 14.5, "c": INK, "ls": 1.34},
          {"t": "又必须是非标准的 DoH 查询流量。", "sz": 14.5, "c": INK, "ls": 1.34}])
    text(s, 7.35, BODY_TOP_N + 2.50, CW - 6.93, 0.40,
         [{"t": "用大白话说：既要又要。", "sz": 14, "b": True, "c": "A5661A"}])
    card(s, LM, 5.28, CW, 1.06, fill=SOFT)
    text(s, LM + 0.34, 5.44, CW - 0.68, 0.74,
         [{"t": "一旦说成「既要又要」，问题就从「想不出办法」，变成了「往哪个方向找办法」。",
           "sz": 13.5, "b": True, "c": NAVY, "ls": 1.28}])
    footer(s, page(), total, "认识 TRIZ")

    # ---------------- 07 对上原理
    s = new_slide(prs)
    header(s, "第一部分 · 认识 TRIZ", "第二步：让矛盾去对上原理")
    trio = [("原理 01", "分割", "分开处理，也叫分而治之。", AMBER),
            ("原理 03", "局部质量", "均匀改成不均匀，让各部分功能不同。", BLUE),
            ("原理 31", "多孔材料", "把要送的东西放进看起来无害的空隙。", GREEN)]
    cw = (CW - 0.34 * 2) / 3
    for i, (k, t, b, col) in enumerate(trio):
        x = LM + i * (cw + 0.34)
        card(s, x, BODY_TOP_N, cw, 2.72)
        rect(s, x, BODY_TOP_N, cw, 0.055, col)
        text(s, x + 0.30, BODY_TOP_N + 0.28, cw - 0.60, 0.30,
             [{"t": k, "sz": 10.5, "b": True, "c": col, "spc": 1.0}])
        text(s, x + 0.30, BODY_TOP_N + 0.66, cw - 0.60, 0.50,
             [{"t": t, "sz": 20, "b": True, "c": NAVY}])
        text(s, x + 0.30, BODY_TOP_N + 1.36, cw - 0.60, 1.10,
             [{"t": b, "sz": 12.5, "c": MUTED, "ls": 1.34}])
    card(s, LM, 4.72, CW, 1.62, fill=SOFT)
    text(s, LM + 0.34, 4.90, CW - 0.68, 0.32,
         [{"t": "另外还点了名的四条", "sz": 10.5, "b": True, "c": MUTED, "spc": 1.0}])
    text(s, LM + 0.34, 5.26, CW - 0.68, 0.44,
         [{"t": "04 非对称      ·      24 中介物      ·      35 参数变化      ·      10 预先作用",
           "sz": 15, "b": True, "c": NAVY, "ls": 1.20}])
    text(s, LM + 0.34, 5.78, CW - 0.68, 0.40,
         [{"t": "原理给的不是答案，是几个可以继续往下追的方向。", "sz": 12, "c": MUTED}])
    footer(s, page(), total, "认识 TRIZ")

    # ---------------- 08 总图
    s = new_slide(prs)
    header(s, "第一部分 · 认识 TRIZ", "整套 TRIZ 是一条回路，FOS 只占其中一步")
    bw, bh = 3.30, 1.05
    lx, rx = LM + 0.10, SW - LM - 0.10 - bw
    ty, by = 1.86, 4.62
    nodes = [("待解决的问题", lx, by, "问题", SOFT2, NAVY),
             ("问题模型", lx, ty, "转化后", SOFT2, NAVY),
             ("解决方案模型", rx, ty, "原理提示", "FDF3E4", "A5661A"),
             ("具体解决方案", rx, by, "应用后", SOFT2, NAVY)]
    for t, x, y, tag, fill, col in nodes:
        card(s, x, y, bw, bh, fill=fill, line=None)
        text(s, x + 0.20, y + 0.16, bw - 0.40, 0.24,
             [{"t": tag, "sz": 9.5, "c": MUTED, "spc": 0.8}])
        text(s, x + 0.20, y + 0.46, bw - 0.40, 0.44,
             [{"t": t, "sz": 16, "b": True, "c": col}])
    box(s, lx + bw / 2 - 0.16, 3.48, 0.32, 0.62, fill=SKY, line=None,
        shape=MSO_SHAPE.UP_ARROW)
    text(s, lx + bw / 2 + 0.24, 3.60, 1.4, 0.30,
         [{"t": "转化", "sz": 11, "b": True, "c": SKY}])
    arrow_down(s, rx + bw / 2 - 0.16, 2.99, 0.32, 0.62, SKY)
    text(s, rx + bw / 2 + 0.24, 3.10, 1.4, 0.30,
         [{"t": "应用", "sz": 11, "b": True, "c": SKY}])
    mid = box(s, lx + bw + 0.34, ty + (bh - 0.60) / 2, (rx - lx - bw - 0.68), 0.60,
              fill=AMBER, line=None, radius=0.28)
    label(mid, [{"t": "TRIZ 工具", "sz": 15, "b": True, "c": WHITE, "spc": 1.0}])
    arrow_right(s, lx + bw + 0.05, ty + (bh - 0.22) / 2, 0.24, 0.22, "C3D2DE")
    arrow_right(s, rx - 0.30, ty + (bh - 0.22) / 2, 0.24, 0.22, "C3D2DE")
    box(s, lx + bw + 0.34, by + bh / 2 - 0.02, (rx - lx - bw - 0.68), 0.035,
        fill="B9C6D2", line=None, shape=MSO_SHAPE.RECTANGLE)
    text(s, lx + bw + 0.34, by + bh / 2 - 0.52, 1.6, 0.34,
         [{"t": "直达：难", "sz": 10.5, "c": "8B9AA8"}])
    text(s, lx + bw + 0.34, by + bh / 2 + 0.10, 3.4, 0.34,
         [{"t": "被划掉的捷径", "sz": 10, "c": "A8B6C4"}])
    card(s, LM, 6.02, CW, 0.62, fill=SOFT)
    text(s, LM + 0.30, 6.16, CW - 0.60, 0.38,
         [{"t": "别从原始问题直接跳到做法。功能导向检索，就住在中间那个叫「TRIZ 工具」的格子里。",
           "sz": 12.5, "b": True, "c": NAVY}])
    footer(s, page(), total, "认识 TRIZ")

    # ---------------- 09 五级发明（条形图）
    s = new_slide(prs)
    header(s, "第一部分 · 认识 TRIZ", "发明分五级，跨领域迁移落在第三级",
           "TRIZ 对专利发明的分级统计，今天只关心第三级这一档。")
    base_x, base_w = 3.10, 4.60
    rows_data = [("第 1 级", "标准化", 32, "用本领域知识做简单改进"),
                 ("第 2 级", "改良化", 45, "专业领域内已知的方法"),
                 ("第 3 级", "模式内创新", 18, "参考其他领域的技术与方法"),
                 ("第 4 级", "模式外创新", 4, "科学原理首次用于本领域"),
                 ("第 5 级", "新发明", 1, "超越当代已有知识")]
    y0, rh, gap = 2.24, 0.62, 0.14
    for i, (lvl, name, pct, note) in enumerate(rows_data):
        yy = y0 + i * (rh + gap)
        hot = (i == 2)
        if hot:
            card(s, LM, yy - 0.09, CW, rh + 0.18, fill="FDF3E4", line=AMBER, lw=1.4)
        text(s, LM + 0.18, yy + 0.14, 1.05, 0.32,
             [{"t": lvl, "sz": 11.5, "b": True, "c": AMBER if hot else MUTED}])
        text(s, LM + 1.28, yy + 0.12, 1.45, 0.38,
             [{"t": name, "sz": 13, "b": True, "c": NAVY}])
        bw2 = max(base_w * pct / 50.0, 0.10)
        box(s, base_x, yy + 0.12, bw2, 0.36, fill=AMBER if hot else "9FBCD4",
            line=None, radius=0.18)
        text(s, base_x + bw2 + 0.14, yy + 0.12, 0.95, 0.36,
             [{"t": ("%d%%" % pct) if pct > 1 else "＜1%", "sz": 13, "b": True,
               "c": AMBER if hot else "7C8C9C"}])
        text(s, 8.15, yy + 0.14, CW + LM - 8.15, 0.34,
             [{"t": note, "sz": 11, "c": MUTED}])
    rect(s, base_x, y0 - 0.06, 0.012, 5 * rh + 4 * gap - 0.10, LINE)
    card(s, LM, 6.02, CW, 0.62, fill=SOFT)
    text(s, LM + 0.30, 6.16, CW - 0.60, 0.38,
         [{"t": "功能导向检索对准的，就是这百分之十八：把别的行业已知的方法，迁到自己的问题里。",
           "sz": 12.5, "b": True, "c": NAVY}])
    footer(s, page(), total, "认识 TRIZ")

    # ---------------- 10 章节页
    divider(prs, "02", "功能导向检索是什么",
            "四个动作，一个顺序：抽功能 → 泛化名词 → 选领先领域 → 下沉并迁回。",
            "FOS 是什么", page(), total)

    # ---------------- 11 四步
    s = new_slide(prs)
    header(s, "第二部分 · 功能导向检索", "把别的行业做熟的结构，搬到你的问题里",
           "四个动作，顺序不能乱。")
    step_row(s, LM, 2.10, CW, 2.28, [
        ("抽功能", "把要解决的问题，说成一句动词短语。"),
        ("泛化名词", "把功能里的名词往上收一层。"),
        ("选领先领域", "挑一个在这件事上被逼得更狠的行业。"),
        ("下沉并迁回", "落到具体物，再把它搬回来。"),
    ], accent=BLUE, title_sz=14, body_sz=11.5)
    card(s, LM, 4.66, CW, 1.68, fill=SOFT)
    rect(s, LM, 4.66, 0.055, 1.68, AMBER)
    text(s, LM + 0.36, 4.84, CW - 0.72, 0.32,
         [{"t": "一句话定义", "sz": 10.5, "b": True, "c": AMBER, "spc": 1.0}])
    text(s, LM + 0.36, 5.22, CW - 0.72, 0.90,
         [{"t": "把别的行业已经做熟的事，搬到你的问题里；不是查重，也不是头脑风暴。",
           "sz": 16, "b": True, "c": NAVY, "ls": 1.30}])
    footer(s, page(), total, "FOS 是什么")

    # ---------------- 12 查重 vs FOS
    s = new_slide(prs)
    header(s, "第二部分 · 功能导向检索", "查重在防守，功能导向检索在找解法")
    cw = (CW - 1.55) / 2
    lx = LM
    for x, ttl, sub, lines, col in [
            (lx, "查重", "有没有人做过一样的事？",
             ["用本领域的术语", "越精确越好"], "8FA0B0"),
            (lx + cw + 1.55, "功能导向检索", "别处有没有人解决过同一个功能？",
             ["用上位词和别人的行话", "越不像本行越好"], BLUE)]:
        card(s, x, BODY_TOP_N, cw, 3.35)
        rect(s, x, BODY_TOP_N, cw, 0.055, col)
        text(s, x + 0.34, BODY_TOP_N + 0.28, cw - 0.68, 0.40,
             [{"t": ttl, "sz": 19, "b": True, "c": NAVY}])
        text(s, x + 0.34, BODY_TOP_N + 0.84, cw - 0.68, 0.80,
             [{"t": sub, "sz": 14, "c": INK, "ls": 1.30}])
        rect(s, x + 0.34, BODY_TOP_N + 1.80, cw - 0.68, 0.012, LINE)
        text(s, x + 0.34, BODY_TOP_N + 1.96, cw - 0.68, 1.10,
             [{"t": lines[0], "sz": 12.5, "c": MUTED, "ls": 1.34},
              {"t": lines[1], "sz": 12.5, "c": MUTED, "ls": 1.34}])
    mid = box(s, lx + cw + 0.28, BODY_TOP_N + 1.30, 0.98, 0.44, fill=SOFT2,
              line=None, radius=0.5)
    label(mid, [{"t": "同一批文献", "sz": 9, "b": True, "c": BLUE}])
    card(s, LM, 5.42, CW, 0.98, fill=SOFT)
    text(s, LM + 0.34, 5.58, CW - 0.68, 0.66,
         [{"t": "两边搜索的方向几乎是相反的。拿本行的词去搜别的行，当然搜不到东西。",
           "sz": 13.5, "b": True, "c": NAVY, "ls": 1.28}])
    footer(s, page(), total, "FOS 是什么")

    # ---------------- 13 约束不同结构不同
    s = new_slide(prs)
    header(s, "第二部分 · 功能导向检索", "约束不同，同一个功能会长出不同结构",
           "功能相同：防止超量使用。")
    trio = [("软件", "计数器 ＋ 权限", "因为程序随时可以改", BLUE),
            ("核工业", "看已经烧掉多少", "因为烧掉的部分改不回来", AMBER),
            ("医院", "收回空安瓿再数", "因为药用完就没了", GREEN)]
    cw = (CW - 0.34 * 2) / 3
    for i, (t, s1, s2, col) in enumerate(trio):
        x = LM + i * (cw + 0.34)
        card(s, x, BODY_TOP, cw, 3.10)
        rect(s, x, BODY_TOP, cw, 0.055, col)
        text(s, x + 0.32, BODY_TOP + 0.30, cw - 0.64, 0.40,
             [{"t": t, "sz": 15, "b": True, "c": col}])
        text(s, x + 0.32, BODY_TOP + 0.88, cw - 0.64, 0.85,
             [{"t": s1, "sz": 16.5, "b": True, "c": NAVY, "ls": 1.26}])
        text(s, x + 0.32, BODY_TOP + 1.96, cw - 0.64, 0.85,
             [{"t": s2, "sz": 12, "c": MUTED, "ls": 1.32}])
    card(s, LM, 5.52, CW, 0.95, fill=SOFT)
    text(s, LM + 0.34, 5.68, CW - 0.68, 0.62,
         [{"t": "功能是同一句话，结构完全不同，差别全在约束上。",
           "sz": 13.5, "b": True, "c": NAVY}])
    footer(s, page(), total, "FOS 是什么")

    # ---------------- 14 搬结构
    s = new_slide(prs)
    header(s, "第二部分 · 功能导向检索", "要搬的不是功能，是被约束逼出来的结构")
    cw = (CW - 1.10) / 2
    card(s, LM, BODY_TOP_N, cw, 3.30, fill=SOFT)
    text(s, LM + 0.34, BODY_TOP_N + 0.26, cw - 0.68, 0.30,
         [{"t": "功能", "sz": 10.5, "b": True, "c": MUTED, "spc": 1.2}])
    text(s, LM + 0.34, BODY_TOP_N + 0.66, cw - 0.68, 0.60,
         [{"t": "防止超量使用", "sz": 21, "b": True, "c": NAVY}])
    text(s, LM + 0.34, BODY_TOP_N + 1.52, cw - 0.68, 1.40,
         [{"t": "抽到最后，大家得到的", "sz": 13, "c": MUTED, "ls": 1.36},
          {"t": "往往是同一句话。", "sz": 13, "c": MUTED, "ls": 1.36},
          {"t": "谁都会写。", "sz": 13, "b": True, "c": CORAL, "ls": 1.36}])
    x2 = LM + cw + 1.10
    card(s, x2, BODY_TOP_N, cw, 3.30, fill="FDF3E4", line=AMBER, lw=1.4)
    text(s, x2 + 0.34, BODY_TOP_N + 0.26, cw - 0.68, 0.30,
         [{"t": "结构", "sz": 10.5, "b": True, "c": "A5661A", "spc": 1.2}])
    text(s, x2 + 0.34, BODY_TOP_N + 0.66, cw - 0.68, 0.60,
         [{"t": "计数器 · 燃耗 · 交回空瓶", "sz": 16, "b": True, "c": NAVY}])
    text(s, x2 + 0.34, BODY_TOP_N + 1.52, cw - 0.68, 1.40,
         [{"t": "落下来是一个部件、一个字段，", "sz": 13, "c": "8A5A1C", "ls": 1.36},
          {"t": "或者一台状态机。", "sz": 13, "c": "8A5A1C", "ls": 1.36},
          {"t": "这才是值钱的东西。", "sz": 13, "b": True, "c": "A5661A", "ls": 1.36}])
    arrow_right(s, LM + cw + 0.34, BODY_TOP_N + 1.36, 0.42, 0.52, AMBER)
    text(s, LM, 5.72, CW, 0.40,
         [{"t": "后面所有的步骤，其实都是在找这个结构。", "sz": 13.5, "b": True, "c": NAVY}])
    footer(s, page(), total, "FOS 是什么")

    # ---------------- 15 四级阶梯
    s = new_slide(prs)
    header(s, "第二部分 · 功能导向检索", "四级阶梯写下来，别人才复现得了你的检索")
    steps = [("本领域的词", "离线槽位", "你自己行业里的说法", "9FBCD4"),
             ("上位词", "有限次预发的消耗单元", "去掉领域色彩，往上收一层", "7FA8C6"),
             ("领先领域", "航天", "在这件事上被逼得更狠", "4E93C4"),
             ("具体物", "在轨消耗品包 ＋ 补给窗口", "能看见、能拆开的东西", AMBER)]
    cw = (CW - 0.30 * 3) / 4
    for i, (t, s1, s2, col) in enumerate(steps):
        x = LM + i * (cw + 0.30)
        card(s, x, 2.06, cw, 2.86)
        rect(s, x, 2.06, cw, 0.055, col)
        text(s, x + 0.26, 2.26, cw - 0.52, 0.30,
             [{"t": "第 %d 级" % (i + 1), "sz": 10, "b": True, "c": col, "spc": 1.0}])
        text(s, x + 0.26, 2.60, cw - 0.52, 0.40,
             [{"t": t, "sz": 14, "b": True, "c": NAVY}])
        text(s, x + 0.26, 3.08, cw - 0.52, 1.05,
             [{"t": s1, "sz": 12.5, "b": True, "c": INK, "ls": 1.30}])
        text(s, x + 0.26, 4.20, cw - 0.52, 0.62,
             [{"t": s2, "sz": 10.5, "c": MUTED, "ls": 1.26}])
        if i < 3:
            arrow_right(s, x + cw + 0.02, 3.32, 0.26, 0.26, "B9C6D2")
    card(s, LM, 5.22, CW, 1.28, fill=SOFT)
    text(s, LM + 0.34, 5.40, CW - 0.68, 0.32,
         [{"t": "自左向右，就是一条检索路径", "sz": 10.5, "b": True, "c": MUTED, "spc": 1.0}])
    text(s, LM + 0.34, 5.74, CW - 0.68, 0.60,
         [{"t": "少写一级，别人就没法判断你有没有跳步。", "sz": 14, "b": True, "c": NAVY}])
    footer(s, page(), total, "FOS 是什么")

    # ---------------- 16 泛化粒度
    s = new_slide(prs)
    header(s, "第二部分 · 功能导向检索", "泛化的粒度：太浅搜不到，太深等于没搜")
    items = [("太浅", "瓶子  →  罐子", "还在包装行业里打转，搜出来还是老面孔。", CORAL, "FBF0ED"),
             ("太深", "瓶子  →  东西", "什么都搜得到，等于没搜。", CORAL, "FBF0ED"),
             ("合适", "瓶子  →  容器",
              "说出这个词，马上能举出三个行业的实物：疫苗瓶、太空舱、乏燃料罐。", GREEN, "EFF7F3")]
    for i, (tag, path, note, col, bg) in enumerate(items):
        yy = 1.98 + i * 1.36
        card(s, LM, yy, CW, 1.26, fill=bg, line=None)
        rect(s, LM, yy, 0.055, 1.26, col)
        section_tag(s, LM + 0.32, yy + 0.20, 0.92, tag, fill=WHITE, color=col, sz=11)
        text(s, LM + 1.42, yy + 0.16, 3.30, 0.48,
             [{"t": path, "sz": 17, "b": True, "c": NAVY}])
        text(s, LM + 4.96, yy + 0.16, CW - 5.30, 0.95,
             [{"t": note, "sz": 12.5, "c": INK, "ls": 1.32}])
    card(s, LM, 6.06, CW, 0.58, fill=SOFT)
    text(s, LM + 0.30, 6.16, CW - 0.60, 0.36,
         [{"t": "标准很硬：举不出三个行业的实物，就再收一层，或者退回一层。",
           "sz": 12, "b": True, "c": NAVY}])
    footer(s, page(), total, "FOS 是什么")

    # ---------------- 17 矛盾
    s = new_slide(prs)
    header(s, "第二部分 · 功能导向检索", "这一步的矛盾：跳级和过度，都会把检索做空")
    cw = (CW - 0.60) / 2
    for i, (t, s1, s2) in enumerate([
            ("跳级", "只拿领域词去搜", "回来的全是同构方案，没有新结构。"),
            ("过度泛化", "上位词大到无法下沉", "最后写不出一句可执行的结构。")]):
        x = LM + i * (cw + 0.60)
        card(s, x, BODY_TOP_N, cw, 2.80)
        text(s, x + 0.36, BODY_TOP_N + 0.28, cw - 0.72, 0.40,
             [{"t": t, "sz": 18, "b": True, "c": CORAL}])
        text(s, x + 0.36, BODY_TOP_N + 0.86, cw - 0.72, 0.46,
             [{"t": s1, "sz": 15, "b": True, "c": NAVY}])
        rect(s, x + 0.36, BODY_TOP_N + 1.50, cw - 0.72, 0.012, LINE)
        text(s, x + 0.36, BODY_TOP_N + 1.68, cw - 0.72, 0.90,
             [{"t": s2, "sz": 12.5, "c": MUTED, "ls": 1.34}])
    box(s, LM + cw + 0.18, BODY_TOP_N + 1.16, 0.24, 0.48, fill="C6D3DE",
        line=None, radius=0.4)
    card(s, LM, 5.02, CW, 1.44, fill="FDF3E4", line=AMBER, lw=1.4)
    text(s, LM + 0.36, 5.20, CW - 0.72, 0.32,
         [{"t": "两条路看起来相反，结果一样", "sz": 10.5, "b": True, "c": "A5661A", "spc": 1.0}])
    text(s, LM + 0.36, 5.54, CW - 0.72, 0.76,
         [{"t": "都是把检索做空。合适的抽象夹在中间：说出这个词，还能点出三个行业的实物。",
           "sz": 14, "b": True, "c": NAVY, "ls": 1.30}])
    footer(s, page(), total, "FOS 是什么")

    # ---------------- 18 领先领域
    s = new_slide(prs)
    header(s, "第二部分 · 功能导向检索", "领先领域，是被迫把这件事做得更狠的地方",
           "别挑听起来先进的行业，要挑被逼得更极端的。")
    trio = [("航天", "坏了没法派人去修", "地球独立", BLUE),
            ("核工业", "记录改不了", "用累计消耗证明", AMBER),
            ("医院", "必须能追责", "交回实物再核销", GREEN)]
    cw = (CW - 0.34 * 2) / 3
    for i, (t, s1, s2, col) in enumerate(trio):
        x = LM + i * (cw + 0.34)
        card(s, x, BODY_TOP, cw, 2.62)
        text(s, x + 0.32, BODY_TOP + 0.28, cw - 0.64, 0.40,
             [{"t": t, "sz": 17, "b": True, "c": col}])
        text(s, x + 0.32, BODY_TOP + 0.84, cw - 0.64, 0.70,
             [{"t": s1, "sz": 12.5, "c": MUTED, "ls": 1.32}])
        arrow_down(s, x + 0.34, BODY_TOP + 1.60, 0.26, 0.34, col)
        text(s, x + 0.32, BODY_TOP + 2.02, cw - 0.64, 0.44,
             [{"t": s2, "sz": 13.5, "b": True, "c": NAVY}])
    card(s, LM, 5.10, CW, 1.36, fill=SOFT)
    text(s, LM + 0.34, 5.28, CW - 0.68, 0.32,
         [{"t": "选领域时只问一句", "sz": 10.5, "b": True, "c": MUTED, "spc": 1.0}])
    text(s, LM + 0.34, 5.62, CW - 0.68, 0.66,
         [{"t": "在这个功能上，谁被逼得比我们更极端？被逼得更极端的地方，才更可能长出你没见过的结构。",
           "sz": 14, "b": True, "c": NAVY, "ls": 1.28}])
    footer(s, page(), total, "FOS 是什么")

    # ---------------- 19 章节页
    divider(prs, "03", "走完一个例子",
            "疫苗瓶监视器：从一句「已用标记」，走到一个只增不减的档位。",
            "完整例子", page(), total)

    # ---------------- 20 三个东西
    s = new_slide(prs)
    header(s, "第三部分 · 完整例子", "三个不相干的东西，在解同一件事",
           "没有联网、没有第三方在场，怎么证明一件事发生过，而且只发生过规定的次数？")
    trio = [("疫苗瓶", "小方块随累计受热慢慢变深", "冷链", BLUE),
            ("乏燃料", "安全不看还剩多少，看已经烧掉多少", "核工业", AMBER),
            ("空安瓿", "用完必须交回来，核对批号和数量", "医院", GREEN)]
    cw = (CW - 0.34 * 2) / 3
    for i, (t, b, tag, col) in enumerate(trio):
        x = LM + i * (cw + 0.34)
        card(s, x, BODY_TOP, cw, 2.55)
        rect(s, x, BODY_TOP, cw, 0.055, col)
        for k in range(3):
            box(s, x + 0.34 + k * 0.26, BODY_TOP + 0.32, 0.20, 0.20,
                fill=col if k <= i else "DCE5ED", line=None, radius=0.2)
        text(s, x + 0.32, BODY_TOP + 0.78, cw - 0.64, 0.40,
             [{"t": t, "sz": 18, "b": True, "c": NAVY}])
        text(s, x + 0.32, BODY_TOP + 1.30, cw - 0.64, 0.95,
             [{"t": b, "sz": 12.5, "c": MUTED, "ls": 1.34}])
        text(s, x + 0.32, BODY_TOP + 2.08, cw - 0.64, 0.30,
             [{"t": tag, "sz": 10, "b": True, "c": col, "spc": 1.0}])
    card(s, LM, 5.02, CW, 1.44, fill=SOFT)
    text(s, LM + 0.34, 5.20, CW - 0.68, 0.32,
         [{"t": "三个行业，一个共同的问题", "sz": 10.5, "b": True, "c": MUTED, "spc": 1.0}])
    text(s, LM + 0.34, 5.54, CW - 0.68, 0.76,
         [{"t": "它们都在做同一件事：留下证据，而且这份证据不能被抹掉、不能被改回去。",
           "sz": 14, "b": True, "c": NAVY, "ls": 1.28}])
    footer(s, page(), total, "完整例子")

    # ---------------- 21 起点
    s = new_slide(prs)
    header(s, "第三部分 · 完整例子", "起点：二元的已用标记，证明不了「只用过一次」",
           "完整例子只走一个：离线槽位。")
    card(s, LM, BODY_TOP, 6.10, 2.55)
    text(s, LM + 0.34, BODY_TOP + 0.22, 5.40, 0.30,
         [{"t": "现在的做法", "sz": 10.5, "b": True, "c": MUTED, "spc": 1.0}])
    b1 = box(s, LM + 0.36, BODY_TOP + 0.70, 2.30, 1.05, fill=SOFT2, line=LINE, radius=0.10)
    label(b1, [{"t": "未用", "sz": 15, "b": True, "c": NAVY}])
    arrow_right(s, LM + 2.86, BODY_TOP + 1.04, 0.60, 0.38, "B9C6D2")
    b2 = box(s, LM + 3.62, BODY_TOP + 0.70, 2.30, 1.05, fill="F6E2DC", line=CORAL,
             radius=0.10, lw=1.2)
    label(b2, [{"t": "已用", "sz": 15, "b": True, "c": CORAL}])
    text(s, LM + 0.34, BODY_TOP + 1.94, 5.40, 0.42,
         [{"t": "一旦有人把它改回「未用」，外观上看不出任何差别。",
           "sz": 12.5, "c": MUTED}])
    x2 = LM + 6.50
    card(s, x2, BODY_TOP, CW - 6.50, 2.55, fill="FDF3E4", line=AMBER, lw=1.4)
    text(s, x2 + 0.34, BODY_TOP + 0.22, CW - 7.18, 0.30,
         [{"t": "我们要找的功能", "sz": 10.5, "b": True, "c": "A5661A", "spc": 1.0}])
    text(s, x2 + 0.34, BODY_TOP + 0.70, CW - 7.18, 1.40,
         [{"t": "一份累计的、", "sz": 17, "b": True, "c": NAVY, "ls": 1.34},
          {"t": "不可逆的使用指示。", "sz": 17, "b": True, "c": NAVY, "ls": 1.34}])
    text(s, x2 + 0.34, BODY_TOP + 2.00, CW - 7.18, 0.40,
         [{"t": "注意两个词：累计，不可逆。", "sz": 11.5, "c": "8A5A1C"}])
    card(s, LM, 5.02, CW, 1.44, fill=SOFT)
    text(s, LM + 0.34, 5.20, CW - 0.68, 0.32,
         [{"t": "问题就在这里", "sz": 10.5, "b": True, "c": MUTED, "spc": 1.0}])
    text(s, LM + 0.34, 5.54, CW - 0.68, 0.76,
         [{"t": "二元的标记只记得「有没有」，记不住「几次」。所以要找的不是标记，是一份带历史的状态。",
           "sz": 14, "b": True, "c": NAVY, "ls": 1.28}])
    footer(s, page(), total, "完整例子")

    # ---------------- 22 阶梯
    s = new_slide(prs)
    header(s, "第三部分 · 完整例子", "沿阶梯走一遍：从「已用标记」到「疫苗瓶监视器」")
    data = [("01", "已用标记", "本领域的词", "离线槽位上的一个状态位", "9FBCD4"),
            ("02", "累计的、不可逆的指示", "上位词", "去掉「槽位」和「标记」，只留功能", "7FA8C6"),
            ("03", "疫苗冷链", "领先领域", "热稳定性被管得极死，容不得模糊", "4E93C4"),
            ("04", "疫苗瓶监视器", "具体物", "一张贴纸，随累计受热变深，分成四档", AMBER)]
    for i, (n, t, tag, note, col) in enumerate(data):
        yy = 1.98 + i * 1.16
        card(s, LM, yy, CW, 1.02)
        rect(s, LM, yy, 0.055, 1.02, col)
        text(s, LM + 0.30, yy + 0.30, 0.55, 0.40,
             [{"t": n, "sz": 15, "b": True, "c": col}])
        text(s, LM + 0.95, yy + 0.14, 4.10, 0.42,
             [{"t": t, "sz": 15.5, "b": True, "c": NAVY}])
        text(s, LM + 0.95, yy + 0.58, 4.10, 0.32,
             [{"t": tag, "sz": 10, "b": True, "c": col, "spc": 0.8}])
        text(s, LM + 5.30, yy + 0.26, CW - 5.70, 0.52,
             [{"t": note, "sz": 12.5, "c": MUTED, "ls": 1.26}])
    footer(s, page(), total, "完整例子")

    # ---------------- 23 阈值 vs 累计
    s = new_slide(prs)
    header(s, "第三部分 · 完整例子", "阈值式留不住历史，累计式留得住")
    cw = (CW - 0.60) / 2
    card(s, LM, BODY_TOP, cw, 3.05)
    text(s, LM + 0.34, BODY_TOP + 0.26, cw - 0.68, 0.40,
         [{"t": "阈值式", "sz": 17, "b": True, "c": CORAL}])
    b1 = box(s, LM + 0.36, BODY_TOP + 0.82, 1.75, 0.72, fill=SOFT2, line=LINE, radius=0.12)
    label(b1, [{"t": "未用", "sz": 13, "b": True, "c": NAVY}])
    b2 = box(s, LM + 2.30, BODY_TOP + 0.82, 1.75, 0.72, fill="F6E2DC", line=CORAL,
             radius=0.12, lw=1.2)
    label(b2, [{"t": "已用", "sz": 13, "b": True, "c": CORAL}])
    text(s, LM + 0.36, BODY_TOP + 1.74, cw - 0.72, 1.05,
         [{"t": "超过就报警，但改回「未用」之后，", "sz": 12.5, "c": MUTED, "ls": 1.34},
          {"t": "外观上看不出被动过。", "sz": 12.5, "c": MUTED, "ls": 1.34}])
    x2 = LM + cw + 0.60
    card(s, x2, BODY_TOP, cw, 3.05, fill="FDF3E4", line=AMBER, lw=1.4)
    text(s, x2 + 0.34, BODY_TOP + 0.26, cw - 0.68, 0.40,
         [{"t": "累计式", "sz": 17, "b": True, "c": "A5661A"}])
    tones = ["EFF7F3", "CFE6DA", "A8D2BE", "7CBDA1"]
    for k in range(4):
        bx = box(s, x2 + 0.36 + k * 1.12, BODY_TOP + 0.82, 1.00, 0.72,
                 fill=tones[k], line=None, radius=0.12)
        label(bx, [{"t": str(k), "sz": 15, "b": True, "c": "1F4A3A"}])
    text(s, x2 + 0.36, BODY_TOP + 1.74, cw - 0.72, 1.05,
         [{"t": "只增不减，每一档对应不同处置动作，", "sz": 12.5, "c": "8A5A1C", "ls": 1.34},
          {"t": "断网也能当场读出来。", "sz": 12.5, "c": "8A5A1C", "ls": 1.34}])
    card(s, LM, 5.52, CW, 0.95, fill=SOFT)
    text(s, LM + 0.34, 5.68, CW - 0.68, 0.62,
         [{"t": "多出来的不是一句新功能，是一个带着历史的状态。",
           "sz": 13.5, "b": True, "c": NAVY}])
    footer(s, page(), total, "完整例子")

    # ---------------- 24 迁回
    s = new_slide(prs)
    header(s, "第三部分 · 完整例子", "迁回来的不是贴纸，是只增不减的档位")
    cw = (CW - 1.00) / 2
    card(s, LM, BODY_TOP, cw, 2.75, fill=SOFT)
    text(s, LM + 0.34, BODY_TOP + 0.22, cw - 0.68, 0.30,
         [{"t": "迁之前", "sz": 10.5, "b": True, "c": MUTED, "spc": 1.0}])
    text(s, LM + 0.34, BODY_TOP + 0.62, cw - 0.68, 0.46,
         [{"t": "一个布尔值", "sz": 20, "b": True, "c": NAVY}])
    text(s, LM + 0.34, BODY_TOP + 1.30, cw - 0.68, 1.05,
         [{"t": "用过就置位，改回去就没了。", "sz": 12.5, "c": MUTED, "ls": 1.34}])
    x2 = LM + cw + 1.00
    card(s, x2, BODY_TOP, cw, 2.75, fill="FDF3E4", line=AMBER, lw=1.4)
    text(s, x2 + 0.34, BODY_TOP + 0.22, cw - 0.68, 0.30,
         [{"t": "迁回来", "sz": 10.5, "b": True, "c": "A5661A", "spc": 1.0}])
    text(s, x2 + 0.34, BODY_TOP + 0.62, cw - 0.68, 0.46,
         [{"t": "0–3 单调档位", "sz": 20, "b": True, "c": NAVY}])
    text(s, x2 + 0.34, BODY_TOP + 1.30, cw - 0.68, 1.05,
         [{"t": "0 档正常，3 档只能核销，不能再新增。", "sz": 12.5, "c": "8A5A1C", "ls": 1.34}])
    arrow_right(s, LM + cw + 0.30, BODY_TOP + 1.10, 0.40, 0.50, AMBER)
    card(s, LM, 4.98, CW, 1.48, fill=SOFT)
    rect(s, LM, 4.98, 0.055, 1.48, AMBER)
    text(s, LM + 0.36, 5.16, CW - 0.72, 0.32,
         [{"t": "多出来的效果", "sz": 10.5, "b": True, "c": AMBER, "spc": 1.0}])
    text(s, LM + 0.36, 5.52, CW - 0.72, 0.80,
         [{"t": "档位本身，就是现场可读的审计证据：不联网，也看得出来被透支过。",
           "sz": 14.5, "b": True, "c": NAVY, "ls": 1.28}])
    footer(s, page(), total, "完整例子")

    # ---------------- 25 三关
    s = new_slide(prs)
    header(s, "第三部分 · 完整例子", "迁回要过三关：结构、前提、可观测效果")
    for i, (t, b) in enumerate([
            ("结构", "能不能落到部件、字段或状态机？这里就是单调档位，0 到 3，只增不减。"),
            ("前提", "它在原领域靠什么成立，你这里还有没有？断网时也要留下看得见的历史。"),
            ("效果", "有没有一个原来测不到、现在能测的量？现场能读出被透支过。")]):
        yy = 1.98 + i * 1.10
        card(s, LM, yy, CW, 0.96)
        numdot(s, LM + 0.28, yy + 0.29, 0.38, str(i + 1), fill=GREEN, sz=11)
        text(s, LM + 0.86, yy + 0.28, 1.30, 0.40,
             [{"t": t, "sz": 15, "b": True, "c": NAVY}])
        text(s, LM + 2.28, yy + 0.24, CW - 2.60, 0.62,
             [{"t": b, "sz": 12.5, "c": MUTED, "ls": 1.26}])
    card(s, LM, 5.42, CW, 1.06, fill="FBF0ED", line=CORAL, lw=1.2)
    text(s, LM + 0.34, 5.58, 2.30, 0.34,
         [{"t": "顺便记一条边界", "sz": 10.5, "b": True, "c": CORAL, "spc": 0.8}])
    text(s, LM + 2.70, 5.56, CW - 3.04, 0.80,
         [{"t": "燃耗信用敢用「已经烧掉多少」，是因为背后有物理定律兜底；软件里没有这条定律，就不能直接搬。",
           "sz": 12.5, "c": "7A3428", "ls": 1.30}])
    footer(s, page(), total, "完整例子")

    # ---------------- 26 六步法
    s = new_slide(prs)
    header(s, "第三部分 · 完整例子", "六步法：散会就能做出一张迁移卡")
    six = [("列出名词", "5 个本领域关键词"),
           ("写上位词", "用三个行业实物验粒度"),
           ("选领域", "挑被迫做得更狠的"),
           ("下沉", "写清它对付哪条约束"),
           ("检索核实", "标清核实了没有"),
           ("写成迁移卡", "出处 ＋ 三道关")]
    cw = (CW - 0.22 * 2) / 3
    for i, (t, b) in enumerate(six):
        r, c = divmod(i, 3)
        x = LM + c * (cw + 0.22)
        yy = 1.98 + r * 1.58
        card(s, x, yy, cw, 1.42)
        rect(s, x, yy, 0.045, 1.42, BLUE)
        text(s, x + 0.26, yy + 0.16, 0.60, 0.32,
             [{"t": "0%d" % (i + 1), "sz": 11, "b": True, "c": BLUE, "spc": 0.6}])
        text(s, x + 0.26, yy + 0.54, cw - 0.50, 0.36,
             [{"t": t, "sz": 14.5, "b": True, "c": NAVY}])
        text(s, x + 0.26, yy + 0.98, cw - 0.50, 0.34,
             [{"t": b, "sz": 11, "c": MUTED}])
    card(s, LM, 5.24, CW, 1.24, fill=SOFT)
    text(s, LM + 0.34, 5.42, CW - 0.68, 0.32,
         [{"t": "顺序不要反", "sz": 10.5, "b": True, "c": AMBER, "spc": 1.0}])
    text(s, LM + 0.34, 5.76, CW - 0.68, 0.58,
         [{"t": "先泛化，再挑领域。反了，很容易挑到听起来高级、其实无关的行业。",
           "sz": 13.5, "b": True, "c": NAVY, "ls": 1.28}])
    footer(s, page(), total, "完整例子")

    # ---------------- 27 五个错误
    s = new_slide(prs)
    header(s, "第四部分 · 收束", "五个坑，最常见的是跳过泛化")
    errs = [("只搜领域词，不泛化名词", "回来的全是同构方案"),
            ("上位词大到无法下沉", "写不出可执行的结构"),
            ("停在「借鉴某某经验」", "这句话不可执行"),
            ("忽略原领域的前提", "没兜底的机制被直接搬来"),
            ("只记机会，不记威胁", "检索变成了自我表扬")]
    for i, (t, b) in enumerate(errs):
        yy = 1.96 + i * 0.90
        card(s, LM, yy, CW, 0.78)
        numdot(s, LM + 0.26, yy + 0.20, 0.38, str(i + 1), fill=CORAL, sz=11)
        text(s, LM + 0.84, yy + 0.22, 6.20, 0.36,
             [{"t": t, "sz": 14, "b": True, "c": NAVY}])
        text(s, LM + 7.30, yy + 0.24, CW - 7.60, 0.36,
             [{"t": b, "sz": 12, "c": MUTED}])
    text(s, LM, 6.44, CW, 0.34,
         [{"t": "第一条最普遍：功能抽完直接去搜，什么新东西都没有。", "sz": 11.5, "c": MUTED}])
    footer(s, page(), total, "收束")

    # ---------------- 28 威胁
    s = new_slide(prs)
    header(s, "第四部分 · 收束", "搜到同构的标准，要把它记成威胁",
           "3GPP 在线计费：网元先申请配额，用尽就终止会话。")
    cw = (CW - 0.60) / 2
    card(s, LM, BODY_TOP, cw, 2.55)
    text(s, LM + 0.34, BODY_TOP + 0.22, cw - 0.68, 0.32,
         [{"t": "功能上几乎一样", "sz": 10.5, "b": True, "c": BLUE, "spc": 0.8}])
    for k, t in enumerate(["预发额度", "计数", "用尽即停"]):
        sp = box(s, LM + 0.36, BODY_TOP + 0.70 + k * 0.56, 2.85, 0.44,
                 fill=SOFT2, line=None, radius=0.25)
        label(sp, [{"t": t, "sz": 12, "b": True, "c": NAVY}])
    x2 = LM + cw + 0.60
    card(s, x2, BODY_TOP, cw, 2.55, fill="FDF3E4", line=AMBER, lw=1.4)
    text(s, x2 + 0.34, BODY_TOP + 0.22, cw - 0.68, 0.32,
         [{"t": "差在有没有实时通道", "sz": 10.5, "b": True, "c": "A5661A", "spc": 0.8}])
    text(s, x2 + 0.34, BODY_TOP + 0.70, cw - 0.68, 1.55,
         [{"t": "对方依赖实时回传，", "sz": 13.5, "c": "8A5A1C", "ls": 1.36},
          {"t": "完全断连的时候，", "sz": 13.5, "c": "8A5A1C", "ls": 1.36},
          {"t": "这条标准对不上。", "sz": 13.5, "b": True, "c": "A5661A", "ls": 1.36}])
    card(s, LM, 5.02, CW, 1.44, fill=SOFT)
    text(s, LM + 0.34, 5.20, CW - 0.68, 0.32,
         [{"t": "正确做法", "sz": 10.5, "b": True, "c": MUTED, "spc": 1.0}])
    text(s, LM + 0.34, 5.54, CW - 0.68, 0.76,
         [{"t": "记成威胁：不要当成一条新灵感，也不要假装没看见。",
           "sz": 14, "b": True, "c": NAVY, "ls": 1.28}])
    footer(s, page(), total, "收束")

    # ---------------- 29 三句话
    s = new_slide(prs)
    header(s, "第四部分 · 收束", "带走三句话，就能开始")
    three = [("泛化的是名词，不是功能", "这决定你搜什么。", BLUE),
             ("下沉之后，再问它对付什么约束", "这决定值不值得搬。", AMBER),
             ("迁移卡要有出处，也要有三道关", "否则它只是一篇灵感笔记。", GREEN)]
    for i, (t, b, col) in enumerate(three):
        yy = 2.02 + i * 1.42
        card(s, LM, yy, CW, 1.26)
        rect(s, LM, yy, 0.055, 1.26, col)
        text(s, LM + 0.36, yy + 0.20, 0.62, 0.40,
             [{"t": "0%d" % (i + 1), "sz": 13, "b": True, "c": col, "spc": 0.6}])
        text(s, LM + 1.10, yy + 0.16, CW - 1.50, 0.44,
             [{"t": t, "sz": 17, "b": True, "c": NAVY}])
        text(s, LM + 1.10, yy + 0.70, CW - 1.50, 0.40,
             [{"t": b, "sz": 12, "c": MUTED}])
    footer(s, page(), total, "收束")

    # ---------------- 30 结尾
    s = new_slide(prs)
    rect(s, 0, 0, SW, SH, NAVY, name="end-bg")
    rect(s, 0, SH - 0.10, SW, 0.10, AMBER)
    text(s, LM + 0.02, 1.60, 8.4, 0.32,
         [{"t": "散会前，只留一个问题", "sz": 12, "c": "9FBBD4", "spc": 2.0}])
    text(s, LM, 2.04, 9.8, 1.65,
         [{"t": "你手上哪个词，", "sz": 40, "b": True, "c": WHITE, "ls": 1.16},
          {"t": "值得泛化一次？", "sz": 40, "b": True, "c": AMBER, "ls": 1.16}])
    rect(s, LM + 0.04, 4.00, 1.10, 0.05, AMBER)
    card(s, LM, 4.32, 8.60, 1.32, fill="17395A", line=None)
    text(s, LM + 0.40, 4.50, 7.80, 0.32,
         [{"t": "最小动作", "sz": 10.5, "b": True, "c": "9FBBD4", "spc": 1.2}])
    text(s, LM + 0.40, 4.86, 7.80, 0.60,
         [{"t": "写五个名词，各写一个上位词，再挑一个领域往下沉一层。",
           "sz": 15, "b": True, "c": WHITE, "ls": 1.26}])
    text(s, LM, 6.14, 9.0, 0.34,
         [{"t": "谢谢大家", "sz": 15, "b": True, "c": AMBER, "spc": 1.4}])
    page()

    return prs, p[0]


def main():
    sys.path.insert(0, str(OUT_DIR))
    from notes_text import NOTES
    prs, n = build()
    assert len(NOTES) == n, "notes 条数 %d 与页数 %d 不一致" % (len(NOTES), n)
    for i, slide in enumerate(prs.slides):
        slide.notes_slide.notes_text_frame.text = NOTES[i]
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    prs.save(str(PPTX_PATH))
    chars = sum(len(t) for t in NOTES)
    print("PAGES", n)
    print("NOTES_CHARS", chars, "EST_MIN", round(chars / 250.0, 1))
    print("SAVED", PPTX_PATH)


if __name__ == "__main__":
    main()
