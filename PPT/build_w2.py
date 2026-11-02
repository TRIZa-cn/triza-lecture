# -*- coding: utf-8 -*-
"""Workshop 2 slides builder: deck A (contradiction theory) + deck B (box/contradiction/On Contradiction)."""
import math
from pathlib import Path
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import nsdecls, qn
from pptx.oxml import parse_xml

D = Path(r"D:\doc\ZGCA\Lectures\TRIZ\PPT")

FONT = "微软雅黑"
INK = RGBColor(0x1A, 0x1D, 0x23)
INK2 = RGBColor(0x56, 0x5C, 0x66)
INK3 = RGBColor(0x8A, 0x90, 0x9A)
PAPER = RGBColor(0xFA, 0xF9, 0xF6)
RED = RGBColor(0xB0, 0x3A, 0x2B)
BLUE = RGBColor(0x2C, 0x5F, 0x8A)
AMBER = RGBColor(0xA9, 0x6E, 0x24)
GREEN = RGBColor(0x35, 0x69, 0x4A)
LINE = RGBColor(0xDD, 0xD8, 0xCE)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
Z0 = RGBColor(0xF2, 0xEF, 0xE8)
Z1 = RGBColor(0xFC, 0xFB, 0xF8)

SW, SH = 13.333, 7.5
ML, MR = 0.85, 0.85
CW = SW - ML - MR


def _set_ea(run, name):
    rPr = run._r.get_or_add_rPr()
    for tag in ("a:ea", "a:cs"):
        el = rPr.find(qn(tag))
        if el is None:
            el = parse_xml('<%s %s typeface="%s"/>' % (tag, nsdecls("a"), name))
            rPr.append(el)
        else:
            el.set("typeface", name)


def rich(p, text, size, color=INK, bold=False, accent=RED, spc=None):
    for i, seg in enumerate(text.split("**")):
        if seg == "":
            continue
        r = p.add_run()
        r.text = seg
        r.font.name = FONT
        r.font.size = Pt(size)
        if i % 2 == 1:
            r.font.bold = True
            r.font.color.rgb = accent
        else:
            r.font.bold = bold
            r.font.color.rgb = color
        if spc:
            r._r.get_or_add_rPr().set("spc", str(spc))
        _set_ea(r, FONT)
    return p


def textbox(slide, left, top, width, height, anchor=MSO_ANCHOR.TOP):
    tb = slide.shapes.add_textbox(Inches(left), Inches(top), Inches(width), Inches(height))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    for m in ("margin_left", "margin_right", "margin_top", "margin_bottom"):
        setattr(tf, m, 0)
    return tf


def para(tf, text, size, color=INK, bold=False, accent=RED, space_before=0,
         space_after=6, align=PP_ALIGN.LEFT, first=False, line=None):
    p = tf.paragraphs[0] if first else tf.add_paragraph()
    p.alignment = align
    p.space_before = Pt(space_before)
    p.space_after = Pt(space_after)
    if line:
        p.line_spacing = line
    rich(p, text, size, color, bold, accent)
    return p


def bullet(tf, text, size=19, color=INK, accent=RED, marker="▸", first=False,
           space_after=11, line=1.22, indent=0.30):
    p = tf.paragraphs[0] if first else tf.add_paragraph()
    p.space_after = Pt(space_after)
    p.line_spacing = line
    p.left_indent = Inches(indent)
    p.first_line_indent = Inches(-indent)
    r = p.add_run()
    r.text = marker + "  "
    r.font.name = FONT
    r.font.size = Pt(size)
    r.font.color.rgb = accent
    r.font.bold = True
    _set_ea(r, FONT)
    rich(p, text, size, color, accent=accent)
    return p


def rect(slide, left, top, width, height, fill, line_col=None,
         shape=MSO_SHAPE.RECTANGLE, lw=1.0):
    s = slide.shapes.add_shape(shape, Inches(left), Inches(top), Inches(width), Inches(height))
    if fill is None:
        s.fill.background()
    else:
        s.fill.solid()
        s.fill.fore_color.rgb = fill
    if line_col is None:
        s.line.fill.background()
    else:
        s.line.color.rgb = line_col
        s.line.width = Pt(lw)
    tf = s.text_frame
    tf.word_wrap = True
    tf.margin_left = Inches(0.18)
    tf.margin_right = Inches(0.18)
    tf.margin_top = Inches(0.10)
    tf.margin_bottom = Inches(0.10)
    return s


def page(prs, kicker, title, accent=RED, notes="", rule=True):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    bg = s.background.fill
    bg.solid()
    bg.fore_color.rgb = PAPER
    if kicker:
        tf = textbox(s, ML, 0.52, CW, 0.3)
        rich(tf.paragraphs[0], kicker, 11, INK3, bold=True, accent=accent, spc=260)
    tf = textbox(s, ML, 0.88, CW, 0.85)
    para(tf, title, 30, INK, bold=True, accent=accent, first=True, space_after=0, line=1.05)
    if rule:
        rect(s, ML, 1.86, 1.05, 0.045, accent)
    if notes:
        s.notes_slide.notes_text_frame.text = notes
    return s


def footer(s, idx, label, accent=RED):
    tf = textbox(s, ML, 6.98, CW - 1.2, 0.3)
    rich(tf.paragraphs[0], label, 10, INK3, accent=accent, spc=120)
    tf2 = textbox(s, SW - MR - 0.7, 6.98, 0.7, 0.3)
    para(tf2, "%02d" % idx, 10, INK3, first=True, align=PP_ALIGN.RIGHT)


def band(s, text, top, accent=RED, size=15, height=0.62):
    r = rect(s, ML + 0.12, top, CW - 0.12, height, None, line_col=None)
    r.text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
    r.text_frame.margin_left = Inches(0.10)
    rich(r.text_frame.paragraphs[0], text, size, INK, accent=accent)
    rect(s, ML, top, 0.055, height, accent)
    rect(s, ML + 0.1, top, CW - 0.1, 0.007, LINE)
    rect(s, ML + 0.1, top + height, CW - 0.1, 0.007, LINE)


def table(s, rows, left, top, width, col_w, size=15, head_size=13.5,
          head_fill=INK, head_col=WHITE, row_h=0.42, head_h=0.44, accent=RED,
          bold_first=True):
    nrow, ncol = len(rows), len(rows[0])
    shp = s.shapes.add_table(nrow, ncol, Inches(left), Inches(top),
                             Inches(width), Inches(head_h + row_h * (nrow - 1)))
    tbl = shp.table
    tbl.first_row = True
    tbl.horz_banding = False
    tbl.last_row = False
    tbl.first_col = False
    tbl.last_col = False
    for i, w in enumerate(col_w):
        tbl.columns[i].width = Inches(w)
    tbl.rows[0].height = Inches(head_h)
    for i in range(1, nrow):
        tbl.rows[i].height = Inches(row_h)
    for ri, row in enumerate(rows):
        for ci, val in enumerate(row):
            c = tbl.cell(ri, ci)
            c.margin_left = Inches(0.10)
            c.margin_right = Inches(0.08)
            c.margin_top = Inches(0.04)
            c.margin_bottom = Inches(0.04)
            c.vertical_anchor = MSO_ANCHOR.MIDDLE
            c.fill.solid()
            if ri == 0:
                c.fill.fore_color.rgb = head_fill
            else:
                c.fill.fore_color.rgb = Z0 if (ri - 1) % 2 == 0 else Z1
            tf = c.text_frame
            tf.word_wrap = True
            p = tf.paragraphs[0]
            p.line_spacing = 1.08
            if ri == 0:
                rich(p, val, head_size, head_col, bold=True, accent=accent)
            else:
                rich(p, val, size, INK, bold=(bold_first and ci == 0), accent=accent)
    return tbl


def cover(prs, kicker, title, subtitle, meta, accent=RED, notes=""):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    bg = s.background.fill
    bg.solid()
    bg.fore_color.rgb = PAPER
    rect(s, 0, 0, 0.16, SH, accent)
    tf = textbox(s, 1.15, 1.68, 10.4, 0.35)
    rich(tf.paragraphs[0], kicker, 12, INK3, bold=True, accent=accent, spc=300)
    tf = textbox(s, 1.15, 2.12, 10.9, 1.5)
    para(tf, title, 44, INK, bold=True, accent=accent, first=True, line=1.08, space_after=0)
    rect(s, 1.18, 3.98, 1.5, 0.05, accent)
    tf = textbox(s, 1.15, 4.28, 10.4, 1.0)
    para(tf, subtitle, 20, INK2, first=True, line=1.4, space_after=0, accent=accent)
    tf = textbox(s, 1.15, 6.10, 10.4, 0.8)
    para(tf, meta, 12, INK3, first=True, line=1.45, space_after=0)
    if notes:
        s.notes_slide.notes_text_frame.text = notes
    return s

# ============================================================ DECK A
def build_a():
    prs = Presentation()
    prs.slide_width = Inches(SW)
    prs.slide_height = Inches(SH)
    LB = "TRIZ 创新构思法研讨班  ·  第 2 次 Workshop  ·  A 讲"

    cover(prs, "TRIZ 创新构思法研讨班 · 第 2 次 WORKSHOP · A 讲",
          "矛盾的解剖学",
          "为什么“难”往往不是没想法，而是两个正确的要求在打架",
          "刘文印  ·  北京中关村学院\nA 讲 15 分钟：先讲理论，再用三个例子落地",
          RED,
          "开场。今天不讲工具清单，只讲一件事：TRIZ 怎么看待一个难题。"
          "十五分钟里我要做的，是把“矛盾”这个词从一句口号，变成可以写在纸上的结构。"
          "因为大部分人卡住，不是因为不聪明，而是因为从来没有把问题写成一句能被推翻的话。")

    s = page(prs, "00  ·  开场测试", "先看三个真实的困境", RED,
             "三个例子，分别来自能源、安全、智能体。读的时候请注意：每一个都不是“想不到办法”，"
             "而是两个都对的要求在互相牵制。这就是矛盾在日常生活里的样子——它不神秘，它很具体。")
    tf = textbox(s, ML, 2.18, CW, 3.3)
    items = [
        "**电池**：要跑一整天，要 15 分钟充满，还不能起火。",
        "**入侵检测模型**：要把攻击抓全，又几乎不许误报。",
        "**智能体**：要自动把事情做完，又要每一步都能被追责。",
    ]
    for i, t in enumerate(items):
        bullet(tf, t, 21, INK, RED, first=(i == 0), space_after=18)
    band(s, "这三件事都不是“想不出来”，而是**两个都正确的要求在打架**。", 5.55, RED, 16, 0.66)
    footer(s, 2, LB, RED)

    s = page(prs, "01  ·  定义", "什么叫一个矛盾", RED,
             "TRIZ 对矛盾的定义很窄：改善一个参数，会恶化另一个参数。注意是两个可测量的参数，"
             "不是一个模糊的不满。判据也很干脆——你能不能把它写成“提高 X，同时不牺牲 Y”。"
             "写得出来，就是矛盾；写不出来，那多半还只是抱怨。")
    tf = textbox(s, ML, 2.16, 6.15, 3.6)
    bullet(tf, "改善 **A**，却恶化了 **B**。A、B 都是**可测量**的参数。", 20, INK, RED, first=True, space_after=15)
    bullet(tf, "不是“资源不够”，也不是“技术不成熟”，是**两个正确目标互相拉扯**。", 20, INK, RED, space_after=15)
    bullet(tf, "矛盾不是一个抱怨，是**一个结构** —— 可以被写下来、被检验、被反驳。", 19, INK, RED, space_after=0)
    b = rect(s, 7.35, 2.14, 5.0, 3.3, None, LINE, MSO_SHAPE.ROUNDED_RECTANGLE)
    b.text_frame.margin_left = Inches(0.30)
    b.text_frame.margin_top = Inches(0.26)
    t = b.text_frame
    rich(t.paragraphs[0], "一句判据", 12, INK3, bold=True, accent=RED, spc=200)
    para(t, "你能把它写成“提高 X，同时不牺牲 Y”吗？", 18, INK, accent=RED, space_before=12, line=1.5)
    para(t, "写得出来 → 矛盾\n写不出来 → 还只是抱怨", 14, INK2, accent=RED, space_before=14, line=1.55)
    footer(s, 3, LB, RED)

    s = page(prs, "02  ·  分类", "矛盾的三种形态", RED,
             "TRIZ 主要管前两种。技术矛盾是“两个参数打架”，物理矛盾是“同一个东西必须同时是两样”。"
             "第三行要老实说：部门之间的目标冲突，TRIZ 没有专用工具，别硬套。")
    table(s, [
        ["形态", "怎么说", "例子", "TRIZ 工具"],
        ["技术矛盾", "改善 A，恶化 B", "散热要好就得厚，要薄就积热", "矛盾矩阵 + 40 条发明原理"],
        ["物理矛盾", "同一参数既要 A 又要非 A", "同一个判定阈值既要低又要高", "分离原理（时间/空间/条件/尺度）"],
        ["管理矛盾", "两个部门的目标冲突", "交付要快 vs 质量要稳", "不在 TRIZ 范围内，别硬套"],
    ], ML, 2.16, CW, [1.85, 3.05, 3.45, 3.28], size=15, row_h=0.88, head_h=0.46)
    band(s, "技术矛盾给你**方向**，物理矛盾给你**刀刃**。方案之所以平庸，多半是停在了第一行。", 5.66, RED, 16, 0.64)
    footer(s, 4, LB, RED)

    s = page(prs, "03  ·  关键一步", "从技术矛盾走到物理矛盾", RED,
             "这一步是整个方法的分水岭。技术矛盾是症状，物理矛盾是病灶。怎么做？连续问“为什么必须二选一”，"
             "问到不能再用“因为”回答为止。看这个机翼的例子：航程不够要多装油，油装进机翼，"
             "但机翼变厚就增加阻力。追问到底，落到一句话上——机翼那个位置，既要厚，又要薄。")
    tf = textbox(s, ML, 2.14, CW, 1.15)
    para(tf, "技术矛盾", 12, INK3, bold=True, accent=RED, first=True, space_after=5)
    para(tf, "航程不够 → 要多装燃油 → 油装进机翼 → 机翼变厚 → 阻力上升", 19, INK, accent=RED, space_after=0)
    tf = textbox(s, ML, 3.30, CW, 1.15)
    para(tf, "追问五次“为什么必须二选一”", 12, INK3, bold=True, accent=RED, first=True, space_after=5)
    para(tf, "→ **物理矛盾：机翼那个位置，既要“厚”（装油），又要“薄”（减阻）**", 19, INK, accent=RED, space_after=0)
    tf = textbox(s, ML, 4.46, CW, 1.15)
    para(tf, "分离", 12, INK3, bold=True, accent=RED, first=True, space_after=5)
    para(tf, "→ 让机翼内部成为油箱本身（整体油箱）：外形不变，容积变大", 19, INK, accent=RED, space_after=0)
    band(s, "技术矛盾是**症状**，物理矛盾是**病灶**。只治症状，方案就只能是折中。", 5.86, RED, 16, 0.62)
    footer(s, 5, LB, RED)

    s = page(prs, "04  ·  工具", "矛盾矩阵：四步，不许跳步", RED,
             "矩阵是个查表工具：横轴是你要改善的参数，纵轴是会被恶化的参数，格子里是几条发明原理的编号。"
             "用法四步。第三步最容易偷懒——很多人拿到编号就去搜“分割是什么意思”，然后套一个模板。"
             "第四步是纪律：只挑一两条做可验证的方案，其余存档。查不到就写清楚为什么查不到，绝不硬凑。")
    tf = textbox(s, ML, 2.12, 7.35, 4.0)
    steps = [
        ("① 写成参数对", "把问题翻译成 39（现代 48）个通用工程参数里的两个"),
        ("② 查格子", "矩阵交点给出 3–4 条发明原理编号"),
        ("③ 追问语义", "每条原理在我的场景里，到底对应什么具体动作"),
        ("④ 收敛", "选 1–2 条做可验证方案，其余归档；查不到就说查不到"),
    ]
    for i, (h, d) in enumerate(steps):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.space_after = Pt(14)
        p.line_spacing = 1.2
        rich(p, h + "　", 19, INK, bold=True, accent=RED)
        rich(p, d, 16, INK2, accent=RED)
    b = rect(s, 8.55, 2.12, 3.8, 3.3, None, LINE, MSO_SHAPE.ROUNDED_RECTANGLE)
    b.text_frame.margin_left = Inches(0.26)
    b.text_frame.margin_top = Inches(0.24)
    t = b.text_frame
    rich(t.paragraphs[0], "一个必须说清的边界", 11.5, INK3, bold=True, accent=RED, spc=150)
    para(t, "矩阵给的是**方向**，不是答案。", 16, INK, accent=RED, space_before=10, line=1.45)
    para(t, "它告诉你该往哪类改动上试，不给公式、参数、阈值。", 13.5, INK2, accent=RED, space_before=10, line=1.5)
    footer(s, 6, LB, RED)

    s = page(prs, "05  ·  例子一 · 工程", "笔记本散热：矛盾没有被折中，是被搬走了", RED,
             "第一个例子。笔记本要散热好，风扇要大、风道要厚；但用户要薄。这是典型的技术矛盾。"
             "查矩阵会拿到分割、抽取、动态化、中介物。真正落地的方案是热管：不在 CPU 原地散热，"
             "把热搬到机身边缘的鳍片上再吹走。注意——矛盾并没有被折中，它是被换了位置。"
             "这就是空间分离的一次具体使用。")
    tf = textbox(s, ML, 2.14, 5.55, 3.5)
    bullet(tf, "**技术矛盾**：散热能力 ↑ ↔ 机身厚度 ↓", 19, INK, RED, first=True, space_after=14)
    bullet(tf, "查矩阵 → 分割 / 抽取 / 动态化 / 中介物", 19, INK, RED, space_after=14)
    bullet(tf, "落地：**热管**把热从 CPU “搬”到边缘鳍片再吹走", 19, INK, RED, space_after=14)
    bullet(tf, "本质：不在原地散热，而是**换一个位置**散热", 19, INK, RED, space_after=0)
    b = rect(s, 6.95, 2.14, 5.4, 3.3, None, LINE, MSO_SHAPE.ROUNDED_RECTANGLE)
    b.text_frame.margin_left = Inches(0.28)
    b.text_frame.margin_top = Inches(0.26)
    t = b.text_frame
    rich(t.paragraphs[0], "判据 —— 这是解，还是折中？", 11.5, INK3, bold=True, accent=RED, spc=140)
    para(t, "折中：风扇小一点、厚度加一点，双方各让一步。矛盾还在，只是被压住了。", 14.5, INK, accent=RED, space_before=10, line=1.5)
    para(t, "分离：热不再必须在 CPU 附近消散。**“既要又要”被拆到两个位置**，各自都不再勉强。", 14.5, INK, accent=RED, space_before=10, line=1.5)
    footer(s, 7, LB, RED)

    s = page(prs, "06  ·  解法机器", "分离原理：四把刀", RED,
             "分离是 TRIZ 对付物理矛盾的总办法：让“既要又要”在不同条件下分别成立，而不是同时成立。"
             "四个维度——时间、空间、条件、尺度。雨伞只在雨天有用，是时间；刀刃硬刀背韧，是空间；"
             "安全气囊只在碰撞瞬间出现，是条件；羽绒被宏观蓬松、微观锁住空气，是尺度。"
             "记住这四把刀，比背四十条原理更管用。")
    table(s, [
        ["维度", "怎么分", "一句话例子"],
        ["时间分离", "需要时才出现", "雨伞只在雨天有用；气囊只在碰撞时充气"],
        ["空间分离", "不同位置不同性质", "刀刃要硬，刀背要韧"],
        ["条件分离", "满足条件才切换", "同一个阈值，按业务线、按时段分别设"],
        ["尺度分离", "宏观与微观各司其职", "羽绒被：羽毛宏观蓬松，绒毛微观锁气"],
    ], ML, 2.16, CW, [2.05, 3.5, 6.06], size=16, row_h=0.74, head_h=0.46)
    band(s, "分离的本质：**不是让两个要求同时满足，而是让它们在不同条件下各自成立**。", 6.06, RED, 15.5, 0.60)
    footer(s, 8, LB, RED)

    s = page(prs, "07  ·  例子二 · AI", "召回与误报：同一把尺子，分条件用", RED,
             "第二个例子来自我们自己的论文。入侵检测要提高召回，误报就上升，这是技术矛盾。"
             "写成物理矛盾就是：同一个判定阈值，既要低到抓得住，又要高到不冤枉。"
             "解法是条件分离——按时段、按业务线、按置信度分层，各用各的阈值。"
             "同一批 2040 个决策上，safety-first、calibrated V1、balanced V2 就是三种不同的分离策略。"
             "关键细节：这三个策略的取舍写在实验之前，不是事后挑一个好看的。")
    tf = textbox(s, ML, 2.14, 6.05, 3.4)
    bullet(tf, "**技术矛盾**：攻击召回 ↑ ↔ 误报率 ↑", 19, INK, RED, first=True, space_after=14)
    bullet(tf, "**物理矛盾**：同一个阈值既要低（抓全）又要高（不冤枉）", 17.5, INK, RED, space_after=14)
    bullet(tf, "**条件分离**：按时段 / 业务线 / 置信度分层设阈值", 18.5, INK, RED, space_after=14)
    bullet(tf, "三种策略 = 三种分离方式，取舍事先声明", 19, INK, RED, space_after=0)
    b = rect(s, 7.25, 2.14, 5.1, 3.28, None, LINE, MSO_SHAPE.ROUNDED_RECTANGLE)
    b.text_frame.margin_left = Inches(0.26)
    b.text_frame.margin_top = Inches(0.24)
    t = b.text_frame
    rich(t.paragraphs[0], "真实数据（TRIZa 论文）", 11.5, INK3, bold=True, accent=RED, spc=140)
    para(t, "同一批 **2040** 个资格判定决策，三种控制策略：", 13.5, INK2, accent=RED, space_before=9, line=1.45)
    para(t, "safety-first　→ 几乎全部拒绝，含合格者", 13, INK, accent=RED, space_before=8)
    para(t, "calibrated V1　→ 接受多，召回高", 13, INK, accent=RED, space_before=6)
    para(t, "balanced V2　→ 居中", 13, INK, accent=RED, space_before=6)
    para(t, "三者并非独立重复，不能当排名读", 11.5, INK3, accent=RED, space_before=10, line=1.4)
    footer(s, 9, LB, RED)

    s = page(prs, "08  ·  核心答案", "矛盾说到多深，创新就走到多高", RED,
             "这一页是本讲的答案。阿奇舒勒按发明程度把专利分五级：一级不需要解决矛盾，二级解决技术矛盾，"
             "三级解决物理矛盾，四级靠 ARIZ 让矛盾一次性消失，五级是全新发现。"
             "比例是 32、45、18、4、不到 1。最要紧的是“矛盾处理”那一列——"
             "它说明创新等级不是由投入多少决定的，是由你处理矛盾的能力决定的。")
    table(s, [
        ["等级", "特征", "矛盾处理", "占比"],
        ["1 标准化", "本领域常规改进，不改结构", "不考虑矛盾", "32%"],
        ["2 改良化", "本行业已知方法 + 微调", "技术矛盾", "45%"],
        ["3 模式内创新", "引入行业外方法，大幅改善", "物理矛盾", "18%"],
        ["4 模式外创新", "已公开原理首次落地，性能量级跃升", "ARIZ，矛盾一次性消失", "4%"],
        ["5 新发现", "新科学、新产业", "超越当代知识", "<1%"],
    ], ML, 2.14, CW, [2.35, 5.15, 3.05, 1.06], size=15, row_h=0.66, head_h=0.44)
    band(s, "创新等级不取决于投入多少，取决于**你把矛盾处理到哪一层**。停在折中，就永远在一级和二级。", 6.06, RED, 15, 0.58)
    footer(s, 10, LB, RED)

    s = page(prs, "09  ·  警告", "折中不是解，是欠下的账", RED,
             "必须给一个警告：折中不是解决矛盾。双方各让一步，那个“既要又要”还在，只是被压住了，"
             "条件一变就会回来。TRIZ 的立场很硬：折中是没解出矛盾时交的税。"
             "怎么判断手里的方案是解还是折中？问一句——落地之后，那个“既要又要”还在不在。还在，就是折中。")
    tf = textbox(s, ML, 2.16, 6.2, 3.3)
    bullet(tf, "折中 = 双方各让一步 = **矛盾还在，只是被压住**", 20, INK, RED, first=True, space_after=17)
    bullet(tf, "条件一变，它会原样回来，而且常常更贵", 20, INK, RED, space_after=17)
    bullet(tf, "TRIZ 的立场：折中是**没解出矛盾时交的税**", 20, INK, RED, space_after=0)
    b = rect(s, 7.45, 2.14, 4.9, 3.2, None, LINE, MSO_SHAPE.ROUNDED_RECTANGLE)
    b.text_frame.margin_left = Inches(0.30)
    b.text_frame.margin_top = Inches(0.30)
    t = b.text_frame
    rich(t.paragraphs[0], "一个自检问题", 11.5, INK3, bold=True, accent=RED, spc=160)
    para(t, "方案落地以后，那个“又要…又要…”还在不在？", 18, INK, accent=RED, space_before=12, line=1.55)
    para(t, "在 → 折中\n不在 → 才叫解", 15, INK2, accent=RED, space_before=14, line=1.5)
    footer(s, 11, LB, RED)

    s = page(prs, "10  ·  小结", "三句话", RED,
             "收一下。第一，矛盾不是障碍，它是问题被你理解到位的证据。第二，技术矛盾给方向，物理矛盾给刀刃，"
             "分离给解法——这三件套缺一件，方案就会退化成折中。第三，你能把矛盾写多准，创新就能走多远。"
             "下一讲我们把它和另外两种“约束观”放在一起看，包括《矛盾论》。")
    tf = textbox(s, ML, 2.2, CW, 3.3)
    pts = [
        ("01", "矛盾不是障碍，是问题**被理解到位**的证据。"),
        ("02", "技术矛盾给方向，物理矛盾给刀刃，**分离**给解法。"),
        ("03", "你能把矛盾写多准，你的创新就能走多远。"),
    ]
    for i, (n, t) in enumerate(pts):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.space_after = Pt(26)
        p.line_spacing = 1.3
        rich(p, n + "　", 15, RED, bold=True, accent=RED, spc=100)
        rich(p, t, 21, INK, accent=RED)
    band(s, "下一讲 B：《盒子、矛盾与〈矛盾论〉》—— 约束到底是**提取**矛盾，还是**注入**矛盾？", 5.72, BLUE, 15.5, 0.66)
    footer(s, 12, LB, RED)

    out = D / "TRIZ_W2A_矛盾解剖_15min.pptx"
    prs.save(str(out))
    return out

# ============================================================ DECK B
def build_b():
    prs = Presentation()
    prs.slide_width = Inches(SW)
    prs.slide_height = Inches(SH)
    LB = "TRIZ 创新构思法研讨班  ·  第 2 次 Workshop  ·  B 讲"

    cover(prs, "TRIZ 创新构思法研讨班 · 第 2 次 WORKSHOP · B 讲",
          "盒子、矛盾与《矛盾论》",
          "三种“约束观”的一次对话，以及一次诚实的自我评估",
          "刘文印  ·  北京中关村学院\nB 讲 15 分钟：三方对照 → 两个例子 → 一个评估",
          BLUE,
          "这一讲做三件事。第一，把《盒子之内》那本书的核心方法，用 TRIZ 的语言重新说一遍。"
          "第二，把它和《矛盾论》放在一起对照，说清哪里真通、哪里只是像。"
          "第三，对我们自己那份研究报告做一次独立评估——它到底有没有创意。三件事我都尽量说得不客气。")

    s = page(prs, "00  ·  一句话摆开", "三个命题，三件不同的事", BLUE,
             "先一句话把三者摆开。TRIZ 的立场是：矛盾本来就藏在系统里，任务是把它们提取出来。"
             "那本书的立场正相反：如果问题里没有矛盾，你就自己造一个。"
             "《矛盾论》关心的不是怎么解，是判断哪个矛盾是主要的。所以三者不是同一件事。"
             "但它们共用一台机器。")
    cols = [
        ("TRIZ", "从问题里**提取**矛盾", "矛盾已经在系统里，任务是找到它、说准它", RED),
        ("《盒子之内》", "往问题里**注入**矛盾", "没有矛盾就造一个：排除 + 引导，人为设约束", AMBER),
        ("《矛盾论》", "判断**哪个**矛盾是主要的", "抓主要矛盾，看矛盾的主要方面，具体条件具体分析", BLUE),
    ]
    for i, (h, t1, t2, col) in enumerate(cols):
        x = ML + i * (CW / 3 + 0.02)
        w = CW / 3 - 0.18
        b = rect(s, x, 2.2, w, 2.62, None, LINE, MSO_SHAPE.ROUNDED_RECTANGLE)
        b.text_frame.margin_left = Inches(0.26)
        b.text_frame.margin_top = Inches(0.26)
        tf = b.text_frame
        rich(tf.paragraphs[0], h, 12, col, bold=True, spc=140)
        para(tf, t1, 19, INK, accent=col, space_before=12, line=1.35)
        para(tf, t2, 12.5, INK2, accent=col, space_before=12, line=1.55)
    band(s, "三者不是同一个理论，但**共用一台解法机器**。这句话是今天全部内容的地基。", 5.28, BLUE, 16, 0.64)
    footer(s, 2, LB, BLUE)

    s = page(prs, "01  ·  书在说什么", "“跳出盒子”是个坏建议", BLUE,
             "这本书的主张很不客气：跳出盒子这个建议搞反了。它举了三组证据。"
             "最著名的是门捷列夫——课本说周期表是梦里排出来的，科学史学家戈尔丁在权威传记里直接说不是。"
             "真实原因是出版合同：他已经写完第一卷，只覆盖了 63 种元素里的 8 种，"
             "剩下的 55 种必须塞进第二卷有限的篇幅。页数和死线逼他放弃了逐个描述的老路。"
             "注意这里：创意不是来自无限自由，是来自一份合同。")
    tf = textbox(s, ML, 2.14, 6.35, 3.6)
    bullet(tf, "课本版本：门捷列夫在梦里排出了周期表", 18.5, INK, AMBER, first=True, space_after=13)
    bullet(tf, "史实版本：戈尔丁在权威传记里说**“他显然不是在梦中发现的”**", 17, INK, AMBER, space_after=13)
    bullet(tf, "真实约束：两卷本合同 + 第一卷只写了 8 种元素 + 剩余 55 种要塞进有限篇幅", 16.5, INK, AMBER, space_after=13)
    bullet(tf, "结论：**创意不是来自自由，是来自一份合同和一条死线**", 18.5, INK, AMBER, space_after=0)
    b = rect(s, 7.55, 2.14, 4.8, 3.3, None, LINE, MSO_SHAPE.ROUNDED_RECTANGLE)
    b.text_frame.margin_left = Inches(0.28)
    b.text_frame.margin_top = Inches(0.26)
    t = b.text_frame
    rich(t.paragraphs[0], "同方向的其他证据", 11.5, INK3, bold=True, accent=AMBER, spc=140)
    para(t, "玩具设计实验：自由挑零件的组 vs 五个零件必须全用 —— 约束组更有创意", 13, INK, accent=AMBER, space_before=10, line=1.55)
    para(t, "斯托克斯的美术课：命题作画的作品，比自由创作更有趣", 13, INK, accent=AMBER, space_before=8, line=1.55)
    footer(s, 3, LB, BLUE)

    s = page(prs, "02  ·  方法", "配对约束 = 排除 + 引导，缺一不可", BLUE,
             "书里的方法是斯托克斯的“配对约束”，两个动作。第一个是排除：主动禁掉你最擅长、最顺手的那条路，"
             "切断最省力路径。第二个是引导：同时给一条具体、可执行的新方向。为什么必须成对？"
             "只排除，人会无所适从；只引导，旧习惯会不断把你拉回老路。这一点非常关键，"
             "因为 TRIZ 里有一模一样的一句话。")
    b1 = rect(s, ML, 2.16, 5.5, 1.45, None, LINE, MSO_SHAPE.ROUNDED_RECTANGLE)
    b1.text_frame.margin_left = Inches(0.28)
    b1.text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
    t = b1.text_frame
    rich(t.paragraphs[0], "排除", 14, AMBER, bold=True, spc=200)
    para(t, "禁止你有最擅长、最顺手的那条路", 17.5, INK, accent=AMBER, space_before=6)
    b2 = rect(s, ML, 3.84, 5.5, 1.45, None, LINE, MSO_SHAPE.ROUNDED_RECTANGLE)
    b2.text_frame.margin_left = Inches(0.28)
    b2.text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
    t = b2.text_frame
    rich(t.paragraphs[0], "引导", 14, AMBER, bold=True, spc=200)
    para(t, "同时给一条具体、可执行的新方向", 17.5, INK, accent=AMBER, space_before=6)
    tf = textbox(s, 6.75, 2.16, 5.6, 3.2)
    para(tf, "为什么必须成对", 11.5, INK3, bold=True, accent=AMBER, first=True, space_after=9)
    para(tf, "只排除 → 无所适从，人还是会偷偷绕回去", 16, INK, accent=AMBER, space_after=9, line=1.5)
    para(tf, "只引导 → 旧习惯会不断把你拉回老路", 16, INK, accent=AMBER, space_after=16, line=1.5)
    para(tf, "**排除 + 引导 = 一个被迫成立的新问题**", 17, INK, accent=AMBER, space_after=0, line=1.5)
    band(s, "这不是心理学花招。它和 TRIZ 的“物理矛盾 + 分离”是**同一个结构的两种叫法**。", 5.64, AMBER, 15.5, 0.64)
    footer(s, 4, LB, BLUE)
    s = page(prs, "03  ·  四个维度", "往哪个方向设约束", BLUE,
             "斯托克斯还给了四个维度，说的都是“往哪设”。目标约束改“做什么”：莫奈不还原客观世界，"
             "改成画我这一刻看到的光。资源约束改“用什么”：布伦屋只给小预算，逼导演在叙事张力上下功夫。"
             "方法约束改“怎么做”：村上春树写《且听风吟》先用英语写，因为词汇不够只能写短句，再译回日语。"
             "主体约束改“核心对象”：戴森给风扇下了一条硬规矩，不许有扇叶。")
    cells = [
        ("目标约束", "改“做什么”", "莫奈：不画物体，画光线；拒绝用黑色画阴影", RED),
        ("资源约束", "改“用什么”", "布伦屋：只给小预算，逼导演在叙事张力上使劲", AMBER),
        ("方法约束", "改“怎么做”", "村上春树：弃母语，先用英语写《且听风吟》", BLUE),
        ("主体约束", "改“核心对象”", "戴森：风扇不许有扇叶；苏斯博士：只准用 50 个词", GREEN),
    ]
    for i, (h, sub, ex, col) in enumerate(cells):
        col_i, row_i = i % 2, i // 2
        x = ML + col_i * (CW / 2 + 0.03)
        y = 2.14 + row_i * 1.75
        w = CW / 2 - 0.18
        b = rect(s, x, y, w, 1.52, None, LINE, MSO_SHAPE.ROUNDED_RECTANGLE)
        b.text_frame.margin_left = Inches(0.26)
        b.text_frame.margin_top = Inches(0.18)
        tf = b.text_frame
        pb = tf.paragraphs[0]
        rich(pb, h, 18, INK, bold=True, accent=col)
        rich(pb, "　" + sub, 12, INK3, accent=col, spc=80)
        para(tf, ex, 13, INK2, accent=col, space_before=9, line=1.5)
    band(s, "四个维度回答“往哪设”，不回答“设多硬”。**硬度**才是决定创新等级的那个变量。", 5.68, BLUE, 15.5, 0.62)
    footer(s, 5, LB, BLUE)

    s = page(prs, "04  ·  核心论点", "配对约束，就是人为植入一个物理矛盾", BLUE,
             "这一页是今天最重要的一页。TRIZ 假设矛盾本来就在系统里，你的任务是提取。"
             "这本书的主张正相反：如果没有矛盾，你就造一个。看苏斯博士那个赌约——只用 50 个词写一本书。"
             "用 TRIZ 的语言重述：方法约束排除了长句和生僻词，目标约束又要求故事仍然完整有趣。"
             "两条合起来就是一个物理矛盾：同一本书既要极简，又要完整。解法是把复杂度转移到韵律和重复结构上。"
             "这就是分离。所以两者不是相似，是同一个机制的两种用法。")
    tf = textbox(s, ML, 2.2, 6.0, 1.4)
    para(tf, "TRIZ：矛盾本来就在 → **提取** → 分离求解", 17.5, INK, accent=RED, first=True, space_after=10, line=1.4)
    para(tf, "盒子：矛盾不在 → **注入** → 分离求解", 17.5, INK, accent=AMBER, space_after=0, line=1.4)
    b = rect(s, 7.15, 2.14, 5.2, 3.42, None, LINE, MSO_SHAPE.ROUNDED_RECTANGLE)
    b.text_frame.margin_left = Inches(0.26)
    b.text_frame.margin_top = Inches(0.24)
    t = b.text_frame
    rich(t.paragraphs[0], "翻译练习 · 苏斯博士的 50 个词", 11.5, INK3, bold=True, accent=AMBER, spc=120)
    para(t, "排除：长句、生僻词、复杂叙事", 13, INK, accent=AMBER, space_before=10, line=1.5)
    para(t, "引导：仍然要讲一个完整、好笑的故事", 13, INK, accent=AMBER, space_before=7, line=1.5)
    para(t, "→ 物理矛盾：同一文本既要极简又要完整", 13.5, INK, accent=AMBER, space_before=10, line=1.5)
    para(t, "→ 分离：把复杂度转移到韵律与重复结构", 13.5, INK, accent=AMBER, space_before=7, line=1.5)
    band(s, "结论：**相通，而且比“相似”更强** —— 它们是同一台机器的两种输入方式。", 5.72, AMBER, 16, 0.64)
    footer(s, 6, LB, BLUE)

    s = page(prs, "05  ·  为什么要“造”", "没有矛盾的问题，注定只能做出一级创新", BLUE,
             "把上一讲那张发明等级表接过来，就能解释这本书为什么有效。"
             "32% 的问题本身不含矛盾，改进只能停留在一级。一旦你人为制造出一个精确的矛盾，"
             "问题就被迫升到二三级，因为不分离就没有答案。所以约束的作用不是让你更自由，"
             "是把你的问题抬高一档。这是这本书对 TRIZ 最有价值的一句话，反过来也成立。")
    table(s, [
        ["如果没有约束", "如果注入配对约束"],
        ["问题本来不含矛盾 → 只能做常规改进（一级）", "人为制造出精确矛盾 → 被迫进入二三级"],
        ["方案空间太大 → 在熟悉的路径上反复微调", "大量常规方案被排除 → 只剩陌生区域"],
        ["产出容易，但“似曾相识”", "产出困难，但一旦出来就是新的"],
    ], ML, 2.16, CW, [5.75, 5.86], size=15, row_h=0.88, head_h=0.46, accent=BLUE)
    band(s, "约束的价值不是让你更自由，是**把你的问题抬高一档**。", 5.74, BLUE, 17, 0.64)
    footer(s, 7, LB, BLUE)
    s = page(prs, "06  ·  三方对照", "《矛盾论》与 TRIZ：哪里真通，哪里只是像", BLUE,
             "现在说《矛盾论》。它和 TRIZ 确实有方法上的联系，但必须一条一条地说。"
             "主要矛盾，在 TRIZ 里对应的是因果链分析找到的关键矛盾——但关注点不同："
             "矛盾论讲抓哪个，TRIZ 讲解怎么解。具体条件具体分析，和条件分离高度相通。"
             "矛盾的转化，对应分离之后矛盾消失，但 TRIZ 多要求一条：可验证。"
             "不同质的矛盾用不同质的方法，对应矛盾类型与工具的匹配。"
             "最值得注意的是第二行——矛盾的主要方面，TRIZ 里没有直接对应物。")
    table(s, [
        ["《矛盾论》", "TRIZ 里的对应", "差别在哪"],
        ["主要矛盾", "因果链分析找到的关键矛盾", "矛盾论答“抓哪个”，TRIZ 答“怎么解”"],
        ["矛盾的主要方面", "**没有直接对应物**", "TRIZ 的矛盾是对称的；谁主导常要人工判断"],
        ["具体条件具体分析", "条件分离", "高度相通，几乎是同一条规矩"],
        ["矛盾的转化", "分离后矛盾消失（四级）", "相通，但 TRIZ 多要一条：可验证"],
        ["不同质的矛盾用不同质的方法", "矛盾类型 → 工具匹配", "相通"],
    ], ML, 2.14, CW, [3.45, 4.1, 4.06], size=14, row_h=0.70, head_h=0.44, accent=BLUE)
    band(s, "第二行是一个**开放问题**，不是结论：TRIZ 至今没有“主要方面”这个变量。", 6.06, BLUE, 14.5, 0.58)
    footer(s, 8, LB, BLUE)

    s = page(prs, "07  ·  边界", "把话说死：三者不是同一个理论", BLUE,
             "这一页是自我约束。我们的研究报告里写得很清楚：三者可以配合，但不是同一理论；"
             "本轮没有核得 TRIZ 直接源自《矛盾论》的原始证据，因此不以相似性推断传承。"
             "还有一句更重要——哲学启发不能替代协议或算法证明。"
             "为什么我要强调这个？因为我们自己最容易犯的错，就是把“看起来像”写成“其实同源”。")
    tf = textbox(s, ML, 2.2, CW, 3.4)
    bullet(tf, "三者可以配合，但**不是同一理论**。", 21, INK, BLUE, first=True, space_after=18)
    bullet(tf, "**未核得** TRIZ 直接源自《矛盾论》的原始证据 —— 不以相似性推断传承。", 19.5, INK, BLUE, space_after=18)
    bullet(tf, "**哲学启发不能替代协议或算法证明。**", 21, INK, BLUE, space_after=18)
    bullet(tf, "“主要”只用于安排研究与优化的**顺序**，不降低任何硬条件。", 19.5, INK, BLUE, space_after=0)
    band(s, "这三条不是客气话，是我们给自己设的**约束** —— 顺便也是配对约束的一个实例。", 5.86, BLUE, 15, 0.62)
    footer(s, 9, LB, BLUE)

    s = page(prs, "08  ·  例子", "把三类约束翻译成 TRIZ 语言", BLUE,
             "做一个翻译练习。苏斯博士的 50 个词是方法约束，落到物理矛盾——既要极简又要完整，用分割和局部质量。"
             "戴森的无叶风扇是主体约束，落到物理矛盾——要有风又不能有叶片，用抽取和中介物："
             "把产生气流的部件藏进底座，只留下一个环。村上春树弃母语，也是方法约束。"
             "你看，同一张表可以把文学、制造、写作装进同一个框架。")
    table(s, [
        ["书里的例子", "约束类型", "翻译成的矛盾", "TRIZ 工具"],
        ["苏斯博士 · 50 个词", "方法约束", "同一文本既要极简又要完整", "分割 / 局部质量"],
        ["戴森 · 无叶风扇", "主体约束", "要有风，又不能有叶片", "抽取 / 中介物"],
        ["村上春树 · 弃母语", "方法约束", "要表达丰富，又要语言极简", "另一个维度的分离"],
        ["布伦屋 · 只给小预算", "资源约束", "要抓人，又没钱堆明星和大场面", "抽取：只保留叙事张力"],
    ], ML, 2.16, CW, [3.15, 2.2, 3.9, 2.38], size=14, row_h=0.84, head_h=0.44, accent=BLUE)
    band(s, "同一张表，能把文学、制造、写作装进**同一个框架** —— 这才是它可迁移的原因。", 6.06, BLUE, 14.5, 0.58)
    footer(s, 10, LB, BLUE)
    s = page(prs, "09  ·  例子 · 我们自己的", "补救动作会毁掉它自己依赖的证据", BLUE,
             "现在用这套语言说我们自己的问题。智能体在授权流程里出错后要补救：补签、重新取证、换模板。"
             "但补救本身会更新版本，让已有证据失效；同时还可能卡在不可控的等待上。"
             "用 TRIZ 说：技术矛盾是安全性和可用性打架；物理矛盾是同一份凭据既要已批准又要已轮换。"
             "分离的答案是按时间分离——批准绑定版本号，轮换即失效。这不是 bug，是矛盾的必然结果。")
    tf = textbox(s, ML, 2.16, 6.15, 3.5)
    bullet(tf, "现象：补救会更新版本 → **已有证据失效**；且可能卡在不可控等待", 16.5, INK, BLUE, first=True, space_after=14)
    bullet(tf, "技术矛盾：**增强安全**（多签多验）↔ **可用性**（预算与等待）", 16.5, INK, BLUE, space_after=14)
    bullet(tf, "物理矛盾：同一凭据既要**已批准**（可执行）又要**已轮换**（安全）", 16.5, INK, BLUE, space_after=14)
    bullet(tf, "分离：按**时间**分离 —— 批准绑定版本，轮换即失效", 16.5, INK, BLUE, space_after=0)
    b = rect(s, 7.3, 2.14, 5.05, 3.3, None, LINE, MSO_SHAPE.ROUNDED_RECTANGLE)
    b.text_frame.margin_left = Inches(0.26)
    b.text_frame.margin_top = Inches(0.24)
    t = b.text_frame
    rich(t.paragraphs[0], "四组反例指向同一件事", 11.5, INK3, bold=True, accent=BLUE, spc=120)
    para(t, "批准 8 / 查询 1 / 观察 TTL 2：先查再批被校验器拒绝", 12.5, INK, accent=BLUE, space_before=10, line=1.6)
    para(t, "批准必须先拿到同一观察：旧模型下不可行", 12.5, INK, accent=BLUE, space_before=7, line=1.6)
    para(t, "TTL 改 20 才可行：换了 fixture，不是延长真实保证", 12.5, INK, accent=BLUE, space_before=7, line=1.6)
    para(t, "等待变区间 1–10：旧模型直接拒绝区间输入", 12.5, INK, accent=BLUE, space_before=7, line=1.6)
    footer(s, 11, LB, BLUE)

    s = page(prs, "10  ·  独立评估", "这份报告的创意，在哪，不在哪", BLUE,
             "最后做一次独立评估，尽量不客气。有创意的地方两处：一是把一个哲学命题降到一个可反证的工程缺口——"
             "补救的副作用加上不可控等待，这两样以前是分开建模的；二是它用四组反例自己排除了“普通排序”这个伪增量，"
             "也就是说，它先证伪了自己最容易吹的那部分。不够的地方也很清楚：停在问题定义，没有实现动作效应求解器，"
             "没有真实人签、没有真实等待分布、没有人机实验。按发明等级类比，大约在一级末到二级初。")
    b1 = rect(s, ML, 2.14, 5.55, 2.55, None, GREEN, MSO_SHAPE.ROUNDED_RECTANGLE)
    b1.text_frame.margin_left = Inches(0.26)
    b1.text_frame.margin_top = Inches(0.20)
    t = b1.text_frame
    rich(t.paragraphs[0], "有创意", 12, GREEN, bold=True, spc=180)
    para(t, "① 把哲学命题降到一个**可反证**的工程缺口", 13.5, INK, accent=GREEN, space_before=9, line=1.5)
    para(t, "② 用 4/4 反例**自己**排除了“普通排序”这个伪增量", 13.5, INK, accent=GREEN, space_before=7, line=1.5)
    para(t, "③ 把“证据失效”与“不可控等待”联合考虑", 13.5, INK, accent=GREEN, space_before=7, line=1.5)
    b2 = rect(s, 6.85, 2.14, 5.5, 2.55, None, RED, MSO_SHAPE.ROUNDED_RECTANGLE)
    b2.text_frame.margin_left = Inches(0.26)
    b2.text_frame.margin_top = Inches(0.20)
    t = b2.text_frame
    rich(t.paragraphs[0], "不够", 12, RED, bold=True, spc=180)
    para(t, "① 停在**问题定义**，没有实现动作效应求解器", 13.5, INK, accent=RED, space_before=9, line=1.5)
    para(t, "② 无真实人签、无真实等待分布、无人机实验", 13.5, INK, accent=RED, space_before=7, line=1.5)
    para(t, "③ 按发明等级类比，约在**一级末 / 二级初**", 13.5, INK, accent=RED, space_before=7, line=1.5)
    band(s, "提出矛盾是创新的一半，分离出解法是另一半 —— 这份报告目前只有前一半。", 5.06, BLUE, 17, 0.66)
    footer(s, 12, LB, BLUE)

    s = page(prs, "11  ·  收束", "三句话，带回去用", BLUE,
             "收三句。第一，TRIZ 提取矛盾，盒子注入矛盾，矛盾论给矛盾排序——三件事不同，"
             "但都指向同一个动作：分离。第二，哲学给视角，机制给证据；不要用相似性冒充同源，"
             "也不要用同源冒充证明。第三，判断一个想法有没有创意，最硬的判据是："
             "它有没有把“既要又要”逼到必须分离的地步。如果逼到了，哪怕方案还没出来，它已经是有价值的了。")
    pts = [
        ("01", "TRIZ **提取**矛盾，盒子**注入**矛盾，《矛盾论》**排序**矛盾。"),
        ("02", "三者共用一台机器：**分离**。哲学给视角，机制给证据。"),
        ("03", "判断创意，看它有没有把“既要又要”**逼到必须分离**的地步。"),
    ]
    tf = textbox(s, ML, 2.2, CW, 3.3)
    for i, (n, t) in enumerate(pts):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.space_after = Pt(26)
        p.line_spacing = 1.3
        rich(p, n + "　", 15, BLUE, bold=True, accent=BLUE, spc=100)
        rich(p, t, 20.5, INK, accent=BLUE)
    band(s, "下一讲预告：功能分析 —— 把“矛盾在哪”从一段文字，变成一张图。", 5.72, BLUE, 15, 0.62)
    footer(s, 13, LB, BLUE)

    out = D / "TRIZ_W2B_盒子与矛盾论_15min.pptx"
    prs.save(str(out))
    return out

def audit(path):
    prs = Presentation(str(path))
    n = len(prs.slides._sldIdLst)
    minf, notes, warns = 99.0, [], 0
    for si, s in enumerate(prs.slides, 1):
        nt = s.notes_slide.notes_text_frame.text.strip() if s.has_notes_slide else ""
        notes.append(len(nt))
        for sh in s.shapes:
            if not sh.has_text_frame:
                continue
            w_in = sh.width / 914400.0
            h_in = sh.height / 914400.0
            need = 0.0
            for p in sh.text_frame.paragraphs:
                txt = "".join(r.text for r in p.runs)
                if not txt.strip():
                    continue
                sz = max([r.font.size.pt for r in p.runs if r.font.size] or [18])
                minf = min(minf, sz)
                cpl = max(1.0, (w_in - 0.36) * 72.0 / (sz * 1.02))
                lines = math.ceil(len(txt) / cpl)
                need += lines * sz * 1.45 / 72.0
            if need > h_in * 1.30 and h_in > 0.4:
                warns += 1
                print("   ! overflow? slide %-2d need %.2f box %.2f :: %s"
                      % (si, need, h_in,
                         "".join(r.text for r in sh.text_frame.paragraphs[0].runs)[:24]))
    est = sum(notes) / 235.0
    print("%-34s slides=%2d min_font=%.1f notes=%d est=%.1fmin warns=%d"
          % (path.name, n, minf, sum(notes), est, warns))
    return n, minf, est


if __name__ == "__main__":
    a = build_a()
    b = build_b()
    print("built:", a.name)
    print("built:", b.name)
    for f in (a, b):
        audit(f)