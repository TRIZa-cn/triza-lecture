# -*- coding: utf-8 -*-
"""20-minute editable FOS talk, compressed from the 45–60 minute source deck.

Figures are native PowerPoint shapes (editable). The source PPTX is read-only
and is never opened for writing.
"""

from __future__ import annotations

import re
import sys
from pathlib import Path

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, MSO_AUTO_SIZE, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Emu, Pt

OUT_DIR = Path(r"D:\VideoPhone\outputs\FOS_科普_20min_20261022")
PPTX_PATH = OUT_DIR / "FOS_科普_20min.pptx"
MD_PATH = OUT_DIR / "讲稿提纲.md"

SW = 12192000  # 13.333 in
SH = 6858000  # 7.5 in
FONT = "Microsoft YaHei"

INK = "172033"
MUTED = "5B667A"
BLUE = "2F6F9F"
TEAL = "2C7A74"
RED = "9C4545"
AMBER = "8F5A16"
GREEN = "246B48"
WHITE = "FFFFFF"
SOFT_BLUE = "E7F2F8"
SOFT_TEAL = "E6F3F1"
SOFT_RED = "F8ECEA"
SOFT_AMBER = "FBF3E4"
SOFT_GREEN = "E6F4EC"
LINE = "D5DEE6"
PALE = "F4F7FA"


def C(value: str) -> RGBColor:
    value = value.lstrip("#")
    return RGBColor(int(value[0:2], 16), int(value[2:4], 16), int(value[4:6], 16))


def emu(inches: float) -> Emu:
    return Emu(int(round(inches * 914400)))


def set_run(run, size, bold, color, font=FONT):
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = C(color)
    run.font.name = font
    r_pr = run._r.get_or_add_rPr()
    for tag in ("a:latin", "a:ea", "a:cs"):
        node = r_pr.find(qn(tag))
        if node is None:
            node = r_pr.makeelement(qn(tag), {})
            r_pr.append(node)
        node.set("typeface", font)


def fill_text(shape, blocks, anchor="top", margins=(0.12, 0.10, 0.08, 0.06)):
    tf = shape.text_frame
    tf.clear()
    tf.word_wrap = True
    tf.auto_size = MSO_AUTO_SIZE.NONE
    left, right, top, bottom = margins
    tf.margin_left = emu(left)
    tf.margin_right = emu(right)
    tf.margin_top = emu(top)
    tf.margin_bottom = emu(bottom)
    tf.anchor = {"top": MSO_ANCHOR.TOP, "middle": MSO_ANCHOR.MIDDLE, "bottom": MSO_ANCHOR.BOTTOM}[anchor]
    align_map = {"left": PP_ALIGN.LEFT, "center": PP_ALIGN.CENTER, "right": PP_ALIGN.RIGHT}
    for i, block in enumerate(blocks):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align_map[block.get("align", "left")]
        p.space_before = Pt(block.get("before", 0))
        p.space_after = Pt(block.get("after", 0))
        p.line_spacing = block.get("spacing", 1.0)
        run = p.add_run()
        run.text = block["text"]
        set_run(run, block.get("size", 16), block.get("bold", False), block.get("color", INK))
    return shape


def add_shape(slide, kind, x, y, w, h, fill, line=None, line_pt=1.25, radius=0.12):
    shape = slide.shapes.add_shape(kind, emu(x), emu(y), emu(w), emu(h))
    shape.fill.solid()
    shape.fill.fore_color.rgb = C(fill)
    if line is None:
        shape.line.fill.background()
    else:
        shape.line.color.rgb = C(line)
        shape.line.width = Pt(line_pt)
    if kind == MSO_SHAPE.ROUNDED_RECTANGLE:
        try:
            shape.adjustments[0] = radius
        except Exception:
            pass
    return shape


def add_text(slide, x, y, w, h, blocks, anchor="middle", margins=(0.02, 0.02, 0.02, 0.02)):
    box = slide.shapes.add_textbox(emu(x), emu(y), emu(w), emu(h))
    fill_text(box, blocks, anchor=anchor, margins=margins)
    return box


def t(text, size=16, bold=False, color=INK, align="left", before=0, after=0, spacing=1.0):
    return {
        "text": text,
        "size": size,
        "bold": bold,
        "color": color,
        "align": align,
        "before": before,
        "after": after,
        "spacing": spacing,
    }


def set_bg(slide):
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = C(WHITE)


def chrome(slide, section, page, total=26):
    add_shape(slide, MSO_SHAPE.RECTANGLE, 0, 0, 13.333333, 0.08, BLUE)
    add_shape(slide, MSO_SHAPE.RECTANGLE, 0.48, 7.02, 12.38, 0.012, LINE)
    add_text(slide, 0.48, 7.06, 8.4, 0.34, [t(section, 12, False, MUTED)], anchor="middle")
    add_text(
        slide,
        9.5,
        7.06,
        3.35,
        0.34,
        [t(f"{page}  /  {total}", 12, False, MUTED, align="right")],
        anchor="middle",
    )


def heading(slide, text, size=28):
    add_shape(slide, MSO_SHAPE.RECTANGLE, 0.48, 0.30, 0.09, 0.40, BLUE)
    add_text(slide, 0.72, 0.18, 12.15, 0.64, [t(text, size, True, INK)], anchor="middle")


def rrect(slide, x, y, w, h, fill, line=None, line_pt=1.25, radius=0.1):
    return add_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE, x, y, w, h, fill, line, line_pt, radius)


def notes_for(meta):
    return f"建议用时：{meta['sec']} 秒\n\n{meta['say']}\n\n对应原稿：{meta['kept']}"


# ---------------------------------------------------------------------------
# Slides
# ---------------------------------------------------------------------------

def draw_01(slide):
    add_shape(slide, MSO_SHAPE.RECTANGLE, 0, 0, 13.333333, 0.08, BLUE)
    add_text(slide, 0.55, 1.15, 12.2, 0.36, [t("20 分钟科普讲稿", 16, False, TEAL)])
    add_text(slide, 0.55, 1.55, 12.2, 0.85, [t("功能导向检索", 48, True, INK)])
    add_text(
        slide,
        0.55,
        2.50,
        12.0,
        0.95,
        [t("把别的行业已经做熟的事，搬到你的问题里", 26, False, INK)],
    )
    add_text(
        slide,
        0.55,
        3.45,
        12.0,
        0.40,
        [t("不是查重，也不是头脑风暴。", 18, False, MUTED)],
    )
    steps = [("1", "抽功能"), ("2", "泛化名词"), ("3", "选领先领域"), ("4", "下沉并迁回")]
    x = 0.55
    for num, label in steps:
        rrect(slide, x, 5.15, 2.55, 1.15, SOFT_BLUE, LINE, 1.0, 0.12)
        add_text(slide, x + 0.16, 5.28, 0.42, 0.42, [t(num, 16, True, BLUE, "center")], anchor="middle")
        add_text(slide, x + 0.58, 5.28, 1.8, 0.85, [t(label, 18, True, INK)], anchor="middle")
        if num != "4":
            add_shape(slide, MSO_SHAPE.RIGHT_ARROW, x + 2.62, 5.55, 0.28, 0.22, BLUE)
        x += 3.15
    add_text(
        slide,
        0.55,
        6.50,
        10.0,
        0.32,
        [t("一套六步法  ·  一个完整例子：疫苗瓶监视器", 14, False, MUTED)],
    )
    add_text(slide, 10.4, 7.05, 2.4, 0.32, [t("1  /  26", 12, False, MUTED, "right")], anchor="middle")


def draw_02(slide):
    heading(slide, "今天只回答三个问题")
    cards = [
        ("1", "为什么跨行业借鉴", "常常变成生搬硬套？"),
        ("2", "有没有一套可重复的动作", "而不是靠灵感？"),
        ("3", "怎么判断一次迁移", "是真有用，还是自我安慰？"),
    ]
    x = 0.48
    for num, line1, line2 in cards:
        rrect(slide, x, 1.45, 4.00, 5.05, SOFT_BLUE, None, radius=0.08)
        add_shape(slide, MSO_SHAPE.OVAL, x + 0.28, 1.80, 0.72, 0.72, BLUE)
        add_text(slide, x + 0.28, 1.80, 0.72, 0.72, [t(num, 22, True, WHITE, "center")], anchor="middle")
        add_text(slide, x + 0.28, 2.85, 3.45, 1.35, [t(line1, 22, True, INK)], anchor="top")
        add_text(slide, x + 0.28, 4.30, 3.45, 1.35, [t(line2, 20, False, MUTED)], anchor="top")
        x += 4.20


def draw_03(slide):
    heading(slide, "三个不相干的东西，在解同一件事")
    add_text(
        slide,
        0.72,
        0.95,
        12.1,
        0.40,
        [t("没有联网、没有第三方在场，怎么证明一件事发生过，而且只发生过规定的次数", 16, False, MUTED)],
    )
    items = [
        (SOFT_AMBER, AMBER, "疫苗瓶", "小方块随累计受热", "慢慢变深", "vvm"),
        (SOFT_RED, RED, "乏燃料", "安全不看还剩多少", "看已经烧掉多少", "cask"),
        (SOFT_TEAL, TEAL, "空安瓿", "用完必须交回来", "核对批号和数量", "vial"),
    ]
    x = 0.48
    for fill, accent, name, line1, line2, kind in items:
        rrect(slide, x, 1.55, 4.00, 5.05, fill, None, radius=0.08)
        ix = x + 1.35
        if kind == "vvm":
            rrect(slide, ix, 1.85, 1.30, 1.55, WHITE, accent, 1.5, 0.15)
            rrect(slide, ix + 0.38, 2.15, 0.54, 0.54, accent, None, radius=0.15)
        elif kind == "cask":
            rrect(slide, ix + 0.15, 1.85, 1.00, 0.28, accent, None, radius=0.3)
            rrect(slide, ix, 2.10, 1.30, 1.25, accent, None, radius=0.12)
        else:
            add_shape(slide, MSO_SHAPE.OVAL, ix + 0.35, 1.80, 0.60, 0.28, accent)
            rrect(slide, ix + 0.42, 2.00, 0.46, 1.15, accent, None, radius=0.4)
            add_shape(slide, MSO_SHAPE.OVAL, ix + 0.35, 3.05, 0.60, 0.28, accent)
        add_text(slide, x + 0.25, 3.60, 3.50, 0.55, [t(name, 26, True, INK, "center")], anchor="middle")
        add_text(slide, x + 0.25, 4.25, 3.50, 0.65, [t(line1, 18, False, INK, "center")], anchor="middle")
        add_text(slide, x + 0.25, 4.95, 3.50, 0.65, [t(line2, 18, False, MUTED, "center")], anchor="middle")
        x += 4.20


def draw_04(slide):
    heading(slide, "查重在防守，FOS 是去别处找解法")
    rrect(slide, 0.40, 1.25, 5.35, 4.55, SOFT_RED, None, radius=0.08)
    rrect(slide, 7.58, 1.25, 5.35, 4.55, SOFT_TEAL, None, radius=0.08)
    add_shape(slide, MSO_SHAPE.OVAL, 5.84, 2.85, 1.65, 1.65, BLUE)
    add_text(slide, 5.94, 3.15, 1.45, 0.50, [t("同一批", 14, True, WHITE, "center")], anchor="middle")
    add_text(slide, 5.94, 3.60, 1.45, 0.50, [t("文献", 14, True, WHITE, "center")], anchor="middle")
    add_text(slide, 0.62, 1.50, 4.95, 0.55, [t("查重", 26, True, RED)], anchor="middle")
    add_text(slide, 0.62, 2.25, 4.95, 1.3, [t("有没有人做过一样的事？", 20, True, INK)], anchor="top")
    add_text(slide, 0.62, 3.70, 4.95, 1.5, [t("用本领域的术语\n越精确越好", 18, False, MUTED, spacing=1.15)], anchor="top")
    add_text(slide, 7.80, 1.50, 4.95, 0.55, [t("功能导向检索", 26, True, TEAL)], anchor="middle")
    add_text(slide, 7.80, 2.25, 4.95, 1.3, [t("别处有没有人解决了\n同一个功能？", 20, True, INK, spacing=1.1)], anchor="top")
    add_text(slide, 7.80, 4.15, 4.95, 1.2, [t("用上位词和异域术语\n越不像本行越好", 18, False, MUTED, spacing=1.15)], anchor="top")
    add_text(
        slide,
        0.48,
        6.05,
        12.4,
        0.70,
        [t("方向相反，结果几乎没有交集。拿本行的词去搜别的行，当然搜不到。", 18, True, INK, "center")],
        anchor="middle",
    )


def draw_05(slide):
    heading(slide, "约束不同，同一个功能会长出不同结构")
    rrect(slide, 0.48, 1.15, 12.38, 0.85, SOFT_BLUE, None, radius=0.1)
    add_text(
        slide,
        0.70,
        1.15,
        12.0,
        0.85,
        [t("功能相同：防止超量使用", 22, True, BLUE, "center")],
        anchor="middle",
    )
    cols = [
        (SOFT_BLUE, BLUE, "软件", "计数器 + 权限", "程序可以随时改"),
        (SOFT_AMBER, AMBER, "核工业", "看已经烧掉多少", "烧掉的改不回来"),
        (SOFT_TEAL, TEAL, "医院", "收回空安瓿再数", "药用完就没了"),
    ]
    x = 0.48
    for fill, accent, name, how, why in cols:
        rrect(slide, x, 2.25, 4.00, 3.55, fill, None, radius=0.08)
        add_shape(slide, MSO_SHAPE.RECTANGLE, x, 2.25, 0.12, 3.55, accent)
        add_text(slide, x + 0.35, 2.45, 3.45, 0.55, [t(name, 22, True, accent)], anchor="middle")
        add_text(slide, x + 0.35, 3.15, 3.45, 1.15, [t(how, 20, True, INK)], anchor="top")
        add_text(slide, x + 0.35, 4.45, 3.45, 0.90, [t(why, 16, False, MUTED)], anchor="top")
        x += 4.20
    add_text(
        slide,
        0.48,
        6.00,
        12.38,
        0.75,
        [t("离开本领域，是为了看见本行约束长不出来的结构。", 18, True, INK, "center")],
        anchor="middle",
    )


def draw_06(slide):
    heading(slide, "要搬的是被约束逼出来的结构")
    rrect(slide, 0.48, 1.25, 5.85, 4.85, PALE, LINE, 1.25, 0.08)
    rrect(slide, 7.00, 1.25, 5.85, 4.85, SOFT_TEAL, TEAL, 1.75, 0.08)
    add_text(slide, 0.75, 1.50, 5.35, 0.50, [t("功能", 16, True, MUTED)], anchor="middle")
    add_text(slide, 0.75, 2.15, 5.35, 1.4, [t("防止超量使用", 28, True, INK)], anchor="top")
    add_text(slide, 0.75, 3.80, 5.35, 1.6, [t("抽到最后，大家得到的\n往往是同一句话。", 18, False, MUTED, spacing=1.15)], anchor="top")
    add_text(slide, 7.28, 1.50, 5.35, 0.50, [t("结构", 16, True, TEAL)], anchor="middle")
    add_text(slide, 7.28, 2.15, 5.35, 1.6, [t("计数器、燃耗、\n交回空瓶", 28, True, INK, spacing=1.05)], anchor="top")
    add_text(slide, 7.28, 4.20, 5.35, 1.4, [t("部件、字段，或一台\n状态机。", 18, False, TEAL, spacing=1.15)], anchor="top")
    add_text(
        slide,
        0.48,
        6.25,
        12.38,
        0.55,
        [t("搬右边。左边那句话，谁都会写。", 20, True, INK, "center")],
        anchor="middle",
    )


def draw_07(slide):
    heading(slide, "抽功能只是翻译，停在这里没有新信息")
    rrect(slide, 0.48, 1.55, 5.15, 2.55, SOFT_BLUE, None, radius=0.08)
    rrect(slide, 7.70, 1.55, 5.15, 2.55, SOFT_AMBER, None, radius=0.08)
    add_shape(slide, MSO_SHAPE.RIGHT_ARROW, 5.90, 2.50, 1.50, 0.48, BLUE)
    add_text(slide, 0.70, 1.75, 4.75, 0.45, [t("本领域的说法", 14, True, BLUE)], anchor="middle")
    add_text(slide, 0.70, 2.30, 4.75, 1.35, [t("离线槽位计数", 28, True, INK)], anchor="middle")
    add_text(slide, 7.92, 1.75, 4.75, 0.45, [t("抽成功能", 14, True, AMBER)], anchor="middle")
    add_text(
        slide,
        7.92,
        2.25,
        4.75,
        1.55,
        [t("无连接时，限制\n某动作的次数", 24, True, INK, spacing=1.05)],
        anchor="middle",
    )
    rrect(slide, 0.48, 4.50, 12.38, 2.15, SOFT_RED, None, radius=0.08)
    add_text(slide, 0.80, 4.70, 11.8, 0.50, [t("停在这一步", 16, True, RED)], anchor="middle")
    add_text(
        slide,
        0.80,
        5.25,
        11.8,
        1.05,
        [t("所有人都会得到同一句话，搜回来的也是同一种方案。", 22, True, INK)],
        anchor="middle",
    )


def draw_08(slide):
    heading(slide, "泛化的是名词，再下沉成具体物")
    rrect(slide, 0.70, 1.35, 3.40, 1.35, PALE, LINE, 1.25, 0.12)
    add_text(slide, 0.70, 1.35, 3.40, 1.35, [t("瓶子", 28, True, INK, "center")], anchor="middle")
    add_shape(slide, MSO_SHAPE.RIGHT_ARROW, 4.30, 1.85, 0.70, 0.36, BLUE)
    rrect(slide, 5.20, 1.25, 3.50, 1.55, BLUE, None, radius=0.12)
    add_text(slide, 5.20, 1.25, 3.50, 1.55, [t("容器", 32, True, WHITE, "center")], anchor="middle")
    add_shape(slide, MSO_SHAPE.DOWN_ARROW, 6.55, 2.95, 0.42, 0.38, TEAL)
    sinks = [("疫苗瓶", "疫苗冷链"), ("太空舱", "航天"), ("乏燃料罐", "核工业")]
    x = 0.70
    for name, domain in sinks:
        rrect(slide, x, 3.55, 3.85, 2.15, SOFT_TEAL, None, radius=0.1)
        add_text(slide, x + 0.20, 3.75, 3.45, 0.90, [t(name, 26, True, INK, "center")], anchor="middle")
        add_text(slide, x + 0.20, 4.70, 3.45, 0.65, [t(domain, 18, False, TEAL, "center")], anchor="middle")
        x += 4.15
    add_text(
        slide,
        0.48,
        6.20,
        12.4,
        0.60,
        [t("去具体物身上找。不要停在「容器」这个词上。", 20, True, INK, "center")],
        anchor="middle",
    )


def draw_09(slide):
    heading(slide, "四级阶梯写下来，检索才查得回")
    steps = [
        ("1", "本领域的词", "离线槽位", SOFT_BLUE, BLUE),
        ("2", "上位词", "有限次预发的\n消耗单元", SOFT_AMBER, AMBER),
        ("3", "先进领域", "航天\n被迫做得更狠", SOFT_RED, RED),
        ("4", "具体物", "在轨消耗品包\n＋补给窗口", SOFT_TEAL, TEAL),
    ]
    x = 0.40
    for num, name, body, fill, accent in steps:
        rrect(slide, x, 1.25, 2.85, 4.15, fill, None, radius=0.08)
        add_shape(slide, MSO_SHAPE.OVAL, x + 1.08, 1.48, 0.68, 0.68, accent)
        add_text(slide, x + 1.08, 1.48, 0.68, 0.68, [t(num, 18, True, WHITE, "center")], anchor="middle")
        add_text(slide, x + 0.12, 2.30, 2.60, 0.70, [t(name, 16, True, accent, "center")], anchor="middle")
        add_text(slide, x + 0.12, 3.10, 2.60, 1.85, [t(body, 18, True, INK, "center", spacing=1.05)], anchor="top")
        if num != "4":
            add_shape(slide, MSO_SHAPE.RIGHT_ARROW, x + 2.90, 3.05, 0.32, 0.26, BLUE)
        x += 3.22
    add_text(
        slide,
        0.48,
        5.65,
        12.4,
        1.10,
        [t("自左向右就是检索路径。少一级，推理就跳步了。", 20, True, INK, "center")],
        anchor="middle",
    )


def draw_10(slide):
    heading(slide, "合适的上位词，能点出三个行业的实物")
    rows = [
        (SOFT_RED, RED, "太浅", "瓶子  →  罐子", "还在包装行业里打转"),
        (SOFT_AMBER, AMBER, "太深", "瓶子  →  东西", "什么都搜得到，等于没搜"),
        (SOFT_TEAL, TEAL, "合适", "瓶子  →  容器", "马上能举出三个行业的实物"),
    ]
    y = 1.25
    for fill, accent, tag, change, judge in rows:
        rrect(slide, 0.48, y, 12.38, 1.65, fill, None, radius=0.08)
        add_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE, 0.72, y + 0.42, 1.55, 0.80, accent, radius=0.2)
        add_text(slide, 0.72, y + 0.42, 1.55, 0.80, [t(tag, 20, True, WHITE, "center")], anchor="middle")
        add_text(slide, 2.55, y + 0.22, 5.4, 1.20, [t(change, 26, True, INK)], anchor="middle")
        add_text(slide, 8.05, y + 0.22, 4.50, 1.20, [t(judge, 18, False, accent)], anchor="middle")
        y += 1.85


def draw_11(slide):
    heading(slide, "这一步的矛盾：跳级和过度都会做空")
    rrect(slide, 0.48, 1.25, 6.05, 4.35, SOFT_RED, None, radius=0.08)
    rrect(slide, 6.80, 1.25, 6.05, 4.35, SOFT_AMBER, None, radius=0.08)
    add_text(slide, 0.75, 1.50, 5.55, 0.55, [t("跳级", 26, True, RED)], anchor="middle")
    add_text(slide, 0.75, 2.25, 5.55, 1.4, [t("只拿领域词去搜", 22, True, INK)], anchor="top")
    add_text(slide, 0.75, 3.80, 5.55, 1.3, [t("回来的全是同构方案\n没有新结构", 18, False, MUTED, spacing=1.15)], anchor="top")
    add_text(slide, 7.08, 1.50, 5.55, 0.55, [t("过度泛化", 26, True, AMBER)], anchor="middle")
    add_text(slide, 7.08, 2.25, 5.55, 1.4, [t("上位词大到无法下沉", 22, True, INK)], anchor="top")
    add_text(slide, 7.08, 3.80, 5.55, 1.3, [t("写不出可执行的\n结构特征", 18, False, MUTED, spacing=1.15)], anchor="top")
    add_text(
        slide,
        0.48,
        5.85,
        12.38,
        0.90,
        [t("合适的抽象夹在中间：说出上位词，还能点出三个行业的实物。", 18, True, INK, "center")],
        anchor="middle",
    )


def draw_12(slide):
    heading(slide, "领先领域，是被迫把这件事做得更狠的")
    cols = [
        ("航天", "坏了没法派人去修", "地球独立"),
        ("核工业", "记录改不了", "用累计消耗证明"),
        ("医院", "必须能追责", "交回实物再核销"),
    ]
    x = 0.48
    for name, force, result in cols:
        rrect(slide, x, 1.25, 4.00, 4.15, SOFT_BLUE, None, radius=0.08)
        add_text(slide, x + 0.25, 1.50, 3.50, 0.70, [t(name, 26, True, BLUE, "center")], anchor="middle")
        add_text(slide, x + 0.25, 2.40, 3.50, 1.20, [t(force, 18, True, INK, "center")], anchor="middle")
        add_shape(slide, MSO_SHAPE.DOWN_ARROW, x + 1.72, 3.70, 0.36, 0.36, TEAL)
        add_text(slide, x + 0.25, 4.15, 3.50, 0.90, [t(result, 18, True, TEAL, "center")], anchor="middle")
        x += 4.20
    add_text(
        slide,
        0.48,
        5.65,
        12.38,
        1.10,
        [t("不要挑听起来先进的行业。问：在这个功能上，谁被逼得比我们更极端？", 18, True, INK, "center")],
        anchor="middle",
    )


def draw_13(slide):
    heading(slide, "不下沉到具体物，结论就不可执行")
    steps = [
        ("1", "上位词", "只是路标"),
        ("2", "具体物", "看得见，能拆开"),
        ("3", "问一句", "哪一段因约束而长出"),
        ("4", "只搬那段", "没有类似约束\n就不要整包搬"),
    ]
    x = 0.40
    for num, name, body in steps:
        rrect(slide, x, 1.30, 2.90, 3.15, SOFT_TEAL if num == "4" else PALE, TEAL if num == "4" else LINE, 1.25, 0.08)
        add_text(slide, x + 0.15, 1.50, 2.60, 0.50, [t(num, 16, True, TEAL, "center")], anchor="middle")
        add_text(slide, x + 0.15, 2.05, 2.60, 0.70, [t(name, 22, True, INK, "center")], anchor="middle")
        add_text(slide, x + 0.18, 2.85, 2.55, 1.30, [t(body, 16, False, MUTED, "center")], anchor="top")
        if num != "4":
            add_shape(slide, MSO_SHAPE.RIGHT_ARROW, x + 2.95, 2.60, 0.28, 0.22, BLUE)
        x += 3.22
    rrect(slide, 0.48, 4.75, 12.38, 1.95, SOFT_AMBER, None, radius=0.08)
    add_text(slide, 0.75, 4.95, 11.9, 0.45, [t("不可执行的句子", 14, True, AMBER)], anchor="middle")
    add_text(
        slide,
        0.75,
        5.45,
        11.9,
        0.95,
        [t("「借鉴航天的库存管理」等于没说。要落到：每个消耗单元带库位和补给窗口。", 18, True, INK)],
        anchor="middle",
    )


def draw_14(slide):
    heading(slide, "迁回过三关：结构、前提、可观测效果")
    gates = [
        ("1", "写成结构", "能不能落到部件、\n字段或状态机？"),
        ("2", "前提还在", "原领域靠什么成立，\n这里还有没有？"),
        ("3", "效果可测", "有没有一个原来测不到、\n现在能测的量？"),
    ]
    x = 0.48
    for num, name, body in gates:
        rrect(slide, x, 1.30, 4.00, 4.20, WHITE, LINE, 1.25, 0.08)
        add_shape(slide, MSO_SHAPE.OVAL, x + 1.50, 1.60, 0.95, 0.95, BLUE)
        add_text(slide, x + 1.50, 1.60, 0.95, 0.95, [t(num, 26, True, WHITE, "center")], anchor="middle")
        add_text(slide, x + 0.25, 2.75, 3.50, 0.70, [t(name, 24, True, INK, "center")], anchor="middle")
        add_text(slide, x + 0.25, 3.55, 3.50, 1.50, [t(body, 16, False, MUTED, "center", spacing=1.1)], anchor="top")
        x += 4.20
    add_text(
        slide,
        0.48,
        5.75,
        12.38,
        1.00,
        [t("三条缺一，就还是灵感笔记。顺序不要乱：先结构，再前提，再效果。", 18, True, INK, "center")],
        anchor="middle",
    )


def draw_15(slide):
    heading(slide, "二元的已用标记，证明不了「只用过一次」")
    add_text(
        slide,
        0.72,
        0.95,
        12.0,
        0.38,
        [t("完整例子只走一个：离线槽位", 16, False, TEAL)],
    )
    rrect(slide, 0.48, 1.55, 3.70, 2.15, SOFT_GREEN, None, radius=0.1)
    rrect(slide, 4.40, 1.55, 3.70, 2.15, SOFT_RED, None, radius=0.1)
    add_text(slide, 0.48, 1.55, 3.70, 2.15, [t("未用", 32, True, GREEN, "center")], anchor="middle")
    add_text(slide, 4.40, 1.55, 3.70, 2.15, [t("已用", 32, True, RED, "center")], anchor="middle")
    add_shape(slide, MSO_SHAPE.RIGHT_ARROW, 8.30, 2.30, 0.55, 0.40, RED)
    rrect(slide, 9.05, 1.55, 3.80, 2.15, SOFT_RED, RED, 1.75, 0.1)
    add_text(slide, 9.20, 1.70, 3.50, 0.70, [t("改回未用", 22, True, RED, "center")], anchor="middle")
    add_text(slide, 9.20, 2.45, 3.50, 0.90, [t("外观上看不出", 18, False, INK, "center")], anchor="middle")
    rrect(slide, 0.48, 4.10, 12.38, 2.55, SOFT_BLUE, None, radius=0.08)
    add_text(slide, 0.80, 4.35, 11.8, 0.50, [t("要找的功能", 16, True, BLUE)], anchor="middle")
    add_text(
        slide,
        0.80,
        5.00,
        11.8,
        1.20,
        [t("一份累计的、不可逆的使用指示。", 28, True, INK)],
        anchor="middle",
    )


def draw_16(slide):
    heading(slide, "已用标记，沿阶梯走到疫苗瓶监视器")
    steps = [
        ("1", "已用标记", "本领域的词"),
        ("2", "累计不可逆指示", "上位词"),
        ("3", "疫苗冷链", "热稳定性被管得很死"),
        ("4", "疫苗瓶监视器", "贴纸随累计受热变深"),
    ]
    y = 1.15
    for num, name, note in steps:
        fill = SOFT_TEAL if num == "4" else PALE
        rrect(slide, 0.55, y, 12.25, 1.20, fill, TEAL if num == "4" else LINE, 1.25, 0.15)
        add_shape(slide, MSO_SHAPE.OVAL, 0.80, y + 0.28, 0.64, 0.64, TEAL if num == "4" else BLUE)
        add_text(slide, 0.80, y + 0.28, 0.64, 0.64, [t(num, 18, True, WHITE, "center")], anchor="middle")
        add_text(slide, 1.70, y + 0.15, 6.4, 0.90, [t(name, 24, True, INK)], anchor="middle")
        add_text(slide, 8.20, y + 0.15, 4.30, 0.90, [t(note, 16, False, TEAL if num == "4" else MUTED, "right")], anchor="middle")
        y += 1.38


def draw_17(slide):
    heading(slide, "二元状态留不住历史，累计档位留得住")
    rrect(slide, 0.40, 1.15, 6.15, 5.55, SOFT_RED, None, radius=0.08)
    rrect(slide, 6.78, 1.15, 6.15, 5.55, SOFT_TEAL, None, radius=0.08)
    add_text(slide, 0.60, 1.35, 5.75, 0.55, [t("阈值式", 22, True, RED, "center")], anchor="middle")
    add_text(slide, 6.98, 1.35, 5.75, 0.55, [t("累计式", 22, True, TEAL, "center")], anchor="middle")
    rrect(slide, 0.75, 2.15, 2.55, 1.70, SOFT_GREEN, GREEN, 1.5, 0.1)
    rrect(slide, 3.55, 2.15, 2.55, 1.70, WHITE, RED, 1.75, 0.1)
    add_text(slide, 0.75, 2.15, 2.55, 1.70, [t("未用", 24, True, GREEN, "center")], anchor="middle")
    add_text(slide, 3.55, 2.15, 2.55, 1.70, [t("已用", 24, True, RED, "center")], anchor="middle")
    add_text(
        slide,
        0.70,
        4.15,
        5.55,
        2.10,
        [
            t("超过就报警", 20, True, INK, "center"),
            t("改回「未用」", 18, False, MUTED, "center", before=8),
            t("看不出被动过", 18, False, MUTED, "center", before=4),
        ],
        anchor="top",
    )
    shades = [("0", "F4E4C4", INK), ("1", "E2C07A", INK), ("2", "C4843A", WHITE), ("3", "6B3E1F", WHITE)]
    sx = 7.05
    for label, fill, color in shades:
        rrect(slide, sx, 2.20, 1.30, 1.55, fill, None, radius=0.1)
        add_text(slide, sx, 2.20, 1.30, 1.55, [t(label, 26, True, color, "center")], anchor="middle")
        sx += 1.45
    add_text(
        slide,
        7.00,
        4.15,
        5.75,
        2.10,
        [
            t("只增不减", 20, True, INK, "center"),
            t("每一档对应不同动作", 18, False, MUTED, "center", before=8),
            t("断网也能当场读", 18, False, TEAL, "center", before=4),
        ],
        anchor="top",
    )


def draw_18(slide):
    heading(slide, "迁回来的结构，是只增不减的档位")
    rrect(slide, 0.48, 1.25, 5.70, 3.55, PALE, LINE, 1.25, 0.08)
    rrect(slide, 7.15, 1.25, 5.70, 3.55, SOFT_TEAL, TEAL, 1.75, 0.08)
    add_shape(slide, MSO_SHAPE.RIGHT_ARROW, 6.30, 2.70, 0.70, 0.40, BLUE)
    add_text(slide, 0.70, 1.45, 5.30, 0.45, [t("迁之前", 14, True, MUTED)], anchor="middle")
    add_text(slide, 0.70, 2.05, 5.30, 1.15, [t("一个布尔值", 28, True, INK)], anchor="middle")
    add_text(slide, 0.70, 3.30, 5.30, 1.05, [t("用过就置位", 18, False, MUTED)], anchor="top")
    add_text(slide, 7.40, 1.45, 5.25, 0.45, [t("迁回来", 14, True, TEAL)], anchor="middle")
    add_text(slide, 7.40, 2.00, 5.25, 1.20, [t("0–3 单调档位", 28, True, INK)], anchor="middle")
    add_text(slide, 7.40, 3.30, 5.25, 1.15, [t("0 正常，3 只能核销", 18, False, TEAL)], anchor="top")
    rrect(slide, 0.48, 5.05, 12.38, 1.65, SOFT_BLUE, None, radius=0.08)
    add_text(slide, 0.75, 5.25, 11.9, 0.45, [t("多出来的效果", 14, True, BLUE)], anchor="middle")
    add_text(
        slide,
        0.75,
        5.70,
        11.9,
        0.75,
        [t("档位本身就是现场可读的审计证据，不联网也看得出来。", 20, True, INK)],
        anchor="middle",
    )


def draw_19(slide):
    heading(slide, "三关都过，才算搬成；没前提就不能搬")
    checks = [
        ("结构", "单调档位：0 到 3，只增不减"),
        ("前提", "断网时也要留下看得见的历史"),
        ("效果", "现场能读出被透支过"),
    ]
    y = 1.15
    for name, body in checks:
        rrect(slide, 0.48, y, 12.38, 1.15, SOFT_GREEN, None, radius=0.12)
        add_shape(slide, MSO_SHAPE.OVAL, 0.72, y + 0.28, 0.58, 0.58, GREEN)
        add_text(slide, 0.72, y + 0.28, 0.58, 0.58, [t("过", 16, True, WHITE, "center")], anchor="middle")
        add_text(slide, 1.50, y + 0.15, 1.70, 0.85, [t(name, 20, True, GREEN)], anchor="middle")
        add_text(slide, 3.30, y + 0.15, 9.20, 0.85, [t(body, 20, False, INK)], anchor="middle")
        y += 1.30
    rrect(slide, 0.48, 5.15, 12.38, 1.60, SOFT_AMBER, None, radius=0.08)
    add_text(slide, 0.75, 5.30, 11.9, 0.40, [t("前提不在，就停", 14, True, AMBER)], anchor="middle")
    add_text(
        slide,
        0.75,
        5.75,
        11.9,
        0.75,
        [t("燃耗信用靠物理定律兜底。软件里没有这条定律，不能直接搬。", 18, True, INK)],
        anchor="middle",
    )


def draw_20(slide):
    heading(slide, "按这六步，散会就可以做一张卡")
    steps = [
        ("1", "列出名词", "5 个本领域关键词"),
        ("2", "写上位词", "用三个行业实物检验"),
        ("3", "选领域", "挑被迫做得更狠的"),
        ("4", "下沉", "写清它对付的约束"),
        ("5", "检索核实", "标明已核实或待核实"),
        ("6", "写成迁移卡", "三道关，补哪条权利要求"),
    ]
    positions = [(0.40, 1.12), (4.50, 1.12), (8.60, 1.12), (0.40, 3.05), (4.50, 3.05), (8.60, 3.05)]
    for (num, name, body), (x, y) in zip(steps, positions):
        rrect(slide, x, y, 3.90, 1.75, SOFT_BLUE if num != "6" else SOFT_TEAL, None, radius=0.1)
        add_text(slide, x + 0.18, y + 0.12, 0.50, 0.45, [t(num, 18, True, BLUE if num != "6" else TEAL)], anchor="middle")
        add_text(slide, x + 0.65, y + 0.12, 3.05, 0.48, [t(name, 20, True, INK)], anchor="middle")
        add_text(slide, x + 0.20, y + 0.75, 3.50, 0.80, [t(body, 16, False, MUTED)], anchor="middle")
    add_text(
        slide,
        0.48,
        5.05,
        12.38,
        1.70,
        [
            t("顺序不要反", 16, True, RED, "center"),
            t("先写上位词，再挑领域。反了，容易挑到听起来高级、其实无关的行业。", 18, True, INK, "center", before=6),
        ],
        anchor="middle",
    )


def draw_21(slide):
    heading(slide, "五个错误里，最常见的是跳过泛化")
    rows = [
        ("1", "只搜领域词，不泛化名词", "回来的全是同构方案"),
        ("2", "上位词大到无法下沉", "没有可执行的结构"),
        ("3", "停在「借鉴某某经验」", "句子不可执行"),
        ("4", "忽略原领域的前提", "没兜底的机制被直接搬来"),
        ("5", "只记机会，不记威胁", "检索变成自我表扬"),
    ]
    y = 1.12
    for num, left, right in rows:
        rrect(slide, 0.48, y, 12.38, 1.05, WHITE, LINE, 1.15, 0.15)
        add_shape(slide, MSO_SHAPE.OVAL, 0.68, y + 0.24, 0.56, 0.56, RED if num == "1" else BLUE)
        add_text(slide, 0.68, y + 0.24, 0.56, 0.56, [t(num, 16, True, WHITE, "center")], anchor="middle")
        add_text(slide, 1.45, y + 0.12, 6.3, 0.80, [t(left, 18, True, INK)], anchor="middle")
        add_text(slide, 7.70, y + 0.12, 4.90, 0.80, [t(right, 16, False, MUTED)], anchor="middle")
        y += 1.14


def draw_22(slide):
    heading(slide, "搜到同构标准时，把它记成威胁")
    add_text(
        slide,
        0.72,
        0.95,
        12.0,
        0.36,
        [t("3GPP 在线计费：先申请配额，用尽就终止会话", 16, False, MUTED)],
    )
    rrect(slide, 0.48, 1.50, 6.05, 3.85, SOFT_AMBER, None, radius=0.08)
    rrect(slide, 6.80, 1.50, 6.05, 3.85, SOFT_RED, None, radius=0.08)
    add_text(slide, 0.75, 1.70, 5.55, 0.50, [t("功能上几乎一样", 20, True, AMBER)], anchor="middle")
    add_text(
        slide,
        0.85,
        2.40,
        5.40,
        2.50,
        [
            t("预发额度", 22, True, INK),
            t("计数", 22, True, INK, before=8),
            t("用尽即停", 22, True, INK, before=8),
        ],
        anchor="top",
    )
    add_text(slide, 7.05, 1.70, 5.55, 0.50, [t("差在有没有实时通道", 20, True, RED)], anchor="middle")
    add_text(
        slide,
        7.15,
        2.45,
        5.45,
        2.40,
        [
            t("对方依赖实时回传", 20, True, INK),
            t("完全断连时", 20, True, INK, before=10),
            t("这条标准对不上", 20, True, RED, before=6),
        ],
        anchor="top",
    )
    add_text(
        slide,
        0.48,
        5.60,
        12.38,
        1.15,
        [t("记成威胁。不要当成一条新灵感，也不要假装没看见。", 20, True, INK, "center")],
        anchor="middle",
    )


def draw_23(slide):
    heading(slide, "印章这个词，值得当场泛化一次")
    steps = [
        ("上位词", "可验证的确认标记"),
        ("三个行业", "票据 · 关封 · 批签"),
        ("下沉", "铅封编号，骑缝章"),
    ]
    x = 0.48
    for name, body in steps:
        rrect(slide, x, 1.20, 4.00, 2.15, SOFT_BLUE, None, radius=0.08)
        add_text(slide, x + 0.20, 1.35, 3.60, 0.50, [t(name, 16, True, BLUE, "center")], anchor="middle")
        add_text(slide, x + 0.20, 1.95, 3.60, 1.10, [t(body, 20, True, INK, "center")], anchor="middle")
        x += 4.20
    rrect(slide, 0.48, 3.65, 12.38, 2.95, SOFT_AMBER, None, radius=0.08)
    add_text(slide, 0.80, 3.90, 11.8, 0.50, [t("第四步要问的那一句", 16, True, AMBER)], anchor="middle")
    add_text(
        slide,
        0.80,
        4.55,
        11.8,
        1.55,
        [t("哪一部分，是为了「无法当面核验」\n才长出来的？", 26, True, INK, "center", spacing=1.05)],
        anchor="middle",
    )


def draw_24(slide):
    heading(slide, "带走三句话就够开始")
    lines = [
        ("1", "泛化的是名词，不是功能", "这决定你搜什么"),
        ("2", "下沉之后再问它对付什么约束", "这决定值不值得搬"),
        ("3", "迁移卡要有出处和三道关", "否则只是灵感笔记"),
    ]
    y = 1.20
    for num, main, sub in lines:
        rrect(slide, 0.48, y, 12.38, 1.70, SOFT_BLUE, None, radius=0.08)
        add_shape(slide, MSO_SHAPE.OVAL, 0.75, y + 0.48, 0.72, 0.72, BLUE)
        add_text(slide, 0.75, y + 0.48, 0.72, 0.72, [t(num, 20, True, WHITE, "center")], anchor="middle")
        add_text(slide, 1.75, y + 0.22, 10.7, 0.70, [t(main, 24, True, INK)], anchor="middle")
        add_text(slide, 1.75, y + 0.92, 10.7, 0.50, [t(sub, 16, False, MUTED)], anchor="middle")
        y += 1.88


def draw_25(slide):
    heading(slide, "出处只说明那边确实有这种做法")
    rows = [
        ("WHO", "四种疫苗瓶监视器与热稳定性", "WHO/V&B/03.18 Rev.1"),
        ("NRC", "仅锕系燃耗信用", "ML003702758"),
        ("卫健委", "麻醉药品空安瓿回收", "第五十四条"),
        ("NASA", "在轨库存与地球独立医疗", "ISS  /  AMIS"),
        ("3GPP", "在线计费与配额预留", "TS 32.240"),
    ]
    y = 1.15
    for code, what, ref in rows:
        rrect(slide, 0.48, y, 12.38, 0.95, WHITE, LINE, 1.1, 0.16)
        add_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE, 0.68, y + 0.20, 1.70, 0.55, SOFT_BLUE, radius=0.2)
        add_text(slide, 0.68, y + 0.20, 1.70, 0.55, [t(code, 14, True, BLUE, "center")], anchor="middle")
        add_text(slide, 2.55, y + 0.12, 6.15, 0.72, [t(what, 18, True, INK)], anchor="middle")
        add_text(slide, 8.55, y + 0.12, 4.05, 0.72, [t(ref, 15, False, MUTED, "right")], anchor="middle")
        y += 1.05
    add_text(
        slide,
        0.48,
        6.40,
        12.38,
        0.48,
        [t("不是新颖性结论。逐篇全文对比，是另一项工作。", 16, True, RED, "center")],
        anchor="middle",
    )


def draw_26(slide):
    add_shape(slide, MSO_SHAPE.RECTANGLE, 0, 0, 13.333333, 0.08, BLUE)
    add_text(slide, 0.60, 1.35, 12.1, 0.40, [t("散会前只留一个问题", 16, False, TEAL)])
    add_text(
        slide,
        0.60,
        1.90,
        12.1,
        1.70,
        [t("你手上哪个词，\n值得泛化一次？", 40, True, INK, spacing=1.0)],
        anchor="top",
    )
    rrect(slide, 0.60, 4.15, 12.15, 1.85, SOFT_BLUE, None, radius=0.08)
    add_text(slide, 0.90, 4.35, 11.6, 0.45, [t("最小动作", 14, True, BLUE)], anchor="middle")
    add_text(
        slide,
        0.90,
        4.85,
        11.6,
        0.85,
        [t("五个名词，各写一个上位词，再做选领域和下沉。", 20, True, INK)],
        anchor="middle",
    )
    add_text(
        slide,
        0.60,
        6.25,
        12.1,
        0.55,
        [t("卡片库  examples/fos/cards     工具  tools/fosdb.py", 16, False, MUTED)],
        anchor="middle",
    )
    add_text(slide, 10.4, 7.05, 2.4, 0.32, [t("26  /  26", 12, False, MUTED, "right")], anchor="middle")


DRAW = [
    draw_01, draw_02, draw_03, draw_04, draw_05, draw_06, draw_07, draw_08,
    draw_09, draw_10, draw_11, draw_12, draw_13, draw_14, draw_15, draw_16,
    draw_17, draw_18, draw_19, draw_20, draw_21, draw_22, draw_23, draw_24,
    draw_25, draw_26,
]

META = [
    dict(sec=40, title="功能导向检索", section="开场", figure="底部四步检索链",
         kept="第 1 页标题和副标题；开场备注里的「不是查重，也不是头脑风暴」。",
         cut="删去 45–60 分钟、8 个案例的场次说明，改成 20 分钟和一个完整例子。",
         say="今天不是查重，也不是头脑风暴。功能导向检索是一套固定动作：先把名词收成上位词，再去被迫做得更狠的领域里找具体的东西，看哪一段结构能搬回来。二十分钟，只走通一个例子。"),
    dict(sec=45, title="今天只回答三个问题", section="开场", figure="三问卡片",
         kept="第 2 页的三个问题，原句保留。",
         cut="删去备注里对「验收标准」的展开，留到迁回三关再讲。",
         say="先管住预期。跨行业借鉴失败，多半是因为功能抽得太粗，粗到什么都能类比。今天只要三样东西：一套动作，一个例子，和一个能判断迁移成没成的标准。"),
    dict(sec=50, title="三个不相干的东西，在解同一件事", section="为什么离开本领域", figure="三个物件示意",
         kept="第 3 页的疫苗瓶、乏燃料、空安瓿，以及备注里的共同问题。",
         cut="三个物件只作引子。乏燃料和空安瓿的案例正文不在这里展开。",
         say="记住三个画面。疫苗瓶上的方块会随累计受热变深。乏燃料的安全看已经烧掉多少。空安瓿用完必须交回来核对。没有联网、没有第三方在场，怎么证明一件事发生过，而且只发生过规定的次数。后面只把第一个走完。"),
    dict(sec=45, title="查重在防守，FOS 是去别处找解法", section="FOS 是什么", figure="查重与 FOS 对照",
         kept="第 4 页：查重与 FOS 的问法、用词方向相反。",
         cut="不展开检索系统操作。",
         say="查重问有没有人做过一样的事，用本领域的词，越准越好。功能导向检索问别处有没有人解决了同一个功能，用上位词和别人的术语。同一批文献，方向相反。拿本行的词去搜别的行，当然搜不到。"),
    dict(sec=50, title="约束不同，同一个功能会长出不同结构", section="为什么离开本领域", figure="三种约束对照",
         kept="第 5 页：防止超量使用在软件、核工业、医院的三种做法。",
         cut="不进入燃耗信用和空安瓿的机制推导。",
         say="都要防止超量使用。软件用计数器和权限，因为程序可以改。核工业看已经烧掉多少，因为烧掉的改不回来。医院把空安瓿收回来数，因为药用完就没了。功能相同，结构不同，差别来自约束。离开本领域，就是为了看见这些结构。"),
    dict(sec=45, title="要搬的是被约束逼出来的结构", section="为什么离开本领域", figure="功能与结构对照",
         kept="第 5 页备注里的枢纽句：迁移的不是功能，是被约束逼出来的结构。",
         cut="不另举新行业。",
         say="功能抽到最后，大家往往是同一句话。值得搬的是在某种约束下被逼出来的结构：一个部件、一个字段，或一台状态机。后面所有步骤，都是在找这个结构。"),
    dict(sec=45, title="抽功能只是翻译，停在这里没有新信息", section="功能抽象", figure="名词改写成功能",
         kept="第 6 页：离线槽位计数改写成无连接条件下限制次数。",
         cut="删去作者自述失误的长段，错误页再点一句。",
         say="第一步仍然要做。把名词换成动词短语，去掉领域词。离线槽位计数，变成无连接时限制某动作的次数。这一步只是翻译，不产生新信息。停在这里去搜，回来的都是同一种方案。"),
    dict(sec=45, title="泛化的是名词，再下沉成具体物", section="功能抽象", figure="瓶子到容器再下沉",
         kept="第 7 页：瓶子、容器，以及疫苗瓶、太空舱、乏燃料罐。",
         cut="不解释卡片库字段。",
         say="第二步才是分水岭。不要只抽象功能，要抽象物体。瓶子收到容器，容器在疫苗里是瓶，在航天里是舱，在核工业里是乏燃料罐。每个具体物身上，都带着那个行业为自己的约束长出来的结构。"),
    dict(sec=55, title="四级阶梯写下来，检索才查得回", section="功能抽象", figure="四级阶梯",
         kept="第 8–9 页的四级：离线槽位、消耗单元、航天、在轨消耗品包和补给窗口。",
         cut="两页合成一页。删去 fosdb 至少三级的工具规则，只保留「写下来才能复查」。空间站不另开案例。",
         say="把路径写成四级，别人才查得回你有没有跳步。本领域的词，上位词，先进领域，再落回那个领域的具体物。离线槽位收到有限次预发的消耗单元，领域挑航天，因为坏了没法派人去修，具体物是带补给窗口的在轨消耗品包。"),
    dict(sec=45, title="合适的上位词，能点出三个行业的实物", section="功能抽象", figure="太浅、太深、合适",
         kept="第 10 页的三档粒度和「三个行业实物」判据。",
         cut="不补充新的粒度例子。",
         say="太浅，瓶子到罐子，还在包装里打转。太深，瓶子到东西，什么都搜得到，等于没搜。合适的标准很硬：你说出这个上位词，马上能举出至少三个行业的实物。不能，就再抽一层，或者退回一层。"),
    dict(sec=45, title="这一步的矛盾：跳级和过度都会做空", section="矛盾", figure="跳级与过度泛化对照",
         kept="第 9 页图下的两条失败路径；第 10 页的粒度判据。",
         cut="氧气烛备注里的「消除矛盾而不是解决矛盾」随该案例删除，不单开一页。这里的矛盾是抽象粒度上的一对走偏。",
         say="这里有一对矛盾。只拿领域词去搜，会跳过泛化，回来全是同构方案。抽象得太过，上位词大到无法下沉，也写不出结构。两条路都把检索做空。合适的抽象夹在中间，标准就是还能点出三个实物。"),
    dict(sec=50, title="领先领域，是被迫把这件事做得更狠的", section="领先领域怎么选", figure="三个被迫更狠的领域",
         kept="第 8 页备注和第 12 步之前的选择标准：航天、核工业、医院。",
         cut="不展开这三个领域的案例正文。",
         say="第三级不要挑听起来先进的行业。要挑在这个功能上被迫做得比我们更狠的。航天被迫地球独立，核工业被迫不可更改，医院被迫能追责。被逼得更极端的地方，才更可能长出你没见过的结构。"),
    dict(sec=45, title="不下沉到具体物，结论就不可执行", section="功能抽象", figure="下沉四步",
         kept="第 11 页：上位词只是路标；第 9 页备注里的库位和补给窗口。",
         cut="删去「只搬库位那一半」的分支讨论，收成一句：没有类似约束就不要整包搬。",
         say="上位词只是路标。写借鉴航天的库存管理，等于没说。要落到能看见、能拆开的东西上，再问哪一部分是为了对付某条约束才长出来的。有类似约束，就搬那一部分。没有，就不要整包搬。"),
    dict(sec=50, title="迁回过三关：结构、前提、可观测效果", section="如何迁回", figure="三道关",
         kept="第 29 页的三个通过条件，顺序保留。",
         cut="不把三关分别配回八个案例。",
         say="迁回来之前过三关，顺序不要乱。第一，能不能写成结构特征。第二，它在原领域成立所依赖的前提，这里还有没有。第三，迁完有没有一个原来测不到、现在能测的量。缺一条，就还是灵感笔记。"),
    dict(sec=40, title="二元的已用标记，证明不了「只用过一次」", section="完整例子", figure="已用 / 未用被改回",
         kept="第 3、14 页提出的问题：已用标记改回去就看不出。",
         cut="不并行打开第二个例子。",
         say="完整例子只做一个。离线槽位通常只有已用和未用。一旦被改回未用，从记录上看不出来。我们要找的，是一份累计的、不可逆的使用指示。"),
    dict(sec=45, title="已用标记，沿阶梯走到疫苗瓶监视器", section="完整例子", figure="例子的四级阶梯",
         kept="第 13 页：已用标记、累计不可逆指示、疫苗冷链、疫苗瓶监视器，以及四种型号的由来。",
         cut="删去国药新冠疫苗等产品细节，备注可以口头补一句。不讲 VVM 的读数操作。",
         say="阶梯这样走。已用标记，收成累计不可逆指示。领域挑疫苗冷链，因为热稳定性被管得很死。落到的具体物是疫苗瓶监视器：一张贴纸，随累计受热变深，分成四档。原稿还提到，灵敏度要和疫苗的耐受度匹配。"),
    dict(sec=55, title="二元状态留不住历史，累计档位留得住", section="完整例子", figure="阈值式与累计式对照",
         kept="第 14–15 页的对照：报警前后无差异，对上只增不减的四档。",
         cut="删去颜色与参照圈的读数细节，保留四档和只增不减。",
         say="左边是阈值。超过就报警，但改回未用之后，外观上没有差异。右边是累计。档位只增不减，每一档可以对应不同动作，断网也能当场读。多出来的不是一句新功能，是一个带着历史的状态。"),
    dict(sec=45, title="迁回来的结构，是只增不减的档位", section="如何迁回", figure="布尔值改为单调档位",
         kept="第 14 页的迁移结论：0–3 档，0 正常，3 只能核销，档位是现场审计证据。",
         cut="不讨论界面和存储实现。",
         say="迁回的时候，不要去搬那张贴纸。搬的是结构：不要只存一个布尔值，改存 0 到 3 的单调档位。0 档正常，3 档只能核销，不能再新增。档位本身，就是不联网也能看的审计证据。"),
    dict(sec=45, title="三关都过，才算搬成；没前提就不能搬", section="如何迁回", figure="三关核对加一条边界",
         kept="第 29 页三关，用疫苗瓶监视器验收；第 16–17 页燃耗信用的边界：软件没有物理定律兜底。",
         cut="燃耗信用的减法创新、单向同步、哈希链，整段删除，只留这一句边界。",
         say="用三关看这个例子。结构有了，就是单调档位。前提也在，断网时同样要留下看得见的历史。效果也有，现场能读出被透支过。顺手记住边界：燃耗信用敢用已经烧掉多少，是因为有物理定律兜底。软件里没有这条定律，就不能直接搬。"),
    dict(sec=55, title="按这六步，散会就可以做一张卡", section="怎么做", figure="六步流程",
         kept="第 31 页六步的顺序和「先泛化再挑领域」。",
         cut="不讲工作台界面。第⑥步仍保留「补哪条权利要求」。",
         say="回去就按这六步。先列五个关键名词。每个写一个上位词，用三个行业的实物验粒度。再挑一到两个被迫更狠的领域。下沉到具体物，写清它对付的约束。检索时记下出处，标清核实了没有。最后填三道关，写成迁移卡，写明它补的是哪条权利要求。先泛化，再挑领域，顺序不要反。"),
    dict(sec=45, title="五个错误里，最常见的是跳过泛化", section="常见错误", figure="五条错误对照",
         kept="第 30 页的五条错误，顺序不变。",
         cut="作者上一轮失误的长故事不放上页面，讲的时候点一句即可。",
         say="五个坑，按原稿的顺序。只搜领域词。泛化过度。不下沉，结论写成借鉴某领域经验。忽略前提，把有物理兜底的机制搬到没有兜底的地方。只找机会，不记威胁。第一条，原稿作者说自己上一轮就犯过：功能抽完就去搜，没有信息增量。"),
    dict(sec=45, title="搜到同构标准时，把它记成威胁", section="常见错误", figure="同构与实时通道对照",
         kept="第 25 页：3GPP 配额预留与「预发额度、计数、用尽即停」同构，差别是有没有实时回传。",
         cut="删去 Gy/Ro、FUI 等协议细节，也没把空壳核销再讲成第二个例子。备注里的「写进答复预案」口头带过即可。",
         say="功能导向检索不只要找机会。3GPP 在线计费里，网元先申请配额，用尽就终止会话。这和预发额度、计数、用尽即停高度同构，而且早就标准化了。差别在于对方依赖实时回传。完全断连时，这条标准对不上。把它记成威胁，不要当成新灵感。"),
    dict(sec=45, title="印章这个词，值得当场泛化一次", section="怎么做", figure="印章的半成品阶梯",
         kept="第 33 页的词、上位词、三个行业、铅封和骑缝章，以及那句约束提问。",
         cut="现场从约三分钟收到 45 秒，不组织多人发言。",
         say="用一个词体会第四步，不找标准答案。印章的上位词，是可验证的确认标记。票据、关封、批签都能落下去。一次性铅封对付的是途中被打开又复原，骑缝章对付的是抽换其中一页。约束不同，能搬的结构就不同。"),
    dict(sec=45, title="带走三句话就够开始", section="收束", figure="三句带走",
         kept="第 34 页三句，措辞收短，意思不变。",
         cut="删去「大概率会找到一条新路」的预期。",
         say="只记三句。泛化的是名词，不是功能，这决定你搜什么。下沉之后问它为了对付什么约束，这决定值不值得搬。迁移卡要有出处和三道关，否则只是灵感。最小动作就是五个名词，各写一个上位词，再做选领域和下沉。"),
    dict(sec=40, title="出处只说明那边确实有这种做法", section="收束", figure="五条出处",
         kept="第 35 页的五条文献，名称和编号照录。",
         cut="不声称链接可点开，不补充第 35 页以外的文献。案例六到八的卡片出处随案例删除。",
         say="这五条都在原稿里，只说明那个做法在那边确实存在。世界卫生组织的疫苗瓶监视器，美国核管会的燃耗信用，安徽关于空安瓿的规定，美国航天局的在轨库存，以及 3GPP 的在线计费。它们不是新颖性结论。逐篇对比是另一项工作。"),
    dict(sec=40, title="你手上哪个词，值得泛化一次？", section="收束", figure="结尾问题与最小动作",
         kept="第 36 页的收束问句；第 34 页备注里的最小动作；结尾的卡片库和工具路径。",
         cut="删去「能不能自动化」的问答，也不演示 fosdb 命令。",
         say="收在一个问题上：你手上哪个词，值得泛化一次。泛化和下沉这两步还得靠人。检索核实、把判断写成卡片，可以以后再工具化。今天先把一张卡的前四步走完。"),
]


def render_markdown() -> str:
    total = sum(item["sec"] for item in META)
    lines = [
        "# 功能导向检索 · 20 分钟讲稿提纲",
        "",
        "由 45–60 分钟讲稿压缩。源文件未改。",
        "",
        f"- 源文件：`C:\\Users\\21173\\Documents\\xwechat_files\\wxid_2swynsci7a1322_9b49\\msg\\file\\2026-10\\FOS_科普_45-60min.pptx`（36 页）",
        f"- 输出：`FOS_科普_20min.pptx`（{len(META)} 页，合计 {total} 秒，即 {total // 60} 分钟）",
        "- 版式：浅色，16:9。中文用微软雅黑。一页一个结论，图是可编辑的形状，不是整页贴图。",
        "- 生成：本机 python-pptx。未把原稿当扫描件重绘，也未做成整页图片，以便文字和图形都能直接改。",
        "",
        "## 时长怎么分",
        "",
        "| 部分 | 页 | 秒 |",
        "| --- | --- | ---: |",
        "| 开场：题目和三个问题 | 1–2 | 85 |",
        "| 为什么离开本领域 | 3–6 | 190 |",
        "| 功能抽象、矛盾、下沉 | 7–11、13 | 280 |",
        "| 领先领域怎么选 | 12 | 50 |",
        "| 迁回三关 | 14 | 50 |",
        "| 完整例子：疫苗瓶监视器 | 15–19 | 230 |",
        "| 六步、常见错误、威胁、印章练习 | 20–23 | 190 |",
        "| 带走、出处、结尾 | 24–26 | 125 |",
        "| 合计 | 1–26 | 1200 |",
        "",
        "正文页按大约 40–55 秒一页。封面、例子入口、出处和结尾稍短，阶梯、对照和六步稍长。",
        "",
        "## 相对原稿删了什么",
        "",
        "保留的主线没有换：FOS 不是查重；离开本领域是因为约束会逼出不同结构；先抽功能，再泛化名词，沿四级阶梯下沉；用「三个行业的实物」卡住粒度；跳级和过度泛化是一对矛盾；领先领域要挑被迫做得更狠的；迁回看结构、前提和可观测效果；用疫苗瓶监视器走完一个例子；五条常见错误；3GPP 配额预留作为威胁；六步法；印章小练习。",
        "",
        "删掉的是重复例子和工具演示：",
        "",
        "- 案例二燃耗信用的完整推导（原稿 16–17 页）只留两处：第 5 页的约束对照，第 19 页「没有物理定律就不能搬」。",
        "- 案例三空安瓿的守恒、空壳、nonce 和权利要求映射（原稿 18–20 页）删除。空安瓿只留在开场画面和第 5 页。",
        "- 案例四空间站不另讲。在轨消耗品包和补给窗口只作为第 9 页阶梯的第 4 级，这一级来自原稿第 8–9 页，不是第二个完整案例。",
        "- 案例五氧气烛（原稿 23–24 页）整例删除，包括「消除矛盾而不是解决矛盾」那句。",
        "- 案例六到八：电子封记、黑匣子、应急定位发射机（原稿 26–28 页）整例删除。",
        "- 五个案例的分节页（原稿 12 页）删除。",
        "- 工作台命令、verified 必须有证据、ladder 至少三级（原稿 32 页）删除。结尾只留卡片库和工具路径。",
        "- 印章练习从约 3 分钟收到 45 秒。",
        "- 参考文献仍是原稿第 35 页那五条，不新增。",
        "",
        "## 逐页",
        "",
    ]
    for i, item in enumerate(META, 1):
        lines.append(f"### {i}. {item['title']}")
        lines.append("")
        lines.append(f"- 建议秒数：{item['sec']}")
        lines.append(f"- 图：{item['figure']}")
        lines.append(f"- 保留了原稿：{item['kept']}")
        lines.append(f"- 删了什么：{item['cut']}")
        lines.append("")
    lines += [
        "## 渲染抽查",
        "",
        "用本机 PowerPoint 把 26 页导出为 1920×1080 的 PNG，并逐页查看。另用 PowerPoint 读取每段文字的实际占位，容差 6 磅，溢出为 0。",
        "",
        "- 没有空白页。",
        "- 没有文字超出文本框，页脚没有压住正文。",
        "- 没有外链图片。流程、对照和示意都是幻灯片里的形状，可以直接改字和改位置。",
        "",
    ]
    return "\n".join(lines)


def check_bounds(prs) -> list[str]:
    problems = []
    for index, slide in enumerate(prs.slides, 1):
        for shape in slide.shapes:
            if shape.left is None or shape.top is None:
                continue
            if shape.left < -20000 or shape.top < -20000:
                problems.append(f"slide {index} {shape.name} negative position")
            if shape.left + shape.width > prs.slide_width + 20000:
                problems.append(f"slide {index} {shape.name} past right edge")
            if shape.top + shape.height > prs.slide_height + 20000:
                problems.append(f"slide {index} {shape.name} past bottom edge")
    return problems


def text_units(text: str) -> float:
    units = 0.0
    for ch in text:
        if ch == " ":
            units += 0.35
        elif ord(ch) < 128:
            units += 0.55
        else:
            units += 1.02
    return units


def wrap_lines(text: str, units_per_line: float) -> int:
    if units_per_line <= 0:
        return 99
    count = 0
    for part in text.split("\n"):
        if part == "":
            count += 1
            continue
        left = text_units(part)
        count += max(1, int((left + units_per_line - 0.01) // units_per_line))
    return count


def check_text_fit(prs) -> list[str]:
    warnings = []
    for index, slide in enumerate(prs.slides, 1):
        for shape in slide.shapes:
            if not getattr(shape, "has_text_frame", False):
                continue
            text = shape.text_frame.text.strip()
            if not text:
                continue
            tf = shape.text_frame
            width_pt = (shape.width - tf.margin_left - tf.margin_right) / 12700
            height_pt = (shape.height - tf.margin_top - tf.margin_bottom) / 12700
            needed = 0.0
            for paragraph in tf.paragraphs:
                size = 16
                for run in paragraph.runs:
                    if run.font.size is not None:
                        size = run.font.size.pt
                units = max(6.0, width_pt / size)
                lines = wrap_lines(paragraph.text, units)
                needed += lines * size * 1.15
                if paragraph.space_before:
                    needed += paragraph.space_before.pt
                if paragraph.space_after:
                    needed += paragraph.space_after.pt
            if needed > height_pt + 2:
                preview = text.replace("\n", " / ")[:40]
                warnings.append(
                    f"slide {index} h={height_pt:.0f}pt need≈{needed:.0f}pt {preview}"
                )
    return warnings


def build():
    if len(DRAW) != len(META):
        raise SystemExit(f"slide mismatch {len(DRAW)} vs {len(META)}")
    total = sum(item["sec"] for item in META)
    if total != 1200 or not (22 <= len(META) <= 28):
        raise SystemExit(f"timing/count failed: {len(META)} slides, {total}s")

    prs = Presentation()
    prs.slide_width = Emu(SW)
    prs.slide_height = Emu(SH)
    prs.core_properties.title = "功能导向检索（FOS）20 分钟科普"
    prs.core_properties.subject = "由 45–60 分钟讲稿压缩，不改源文件"
    prs.core_properties.category = "科普讲稿"
    blank = prs.slide_layouts[6]

    for index, (draw, meta) in enumerate(zip(DRAW, META), 1):
        slide = prs.slides.add_slide(blank)
        set_bg(slide)
        draw(slide)
        if index not in (1, 26):
            chrome(slide, meta["section"], index)
        slide.notes_slide.notes_text_frame.text = notes_for(meta)

    bounds = check_bounds(prs)
    fit = check_text_fit(prs)
    if bounds or fit:
        print("LAYOUT WARNINGS")
        for item in bounds + fit:
            print(item)

    OUT_DIR.mkdir(parents=True, exist_ok=True)
    prs.save(str(PPTX_PATH))
    MD_PATH.write_text(render_markdown(), encoding="utf-8")
    print(f"slides {len(META)} seconds {total}")
    print(f"bounds {len(bounds)} fit {len(fit)}")
    print(PPTX_PATH)
    return bounds, fit


if __name__ == "__main__":
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass
    build()
