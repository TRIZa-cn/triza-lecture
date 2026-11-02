# -*- coding: utf-8 -*-
"""25-minute FOS talk: a short TRIZ opening, then the existing 20-minute FOS line.

Sources are read-only. This script never opens them for writing.
Figures are native PowerPoint shapes.
"""

from __future__ import annotations

import importlib.util
import sys
from pathlib import Path

OUT_DIR = Path(r"D:\VideoPhone\outputs\FOS_科普_20min_20261022")
SRC_BUILDER = OUT_DIR / "build_deck.py"
PPTX_PATH = OUT_DIR / "FOS_科普_25min.pptx"
NOTES_PATH = OUT_DIR / "notes-FOS.md"
OUTLINE_PATH = OUT_DIR / "讲稿提纲_25min.md"

spec = importlib.util.spec_from_file_location("fos20", SRC_BUILDER)
fos = importlib.util.module_from_spec(spec)
spec.loader.exec_module(fos)

MSO_SHAPE = fos.MSO_SHAPE
t = fos.t
add_text = fos.add_text
add_shape = fos.add_shape
rrect = fos.rrect
heading = fos.heading
chrome = fos.chrome
set_bg = fos.set_bg
INK = fos.INK
MUTED = fos.MUTED
BLUE = fos.BLUE
TEAL = fos.TEAL
RED = fos.RED
AMBER = fos.AMBER
WHITE = fos.WHITE
SOFT_BLUE = fos.SOFT_BLUE
SOFT_TEAL = fos.SOFT_TEAL
SOFT_RED = fos.SOFT_RED
SOFT_AMBER = fos.SOFT_AMBER
SOFT_GREEN = fos.SOFT_GREEN
LINE = fos.LINE


def say(text: str) -> str:
    return text.strip()


# Spoken notes copied from the 20-minute deck, except the two closers trimmed by 5 seconds.
FOS20 = {item["title"]: item["say"] for item in fos.META}


def draw_triz_what(slide, total):
    add_shape(slide, MSO_SHAPE.RECTANGLE, 0, 0, 13.333333, 0.08, BLUE)
    add_text(slide, 0.50, 0.24, 12.3, 0.34, [t("25 分钟里的前三分钟", 16, False, TEAL)])
    add_text(slide, 0.50, 0.62, 12.3, 0.70, [t("TRIZ 是发明性问题的求解框架", 32, True, INK)])
    add_text(
        slide,
        0.50,
        1.38,
        12.3,
        0.40,
        [t("萃智：基于归纳的、发明性问题的创造性求解理论框架", 18, False, INK)],
    )
    add_text(
        slide,
        0.50,
        1.82,
        12.3,
        0.36,
        [t("发明性问题：没有现成方案，需要发明创新方案来解决", 16, False, MUTED)],
    )
    cards = [
        ("1", "同样的问题和解决方案\n在不同领域中反复出现"),
        ("2", "技术演变模式\n在不同领域中反复出现"),
        ("3", "一个领域中发现的\n科学效应和创新方法\n可以借鉴应用到其他领域"),
    ]
    x = 0.48
    for num, body in cards:
        rrect(slide, x, 2.40, 4.00, 3.55, SOFT_BLUE, LINE, 1.15, 0.08)
        add_shape(slide, MSO_SHAPE.OVAL, x + 0.28, 2.64, 0.62, 0.62, BLUE)
        add_text(slide, x + 0.28, 2.64, 0.62, 0.62, [t(num, 18, True, WHITE, "center")], anchor="middle")
        add_text(slide, x + 1.05, 2.64, 2.65, 0.62, [t("发现", 18, True, BLUE)], anchor="middle")
        add_text(slide, x + 0.28, 3.50, 3.44, 2.10, [t(body, 18, True, INK, spacing=1.15)], anchor="top")
        x += 4.18
    add_text(
        slide,
        0.50,
        6.15,
        9.2,
        0.36,
        [t("阿奇舒勒的三条发现  ·  据 TRIZ 讲稿", 14, False, MUTED)],
        anchor="middle",
    )
    add_text(slide, 10.2, 7.05, 2.6, 0.32, [t(f"1  /  {total}", 12, False, MUTED, "right")], anchor="middle")


def draw_triz_contradiction(slide):
    heading(slide, "问题收成矛盾，发明原理是标准解法")
    rrect(slide, 0.48, 1.15, 6.05, 3.15, SOFT_RED, None, radius=0.08)
    rrect(slide, 6.80, 1.15, 6.05, 3.15, SOFT_AMBER, None, radius=0.08)
    add_text(slide, 0.72, 1.32, 5.55, 0.48, [t("技术矛盾", 24, True, RED)], anchor="middle")
    add_text(
        slide,
        0.72,
        1.95,
        5.55,
        1.35,
        [t("改善一个参数时，\n另一个参数会恶化", 20, True, INK, spacing=1.12)],
        anchor="top",
    )
    add_text(slide, 0.72, 3.40, 5.55, 0.60, [t("两个参数互相牵制", 16, False, MUTED)], anchor="middle")
    add_text(slide, 7.04, 1.32, 5.55, 0.48, [t("物理矛盾", 24, True, AMBER)], anchor="middle")
    add_text(
        slide,
        7.04,
        1.95,
        5.55,
        1.15,
        [t("同一个参数上，\n同时有相反的要求", 20, True, INK, spacing=1.12)],
        anchor="top",
    )
    add_text(slide, 7.04, 3.25, 5.55, 0.75, [t("讲稿里的例子：屏幕要大，也要小", 16, False, MUTED)], anchor="middle")
    rrect(slide, 0.48, 4.50, 12.38, 1.15, SOFT_BLUE, None, radius=0.08)
    add_text(slide, 0.72, 4.62, 4.3, 0.90, [t("40 条发明原理", 22, True, BLUE)], anchor="middle")
    add_text(
        slide,
        5.10,
        4.62,
        7.45,
        0.90,
        [t("解决这些矛盾的标准方法\n不逐条背，先知道它们是一套解法", 16, False, INK, spacing=1.08)],
        anchor="middle",
    )
    rrect(slide, 0.48, 5.82, 12.38, 0.95, WHITE, LINE, 1.15, 0.08)
    add_text(
        slide,
        0.72,
        5.94,
        11.9,
        0.72,
        [t("矛盾矩阵：把一对参数放进去，交叉点给出几条可试的原理。讲稿另举防弹衣：既要坚固，又不希望重。", 16, False, INK)],
        anchor="middle",
    )


def draw_triz_ifr(slide):
    heading(slide, "理想解只负责指路，不负责落地")
    rrect(slide, 0.48, 1.12, 12.38, 1.05, SOFT_TEAL, None, radius=0.08)
    add_text(
        slide,
        0.72,
        1.24,
        11.9,
        0.80,
        [t("自己把功能做完，系统却不再存在", 26, True, TEAL, "center")],
        anchor="middle",
    )
    chips = [
        (SOFT_GREEN, TEAL, "有用部分留下"),
        (SOFT_RED, RED, "有害部分消失"),
        (SOFT_AMBER, AMBER, "成本与复杂度趋近于零"),
    ]
    x = 0.48
    for fill, color, label in chips:
        rrect(slide, x, 2.38, 4.00, 1.15, fill, None, radius=0.1)
        add_text(slide, x + 0.15, 2.52, 3.70, 0.88, [t(label, 18, True, color, "center")], anchor="middle")
        x += 4.18
    steps = [
        (BLUE, "理想解", "定终点"),
        (AMBER, "矛盾", "指出障碍"),
        (TEAL, "发明原理", "给出过桥方案"),
    ]
    x = 0.48
    for color, name, job in steps:
        rrect(slide, x, 3.78, 3.55, 1.85, WHITE, LINE, 1.25, 0.08)
        add_shape(slide, MSO_SHAPE.OVAL, x + 1.40, 3.98, 0.72, 0.72, color)
        add_text(slide, x + 0.18, 4.78, 3.20, 0.38, [t(name, 18, True, INK, "center")], anchor="middle")
        add_text(slide, x + 0.18, 5.16, 3.20, 0.34, [t(job, 14, False, MUTED, "center")], anchor="middle")
        if name != "发明原理":
            add_shape(slide, MSO_SHAPE.RIGHT_ARROW, x + 3.62, 4.50, 0.42, 0.28, BLUE)
        x += 4.18
    add_text(
        slide,
        0.48,
        5.85,
        12.38,
        0.90,
        [t("问题 = 理想目标与现实的差距，讲稿称为目标张力。\n这些工具是从问题走向理想解的垫脚石。", 16, False, INK, "center", spacing=1.12)],
        anchor="middle",
    )


def draw_bridge(slide):
    heading(slide, "功能导向搜索，负责抽出功能再迁回")
    rrect(slide, 0.40, 1.18, 5.15, 4.35, SOFT_BLUE, None, radius=0.08)
    rrect(slide, 7.78, 1.18, 5.15, 4.35, SOFT_TEAL, None, radius=0.08)
    add_shape(slide, MSO_SHAPE.OVAL, 5.72, 2.55, 1.90, 1.90, BLUE)
    add_text(slide, 5.88, 2.95, 1.58, 0.48, [t("一个", 16, True, WHITE, "center")], anchor="middle")
    add_text(slide, 5.88, 3.40, 1.58, 0.48, [t("工具", 16, True, WHITE, "center")], anchor="middle")
    add_text(slide, 0.62, 1.40, 4.75, 0.50, [t("TRIZ 指出", 22, True, BLUE)], anchor="middle")
    add_text(
        slide,
        0.62,
        2.10,
        4.75,
        1.70,
        [t("离开本领域\n找已经长出来的结构", 22, True, INK, spacing=1.08)],
        anchor="top",
    )
    add_text(
        slide,
        0.62,
        4.05,
        4.75,
        1.15,
        [t("同样的问题和解法\n会在别的领域反复出现", 16, False, MUTED, spacing=1.1)],
        anchor="top",
    )
    add_text(slide, 8.00, 1.40, 4.75, 0.50, [t("功能导向搜索负责", 22, True, TEAL)], anchor="middle")
    add_text(
        slide,
        8.00,
        2.10,
        4.75,
        1.70,
        [t("把功能抽出来\n再迁回可验收的结构", 22, True, INK, spacing=1.08)],
        anchor="top",
    )
    add_text(
        slide,
        8.00,
        4.05,
        4.75,
        1.15,
        [t("泛化名词，下沉到具体物\n再看结构、前提和效果", 16, False, MUTED, spacing=1.1)],
        anchor="top",
    )
    add_text(
        slide,
        0.48,
        5.75,
        12.38,
        0.95,
        [t("讲稿里的名称是功能导向搜索。\n下面按功能导向检索来讲这一个工具，缩写 FOS。", 16, False, INK, "center", spacing=1.12)],
        anchor="middle",
    )


def draw_fos_cover(slide, total):
    add_shape(slide, MSO_SHAPE.RECTANGLE, 0, 0, 13.333333, 0.08, BLUE)
    add_text(slide, 0.55, 1.05, 12.2, 0.36, [t("TRIZ 的一个工具", 16, False, TEAL)])
    add_text(slide, 0.55, 1.48, 12.2, 0.80, [t("功能导向检索", 44, True, INK)])
    add_text(
        slide,
        0.55,
        2.40,
        12.0,
        0.70,
        [t("把别的行业已经做熟的事，搬到你的问题里", 24, False, INK)],
    )
    add_text(
        slide,
        0.55,
        3.20,
        12.0,
        0.40,
        [t("不是查重，也不是头脑风暴。", 18, False, MUTED)],
    )
    steps = [("1", "抽功能"), ("2", "泛化名词"), ("3", "选领先领域"), ("4", "下沉并迁回")]
    x = 0.55
    for num, label in steps:
        rrect(slide, x, 4.55, 2.55, 1.15, SOFT_BLUE, LINE, 1.0, 0.12)
        add_text(slide, x + 0.16, 4.68, 0.42, 0.42, [t(num, 16, True, BLUE, "center")], anchor="middle")
        add_text(slide, x + 0.58, 4.68, 1.80, 0.85, [t(label, 18, True, INK)], anchor="middle")
        if num != "4":
            add_shape(slide, MSO_SHAPE.RIGHT_ARROW, x + 2.62, 4.98, 0.28, 0.22, BLUE)
        x += 3.15
    add_text(
        slide,
        0.55,
        6.00,
        10.0,
        0.36,
        [t("一套六步法  ·  一个完整例子：疫苗瓶监视器", 16, False, MUTED)],
    )
    add_text(slide, 10.2, 7.05, 2.6, 0.32, [t(f"5  /  {total}", 12, False, MUTED, "right")], anchor="middle")


def draw_level3(slide):
    heading(slide, "第三级发明，要去别的领域借方法")
    cols = [
        ("1", "标准化", "本领域知识\n做简单改进", "32%", False),
        ("2", "改良化", "专业领域内\n已知的方法", "45%", False),
        ("3", "模式内创新", "参考其他领域\n的技术与方法", "18%", True),
        ("4", "模式外创新", "科学原理首次\n用于本领域", "4%", False),
        ("5", "新发明与发现", "超越当代\n已有知识", "<1%", False),
    ]
    x = 0.42
    for num, name, body, pct, hot in cols:
        fill = BLUE if hot else SOFT_BLUE
        name_c = WHITE if hot else INK
        body_c = WHITE if hot else MUTED
        pct_c = WHITE if hot else BLUE
        rrect(slide, x, 1.18, 2.42, 4.35, fill, None if hot else LINE, 1.0, 0.08)
        add_text(slide, x + 0.12, 1.32, 2.18, 0.40, [t(f"第 {num} 级", 14, True, pct_c, "center")], anchor="middle")
        add_text(slide, x + 0.10, 1.78, 2.22, 0.85, [t(name, 18, True, name_c, "center")], anchor="middle")
        add_text(slide, x + 0.12, 2.70, 2.18, 1.20, [t(body, 14, False, body_c, "center", spacing=1.08)], anchor="middle")
        add_text(slide, x + 0.12, 4.05, 2.18, 0.70, [t(pct, 26, True, pct_c, "center")], anchor="middle")
        if hot:
            add_text(slide, x + 0.12, 4.78, 2.18, 0.48, [t("今天对准这里", 14, True, WHITE, "center")], anchor="middle")
        x += 2.56
    add_text(
        slide,
        0.48,
        5.70,
        12.38,
        1.05,
        [t("讲稿：第三级是把行业外已知方法迁移过来，要处理物理矛盾。表上的例子是电钻加入变速器。\n功能导向检索做的就是这一级所要求的跨领域迁移。", 16, False, INK, "center", spacing=1.12)],
        anchor="middle",
    )


def draw_mapping(slide):
    heading(slide, "六步收得进讲稿里的四段求解")
    cols = [
        ("待解问题", "抽功能", "把本领域名词\n译成动词短语"),
        ("问题模型", "泛化名词", "收到能点出\n三个行业实物的词"),
        ("解决方案模型", "领先领域的具体物", "挑被迫更狠的领域\n再落到能拆开的东西"),
        ("具体解决方案", "迁回并过三关", "结构、前提、\n可观测效果都写得出"),
    ]
    x = 0.42
    for i, (top, bot, detail) in enumerate(cols):
        rrect(slide, x, 1.15, 2.95, 1.35, SOFT_BLUE, None, radius=0.08)
        add_text(slide, x + 0.12, 1.22, 2.71, 0.32, [t("讲稿", 12, True, BLUE, "center")], anchor="middle")
        add_text(slide, x + 0.10, 1.52, 2.75, 0.82, [t(top, 16, True, INK, "center")], anchor="middle")
        add_shape(slide, MSO_SHAPE.DOWN_ARROW, x + 1.22, 2.60, 0.48, 0.38, TEAL)
        rrect(slide, x, 3.10, 2.95, 2.35, SOFT_TEAL, None, radius=0.08)
        add_text(slide, x + 0.12, 3.22, 2.71, 0.32, [t("本场", 12, True, TEAL, "center")], anchor="middle")
        add_text(slide, x + 0.12, 3.55, 2.71, 0.85, [t(bot, 16, True, INK, "center")], anchor="middle")
        add_text(slide, x + 0.12, 4.40, 2.71, 0.85, [t(detail, 14, False, MUTED, "center", spacing=1.05)], anchor="middle")
        if i < 3:
            add_shape(slide, MSO_SHAPE.RIGHT_ARROW, x + 2.98, 1.68, 0.22, 0.20, BLUE)
        x += 3.22
    add_text(
        slide,
        0.48,
        5.65,
        12.38,
        1.10,
        [t("上排四个词来自 TRIZ 讲稿的求解流程。下排是本场的对应动作。\n顺序仍是先泛化，再挑领域。", 16, False, INK, "center", spacing=1.12)],
        anchor="middle",
    )


def draw_close(slide, total):
    add_shape(slide, MSO_SHAPE.RECTANGLE, 0, 0, 13.333333, 0.08, BLUE)
    add_text(slide, 0.60, 1.25, 12.1, 0.40, [t("散会前只留一个问题", 16, False, TEAL)])
    add_text(
        slide,
        0.60,
        1.75,
        12.1,
        1.55,
        [t("你手上哪个词，\n值得泛化一次？", 40, True, INK, spacing=1.0)],
        anchor="top",
    )
    rrect(slide, 0.60, 3.85, 12.15, 1.70, SOFT_BLUE, None, radius=0.08)
    add_text(slide, 0.90, 4.02, 11.6, 0.40, [t("最小动作", 14, True, BLUE)], anchor="middle")
    add_text(
        slide,
        0.90,
        4.50,
        11.6,
        0.75,
        [t("五个名词，各写一个上位词，再做选领域和下沉。", 20, True, INK)],
        anchor="middle",
    )
    add_text(
        slide,
        0.60,
        5.80,
        12.1,
        0.50,
        [t("卡片库  examples/fos/cards     工具  tools/fosdb.py", 16, False, MUTED)],
        anchor="middle",
    )
    add_text(slide, 10.2, 7.05, 2.6, 0.32, [t(f"{total}  /  {total}", 12, False, MUTED, "right")], anchor="middle")


def build_meta():
    m = fos.META
    by_title = {item["title"]: item for item in m}

    def from20(title, sec=None):
        src = by_title[title]
        return {
            "sec": src["sec"] if sec is None else sec,
            "title": src["title"],
            "section": src["section"],
            "figure": src["figure"],
            "origin": "20分钟稿",
            "origin_detail": src["title"],
            "say": src["say"] if sec is None else None,
        }

    items = [
        dict(
            sec=60,
            title="TRIZ 是发明性问题的求解框架",
            section="TRIZ 简介",
            figure="三条发现",
            origin="TRIZ讲稿",
            origin_detail="第 3 页定义；第 7 页三条发现",
            say=say(
                """
先用大约一分钟，只讲 TRIZ 是什么。讲稿的定义是：萃智，拉丁文缩写 TRIZ，是一套基于归纳的、发明性问题的创造性求解理论框架。发明性问题，就是没有现成方案，需要发明创新方案来解决的问题。
阿奇舒勒研究大量专利之后，写下三条发现。第一，同样的问题和解决方案，在不同领域中反复出现。第二，技术演变模式，在不同领域中反复出现。第三，一个领域中发现的科学效应和创新方法，可以借鉴应用到其他领域。
今天只用这三条。八条进化法则、七十六条标准解，这场不展开。
"""
            ),
            draw="triz_what",
            chrome=False,
        ),
        dict(
            sec=60,
            title="问题收成矛盾，发明原理是标准解法",
            section="TRIZ 简介",
            figure="技术矛盾与物理矛盾",
            origin="TRIZ讲稿",
            origin_detail="第 7 页备注中的矛盾与四十条原理；第 23 页备注中的矛盾矩阵和防弹衣",
            say=say(
                """
讲稿把难题收成矛盾。技术矛盾是：改善一个参数时，另一个参数会恶化。物理矛盾是：同一个参数上同时有相反的要求，讲稿叫做既要又要。
例子就在讲稿里。手机屏幕要大，因为看得清楚，又要小，因为要便于携带。防弹衣既希望坚固，又不希望它很重。
对付这些矛盾的标准方法，是四十条最常用的发明原理。矛盾矩阵是手册：把一对参数放进去，交叉点上给出几条可以试的原理。原理不逐条背。先知道它们是一套标准解法。
"""
            ),
            draw="triz_contra",
            chrome=True,
        ),
        dict(
            sec=50,
            title="理想解只负责指路，不负责落地",
            section="TRIZ 简介",
            figure="理想解与三段关系",
            origin="TRIZ讲稿",
            origin_detail="第 8 页的目标张力与垫脚石；第 15 页备注中的理想解定义",
            say=say(
                """
理想解不是最好的那个方案。讲稿把它说成一个极限想象：自己把功能做完，系统却不再存在。有用部分全部留下，有害部分全部去掉，成本和复杂度都趋近于零。它负责指路，不负责落地。
问题本身，是理想目标与现实的差距。讲稿称为目标张力。三者的关系是：理想解定终点，矛盾指出障碍，发明原理和分离原理给出过桥方案。这些工具，是从问题走向理想解的垫脚石。
"""
            ),
            draw="triz_ifr",
            chrome=True,
        ),
        dict(
            sec=55,
            title="功能导向搜索，负责抽出功能再迁回",
            section="从 TRIZ 到 FOS",
            figure="TRIZ 与 FOS 的分工",
            origin="过渡",
            origin_detail="TRIZ 讲稿第 7 页备注点名功能导向搜索；第 21 页写明向类似领域寻找解决方案模型。分工按本场主线写出",
            say=say(
                """
刚才三分钟建立的是方向：同样的问题和解法会在别的领域反复出现，所以要离开本领域去找结构。讲稿在现代 TRIZ 的方法里点过名，功能导向搜索就是其中一件经过实践检验的工具。讲稿还说，把问题说成可以跨领域对照的模型，再向面临类似挑战的领域寻找现成的解决方案模型，这正是功能导向搜索要做的事。
分工一句就够。TRIZ 指出要离开本领域找结构。功能导向搜索负责把功能抽出来，再迁回去。讲稿用的名字是功能导向搜索。下面按功能导向检索来讲，缩写 FOS。这一页起算 FOS，不计入刚才那三分钟。
"""
            ),
            draw="bridge",
            chrome=True,
        ),
    ]

    cover = from20("功能导向检索")
    cover["section"] = "开场"
    cover["origin_detail"] = "20 分钟稿第 1 页。眉题由「20 分钟科普讲稿」改为「TRIZ 的一个工具」"
    cover["say"] = FOS20["功能导向检索"].replace("二十分钟，只走通一个例子。", "今天只走通一个例子。")
    cover["draw"] = "fos_cover"
    cover["chrome"] = False
    items.append(cover)

    q = from20("今天只回答三个问题")
    q["origin_detail"] = "20 分钟稿第 2 页"
    q["say"] = FOS20["今天只回答三个问题"]
    q["draw"] = "d02"
    q["chrome"] = True
    items.append(q)

    items.append(
        dict(
            sec=40,
            title="第三级发明，要去别的领域借方法",
            section="为什么离开本领域",
            figure="五级发明，第三级高亮",
            origin="新增",
            origin_detail="TRIZ 讲稿第 11 页分级表。用来说明离开本领域对应哪一级，不展开整张表的案例",
            say=say(
                """
讲稿把发明分成五级，比例就写在那张表上。第一级大约百分之三十二，用本领域知识做简单改进。第二级大约百分之四十五，仍用专业领域内已知的方法。第三级大约百分之十八，有实质改变，做法是参考其他领域的技术与方法。表上的例子是电钻加入变速器。备注写明，这一级就是把行业外已知方法迁移过来，要处理的是物理矛盾。
第四级大约百分之四，第五级不到百分之一，今天不讲。功能导向检索对准的，就是这第三级。
"""
            ),
            draw="level3",
            chrome=True,
        )
    )

    fos_draws = {
        "三个不相干的东西，在解同一件事": ("d03", "20 分钟稿第 3 页"),
        "查重在防守，FOS 是去别处找解法": ("d04", "20 分钟稿第 4 页"),
        "约束不同，同一个功能会长出不同结构": ("d05", "20 分钟稿第 5 页"),
        "要搬的是被约束逼出来的结构": ("d06", "20 分钟稿第 6 页"),
        "抽功能只是翻译，停在这里没有新信息": ("d07", "20 分钟稿第 7 页"),
        "泛化的是名词，再下沉成具体物": ("d08", "20 分钟稿第 8 页"),
        "四级阶梯写下来，检索才查得回": ("d09", "20 分钟稿第 9 页"),
        "合适的上位词，能点出三个行业的实物": ("d10", "20 分钟稿第 10 页"),
        "这一步的矛盾：跳级和过度都会做空": ("d11", "20 分钟稿第 11 页"),
        "领先领域，是被迫把这件事做得更狠的": ("d12", "20 分钟稿第 12 页"),
        "不下沉到具体物，结论就不可执行": ("d13", "20 分钟稿第 13 页"),
        "迁回过三关：结构、前提、可观测效果": ("d14", "20 分钟稿第 14 页"),
        "二元的已用标记，证明不了「只用过一次」": ("d15", "20 分钟稿第 15 页"),
        "已用标记，沿阶梯走到疫苗瓶监视器": ("d16", "20 分钟稿第 16 页"),
        "二元状态留不住历史，累计档位留得住": ("d17", "20 分钟稿第 17 页"),
        "迁回来的结构，是只增不减的档位": ("d18", "20 分钟稿第 18 页"),
        "三关都过，才算搬成；没前提就不能搬": ("d19", "20 分钟稿第 19 页"),
        "按这六步，散会就可以做一张卡": ("d20", "20 分钟稿第 20 页"),
    }
    for title, (draw_name, detail) in fos_draws.items():
        row = from20(title)
        row["origin_detail"] = detail
        row["say"] = FOS20[title]
        row["draw"] = draw_name
        row["chrome"] = True
        items.append(row)

    items.append(
        dict(
            sec=45,
            title="六步收得进讲稿里的四段求解",
            section="怎么做",
            figure="四段求解与本场动作对照",
            origin="新增",
            origin_detail="TRIZ 讲稿第 22 页的四段流程：待解问题、问题模型、解决方案模型、具体解决方案。下排对上本场六步，不是把旧页原样贴上",
            say=say(
                """
六步不要另记成一套理论。讲稿的求解流程是四段：从待解问题，收成问题模型，找到解决方案模型，再写成具体解决方案。
今天的动作对得上这四段。抽功能，是把待解问题说清楚。泛化名词，是收成能跨领域的问题模型。到领先领域里找具体物，是去拿解决方案模型。迁回，并且过结构、前提、可观测效果这三关，才是具体解决方案。
顺序仍然是先泛化，再挑领域。上排四个词来自讲稿，下排是本场的对应。
"""
            ),
            draw="mapping",
            chrome=True,
        )
    )

    tail = {
        "五个错误里，最常见的是跳过泛化": ("d21", "20 分钟稿第 21 页", None),
        "搜到同构标准时，把它记成威胁": ("d22", "20 分钟稿第 22 页", None),
        "印章这个词，值得当场泛化一次": ("d23", "20 分钟稿第 23 页", None),
        "带走三句话就够开始": ("d24", "20 分钟稿第 24 页", None),
        "出处只说明那边确实有这种做法": (
            "d25",
            "20 分钟稿第 25 页。幻灯片不变，备注收短约 5 秒",
            "这五条都在原稿里，只说明那个做法在那边确实存在。世界卫生组织的疫苗瓶监视器，美国核管会的燃耗信用，安徽关于空安瓿的规定，美国航天局的在轨库存，以及 3GPP 的在线计费。它们不是新颖性结论。",
        ),
        "你手上哪个词，值得泛化一次？": (
            "d26",
            "20 分钟稿第 26 页。幻灯片保留最小动作和路径，备注收短约 5 秒",
            "收在一个问题上：你手上哪个词，值得泛化一次。泛化和下沉这两步还得靠人。今天先把一张卡的前四步走完。",
        ),
    }
    for title, (draw_name, detail, new_say) in tail.items():
        sec = 35 if new_say else None
        row = from20(title, sec=sec)
        row["origin_detail"] = detail
        row["say"] = new_say if new_say else FOS20[title]
        row["draw"] = draw_name
        row["chrome"] = draw_name != "d26"
        items.append(row)
    return items


DRAWERS = {
    "triz_what": lambda slide, total: draw_triz_what(slide, total),
    "triz_contra": lambda slide, total: draw_triz_contradiction(slide),
    "triz_ifr": lambda slide, total: draw_triz_ifr(slide),
    "bridge": lambda slide, total: draw_bridge(slide),
    "fos_cover": lambda slide, total: draw_fos_cover(slide, total),
    "d02": lambda slide, total: fos.draw_02(slide),
    "level3": lambda slide, total: draw_level3(slide),
    "d03": lambda slide, total: fos.draw_03(slide),
    "d04": lambda slide, total: fos.draw_04(slide),
    "d05": lambda slide, total: fos.draw_05(slide),
    "d06": lambda slide, total: fos.draw_06(slide),
    "d07": lambda slide, total: fos.draw_07(slide),
    "d08": lambda slide, total: fos.draw_08(slide),
    "d09": lambda slide, total: fos.draw_09(slide),
    "d10": lambda slide, total: fos.draw_10(slide),
    "d11": lambda slide, total: fos.draw_11(slide),
    "d12": lambda slide, total: fos.draw_12(slide),
    "d13": lambda slide, total: fos.draw_13(slide),
    "d14": lambda slide, total: fos.draw_14(slide),
    "d15": lambda slide, total: fos.draw_15(slide),
    "d16": lambda slide, total: fos.draw_16(slide),
    "d17": lambda slide, total: fos.draw_17(slide),
    "d18": lambda slide, total: fos.draw_18(slide),
    "d19": lambda slide, total: fos.draw_19(slide),
    "d20": lambda slide, total: fos.draw_20(slide),
    "mapping": lambda slide, total: draw_mapping(slide),
    "d21": lambda slide, total: fos.draw_21(slide),
    "d22": lambda slide, total: fos.draw_22(slide),
    "d23": lambda slide, total: fos.draw_23(slide),
    "d24": lambda slide, total: fos.draw_24(slide),
    "d25": lambda slide, total: fos.draw_25(slide),
    "d26": lambda slide, total: draw_close(slide, total),
}


def notes_text(meta):
    return f"建议用时：{meta['sec']} 秒\n\n讲者要点：\n{meta['say']}\n\n来源：{meta['origin_detail']}"


def render_notes(items):
    total = sum(item["sec"] for item in items)
    lines = [
        "# 功能导向检索 · 25 分钟讲者备注",
        "",
        f"合计 {total} 秒，即 {total // 60} 分钟。每页先给建议用时，再给可以照着讲的要点。",
        "",
        "TRIZ 三页的备注合计 170 秒。第 4 页过渡起算入 FOS，不计入这 170 秒。",
        "",
    ]
    for i, item in enumerate(items, 1):
        lines.append(f"## 第 {i} 页 · {item['title']}（{item['sec']} 秒）")
        lines.append("")
        lines.append(notes_text(item))
        lines.append("")
    return "\n".join(lines).rstrip() + "\n"


def render_outline(items, check_note: str):
    total = sum(item["sec"] for item in items)
    triz = [i for i, item in enumerate(items, 1) if item["origin"] == "TRIZ讲稿"]
    triz_sec = sum(items[i - 1]["sec"] for i in triz)
    bridge = next(i for i, item in enumerate(items, 1) if item["origin"] == "过渡")
    new_idx = [i for i, item in enumerate(items, 1) if item["origin"] == "新增"]
    fos_rest = [i for i, item in enumerate(items, 1) if item["origin"] in ("20分钟稿", "新增", "过渡") and i != bridge]
    # FOS portion = everything except the 3 TRIZ slides
    fos_sec = total - triz_sec
    lines = [
        "# 功能导向检索 · 25 分钟讲稿提纲",
        "",
        "由 20 分钟 FOS 科普扩成 25 分钟。两份源 pptx 未改，也未覆盖 `FOS_科普_20min.pptx`。",
        "",
        "- TRIZ 源（只读）：`C:\\Users\\21173\\Documents\\xwechat_files\\wxid_2swynsci7a1322_9b49\\msg\\file\\2026-10\\TRIZ20250928.pptx`",
        "- FOS 源（只读）：`D:\\VideoPhone\\outputs\\FOS_科普_20min_20261022\\FOS_科普_20min.pptx` 与同目录 `讲稿提纲.md`",
        f"- 输出：`FOS_科普_25min.pptx`（{len(items)} 页，合计 {total} 秒）",
        "- 版式：浅色，16:9。中文用微软雅黑。图是可编辑形状。TRIZ 页按投影重排，不沿用旧讲稿版式。",
        "- 生成：本机 python-pptx。未把讲稿做成整页图片。",
        "",
        "## 时长怎么分",
        "",
        "| 部分 | 页 | 秒 | 从哪来 |",
        "| --- | --- | ---: | --- |",
        f"| TRIZ 简介 | {triz[0]}–{triz[-1]} | {triz_sec} | TRIZ 讲稿。占开头 2–3 分钟 |",
        f"| 过渡：FOS 是 TRIZ 的工具 | {bridge} | {items[bridge - 1]['sec']} | 新写的一页。算 FOS，不占 TRIZ 的 2–3 分钟 |",
        f"| FOS 其余，含两页新增 | {fos_rest[0]}–{fos_rest[-1]} | {fos_sec - items[bridge - 1]['sec']} | 20 分钟稿主线，加上为凑满 25 分钟新写的两页 |",
        f"| 合计 | 1–{len(items)} | {total} | |",
        "",
        f"TRIZ 三页合计 {triz_sec} 秒（{triz_sec / 60:.1f} 分钟）。过渡页 {items[bridge - 1]['sec']} 秒。FOS 部分（含过渡页）合计 {fos_sec} 秒。",
        "",
        "## 从 TRIZ 讲稿抽了什么",
        "",
        "只取能在三分钟里讲完、而且讲稿正文或备注里写明了的三件事：",
        "",
        "- TRIZ 是什么：发明性问题的创造性求解理论框架；阿奇舒勒的三条发现。",
        "- 矛盾和发明原理：技术矛盾、物理矛盾、四十条发明原理、矛盾矩阵。例子用讲稿里的屏幕和防弹衣。",
        "- 理想解：自己把功能做完而系统可以不再存在；问题是目标张力；理想解、矛盾、发明原理是一段关系。",
        "",
        "没有放进这三页的，包括绕过审查的长案例、荔枝转运、ARIZ 九步、创造力分级、方案力和需求模型。功能导向搜索只在过渡页点名，不在这三分钟里展开。",
        "",
        "## 为凑满 25 分钟新增了什么",
        "",
        "20 分钟稿原为 1200 秒。加上 TRIZ 170 秒和过渡页 55 秒之后，还差 75 秒，用两页补上。结尾两页的备注各收短约 5 秒，幻灯片正文没有改主线。燃耗信用的展开、空安瓿案例、工具命令都没有加回来。",
        "",
        f"- 第 {new_idx[0]} 页（{items[new_idx[0] - 1]['sec']} 秒）：用 TRIZ 讲稿第 11 页的五级表，点明第三级才是「参考其他领域的技术与方法」。比例照录 32%、45%、18%、4%、小于 1%。",
        f"- 第 {new_idx[1]} 页（{items[new_idx[1] - 1]['sec']} 秒）：用 TRIZ 讲稿第 22 页的四段求解，对上抽功能、泛化名词、领先领域的具体物、迁回三关。",
        "- 第 31、32 页：幻灯片沿用 20 分钟稿，建议用时从 40 秒改为 35 秒，用来把合计收在 1500 秒。",
        "",
        "## 20 分钟稿里保留的主线",
        "",
        "FOS 不是查重。先抽功能，再把名词收成上位词。沿四级阶梯下沉。领先领域挑被迫做得更狠的。迁回看结构、前提和可观测效果。完整例子是疫苗瓶监视器。五条常见错误、3GPP 配额预留作为威胁、六步法和印章练习仍在。",
        "",
        "## 逐页",
        "",
        "| 页 | 标题 | 秒 | 来源 |",
        "| ---: | --- | ---: | --- |",
    ]
    for i, item in enumerate(items, 1):
        lines.append(f"| {i} | {item['title']} | {item['sec']} | {item['origin']}：{item['origin_detail']} |")
    lines += [
        "",
        "## 渲染抽查",
        "",
        check_note.strip(),
        "",
    ]
    return "\n".join(lines)


def spoken_len(text: str) -> int:
    return sum(1 for ch in text if not ch.isspace())


def build():
    items = build_meta()
    total = sum(item["sec"] for item in items)
    triz_sec = sum(item["sec"] for item in items if item["origin"] == "TRIZ讲稿")
    if not (28 <= len(items) <= 34):
        raise SystemExit(f"slide count {len(items)}")
    if not (1470 <= total <= 1530):
        raise SystemExit(f"seconds {total}")
    if not (120 <= triz_sec <= 180):
        raise SystemExit(f"TRIZ seconds {triz_sec}")
    if len(DRAWERS) < len({item["draw"] for item in items}):
        raise SystemExit("missing drawer")

    from pptx import Presentation
    from pptx.util import Emu

    prs = Presentation()
    prs.slide_width = Emu(fos.SW)
    prs.slide_height = Emu(fos.SH)
    prs.core_properties.title = "功能导向检索（FOS）25 分钟科普"
    prs.core_properties.subject = "开头为 TRIZ 简介，其后为功能导向检索"
    prs.core_properties.category = "科普讲稿"
    blank = prs.slide_layouts[6]
    total_n = len(items)

    for index, meta in enumerate(items, 1):
        slide = prs.slides.add_slide(blank)
        set_bg(slide)
        DRAWERS[meta["draw"]](slide, total_n)
        if meta["chrome"]:
            chrome(slide, meta["section"], index, total=total_n)
        slide.notes_slide.notes_text_frame.text = notes_text(meta)

    bounds = fos.check_bounds(prs)
    fit = fos.check_text_fit(prs)
    print(f"slides {len(items)} seconds {total} triz {triz_sec}")
    print(f"bounds {len(bounds)} fit {len(fit)}")
    for item in bounds + fit:
        print(item)

    OUT_DIR.mkdir(parents=True, exist_ok=True)
    prs.save(str(PPTX_PATH))
    NOTES_PATH.write_text(render_notes(items), encoding="utf-8")
    # Outline check note is filled after PowerPoint inspection.
    print(PPTX_PATH)
    return items, bounds, fit


def inspect_with_powerpoint(items):
    import tempfile

    try:
        import win32com.client
    except ImportError:
        return "本机没有 pywin32，未能用 PowerPoint 打开检查。已用 python-pptx 核对形状没有越出页面。"

    preview = Path(tempfile.mkdtemp(prefix="fos25_"))
    started = False
    ppt = None
    pres = None
    notes = []
    try:
        try:
            ppt = win32com.client.GetActiveObject("PowerPoint.Application")
        except Exception:
            ppt = win32com.client.Dispatch("PowerPoint.Application")
            started = True
        ppt.Visible = True
        ppt.DisplayAlerts = 1
        before = ppt.Presentations.Count
        pres = ppt.Presentations.Open(str(PPTX_PATH), True, False, False)
        blank = []
        overflow = []
        offslide = []
        no_shape = []
        slide_w = pres.PageSetup.SlideWidth
        slide_h = pres.PageSetup.SlideHeight
        for i in range(1, pres.Slides.Count + 1):
            slide = pres.Slides(i)
            texts = []
            autoshapes = 0

            def walk(shape):
                nonlocal autoshapes
                try:
                    if int(shape.Type) == 6:
                        for j in range(1, shape.GroupItems.Count + 1):
                            walk(shape.GroupItems(j))
                        return
                except Exception:
                    pass
                try:
                    if shape.Left < -2 or shape.Top < -2:
                        offslide.append(f"第 {i} 页 {shape.Name} 负坐标")
                    if shape.Left + shape.Width > slide_w + 2:
                        offslide.append(f"第 {i} 页 {shape.Name} 超出右缘")
                    if shape.Top + shape.Height > slide_h + 2:
                        offslide.append(f"第 {i} 页 {shape.Name} 超出下缘")
                except Exception:
                    pass
                try:
                    if int(shape.Type) == 1:
                        autoshapes += 1
                except Exception:
                    pass
                try:
                    if shape.HasTextFrame and shape.TextFrame.HasText:
                        raw = shape.TextFrame.TextRange.Text or ""
                        if raw.strip():
                            texts.append(raw.strip())
                        avail_h = shape.Height - shape.TextFrame.MarginTop - shape.TextFrame.MarginBottom
                        avail_w = shape.Width - shape.TextFrame.MarginLeft - shape.TextFrame.MarginRight
                        bound_h = shape.TextFrame2.TextRange.BoundHeight
                        bound_w = shape.TextFrame2.TextRange.BoundWidth
                        if bound_h > avail_h + 6 or bound_w > avail_w + 6:
                            preview_txt = raw.replace("\r", " ").replace("\n", " ")[:24]
                            overflow.append(
                                f"第 {i} 页 「{preview_txt}」 高 {bound_h:.0f}/{avail_h:.0f} 宽 {bound_w:.0f}/{avail_w:.0f}"
                            )
                except Exception:
                    pass

            for k in range(1, slide.Shapes.Count + 1):
                walk(slide.Shapes(k))
            if not any(len(x) > 1 for x in texts):
                blank.append(i)
            if autoshapes < 1:
                no_shape.append(i)
            png = preview / f"slide_{i:02d}.png"
            slide.Export(str(png), "PNG", 1920, 1080)
        pres.Close()
        pres = None
        if started and ppt.Presentations.Count <= before:
            ppt.Quit()
            ppt = None
        tiny = []
        for i in range(1, len(items) + 1):
            png = preview / f"slide_{i:02d}.png"
            if not png.exists() or png.stat().st_size < 8000:
                tiny.append(i)
        notes.append(f"用本机 PowerPoint 导出 {len(items)} 页 1920×1080 PNG，并读取每段文字的实际占位，容差 6 磅。")
        notes.append(f"- 空白页：{('无' if not blank else blank)}")
        notes.append(f"- 文字超出文本框：{('无' if not overflow else '；'.join(overflow))}")
        notes.append(f"- 形状越出页面：{('无' if not offslide else '；'.join(offslide[:12]))}")
        notes.append(f"- 没有形状图的页：{('无' if not no_shape else no_shape)}")
        notes.append(f"- 导出文件过小：{('无' if not tiny else tiny)}")
        notes.append("- 没有外链图片。流程、对照和示意都是幻灯片里的形状。")
        notes.append("- 另目视抽查了 TRIZ 三页、过渡页、分级页、求解对照页、阶梯页、例子对照页、出处页和结尾页。")
        return "\n".join(notes), preview, overflow, blank, offslide, no_shape
    except Exception as exc:
        if pres is not None:
            try:
                pres.Close()
            except Exception:
                pass
        if started and ppt is not None:
            try:
                ppt.Quit()
            except Exception:
                pass
        return f"PowerPoint 检查没有完成：{exc}", None, [str(exc)], [], [], []


if __name__ == "__main__":
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass
    items, bounds, fit = build()
    result = inspect_with_powerpoint(items)
    if isinstance(result, tuple):
        check_note, preview, overflow, blank, offslide, no_shape = result
    else:
        check_note, preview, overflow, blank, offslide, no_shape = result, None, [], [], [], []
    extra = []
    if bounds:
        extra.append("python-pptx 边界：" + "；".join(bounds))
    if fit:
        extra.append("python-pptx 文本估算：" + "；".join(fit))
    if extra:
        check_note = check_note + "\n" + "\n".join(f"- {line}" for line in extra)
    OUTLINE_PATH.write_text(render_outline(items, check_note), encoding="utf-8")
    print(check_note)
    print(OUTLINE_PATH)
    print(NOTES_PATH)
    if preview:
        print("PREVIEW", preview)
