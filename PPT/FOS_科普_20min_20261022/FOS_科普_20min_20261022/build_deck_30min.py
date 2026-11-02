# -*- coding: utf-8 -*-
"""30-minute FOS talk.

Opening remakes TRIZ lecture slides 2-5 in this deck's own light layout.
The overview figure on source slide 5 gets its own page.
The FOS line after that follows the 25-minute deck.

Sources are read-only. This script never opens them for writing.
"""

from __future__ import annotations

import importlib.util
import sys
from pathlib import Path

OUT_DIR = Path(r"D:\VideoPhone\outputs\FOS_科普_20min_20261022")
SRC_25 = OUT_DIR / "build_deck_25min.py"
PPTX_PATH = OUT_DIR / "FOS_科普_30min.pptx"
NOTES_MD = OUT_DIR / "notes-FOS-30min.md"
NOTES_TXT = OUT_DIR / "notes-FOS-30min.txt"
OUTLINE_PATH = OUT_DIR / "讲稿提纲_30min.md"

spec = importlib.util.spec_from_file_location("fos25", SRC_25)
d25 = importlib.util.module_from_spec(spec)
spec.loader.exec_module(d25)
fos = d25.fos

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
GREEN = fos.GREEN
WHITE = fos.WHITE
SOFT_BLUE = fos.SOFT_BLUE
SOFT_TEAL = fos.SOFT_TEAL
SOFT_RED = fos.SOFT_RED
SOFT_AMBER = fos.SOFT_AMBER
SOFT_GREEN = fos.SOFT_GREEN
LINE = fos.LINE

# Labels actually read from the source slide-5 picture. No tool list sits under  TRIZ工具.
DIAGRAM_MODULES = [
    "待解决的问题",
    "转化",
    "问题模型",
    "TRIZ工具",
    "解决方案模型",
    "应用",
    "最终解决方案",
    "难",
]
DIAGRAM_UNREAD = "「TRIZ工具」下面的工具清单：未辨认。图上这个标签下面没有工具名称。"


def say(text: str) -> str:
    return text.strip()


def draw_cover(slide, total):
    add_shape(slide, MSO_SHAPE.RECTANGLE, 0, 0, 13.333333, 0.08, BLUE)
    add_text(slide, 0.55, 0.85, 12.2, 0.36, [t("30 分钟科普讲稿", 16, False, TEAL)])
    add_text(slide, 0.55, 1.28, 12.2, 0.78, [t("功能导向检索", 44, True, INK)])
    add_text(
        slide,
        0.55,
        2.15,
        12.2,
        0.48,
        [t("先看清它在整套 TRIZ 里的位置，再只讲这一个工具", 22, False, INK)],
    )
    cards = [
        ("1", "萃智理论怎么串", "大约六分钟\n看总图，不讲完整套"),
        ("2", "交到这一个工具", "离开本领域找结构\n再把功能抽出来迁回"),
        ("3", "功能导向检索", "一个完整例子\n疫苗瓶监视器"),
    ]
    x = 0.55
    for num, title, body in cards:
        rrect(slide, x, 3.00, 3.90, 2.35, SOFT_BLUE, LINE, 1.0, 0.1)
        add_shape(slide, MSO_SHAPE.OVAL, x + 0.22, 3.22, 0.52, 0.52, BLUE)
        add_text(slide, x + 0.22, 3.22, 0.52, 0.52, [t(num, 16, True, WHITE, "center")], anchor="middle")
        add_text(slide, x + 0.86, 3.22, 2.80, 0.52, [t(title, 18, True, INK)], anchor="middle")
        add_text(slide, x + 0.22, 3.90, 3.46, 1.20, [t(body, 16, False, MUTED, spacing=1.12)], anchor="top")
        x += 4.15
    add_text(
        slide,
        0.55,
        5.60,
        12.2,
        0.70,
        [t("不是查重，也不是把整套 TRIZ 讲完。", 18, False, MUTED)],
        anchor="middle",
    )
    add_text(slide, 10.2, 7.05, 2.6, 0.32, [t(f"1  /  {total}", 12, False, MUTED, "right")], anchor="middle")


def draw_course(slide):
    heading(slide, "这几分钟，借的是萃智理论简介", size=28)
    rrect(slide, 0.48, 1.15, 12.38, 1.35, SOFT_TEAL, None, radius=0.08)
    add_text(slide, 0.72, 1.28, 11.9, 0.42, [t("刘文印  ·  TRIZ 创新构思法实战课", 22, True, TEAL)], anchor="middle")
    add_text(slide, 0.72, 1.78, 11.9, 0.48, [t("这一段的题目：TRIZ「萃智理论」简介", 18, False, INK)], anchor="middle")
    items = [
        ("1", "它是什么", "发明性问题的求解框架\n以及讲稿里的期待"),
        ("2", "一个例子怎么走", "先把表面问题收成矛盾\n原理只认名字"),
        ("3", "一张总图", "整套方法怎么串\nFOS 落在哪一步"),
    ]
    x = 0.48
    for num, title, body in items:
        rrect(slide, x, 2.75, 4.00, 2.85, WHITE, LINE, 1.15, 0.08)
        add_shape(slide, MSO_SHAPE.OVAL, x + 0.22, 2.98, 0.56, 0.56, BLUE)
        add_text(slide, x + 0.22, 2.98, 0.56, 0.56, [t(num, 16, True, WHITE, "center")], anchor="middle")
        add_text(slide, x + 0.92, 2.98, 2.85, 0.56, [t(title, 18, True, INK)], anchor="middle")
        add_text(slide, x + 0.24, 3.75, 3.52, 1.50, [t(body, 16, False, MUTED, spacing=1.15)], anchor="top")
        x += 4.18
    add_text(
        slide,
        0.48,
        5.80,
        12.38,
        0.90,
        [t("源第 2 页是这门课的封面。联系方式不写入本讲稿。\n下面四页只重做源第 3、4、5 页，版式是新的。", 16, False, INK, "center", spacing=1.12)],
        anchor="middle",
    )


def draw_what(slide):
    heading(slide, "TRIZ 是发明性问题的求解框架")
    rrect(slide, 0.48, 1.12, 12.38, 1.15, SOFT_BLUE, None, radius=0.08)
    add_text(
        slide,
        0.72,
        1.24,
        11.9,
        0.90,
        [t("萃智：基于归纳的、发明性问题的创造性求解理论框架\n拉丁文缩写 TRIZ，也可译成萃思", 18, True, INK, "center", spacing=1.12)],
        anchor="middle",
    )
    cards = [
        (SOFT_TEAL, TEAL, "发明性问题", "没有现成方案\n需要发明创新方案\n来解决"),
        (SOFT_AMBER, AMBER, "等级划分", "讲稿点了这一项\n说后面再讲\n表留到 FOS 那一段"),
        (SOFT_GREEN, GREEN, "讲稿里的期待", "从靠灵感迸发\n走到可以重复执行\n分析也可以交给机器"),
    ]
    x = 0.48
    for fill, color, title, body in cards:
        rrect(slide, x, 2.50, 4.00, 2.95, fill, None, radius=0.08)
        add_text(slide, x + 0.20, 2.68, 3.60, 0.50, [t(title, 20, True, color, "center")], anchor="middle")
        add_text(slide, x + 0.20, 3.30, 3.60, 1.85, [t(body, 16, False, INK, "center", spacing=1.12)], anchor="top")
        x += 4.18
    add_text(
        slide,
        0.48,
        5.65,
        12.38,
        1.10,
        [t("「完全由机器来执行」「可自动批量生产」是讲稿幻灯片上的期待，不是已经做到的结论。\n方法很多。讲稿推崇这一套，是因为它有希望变成固定步骤。", 16, False, INK, "center", spacing=1.12)],
        anchor="middle",
    )


def draw_case(slide):
    heading(slide, "例子先停在矛盾，不进入做法")
    rrect(slide, 0.42, 1.12, 6.15, 4.15, SOFT_AMBER, None, radius=0.08)
    rrect(slide, 6.75, 1.12, 6.15, 4.15, SOFT_RED, None, radius=0.08)
    add_text(slide, 0.66, 1.30, 5.70, 0.40, [t("讲稿里的初始问题", 16, True, AMBER)], anchor="middle")
    add_text(slide, 0.66, 1.80, 5.70, 1.15, [t("如何绕过网络审查", 26, True, INK, "center")], anchor="middle")
    add_text(
        slide,
        0.66,
        3.15,
        5.70,
        1.70,
        [t("这一步的名字：\n问题识别与矛盾转化\n今天只用它演示这一步", 16, False, MUTED, "center", spacing=1.12)],
        anchor="top",
    )
    add_text(slide, 6.99, 1.30, 5.70, 0.40, [t("收成物理矛盾", 16, True, RED)], anchor="middle")
    add_text(
        slide,
        6.99,
        1.82,
        5.70,
        2.35,
        [t("同一个网络数据流\n必须同时是标准的\nHTTPS 网页流量\n又必须是非标准的\nDoH 查询流量", 18, True, INK, "center", spacing=1.05)],
        anchor="top",
    )
    add_text(slide, 6.99, 4.30, 5.70, 0.70, [t("讲稿称为既要又要", 16, False, MUTED, "center")], anchor="middle")
    add_text(
        slide,
        0.48,
        5.45,
        12.38,
        1.30,
        [t("讲稿用大模型帮忙找矛盾、推荐原理，所以不必先把原理背熟。\n降低的是记忆门槛。绕过审查的做法不在这页，也不在后面展开。", 16, False, INK, "center", spacing=1.12)],
        anchor="middle",
    )


def draw_principles(slide):
    heading(slide, "沿着矛盾，讲稿点了原理的名字", size=26)
    cards = [
        ("原理 1", "分割", "讲稿写的是：分割或分离，\n也叫分而治之。\n具体怎么拆，不讲。"),
        ("原理 3", "局部质量", "均匀改为不均匀，\n各部分功能不同。\n已有方案只留名字：域前端。"),
        ("原理 31", "多孔材料", "讲稿的对法：放进看起来\n无害的空隙。\n走哪条通道，不讲。"),
    ]
    x = 0.42
    for kicker, title, body in cards:
        rrect(slide, x, 1.08, 4.05, 3.55, WHITE, LINE, 1.15, 0.08)
        add_text(slide, x + 0.18, 1.22, 3.70, 0.36, [t(kicker, 14, True, BLUE, "center")], anchor="middle")
        add_text(slide, x + 0.18, 1.62, 3.70, 0.55, [t(title, 24, True, INK, "center")], anchor="middle")
        add_text(slide, x + 0.22, 2.35, 3.62, 2.00, [t(body, 16, False, MUTED, "center", spacing=1.12)], anchor="top")
        x += 4.22
    rrect(slide, 0.42, 4.80, 12.48, 1.90, SOFT_BLUE, None, radius=0.08)
    add_text(slide, 0.66, 4.95, 12.00, 0.40, [t("另外点了名，今天不展开", 16, True, BLUE)], anchor="middle")
    add_text(
        slide,
        0.66,
        5.40,
        12.00,
        0.48,
        [t("4 非对称    ·    24 中介物    ·    35 参数变化    ·    10 预先作用", 18, True, INK, "center")],
        anchor="middle",
    )
    add_text(
        slide,
        0.66,
        5.95,
        12.00,
        0.52,
        [t("这些是源第 5 页列出的原理名称。功能导向搜索还不是这一页的主角。", 15, False, MUTED, "center")],
        anchor="middle",
    )


def draw_diagram(slide):
    heading(slide, "整套方法是这条回路，FOS 只占其中一步", size=26)
    # Four boxes in the same places as the source figure.
    rrect(slide, 0.36, 1.02, 3.85, 1.22, SOFT_BLUE, LINE, 1.25, 0.08)
    add_text(slide, 0.50, 1.14, 3.57, 0.98, [t("问题模型", 22, True, INK, "center")], anchor="middle")
    rrect(slide, 9.12, 1.02, 3.85, 1.22, SOFT_TEAL, LINE, 1.25, 0.08)
    add_text(slide, 9.26, 1.14, 3.57, 0.98, [t("解决方案模型", 22, True, INK, "center")], anchor="middle")
    rrect(slide, 0.36, 3.42, 3.85, 1.22, SOFT_AMBER, LINE, 1.25, 0.08)
    add_text(slide, 0.50, 3.54, 3.57, 0.98, [t("待解决的问题", 22, True, INK, "center")], anchor="middle")
    rrect(slide, 9.12, 3.42, 3.85, 1.22, SOFT_GREEN, LINE, 1.25, 0.08)
    add_text(slide, 9.26, 3.54, 3.57, 0.98, [t("最终解决方案", 22, True, INK, "center")], anchor="middle")

    add_shape(slide, MSO_SHAPE.UP_ARROW, 1.95, 2.38, 0.62, 0.90, BLUE)
    add_text(slide, 0.42, 2.48, 1.45, 0.70, [t("转化", 16, True, BLUE, "center")], anchor="middle")

    rrect(slide, 4.55, 1.08, 4.22, 0.48, BLUE, None, radius=0.1)
    add_text(slide, 4.65, 1.10, 4.02, 0.44, [t("TRIZ工具", 18, True, WHITE, "center")], anchor="middle")
    add_shape(slide, MSO_SHAPE.RIGHT_ARROW, 5.55, 1.68, 2.20, 0.40, BLUE)

    add_shape(slide, MSO_SHAPE.DOWN_ARROW, 10.72, 2.38, 0.62, 0.90, TEAL)
    add_text(slide, 11.40, 2.48, 1.45, 0.70, [t("应用", 16, True, TEAL, "center")], anchor="middle")

    add_shape(slide, MSO_SHAPE.LEFT_RIGHT_ARROW, 4.70, 3.62, 3.95, 0.55, SOFT_RED)
    add_text(slide, 4.70, 3.62, 3.95, 0.55, [t("×", 28, True, RED, "center")], anchor="middle")
    add_text(slide, 4.70, 4.22, 3.95, 0.38, [t("难", 18, True, RED, "center")], anchor="middle")

    rrect(slide, 0.36, 4.85, 12.61, 1.95, SOFT_AMBER, None, radius=0.08)
    add_text(
        slide,
        0.56,
        4.98,
        12.22,
        0.85,
        [t("功能导向搜索（FOS）只是「TRIZ工具」里的一个工具，不是整套 TRIZ。", 18, True, INK, "center")],
        anchor="middle",
    )
    add_text(
        slide,
        0.56,
        5.85,
        12.22,
        0.78,
        [t("图上「TRIZ工具」下面没有工具名称。读法：待解决的问题 → 转化 → 问题模型 → TRIZ工具 → 解决方案模型 → 应用 → 最终解决方案。直达标着「难」。", 14, False, MUTED, "center", spacing=1.08)],
        anchor="middle",
    )


def draw_bridge(slide):
    heading(slide, "离开本领域找结构，FOS 负责抽出再迁回", size=26)
    rrect(slide, 0.40, 1.12, 5.15, 4.20, SOFT_BLUE, None, radius=0.08)
    rrect(slide, 7.78, 1.12, 5.15, 4.20, SOFT_TEAL, None, radius=0.08)
    add_shape(slide, MSO_SHAPE.OVAL, 5.72, 2.45, 1.90, 1.90, BLUE)
    add_text(slide, 5.88, 2.85, 1.58, 0.48, [t("一个", 16, True, WHITE, "center")], anchor="middle")
    add_text(slide, 5.88, 3.30, 1.58, 0.48, [t("工具", 16, True, WHITE, "center")], anchor="middle")
    add_text(slide, 0.62, 1.32, 4.75, 0.48, [t("TRIZ 指出", 22, True, BLUE)], anchor="middle")
    add_text(
        slide,
        0.62,
        1.95,
        4.75,
        1.55,
        [t("离开本领域\n找已经长出来的结构", 22, True, INK, spacing=1.08)],
        anchor="top",
    )
    add_text(
        slide,
        0.62,
        3.70,
        4.75,
        1.30,
        [t("总图里，这一步标成\n「TRIZ工具」\n直达被划成难", 16, False, MUTED, spacing=1.08)],
        anchor="top",
    )
    add_text(slide, 8.00, 1.32, 4.75, 0.48, [t("功能导向搜索负责", 20, True, TEAL)], anchor="middle")
    add_text(
        slide,
        8.00,
        1.95,
        4.75,
        1.55,
        [t("把功能抽出来\n再迁回可验收的结构", 22, True, INK, spacing=1.08)],
        anchor="top",
    )
    add_text(
        slide,
        8.00,
        3.70,
        4.75,
        1.30,
        [t("泛化名词，下沉到具体物\n再看结构、前提和效果", 16, False, MUTED, spacing=1.08)],
        anchor="top",
    )
    add_text(
        slide,
        0.48,
        5.50,
        12.38,
        1.25,
        [t("它只是上一页「TRIZ工具」里的一个工具。图上没有把它的名字写出来。\n讲稿叫功能导向搜索。下面按功能导向检索来讲，缩写 FOS。", 16, False, INK, "center", spacing=1.12)],
        anchor="middle",
    )


def draw_fos_cover(slide, page, total):
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
    add_text(slide, 10.2, 7.05, 2.6, 0.32, [t(f"{page}  /  {total}", 12, False, MUTED, "right")], anchor="middle")


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
    old = d25.build_meta()
    # 25-minute items 0-2 are the short TRIZ opening, item 3 is its bridge.
    # Keep the FOS line from the section cover onward.
    fos_items = old[4:]
    bumps = {
        "按这六步，散会就可以做一张卡": 60,
        "出处只说明那边确实有这种做法": 40,
        "你手上哪个词，值得泛化一次？": 40,
    }
    for item in fos_items:
        if item["title"] in bumps:
            item["sec"] = bumps[item["title"]]
        item["triz_page"] = "不是源第 2–5 页"
        item["modules"] = ""

    opening = [
        dict(
            sec=50,
            title="功能导向检索",
            section="开场",
            figure="三块议程",
            origin="封面",
            origin_detail="新写的封面。源第 2 页是课程封面，不拿来当这一页。",
            triz_page="无",
            modules="",
            say=say(
                """
今天三十分钟。先用大约六分钟，看 TRIZ 整套方法怎么串起来。后面才讲功能导向检索。它只是里面的一个工具，不是整套理论。完整例子仍然是疫苗瓶监视器。不是查重，也不把四十条原理逐条讲完。
"""
            ),
            draw="cover",
            chrome=False,
        ),
        dict(
            sec=45,
            title="这几分钟，借的是萃智理论简介",
            section="萃智理论简介",
            figure="三段要讲的事",
            origin="TRIZ讲稿",
            origin_detail="源第 2 页：课程名与「TRIZ萃智理论简介」。微信和邮箱未写入讲稿。",
            triz_page="第 2 页",
            modules="",
            say=say(
                """
这几页不是从功能导向检索讲起。材料来自刘文印的课，题目是 TRIZ 创新构思法实战课，这一段叫萃智理论简介。源页上有微信和邮箱，这里不念。
接下来只做三件事：它是什么，一个例子怎么把问题收成矛盾，再看那张总图怎么把方法串起来。
"""
            ),
            draw="course",
            chrome=True,
        ),
        dict(
            sec=80,
            title="TRIZ 是发明性问题的求解框架",
            section="萃智理论简介",
            figure="定义与三点",
            origin="TRIZ讲稿",
            origin_detail="源第 3 页正文和备注：定义、发明性问题、等级划分、机械化的期待。",
            triz_page="第 3 页",
            modules="",
            say=say(
                """
讲稿的定义是：萃智，拉丁文缩写 TRIZ，也可译成萃思，是一套基于归纳的、发明性问题的创造性求解理论框架。发明性问题，就是没有现成方案，需要发明创新方案来解决。
讲稿还写了发明等级，说后面再讲。分级表留在功能导向检索那一段，用来对第三级，这里不展开。
为什么需要它。讲稿的理由是：解题方法很多，这一套有希望变成可重复的步骤，而不只靠灵感。幻灯片上的期待，是分析、搜索、评估都可以交给机器，甚至希望完全由机器执行，创新可以批量生产。这是讲稿里的理想，不是已经实现的结论。
"""
            ),
            draw="what",
            chrome=True,
        ),
        dict(
            sec=75,
            title="例子先停在矛盾，不进入做法",
            section="例子：收成矛盾",
            figure="初始问题与物理矛盾",
            origin="TRIZ讲稿",
            origin_detail="源第 4 页：初始问题、矛盾转化、物理矛盾原句。备注里的绕过步骤未写入。",
            triz_page="第 4 页",
            modules="",
            say=say(
                """
讲稿接着用一个例子走流程。初始问题写的是：如何绕过网络审查。今天只用它演示第一步，不讲任何绕过做法。
第一步叫问题识别和矛盾转化。讲稿把这个问题收成物理矛盾，也就是既要又要：同一个网络数据流，必须同时是标准的 HTTPS 网页流量，又必须是非标准的 DoH 查询流量。DoH 在这里只是矛盾里的后一个要求，不解释它怎么部署。
讲稿还说，大模型可以帮忙找矛盾、推荐原理，所以不必先把原理背熟。降低的是记忆门槛。问题仍然要先说成矛盾。
"""
            ),
            draw="case",
            chrome=True,
        ),
        dict(
            sec=70,
            title="沿着矛盾，讲稿点了原理的名字",
            section="例子：对上原理",
            figure="三条展开，四条只点名",
            origin="TRIZ讲稿",
            origin_detail="源第 5 页正文列出的原理。备注里的域前端、隧道和代理做法未写入。",
            triz_page="第 5 页正文",
            modules="",
            say=say(
                """
有了矛盾，讲稿从发明原理里点了几条名字。原理 1，分割或分离。原理 3，局部质量：把均匀结构改成不均匀，让各部分功能不同。讲稿把它对到数据流的不同部分扮演不同角色，并点了一个已有方案的名字，域前端。原理 31，多孔材料：讲稿的对法是把要送的内容放进看起来无害的空隙。
另外四个只报名字：4 非对称，24 中介物，35 参数变化，10 预先作用。
这些名字说明讲稿怎么把原理对到例子上。域前端怎么做、空隙走哪条通道，今天都不讲。功能导向搜索还不是这一页的主角。下一页才看整套方法的总图。
"""
            ),
            draw="principles",
            chrome=True,
        ),
        dict(
            sec=120,
            title="整套方法是这条回路，FOS 只占其中一步",
            section="总图",
            figure="四框回路，直达划掉",
            origin="TRIZ讲稿",
            origin_detail="源第 5 页右上角总图。模块名以图中看清的字为准。",
            triz_page="第 5 页右上角总图",
            modules="、".join(DIAGRAM_MODULES),
            say=say(
                """
看这张总图。它不是某一个工具的说明书，是整套走法的简图。左下角是待解决的问题。向上的箭头写着转化，到左上角的问题模型。再向右，箭头上写着 TRIZ工具，到右上角的解决方案模型。然后向下，箭头写着应用，到右下角的最终解决方案。
中间还有一条从问题直接通到最终方案的路，被划掉了，下面写着难。所以图的意思是：不要从原始问题直接跳到做法。先变成问题模型，用 TRIZ 工具找到解决方案模型，再应用回具体问题。
请注意 TRIZ工具 这四个字。它标在从问题模型到解决方案模型的箭头上。图的下面没有列出任何工具名称。这里不是看不清，是没有清单，所以工具名单写成未辨认，不猜矛盾矩阵，也不猜功能导向搜索已经印在图上。
功能导向搜索，缩写 FOS，只是这一步里的一个工具。它不是左下角的问题，不是整条回路，更不是整套 TRIZ。后面要讲的，就是这一个工具怎么把功能抽出来，再迁回去。
"""
            ),
            draw="diagram",
            chrome=True,
        ),
        dict(
            sec=70,
            title="离开本领域找结构，FOS 负责抽出再迁回",
            section="从 TRIZ 到 FOS",
            figure="TRIZ 与 FOS 的分工",
            origin="过渡",
            origin_detail="新写的过渡。左句是 TRIZ 要离开本领域找结构；右句是 FOS 抽出功能再迁回。FOS 的位置按总图的「TRIZ工具」来说，不把它写成图上已印的模块名。",
            triz_page="无（紧接总图）",
            modules="",
            say=say(
                """
上一页把分工画出来了。TRIZ 这套回路要求先离开原始问法，去找解决方案模型。要找的结构往往不在本领域里现成长着，所以要离开本领域。
功能导向搜索负责的，是「TRIZ工具」那一步里的一件事：把功能抽出来，到别的领域找已经被约束逼出来的结构，再迁回来。迁回要看结构、前提和可观测效果。
讲稿用的名字是功能导向搜索。下面按功能导向检索来讲，缩写 FOS。这一页起算 FOS，不计入刚才那六分半。
"""
            ),
            draw="bridge",
            chrome=True,
        ),
    ]
    return opening + fos_items


DRAWERS = {
    "cover": lambda slide, page, total: draw_cover(slide, total),
    "course": lambda slide, page, total: draw_course(slide),
    "what": lambda slide, page, total: draw_what(slide),
    "case": lambda slide, page, total: draw_case(slide),
    "principles": lambda slide, page, total: draw_principles(slide),
    "diagram": lambda slide, page, total: draw_diagram(slide),
    "bridge": lambda slide, page, total: draw_bridge(slide),
    "fos_cover": lambda slide, page, total: draw_fos_cover(slide, page, total),
    "d02": lambda slide, page, total: fos.draw_02(slide),
    "level3": lambda slide, page, total: d25.draw_level3(slide),
    "d03": lambda slide, page, total: fos.draw_03(slide),
    "d04": lambda slide, page, total: fos.draw_04(slide),
    "d05": lambda slide, page, total: fos.draw_05(slide),
    "d06": lambda slide, page, total: fos.draw_06(slide),
    "d07": lambda slide, page, total: fos.draw_07(slide),
    "d08": lambda slide, page, total: fos.draw_08(slide),
    "d09": lambda slide, page, total: fos.draw_09(slide),
    "d10": lambda slide, page, total: fos.draw_10(slide),
    "d11": lambda slide, page, total: fos.draw_11(slide),
    "d12": lambda slide, page, total: fos.draw_12(slide),
    "d13": lambda slide, page, total: fos.draw_13(slide),
    "d14": lambda slide, page, total: fos.draw_14(slide),
    "d15": lambda slide, page, total: fos.draw_15(slide),
    "d16": lambda slide, page, total: fos.draw_16(slide),
    "d17": lambda slide, page, total: fos.draw_17(slide),
    "d18": lambda slide, page, total: fos.draw_18(slide),
    "d19": lambda slide, page, total: fos.draw_19(slide),
    "d20": lambda slide, page, total: fos.draw_20(slide),
    "mapping": lambda slide, page, total: d25.draw_mapping(slide),
    "d21": lambda slide, page, total: fos.draw_21(slide),
    "d22": lambda slide, page, total: fos.draw_22(slide),
    "d23": lambda slide, page, total: fos.draw_23(slide),
    "d24": lambda slide, page, total: fos.draw_24(slide),
    "d25": lambda slide, page, total: fos.draw_25(slide),
    "d26": lambda slide, page, total: draw_close(slide, total),
}


def notes_text(meta):
    return f"建议用时：{meta['sec']} 秒\n\n讲者要点：\n{meta['say']}\n\n来源：{meta['origin_detail']}"


def render_notes(items):
    total = sum(item["sec"] for item in items)
    triz_sec = sum(item["sec"] for item in items if item["origin"] == "TRIZ讲稿")
    lines = [
        "# 功能导向检索 · 30 分钟讲者备注",
        "",
        f"合计 {total} 秒，即 {total // 60} 分钟。每页先给建议用时，再给可以照着讲的要点。",
        "",
        f"封面 50 秒单列。TRIZ 部分（源第 2–5 页重做，含总图）合计 {triz_sec} 秒。过渡页起算入 FOS，不计入这 {triz_sec} 秒。",
        "",
        f"总图里看清的字：{'、'.join(DIAGRAM_MODULES)}。{DIAGRAM_UNREAD}",
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
    triz_idx = [i for i, item in enumerate(items, 1) if item["origin"] == "TRIZ讲稿"]
    triz_sec = sum(items[i - 1]["sec"] for i in triz_idx)
    cover_i = next(i for i, item in enumerate(items, 1) if item["origin"] == "封面")
    bridge_i = next(i for i, item in enumerate(items, 1) if item["origin"] == "过渡")
    fos_sec = total - triz_sec - items[cover_i - 1]["sec"]
    lines = [
        "# 功能导向检索 · 30 分钟讲稿提纲",
        "",
        "由 25 分钟 FOS 科普扩成 30 分钟。源 pptx 未改，也未覆盖 `FOS_科普_20min.pptx` 和 `FOS_科普_25min.pptx`。",
        "",
        "- TRIZ 源（只读）：`C:\\Users\\21173\\Documents\\xwechat_files\\wxid_2swynsci7a1322_9b49\\msg\\file\\2026-10\\TRIZ20250928.pptx`",
        "- 版式沿用：`D:\\VideoPhone\\outputs\\FOS_科普_20min_20261022\\FOS_科普_25min.pptx`。浅色，16:9，微软雅黑，可编辑形状。没有把 TRIZ 源的旧版式贴进来。",
        f"- 输出：`FOS_科普_30min.pptx`（{len(items)} 页，合计 {total} 秒）",
        "- 生成：本机 python-pptx。未安装大型软件。",
        "",
        "## 时长怎么分",
        "",
        "| 部分 | 页 | 秒 | 从哪来 |",
        "| --- | --- | ---: | --- |",
        f"| 封面 | {cover_i} | {items[cover_i - 1]['sec']} | 新写。不计入 TRIZ 的 5–7 分钟 |",
        f"| TRIZ（含总图） | {triz_idx[0]}–{triz_idx[-1]} | {triz_sec} | 源第 2–5 页重做。{triz_sec / 60:.1f} 分钟，落在 5–7 分钟里 |",
        f"| FOS（含过渡） | {bridge_i}–{len(items)} | {fos_sec} | 过渡页新写。其后沿用 25 分钟稿主线 |",
        f"| 合计 | 1–{len(items)} | {total} | |",
        "",
        f"封面 {items[cover_i - 1]['sec']} 秒。TRIZ 五页 {triz_sec} 秒。过渡页起，其余 {fos_sec} 秒是 FOS。",
        "",
        "## 总图里实际辨认出的模块",
        "",
        "源第 5 页右上角是一张嵌入图片（约 694×310）。放大后看清的字如下，没有另补模块名。",
        "",
        "- 方框：待解决的问题；问题模型；解决方案模型；最终解决方案。",
        "- 箭头上的字：转化（向上，从待解决的问题到问题模型）；TRIZ工具（向右，从问题模型到解决方案模型）；应用（向下，从解决方案模型到最终解决方案）。",
        "- 中间一条双向箭头被划掉，下面的字是：难。",
        "",
        f"- {DIAGRAM_UNREAD}",
        "- 功能导向搜索在图中的位置：图上没有 FOS 这三个字，也没有单独的方框。讲稿写明，它只是「TRIZ工具」这一步里的一个工具，不是整套 TRIZ。",
        "",
        "## 源第 2–5 页重做成了什么",
        "",
        "- 第 2 页：课程名和「萃智理论简介」。联系方式没有抄上幻灯片。",
        "- 第 3 页：定义、发明性问题、等级划分只点名、机械化是讲稿里的期待。",
        "- 第 4 页：初始问题如何收成物理矛盾。备注里的绕过步骤没有写入。",
        "- 第 5 页正文：原理 1、3、31 的名称和讲稿上的一句对法；4、24、35、10 只点名。备注里的具体做法没有写入。",
        "- 第 5 页右上角总图：单独一页，按上面辨认到的模块讲回路。",
        "",
        "燃耗信用的展开、空安瓿案例和工具命令没有加回来。25 分钟稿里原有的一句边界、出处名单和结尾路径仍保留。",
        "",
        "## 逐页",
        "",
        "| 页 | 标题 | 秒 | 来自 TRIZ 源 | 总图模块 |",
        "| ---: | --- | ---: | --- | --- |",
    ]
    for i, item in enumerate(items, 1):
        modules = item.get("modules") or "—"
        lines.append(
            f"| {i} | {item['title']} | {item['sec']} | {item.get('triz_page', '无')} | {modules} |"
        )
    lines += [
        "",
        "## 渲染抽查",
        "",
        check_note.strip(),
        "",
    ]
    return "\n".join(lines)


def build():
    items = build_meta()
    total = sum(item["sec"] for item in items)
    triz_sec = sum(item["sec"] for item in items if item["origin"] == "TRIZ讲稿")
    if not (1760 <= total <= 1840):
        raise SystemExit(f"seconds {total}")
    if not (300 <= triz_sec <= 420):
        raise SystemExit(f"TRIZ seconds {triz_sec}")
    missing = sorted({item["draw"] for item in items} - set(DRAWERS))
    if missing:
        raise SystemExit(f"missing drawer {missing}")

    from pptx import Presentation
    from pptx.util import Emu

    prs = Presentation()
    prs.slide_width = Emu(fos.SW)
    prs.slide_height = Emu(fos.SH)
    prs.core_properties.title = "功能导向检索（FOS）30 分钟科普"
    prs.core_properties.subject = "开头重做 TRIZ 第 2–5 页，其后为功能导向检索"
    prs.core_properties.category = "科普讲稿"
    blank = prs.slide_layouts[6]
    total_n = len(items)

    for index, meta in enumerate(items, 1):
        slide = prs.slides.add_slide(blank)
        set_bg(slide)
        DRAWERS[meta["draw"]](slide, index, total_n)
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
    notes = render_notes(items)
    NOTES_MD.write_text(notes, encoding="utf-8")
    NOTES_TXT.write_text(notes, encoding="utf-8")
    print(PPTX_PATH)
    return items, bounds, fit


def inspect_with_powerpoint(items):
    import tempfile

    try:
        import win32com.client
    except ImportError:
        return "本机没有 pywin32，未能用 PowerPoint 打开检查。已用 python-pptx 核对形状没有越出页面。", None, [], [], [], []

    preview = Path(tempfile.mkdtemp(prefix="fos30_"))
    started = False
    ppt = None
    pres = None
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
        notes = [
            f"用本机 PowerPoint 导出 {len(items)} 页 1920×1080 PNG，并读取每段文字的实际占位，容差 6 磅。",
            f"- 空白页：{('无' if not blank else blank)}",
            f"- 文字超出文本框：{('无' if not overflow else '；'.join(overflow))}",
            f"- 形状越出页面：{('无' if not offslide else '；'.join(offslide[:12]))}",
            f"- 没有形状图的页：{('无' if not no_shape else no_shape)}",
            f"- 导出文件过小：{('无' if not tiny else tiny)}",
            "- 总图那一页用可编辑形状重画源第 5 页右上角的四框回路，模块名只使用图中看清的字。",
            "- 没有把 TRIZ 源的旧版式或外链图片贴进来。",
        ]
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
    print(NOTES_MD)
    if preview:
        print("PREVIEW", preview)
