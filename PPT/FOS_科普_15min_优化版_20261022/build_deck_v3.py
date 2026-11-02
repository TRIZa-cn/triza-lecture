# -*- coding: utf-8 -*-
"""35 页版本：全篇最小字号 20pt，闰秒案例作主线，疫苗瓶案例作附录。"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import build_deck_v2 as B  # noqa: E402
from pptx import Presentation  # noqa: E402
from pptx.util import Inches, Pt  # noqa: E402
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN  # noqa: E402
from pptx.enum.shapes import MSO_SHAPE  # noqa: E402

from notes_text_v3 import NOTES  # noqa: E402

C, text, box, label, rect = B.C, B.text, B.box, B.label, B.rect
card, numdot, new_slide = B.card, B.numdot, B.new_slide
PRINCIPLES = B.PRINCIPLES
NAVY, INK, BLUE, SKY = B.NAVY, B.INK, B.BLUE, B.SKY
AMBER, CORAL, GREEN, MUTED = B.AMBER, B.CORAL, B.GREEN, B.MUTED
LINE, SOFT, SOFT2, WHITE = B.LINE, B.SOFT, B.SOFT2, B.WHITE
GREY = "9AA8B8"

OUT = Path(r"D:\doc\ZGCA\Lectures\TRIZ\FOS_科普_15min_优化版_20261022")
PPTX = OUT / "FOS_科普_15min_优化版+闰秒案例.pptx"

SW, SH = 13.3333, 7.5
LM = 0.72
CW = SW - 2 * LM
FOOT_Y = 6.98
BT = 1.98
BT0 = 1.70

HL = {1, 3, 10, 24, 26, 34, 35}

DESC = {
    1: "把对象或状态拆开分别处理", 2: "只取出有用的那一部分", 3: "各处不一样，各管一段",
    4: "打破对称，按需要偏置", 5: "把同类操作并到一起", 6: "一件东西干多种活",
    7: "一层套一层，省空间", 8: "用浮力、升力抵消重量", 9: "提前施加反向的作用",
    10: "提前把该做的做完", 11: "提前备好应急手段", 12: "不抬不落，省掉升降",
    13: "反过来做，或者颠倒过来", 14: "直线换成弧线、球面", 15: "让结构随需要改变",
    16: "做得略多或略少更好用", 17: "从点线面换到另一个维度", 18: "加振动、加振荡",
    19: "由连续改成有节奏地做", 20: "别空转，让有效动作不停", 21: "危险段加速冲过去",
    22: "把坏处变成可利用之处", 23: "用结果反过来调动作", 24: "加一个中间载体来传递",
    25: "让对象自己维护自己", 26: "用便宜的复制品代替原件", 27: "用完就丢，省掉维护",
    28: "换成光、电、声、场", 29: "用气或液体代替固体", 30: "用薄膜隔开内外",
    31: "让材料带孔、带空隙", 32: "用颜色变化传递状态", 33: "相接触的部件用同种材料",
    34: "用完的部件丢掉或复原", 35: "改尺寸、温度、浓度等参数", 36: "利用物态变化吸放能量",
    37: "用热胀冷缩产生动作", 38: "用富氧加快反应", 39: "换成惰性介质保护",
    40: "把几种材料合成一种",
}


def head(s, kicker, title, sub=None):
    rect(s, LM, 0.40, 0.055, 0.40, AMBER)
    text(s, LM + 0.20, 0.37, CW - 0.2, 0.44,
         [{"t": kicker, "sz": 20, "c": MUTED, "spc": 1.0}], name="kicker")
    text(s, LM, 0.80, CW, 0.66,
         [{"t": title, "sz": 30, "b": True, "c": NAVY, "ls": 1.06}], name="title")
    if sub:
        text(s, LM, 1.44, CW, 0.46,
             [{"t": sub, "sz": 20, "c": MUTED, "ls": 1.1}], name="sub")


def foot(s, page, total, section):
    rect(s, LM, FOOT_Y, CW, 0.008, LINE)
    text(s, LM, FOOT_Y + 0.07, 9.0, 0.42, [{"t": section, "sz": 20, "c": GREY}])
    text(s, LM + CW - 2.6, FOOT_Y + 0.07, 2.6, 0.42,
         [{"t": "%02d / %02d" % (page, total), "sz": 20, "c": GREY}], align=PP_ALIGN.RIGHT)


def chip(s, x, y, w, h, t, fill=SOFT2, color=BLUE, sz=20, b=True, radius=0.5):
    sp = box(s, x, y, w, h, fill=fill, line=None, radius=radius)
    label(sp, [{"t": t, "sz": sz, "b": b, "c": color}], margins=(0.12, 0.12, 0.02, 0.02))
    return sp


def ncard(s, x, y, w, h, num, title, body, accent=BLUE, tsz=22, bsz=20):
    card(s, x, y, w, h)
    rect(s, x, y + 0.18, 0.05, h - 0.36, accent)
    text(s, x + 0.26, y + 0.18, w - 0.5, 0.36,
         [{"t": num, "sz": 20, "b": True, "c": accent, "spc": 0.6}])
    text(s, x + 0.26, y + 0.62, w - 0.50, 0.86,
         [{"t": title, "sz": tsz, "b": True, "c": NAVY, "ls": 1.14}])
    text(s, x + 0.26, y + 1.52, w - 0.50, h - 1.66,
         [{"t": body, "sz": bsz, "c": INK, "ls": 1.28}])


def row3(s, y, items, lw=2.7, h=1.18, gap=0.16):
    for i, (lab, body) in enumerate(items):
        yy = y + i * (h + gap)
        card(s, LM, yy, CW, h, fill=SOFT if i % 2 == 0 else WHITE)
        chip(s, LM + 0.22, yy + (h - 0.56) / 2, lw, 0.56, lab, fill=NAVY, color=WHITE, sz=20)
        text(s, LM + 0.22 + lw + 0.30, yy + 0.10, CW - (lw + 0.86), h - 0.20,
             [{"t": body, "sz": 20, "c": INK, "ls": 1.26}], anchor=MSO_ANCHOR.MIDDLE)


def pgrid(s, start):
    y0, rows, cols = 1.62, 5, 4
    rowh = (6.76 - y0 - 4 * 0.08) / rows
    colw = (CW - 3 * 0.12) / cols
    for i in range(20):
        idx = start + i
        r, c = divmod(i, cols)
        x = LM + c * (colw + 0.12)
        y = y0 + r * (rowh + 0.08)
        hl = idx in HL
        card(s, x, y, colw, rowh, fill=SOFT2 if hl else WHITE,
             line=AMBER if hl else LINE, lw=1.6 if hl else 1.0)
        text(s, x + 0.10, y, colw - 0.16, rowh,
             [{"t": "%02d %s" % (idx, PRINCIPLES[idx - 1]), "sz": 20, "b": True,
               "c": AMBER if hl else NAVY, "ls": 1.05},
              {"t": DESC[idx], "sz": 13, "c": MUTED, "ls": 1.05, "sp": 3}],
             anchor=MSO_ANCHOR.MIDDLE)


def build():
    prs = Presentation()
    prs.slide_width = Inches(SW)
    prs.slide_height = Inches(SH)
    S = []

    # ------------------------------------------------ 01 封面
    s = new_slide(prs); S.append(s)
    rect(s, 0, 0, SW, 4.15, NAVY)
    rect(s, 0, 4.15, SW, 0.06, AMBER)
    text(s, LM, 0.66, 10.0, 0.40, [{"t": "科普报告 · 15 分钟", "sz": 20, "c": SKY, "spc": 1.0}])
    text(s, LM, 1.10, 11.6, 1.85,
         [{"t": "功能导向检索：把别的行业做熟的事，搬到你的问题里", "sz": 40, "b": True,
           "c": WHITE, "ls": 1.22}])
    text(s, LM, 3.06, 11.6, 0.52,
         [{"t": "以「闰秒退场」为例，走一遍可重复的四步", "sz": 22, "c": "9FC3DF"}])
    rect(s, LM, 4.74, 0.055, 1.08, AMBER)
    text(s, LM + 0.26, 4.68, 10.8, 0.46,
         [{"t": "FOS · Function-Oriented Search", "sz": 24, "b": True, "c": NAVY}])
    text(s, LM + 0.26, 5.22, 10.8, 0.42,
         [{"t": "TRIZ 四十条发明原理，今天只点亮七条", "sz": 20, "c": MUTED}])
    text(s, LM + 0.26, 5.72, 11.4, 0.42,
         [{"t": "2027 年 5 月 20 日，UTC 停止插入闰秒", "sz": 20, "c": INK}])

    # ------------------------------------------------ 02 三个问题
    s = new_slide(prs); S.append(s)
    head(s, "今天只回答三个问题", "跨行业借鉴，为什么常常搬不动？")
    cw3 = (CW - 2 * 0.25) / 3
    for i, (num, q, a) in enumerate([
            ("问题一", "为什么听起来很美，做起来却像生搬硬套？", "因为搬的往往是那句功能描述，不是结构。"),
            ("问题二", "有没有不靠灵感的重复动作？", "有，四个动作，顺序不能乱。"),
            ("问题三", "怎么判断一次迁移是真有用？", "过三关：结构、前提、可观测效果。")]):
        ncard(s, LM + i * (cw3 + 0.25), BT, cw3, 3.30, num, q, a)
    rect(s, LM, 5.58, CW, 0.62, SOFT2)
    text(s, LM + 0.26, 5.62, CW - 0.5, 0.54,
         [{"t": "先用一个真实的国家级问题，把这三件事串起来。", "sz": 20, "c": INK}],
         anchor=MSO_ANCHOR.MIDDLE)
    foot(s, 2, 35, "开场")

    # ------------------------------------------------ 03 TRIZ 是什么
    s = new_slide(prs); S.append(s)
    head(s, "01 · 先认识 TRIZ", "TRIZ 不是灵感，是一套可重复的动作")
    row3(s, BT0, [
        ("它是什么", "从大量专利里归纳出的、面向发明性问题的求解框架，中文也叫萃智。"),
        ("它做什么", "先把一句模糊的问题收成一个清楚的矛盾，再用原理提示方向。"),
        ("它追求什么", "不是灵感，是可重复：换一个人、换一天，还能得出同样的判断。")])
    rect(s, LM, 5.72, CW, 0.62, SOFT2)
    text(s, LM + 0.26, 5.76, CW - 0.5, 0.54,
         [{"t": "今天只用到其中一小块：功能导向检索。", "sz": 20, "b": True, "c": BLUE}],
         anchor=MSO_ANCHOR.MIDDLE)
    foot(s, 3, 35, "01 · 认识 TRIZ")

    # ------------------------------------------------ 04 / 05 四十条原理
    s = new_slide(prs); S.append(s)
    head(s, "01 · 工具箱 · 四十个提问的角度", "四十条发明原理（一）")
    pgrid(s, 1)
    foot(s, 4, 35, "01 · 认识 TRIZ")

    s = new_slide(prs); S.append(s)
    head(s, "01 · 工具箱 · 同一条原理跨行业都成立", "四十条发明原理（二）")
    pgrid(s, 21)
    foot(s, 5, 35, "01 · 认识 TRIZ")

    # ------------------------------------------------ 06 收成矛盾
    s = new_slide(prs); S.append(s)
    head(s, "01 · 把问题收成矛盾", "闰秒退场：既要跳，又要不跳")
    text(s, LM, BT, CW, 1.30,
         [{"t": "2027 年 5 月 20 日起，UTC 不再插入闰秒。替代方案叫「闰时」：千年一次，一次跳 3600 秒。",
           "sz": 24, "b": True, "c": NAVY, "ls": 1.30}])
    w2 = 5.30
    for i, (tag, col, body) in enumerate([
            ("必须跳", AMBER, "要和地球自转对齐：天文、测绘、深空测控都要 UT1。"),
            ("必须不跳", BLUE, "要能计算、排序、举证：证书、口令、日志都要连续。")]):
        x = LM + i * (CW - w2)
        card(s, x, 3.46, w2, 1.98, fill=SOFT)
        rect(s, x, 3.46, w2, 0.05, col)
        text(s, x + 0.30, 3.62, w2 - 0.60, 0.48, [{"t": tag, "sz": 24, "b": True, "c": col}])
        text(s, x + 0.30, 4.16, w2 - 0.60, 1.05, [{"t": body, "sz": 20, "c": INK, "ls": 1.26}])
    ov = box(s, LM + w2 + 0.19, 3.99, 0.92, 0.92, fill=CORAL, line=None,
             shape=MSO_SHAPE.OVAL, radius=0.5)
    label(ov, [{"t": "矛盾", "sz": 20, "b": True, "c": WHITE}], margins=(0.02, 0.02, 0.01, 0.01))
    rect(s, LM, 5.66, CW, 1.02, SOFT2)
    text(s, LM + 0.30, 5.72, CW - 0.60, 0.90,
         [{"t": "取消闰秒不是消解这个矛盾，而是把妥协方向翻转：成本从计算侧，搬到天文侧。",
           "sz": 21, "c": INK, "ls": 1.24}], anchor=MSO_ANCHOR.MIDDLE)
    foot(s, 6, 35, "01 · 认识 TRIZ")

    # ------------------------------------------------ 07 七条原理
    s = new_slide(prs); S.append(s)
    head(s, "01 · 矛盾对原理", "七条原理，七个可以往下追的方向")
    items = [("01", "分割", "把时标和计数拆开"),
             ("03", "局部质量", "只在高价值链路做区间判定"),
             ("10", "预先作用", "闰时日历提前注入"),
             ("24", "中介物", "用版本化的相关表做中介"),
             ("26", "复制", "日志聚合成一个根"),
             ("34", "抛弃与修复", "把闰秒能力当资产退役"),
             ("35", "参数变化", "时刻从点改成区间")]
    cw4 = (CW - 3 * 0.16) / 4
    for i in range(8):
        r, c = divmod(i, 4)
        x = LM + c * (cw4 + 0.16)
        y = 1.95 + r * (1.80 + 0.18)
        card(s, x, y, cw4, 1.80, fill=SOFT2 if i < 7 else WHITE,
             line=AMBER if i < 7 else LINE, lw=1.6 if i < 7 else 1.0)
        if i < 7:
            num, name, body = items[i]
            text(s, x + 0.16, y + 0.16, cw4 - 0.32, 1.48,
                 [{"t": "%s %s" % (num, name), "sz": 20, "b": True, "c": AMBER, "ls": 1.05},
                  {"t": body, "sz": 20, "c": INK, "ls": 1.20, "sp": 4}])
        else:
            text(s, x + 0.16, y + 0.16, cw4 - 0.32, 1.48,
                 [{"t": "原理不给答案，只给方向。", "sz": 20, "c": MUTED, "ls": 1.24}],
                 anchor=MSO_ANCHOR.MIDDLE)
    foot(s, 7, 35, "01 · 认识 TRIZ")

    # ------------------------------------------------ 08 整套回路
    s = new_slide(prs); S.append(s)
    head(s, "01 · 整套方法", "别从原始问题直接跳到做法")
    bw, bh, by = 2.42, 1.45, 2.62
    x0 = LM + (CW - (4 * bw + 3 * 0.41)) / 2
    labels = ["待解决的问题", "问题模型", "解决方案模型", "具体方案"]
    for i, t in enumerate(labels):
        x = x0 + i * (bw + 0.41)
        fill = NAVY if i == 2 else WHITE
        card(s, x, by, bw, bh, fill=fill, line=LINE if i != 2 else NAVY)
        text(s, x + 0.12, by + 0.10, bw - 0.24, bh - 0.20,
             [{"t": t, "sz": 20, "b": True, "c": WHITE if i == 2 else NAVY, "ls": 1.20}],
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        if i < 3:
            B.arrow_right(s, x + bw + 0.05, by + bh / 2 - 0.18, 0.31, 0.36, SKY)
    for i, t in enumerate(["转化", "TRIZ 工具", "应用"]):
        gx = x0 + i * (bw + 0.41) + bw + 0.205
        text(s, gx - 1.05, 2.14, 2.10, 0.38, [{"t": t, "sz": 20, "c": MUTED}], align=PP_ALIGN.CENTER)
    rect(s, LM, 4.42, CW, 0.66, SOFT2)
    text(s, LM + 0.26, 4.46, CW - 0.5, 0.58,
         [{"t": "FOS 就住在「TRIZ 工具」这个格子里：它不是整条回路，只是其中一步。",
           "sz": 20, "b": True, "c": BLUE}], anchor=MSO_ANCHOR.MIDDLE)
    rect(s, LM, 5.30, CW, 0.74, "FBEEE6")
    text(s, LM + 0.26, 5.34, CW - 0.5, 0.66,
         [{"t": "× 从原始问题直接跳到做法：这条捷径走不通。", "sz": 20, "c": CORAL}],
         anchor=MSO_ANCHOR.MIDDLE)
    foot(s, 8, 35, "01 · 认识 TRIZ")

    # ------------------------------------------------ 09 五级发明
    s = new_slide(prs); S.append(s)
    head(s, "01 · 为什么要离开本领域", "发明分五级，FOS 对准第三级")
    rows = [("一级 · 本领域小改进", 32, "32%", SKY),
            ("二级 · 专业领域已知方法", 45, "45%", SKY),
            ("三级 · 跨领域迁移", 18, "18%", AMBER),
            ("四级 · 新科学原理", 4, "4%", BLUE),
            ("五级 · 全新发现", 0.7, "<1%", BLUE)]
    bx, bmax, sc = LM + 3.60, 5.95, 5.95 / 45.0
    for i, (lab, pct, txt, col) in enumerate(rows):
        y = 1.95 + i * 0.86
        text(s, LM, y + 0.10, 3.45, 0.42, [{"t": lab, "sz": 20, "c": INK}])
        rect(s, bx, y + 0.13, max(0.22, pct * sc), 0.36, col)
        text(s, bx + max(0.22, pct * sc) + 0.16, y + 0.10, 1.20, 0.42,
             [{"t": txt, "sz": 20, "b": True, "c": col}])
    rect(s, LM, 6.14, CW, 0.56, SOFT2)
    text(s, LM + 0.26, 6.16, CW - 0.5, 0.52,
         [{"t": "功能导向检索对准的，就是这 18%。", "sz": 20, "b": True, "c": BLUE}],
         anchor=MSO_ANCHOR.MIDDLE)
    foot(s, 9, 35, "01 · 认识 TRIZ")

    # ------------------------------------------------ 10 章节页 02
    s = new_slide(prs); S.append(s)
    rect(s, 0, 0, SW, SH, NAVY)
    text(s, LM, 1.85, 4.0, 1.70, [{"t": "02", "sz": 96, "b": True, "c": AMBER, "ls": 1.0}])
    text(s, LM, 3.62, 11.0, 0.92, [{"t": "功能导向检索是什么", "sz": 44, "b": True, "c": WHITE}])
    text(s, LM, 4.68, 11.0, 0.52, [{"t": "四个动作，一条顺序", "sz": 22, "c": "9FC3DF"}])
    rect(s, LM, 5.55, 2.6, 0.05, AMBER)

    # ------------------------------------------------ 11 四步
    s = new_slide(prs); S.append(s)
    head(s, "02 · 四个动作", "抽功能 → 泛化 → 选领域 → 迁回")
    cw4 = (CW - 3 * 0.22) / 4
    for i, (num, t, b) in enumerate([
            ("01", "抽功能", "把要解决的问题，抽成一句功能描述。"),
            ("02", "泛化名词", "把功能里的名词往上收一层。"),
            ("03", "选领先领域", "挑一个被逼得比你更狠的领域。"),
            ("04", "下沉并迁回", "到那里找具体结构，再搬回来。")]):
        ncard(s, LM + i * (cw4 + 0.22), BT, cw4, 3.30, num, t, b)
    rect(s, LM, 5.56, CW, 0.62, SOFT2)
    text(s, LM + 0.26, 5.60, CW - 0.5, 0.54,
         [{"t": "顺序不能乱：先泛化，再挑领域。", "sz": 20, "b": True, "c": BLUE}],
         anchor=MSO_ANCHOR.MIDDLE)
    foot(s, 11, 35, "02 · 功能导向检索")

    # ------------------------------------------------ 12 查重 vs FOS
    s = new_slide(prs); S.append(s)
    head(s, "02 · 最容易混的地方", "查重和功能导向检索，方向是反的")
    w2 = 5.72
    for i, (tag, col, lines) in enumerate([
            ("查重", MUTED, ["问：有没有人做过一模一样的事？", "用本领域的术语，越精确越好。", "方向：往里收。"]),
            ("功能导向检索", BLUE, ["问：别的地方解决过同一个功能吗？", "用上位词，用别人的行话。", "方向：往外扩。"])]):
        x = LM + i * (CW - w2)
        card(s, x, BT, w2, 3.50, fill=SOFT if i == 0 else SOFT2)
        rect(s, x, BT, w2, 0.05, col)
        text(s, x + 0.30, BT + 0.20, w2 - 0.60, 0.48, [{"t": tag, "sz": 24, "b": True, "c": col}])
        text(s, x + 0.30, BT + 0.86, w2 - 0.60, 2.30,
             [{"t": t, "sz": 20, "c": INK, "ls": 1.24} for t in lines])
    rect(s, LM, 5.72, CW, 0.62, SOFT2)
    text(s, LM + 0.26, 5.76, CW - 0.5, 0.54,
         [{"t": "同一批文献，两边搜索的方向几乎是相反的。", "sz": 20, "b": True, "c": INK}],
         anchor=MSO_ANCHOR.MIDDLE)
    foot(s, 12, 35, "02 · 功能导向检索")

    # ------------------------------------------------ 13 约束不同
    s = new_slide(prs); S.append(s)
    head(s, "02 · 为什么值得往外看", "同一个功能，约束不同，结构就不同")
    cw3 = (CW - 2 * 0.24) / 3
    for i, (dom, how, why) in enumerate([
            ("软件", "计数器加权限", "因为程序随时可以改。"),
            ("核工业", "看已经烧掉多少", "因为烧掉的部分改不回来。"),
            ("医院", "空安瓿收回来数", "因为药用完就没了。")]):
        x = LM + i * (cw3 + 0.24)
        card(s, x, BT, cw3, 2.62, fill=SOFT)
        text(s, x + 0.26, BT + 0.20, cw3 - 0.52, 0.46, [{"t": dom, "sz": 24, "b": True, "c": NAVY}])
        text(s, x + 0.26, BT + 0.74, cw3 - 0.52, 0.80, [{"t": how, "sz": 20, "c": INK, "ls": 1.22}])
        text(s, x + 0.26, BT + 1.62, cw3 - 0.52, 0.86, [{"t": why, "sz": 20, "c": MUTED, "ls": 1.22}])
    rect(s, LM, 4.92, CW, 1.28, SOFT2)
    text(s, LM + 0.30, 4.98, CW - 0.60, 1.16,
         [{"t": "功能是同一句话，结构完全不同，差别全在约束上。", "sz": 21, "b": True, "c": NAVY, "ls": 1.24},
          {"t": "你离开本领域，就是为了看见这些在本行长不出来的结构。", "sz": 20, "c": INK, "ls": 1.24, "sp": 6}],
         anchor=MSO_ANCHOR.MIDDLE)
    foot(s, 13, 35, "02 · 功能导向检索")

    # ------------------------------------------------ 14 搬结构
    s = new_slide(prs); S.append(s)
    head(s, "02 · 关键判断", "要搬的不是功能描述，是结构")
    card(s, LM, BT, 5.20, 3.05, fill=SOFT)
    text(s, LM + 0.28, BT + 0.20, 4.64, 0.46, [{"t": "不要搬这句", "sz": 22, "b": True, "c": MUTED}])
    text(s, LM + 0.28, BT + 0.80, 4.64, 1.90,
         [{"t": "功能抽到最后，大家得到的都是同一句话，谁都会写。", "sz": 20, "c": INK, "ls": 1.26}])
    x2 = LM + 5.60
    card(s, x2, BT, CW - 5.60, 3.05, fill=SOFT2, line=AMBER, lw=1.6)
    text(s, x2 + 0.28, BT + 0.20, CW - 6.16, 0.46,
         [{"t": "要搬这个", "sz": 22, "b": True, "c": AMBER}])
    for i, t in enumerate(["一个部件", "一个字段", "一台状态机"]):
        chip(s, x2 + 0.28, BT + 0.78 + i * 0.70, 2.30, 0.56, t, fill=WHITE, color=NAVY, sz=20)
    text(s, x2 + 2.80, BT + 0.80, CW - 8.80, 2.00,
         [{"t": "某种约束下被逼出来的具体做法。", "sz": 20, "c": INK, "ls": 1.26}])
    rect(s, LM, 5.30, CW, 0.62, SOFT2)
    text(s, LM + 0.26, 5.34, CW - 0.5, 0.54,
         [{"t": "后面所有的步骤，其实都是在找这个结构。", "sz": 20, "b": True, "c": BLUE}],
         anchor=MSO_ANCHOR.MIDDLE)
    foot(s, 14, 35, "02 · 功能导向检索")

    # ------------------------------------------------ 15 四级阶梯
    s = new_slide(prs); S.append(s)
    head(s, "02 · 把路径写下来", "四级阶梯：一条能复现的检索路径")
    for i, (lv, txt) in enumerate([
            ("一级 本领域词", "离线槽位"),
            ("二级 上位词", "有限次预发的消耗单元"),
            ("三级 领先领域", "航天：坏了没法派人上去修"),
            ("四级 具体物", "带补给窗口的在轨消耗品包")]):
        y = BT + i * 1.02
        card(s, LM, y, CW, 0.95, fill=SOFT if i % 2 == 0 else WHITE)
        chip(s, LM + 0.22, y + 0.20, 2.42, 0.55, lv, fill=NAVY, color=WHITE, sz=20)
        text(s, LM + 2.88, y + 0.10, CW - 3.04, 0.75, [{"t": txt, "sz": 20, "c": INK}],
             anchor=MSO_ANCHOR.MIDDLE)
    rect(s, LM, 6.12, CW, 0.56, SOFT2)
    text(s, LM + 0.26, 6.14, CW - 0.5, 0.52,
         [{"t": "少写一级，别人就没法判断你有没有跳步。", "sz": 20, "c": INK}],
         anchor=MSO_ANCHOR.MIDDLE)
    foot(s, 15, 35, "02 · 功能导向检索")

    # ------------------------------------------------ 16 泛化粒度
    s = new_slide(prs); S.append(s)
    head(s, "02 · 泛化的粒度", "太浅搜不到，太深等于没搜")
    cw3 = (CW - 2 * 0.24) / 3
    for i, (tag, ex, res, col) in enumerate([
            ("太浅", "把「瓶子」收到「罐子」", "还在包装行业里打转。", CORAL),
            ("太深", "把「瓶子」收到「东西」", "什么都搜得到，等于没搜。", CORAL),
            ("合适", "把「瓶子」收到「容器」", "疫苗瓶、太空舱、乏燃料罐。", GREEN)]):
        x = LM + i * (cw3 + 0.24)
        card(s, x, BT, cw3, 2.62, fill=SOFT2 if i == 2 else SOFT, line=AMBER if i == 2 else LINE)
        text(s, x + 0.26, BT + 0.20, cw3 - 0.52, 0.46, [{"t": tag, "sz": 24, "b": True, "c": col}])
        text(s, x + 0.26, BT + 0.74, cw3 - 0.52, 0.80, [{"t": ex, "sz": 20, "c": INK, "ls": 1.22}])
        text(s, x + 0.26, BT + 1.62, cw3 - 0.52, 0.86, [{"t": res, "sz": 20, "c": MUTED, "ls": 1.22}])
    rect(s, LM, 4.92, CW, 0.90, SOFT2)
    text(s, LM + 0.30, 4.96, CW - 0.60, 0.82,
         [{"t": "标准很硬：说出这个上位词，马上能举出三个行业的实物。", "sz": 21, "b": True, "c": NAVY}],
         anchor=MSO_ANCHOR.MIDDLE)
    foot(s, 16, 35, "02 · 功能导向检索")

    # ------------------------------------------------ 17 跳级 vs 过度泛化
    s = new_slide(prs); S.append(s)
    head(s, "02 · 这里有一对矛盾", "两条看起来相反的路，结果一样")
    w2 = 5.72
    for i, (tag, body) in enumerate([
            ("跳级", "只拿领域词去搜，跳过泛化。回来的全是跟你同构的方案，没有新结构。"),
            ("过度泛化", "上位词大到没法往下沉。最后写不出一句能执行的结构。")]):
        x = LM + i * (CW - w2)
        card(s, x, BT, w2, 2.72, fill=SOFT)
        rect(s, x, BT, w2, 0.05, CORAL)
        text(s, x + 0.30, BT + 0.20, w2 - 0.60, 0.46, [{"t": tag, "sz": 24, "b": True, "c": CORAL}])
        text(s, x + 0.30, BT + 0.78, w2 - 0.60, 1.80, [{"t": body, "sz": 20, "c": INK, "ls": 1.26}])
    rect(s, LM, 5.02, CW, 1.18, SOFT2)
    text(s, LM + 0.30, 5.06, CW - 0.60, 1.10,
         [{"t": "两条路看起来相反，结果却一样：都把检索做空。", "sz": 21, "b": True, "c": NAVY, "ls": 1.24},
          {"t": "合适的抽象夹在中间。", "sz": 20, "c": INK, "ls": 1.24, "sp": 4}],
         anchor=MSO_ANCHOR.MIDDLE)
    foot(s, 17, 35, "02 · 功能导向检索")

    # ------------------------------------------------ 18 领先领域
    s = new_slide(prs); S.append(s)
    head(s, "02 · 第三步", "挑被逼得更极端的领域，不挑听起来先进的")
    row3(s, BT, [
        ("航天", "被逼成地球独立：坏了没法派人上去修。"),
        ("核工业", "被逼成记录不可更改：烧掉的就改不回来了。"),
        ("医院", "被逼成必须能追责：每一支药都要能对上号。")])
    rect(s, LM, 5.72, CW, 0.62, SOFT2)
    text(s, LM + 0.26, 5.76, CW - 0.5, 0.54,
         [{"t": "被逼得更极端的地方，才更可能长出你没见过的结构。", "sz": 20, "b": True, "c": BLUE}],
         anchor=MSO_ANCHOR.MIDDLE)
    foot(s, 18, 35, "02 · 功能导向检索")

    # ------------------------------------------------ 19 章节页 03
    s = new_slide(prs); S.append(s)
    rect(s, 0, 0, SW, SH, NAVY)
    text(s, LM, 1.85, 4.0, 1.70, [{"t": "03", "sz": 96, "b": True, "c": AMBER, "ls": 1.0}])
    text(s, LM, 3.62, 11.0, 0.92, [{"t": "用闰秒走一遍", "sz": 44, "b": True, "c": WHITE}])
    text(s, LM, 4.68, 11.0, 0.52, [{"t": "从一句功能，走到一个能落地的结构", "sz": 22, "c": "9FC3DF"}])
    rect(s, LM, 5.55, 2.6, 0.05, AMBER)

    # ------------------------------------------------ 20 三件互不相干的事
    s = new_slide(prs); S.append(s)
    head(s, "03 · 例子", "三件互不相干的事，在解同一个问题")
    cw3 = (CW - 2 * 0.24) / 3
    for i, (dom, body) in enumerate([
            ("证书有效期", "判定是一个绝对时刻去比一个区间，字段里不带时标标识。"),
            ("一次性口令", "三十秒一步，容差最多一步；再放宽，就等于削弱安全。"),
            ("日志与审计", "排序靠时间戳；跨时标一拼，同一个「第几秒」可能不是同一秒。")]):
        x = LM + i * (cw3 + 0.24)
        card(s, x, BT, cw3, 2.80, fill=SOFT)
        text(s, x + 0.26, BT + 0.20, cw3 - 0.52, 0.46, [{"t": dom, "sz": 24, "b": True, "c": NAVY}])
        text(s, x + 0.26, BT + 0.78, cw3 - 0.52, 1.86, [{"t": body, "sz": 20, "c": INK, "ls": 1.24}])
    rect(s, LM, 5.10, CW, 1.10, SOFT2)
    text(s, LM + 0.30, 5.14, CW - 0.60, 1.02,
         [{"t": "它们都在解同一件事：当时间这个基准本身不可靠，", "sz": 21, "b": True, "c": NAVY, "ls": 1.24},
          {"t": "你怎么仍然能证明次序、新鲜度和基准变化。", "sz": 20, "c": INK, "ls": 1.24, "sp": 4}],
         anchor=MSO_ANCHOR.MIDDLE)
    foot(s, 20, 35, "03 · 闰秒案例")

    # ------------------------------------------------ 21 抽功能
    s = new_slide(prs); S.append(s)
    head(s, "03 · 第一步", "把问题抽成一句功能")
    card(s, LM, BT, 4.55, 2.50, fill=SOFT)
    text(s, LM + 0.28, BT + 0.20, 3.99, 0.46, [{"t": "原始问题", "sz": 22, "b": True, "c": MUTED}])
    text(s, LM + 0.28, BT + 0.82, 3.99, 1.30,
         [{"t": "怎么处理闰秒补偿？", "sz": 20, "c": INK, "ls": 1.26}])
    B.arrow_right(s, LM + 4.62, BT + 1.05, 0.46, 0.40, SKY)
    x2 = LM + 5.18
    card(s, x2, BT, CW - 5.18, 2.50, fill=SOFT2, line=BLUE, lw=1.4)
    text(s, x2 + 0.28, BT + 0.20, CW - 5.74, 0.46, [{"t": "功能一句话", "sz": 22, "b": True, "c": BLUE}])
    text(s, x2 + 0.28, BT + 0.82, CW - 5.74, 1.50,
         [{"t": "在基准本身会跳变、而且维护机制被移除之后，仍然能证明次序、新鲜度和基准变化。",
           "sz": 20, "c": INK, "ls": 1.24}])
    rect(s, LM, 4.80, CW, 1.52, SOFT2)
    text(s, LM + 0.30, 4.86, CW - 0.60, 1.40,
         [{"t": "这里刻意回避了「时间」这个词。", "sz": 21, "b": True, "c": NAVY, "ls": 1.24},
          {"t": "一旦限定在时间上，搜索范围就只剩授时和时钟同步，边际收益很低。", "sz": 20, "c": INK, "ls": 1.24, "sp": 6}],
         anchor=MSO_ANCHOR.MIDDLE)
    foot(s, 21, 35, "03 · 闰秒案例")

    # ------------------------------------------------ 22 泛化 + 选领域
    s = new_slide(prs); S.append(s)
    head(s, "03 · 第二、三步", "泛化名词，再挑领域")
    for i, (lv, txt) in enumerate([
            ("一级 本领域词", "闰秒补偿配置"),
            ("二级 上位词", "不连续基准上的次序与新鲜度证明"),
            ("三级 领先领域", "广播电视、深空测控"),
            ("四级 具体物", "帧计数加掉帧规则；双钟加版本化相关表")]):
        y = BT + i * 1.02
        card(s, LM, y, CW, 0.95, fill=SOFT if i % 2 == 0 else WHITE)
        chip(s, LM + 0.22, y + 0.20, 2.42, 0.55, lv, fill=NAVY, color=WHITE, sz=20)
        text(s, LM + 2.88, y + 0.10, CW - 3.04, 0.75, [{"t": txt, "sz": 20, "c": INK}],
             anchor=MSO_ANCHOR.MIDDLE)
    rect(s, LM, 6.12, CW, 0.56, SOFT2)
    text(s, LM + 0.26, 6.14, CW - 0.5, 0.52,
         [{"t": "走到这一步，我们要的那个结构已经能看见了。", "sz": 20, "c": INK}],
         anchor=MSO_ANCHOR.MIDDLE)
    foot(s, 22, 35, "03 · 闰秒案例")

    # ------------------------------------------------ 23 下沉
    s = new_slide(prs); S.append(s)
    head(s, "03 · 第四步", "到领先领域去下沉")
    w2 = 5.72
    for i, (dom, body) in enumerate([
            ("广播电视", "29.97 帧每秒和真实的秒对不上，于是发明了掉帧规则：用计数定义次序，用一条显式规则换算真实时间。"),
            ("深空测控", "地面和航天器的相关关系会事后被修正，所以双钟并行，相关表必须版本化，改了要留痕。")]):
        x = LM + i * (CW - w2)
        card(s, x, BT, w2, 3.20, fill=SOFT)
        rect(s, x, BT, w2, 0.05, AMBER)
        text(s, x + 0.30, BT + 0.20, w2 - 0.60, 0.48, [{"t": dom, "sz": 24, "b": True, "c": NAVY}])
        text(s, x + 0.30, BT + 0.84, w2 - 0.60, 2.20, [{"t": body, "sz": 20, "c": INK, "ls": 1.26}])
    rect(s, LM, 5.60, CW, 0.90, SOFT2)
    text(s, LM + 0.30, 5.64, CW - 0.60, 0.82,
         [{"t": "这不是表面类比：同构的问题，在那两个领域已经有成熟解。", "sz": 21, "b": True, "c": BLUE}],
         anchor=MSO_ANCHOR.MIDDLE)
    foot(s, 23, 35, "03 · 闰秒案例")

    # ------------------------------------------------ 24 迁回
    s = new_slide(prs); S.append(s)
    head(s, "03 · 最有潜力的那条", "计数时标 + 版本化相关表")
    boxes = [("计数时标", "单调，只增，不看墙上时钟"),
             ("版本化相关表", "v1 / v2 / v3，每次修订留痕"),
             ("对外绝对时刻", "按表换算，可复核")]
    bw, bh = 3.45, 1.48
    x0 = LM + (CW - (3 * bw + 2 * 0.45)) / 2
    for i, (t, b) in enumerate(boxes):
        x = x0 + i * (bw + 0.45)
        mid = (i == 1)
        card(s, x, BT, bw, bh, fill=NAVY if mid else WHITE, line=NAVY if mid else LINE)
        text(s, x + 0.22, BT + 0.16, bw - 0.44, 0.44,
             [{"t": t, "sz": 22, "b": True, "c": WHITE if mid else NAVY}])
        text(s, x + 0.22, BT + 0.66, bw - 0.44, 0.80,
             [{"t": b, "sz": 20, "c": "D7E4F0" if mid else INK, "ls": 1.20}])
        if i < 2:
            B.arrow_right(s, x + bw + 0.05, BT + bh / 2 - 0.19, 0.35, 0.38, SKY)
    cw3 = (CW - 2 * 0.24) / 3
    for i, (t, b) in enumerate([
            ("次序不变", "基准怎么跳，都不影响谁先谁后。"),
            ("历史可比", "旧记录用旧版本的表，能重建当时那一刻。"),
            ("改表可发现", "表被改过，能查出来，也能定位到哪一版。")]):
        x = LM + i * (cw3 + 0.24)
        card(s, x, 3.86, cw3, 1.36, fill=SOFT2)
        text(s, x + 0.24, 3.94, cw3 - 0.48, 0.42, [{"t": t, "sz": 21, "b": True, "c": GREEN}])
        text(s, x + 0.24, 4.38, cw3 - 0.48, 0.78, [{"t": b, "sz": 20, "c": INK, "ls": 1.20}])
    rect(s, LM, 5.42, CW, 0.96, SOFT2)
    text(s, LM + 0.30, 5.46, CW - 0.60, 0.88,
         [{"t": "它一次都没有追求「更准」，追求的是可换算、可举证。", "sz": 21, "b": True, "c": AMBER}],
         anchor=MSO_ANCHOR.MIDDLE)
    foot(s, 24, 35, "03 · 闰秒案例")

    # ------------------------------------------------ 25 三关 + 边界
    s = new_slide(prs); S.append(s)
    head(s, "03 · 迁回之前", "过三关，再签字")
    row3(s, BT, [
        ("关一", "能不能写成结构？→ 能：一张带版本号的相关表。"),
        ("关二", "前提还在不在？→ 在：表要进配置管理，改了必须留痕。"),
        ("关三", "有没有新的可测量？→ 有：相关表版本漂移、跨表拼接失败率。")],
         lw=1.25, h=1.05, gap=0.14)
    rect(s, LM, 5.55, CW, 0.98, "FBEEE6")
    text(s, LM + 0.30, 5.59, CW - 0.60, 0.90,
         [{"t": "边界：表被手工改、又没有版本，就会变成全网一致的错误。所以先清点，再上表。",
           "sz": 20, "c": CORAL, "ls": 1.24}], anchor=MSO_ANCHOR.MIDDLE)
    foot(s, 25, 35, "03 · 闰秒案例")

    # ------------------------------------------------ 26 六步法
    s = new_slide(prs); S.append(s)
    head(s, "04 · 带走这套动作", "回去以后，按这六步做")
    cw3 = (CW - 2 * 0.20) / 3
    steps = ["列出五个本领域的关键名词。", "每个名词写一个上位词。",
             "用三个行业的实物验粒度。", "挑一到两个被逼得更狠的领域。",
             "下沉到具体物，写清它对付哪条约束。", "填三关，写成一张迁移卡。"]
    for i, t in enumerate(steps):
        r, c = divmod(i, 3)
        x = LM + c * (cw3 + 0.20)
        y = 1.98 + r * 2.02
        card(s, x, y, cw3, 1.86, fill=SOFT if r == 0 else SOFT2)
        text(s, x + 0.26, y + 0.16, cw3 - 0.52, 0.40,
             [{"t": "0%d" % (i + 1), "sz": 20, "b": True, "c": AMBER, "spc": 0.6}])
        text(s, x + 0.26, y + 0.62, cw3 - 0.52, 1.12,
             [{"t": t, "sz": 20, "c": INK, "ls": 1.24}])
    rect(s, LM, 6.12, CW, 0.56, SOFT2)
    text(s, LM + 0.26, 6.14, CW - 0.5, 0.52,
         [{"t": "先泛化，再挑领域；反了很容易挑到听起来高级、其实无关的行业。", "sz": 20, "c": INK}],
         anchor=MSO_ANCHOR.MIDDLE)
    foot(s, 26, 35, "04 · 带走这套动作")

    # ------------------------------------------------ 27 五个坑
    s = new_slide(prs); S.append(s)
    head(s, "04 · 常见失败", "五个坑，按顺序")
    pits = [("01", "只搜领域词，不泛化名词", "回来的全是同构方案。"),
            ("02", "上位词大到没法下沉", "写不出一句能执行的结构。"),
            ("03", "停在「借鉴某某经验」", "这句话不可执行。"),
            ("04", "忽略原领域的前提", "把有兜底的机制搬到没兜底的地方。"),
            ("05", "只记机会，不记威胁", "检索就变成了自我表扬。")]
    for i, (num, t, b) in enumerate(pits):
        y = 1.95 + i * 0.94
        card(s, LM, y, CW, 0.84, fill=SOFT if i % 2 == 0 else WHITE)
        text(s, LM + 0.26, y + 0.06, 0.70, 0.72, [{"t": num, "sz": 20, "b": True, "c": CORAL}],
             anchor=MSO_ANCHOR.MIDDLE)
        text(s, LM + 1.00, y + 0.06, 4.60, 0.72, [{"t": t, "sz": 20, "b": True, "c": NAVY}],
             anchor=MSO_ANCHOR.MIDDLE)
        text(s, LM + 5.80, y + 0.06, CW - 6.06, 0.72, [{"t": b, "sz": 20, "c": INK}],
             anchor=MSO_ANCHOR.MIDDLE)
    foot(s, 27, 35, "04 · 带走这套动作")

    # ------------------------------------------------ 28 机会 vs 威胁
    s = new_slide(prs); S.append(s)
    head(s, "04 · 第五个坑", "只记机会，检索就变成了自我表扬")
    w2 = 5.72
    cols = [("机会这边", GREEN, ["扫描器", "断言格式", "锚定服务"]),
            ("威胁这边", CORAL, ["相关表被篡改，全网一起错", "退役删错，造成回调事故", "断言过宽，什么都降级"])]
    for i, (tag, col, items) in enumerate(cols):
        x = LM + i * (CW - w2)
        card(s, x, BT, w2, 3.30, fill=SOFT)
        rect(s, x, BT, w2, 0.05, col)
        text(s, x + 0.30, BT + 0.18, w2 - 0.60, 0.46, [{"t": tag, "sz": 24, "b": True, "c": col}])
        for j, it in enumerate(items):
            text(s, x + 0.30, BT + 0.80 + j * 0.72, w2 - 0.60, 0.64,
                 [{"t": "· " + it, "sz": 20, "c": INK, "ls": 1.20}])
    rect(s, LM, 5.62, CW, 0.90, SOFT2)
    text(s, LM + 0.30, 5.66, CW - 0.60, 0.82,
         [{"t": "把右边补上，你才知道这条路会在哪里翻车。", "sz": 21, "b": True, "c": CORAL}],
         anchor=MSO_ANCHOR.MIDDLE)
    foot(s, 28, 35, "04 · 带走这套动作")

    # ------------------------------------------------ 29 三句话
    s = new_slide(prs); S.append(s)
    head(s, "04 · 带走三句话", "如果只记三句")
    for i, t in enumerate(["搬结构，不搬功能。", "先泛化，再挑领域。", "过三关，再签字。"]):
        y = 1.95 + i * 1.44
        card(s, LM, y, CW, 1.24, fill=SOFT if i % 2 == 0 else SOFT2)
        text(s, LM + 0.30, y + 0.10, 1.20, 1.04,
             [{"t": "0%d" % (i + 1), "sz": 40, "b": True, "c": AMBER}],
             anchor=MSO_ANCHOR.MIDDLE)
        text(s, LM + 1.60, y + 0.10, CW - 1.90, 1.04,
             [{"t": t, "sz": 30, "b": True, "c": NAVY}], anchor=MSO_ANCHOR.MIDDLE)
    foot(s, 29, 35, "04 · 带走这套动作")

    # ------------------------------------------------ 30 结尾
    s = new_slide(prs); S.append(s)
    rect(s, 0, 0, SW, 2.05, NAVY)
    rect(s, 0, 2.05, SW, 0.05, AMBER)
    text(s, LM, 0.62, 11.4, 0.46, [{"t": "最后一句话", "sz": 20, "c": SKY, "spc": 1.0}])
    text(s, LM, 1.10, 11.4, 0.62, [{"t": "带着问题来，比带着答案走更有用", "sz": 30, "b": True, "c": WHITE}])
    text(s, LM, 2.55, 11.6, 1.60,
         [{"t": "如果你手上正好有一个「听起来很美、做起来搬不动」的借鉴，", "sz": 28, "c": INK, "ls": 1.30},
          {"t": "欢迎会后拿来，我们一起把它收成一个矛盾。", "sz": 28, "b": True, "c": NAVY, "ls": 1.30, "sp": 8}])
    rect(s, LM, 4.55, 2.2, 0.05, AMBER)
    text(s, LM, 4.78, 8.0, 1.10, [{"t": "谢谢", "sz": 52, "b": True, "c": NAVY}])
    text(s, LM, 5.95, 11.4, 0.44, [{"t": "FOS · 功能导向检索", "sz": 20, "c": MUTED}])
    foot(s, 30, 35, "04 · 带走这套动作")

    # ------------------------------------------------ 31 附录章节页
    s = new_slide(prs); S.append(s)
    rect(s, 0, 0, SW, SH, NAVY)
    text(s, LM, 1.95, 8.0, 0.46, [{"t": "附录 · 第二个例子", "sz": 20, "c": SKY, "spc": 1.0}])
    text(s, LM, 2.45, 11.4, 0.92, [{"t": "疫苗瓶：一瓶药，用过没有", "sz": 40, "b": True, "c": WHITE}])
    text(s, LM, 3.52, 11.4, 0.52, [{"t": "同一个动作，换一个更小的场景", "sz": 22, "c": "9FC3DF"}])
    rect(s, LM, 4.32, 2.6, 0.05, AMBER)
    for i, t in enumerate(["断网", "无第三方", "只能现场读"]):
        chip(s, LM + i * 2.35, 4.70, 2.05, 0.62, t, fill="1B3D5E", color="CFE0EE", sz=20)

    # ------------------------------------------------ 32 三个画面
    s = new_slide(prs); S.append(s)
    head(s, "附录 · 第一步", "三个画面，同一件事")
    cw3 = (CW - 2 * 0.24) / 3
    for i, (dom, body) in enumerate([
            ("疫苗瓶", "瓶身上一小块方块，随累计受热慢慢变深。"),
            ("乏燃料", "安全评估不看还剩多少，看已经烧掉多少。"),
            ("空安瓿", "药用完必须交回来，核对批号和数量。")]):
        x = LM + i * (cw3 + 0.24)
        card(s, x, BT, cw3, 2.62, fill=SOFT)
        text(s, x + 0.26, BT + 0.20, cw3 - 0.52, 0.46, [{"t": dom, "sz": 24, "b": True, "c": NAVY}])
        text(s, x + 0.26, BT + 0.78, cw3 - 0.52, 1.66, [{"t": body, "sz": 20, "c": INK, "ls": 1.24}])
    rect(s, LM, 4.92, CW, 1.28, SOFT2)
    text(s, LM + 0.30, 4.98, CW - 0.60, 1.16,
         [{"t": "没有联网、没有第三方在场，", "sz": 21, "b": True, "c": NAVY, "ls": 1.24},
          {"t": "怎么证明一件事发生过，而且只发生过规定的次数？", "sz": 20, "c": INK, "ls": 1.24, "sp": 6}],
         anchor=MSO_ANCHOR.MIDDLE)
    foot(s, 32, 35, "附录 · 疫苗瓶")

    # ------------------------------------------------ 33 阶梯
    s = new_slide(prs); S.append(s)
    head(s, "附录 · 第二、三步", "把阶梯走一遍")
    for i, (lv, txt) in enumerate([
            ("一级 本领域词", "已用标记"),
            ("二级 上位词", "累计的、不可逆的指示"),
            ("三级 领先领域", "疫苗冷链：热稳定性被管得极死"),
            ("四级 具体物", "疫苗瓶监视器：一张贴纸，随受热变深，分四档")]):
        y = BT + i * 1.02
        card(s, LM, y, CW, 0.95, fill=SOFT if i % 2 == 0 else WHITE)
        chip(s, LM + 0.22, y + 0.20, 2.42, 0.55, lv, fill=NAVY, color=WHITE, sz=20)
        text(s, LM + 2.88, y + 0.10, CW - 3.04, 0.75, [{"t": txt, "sz": 20, "c": INK}],
             anchor=MSO_ANCHOR.MIDDLE)
    rect(s, LM, 6.12, CW, 0.56, SOFT2)
    text(s, LM + 0.26, 6.14, CW - 0.5, 0.52,
         [{"t": "请注意这两个词：累计，不可逆。", "sz": 20, "b": True, "c": AMBER}],
         anchor=MSO_ANCHOR.MIDDLE)
    foot(s, 33, 35, "附录 · 疫苗瓶")

    # ------------------------------------------------ 34 阈值 vs 累计
    s = new_slide(prs); S.append(s)
    head(s, "附录 · 第四步", "阈值式，还是累计式")
    w2 = 5.72
    card(s, LM, BT, w2, 3.20, fill=SOFT)
    rect(s, LM, BT, w2, 0.05, MUTED)
    text(s, LM + 0.30, BT + 0.20, w2 - 0.60, 0.48, [{"t": "阈值式", "sz": 24, "b": True, "c": MUTED}])
    text(s, LM + 0.30, BT + 0.84, w2 - 0.60, 2.10,
         [{"t": "只有一个阈值，超过就报警。改回未用之后，外观上没有差别。",
           "sz": 20, "c": INK, "ls": 1.26}])
    x2 = LM + (CW - w2)
    card(s, x2, BT, w2, 3.20, fill=SOFT2, line=AMBER, lw=1.4)
    rect(s, x2, BT, w2, 0.05, AMBER)
    text(s, x2 + 0.30, BT + 0.20, w2 - 0.60, 0.48, [{"t": "累计式", "sz": 24, "b": True, "c": AMBER}])
    text(s, x2 + 0.30, BT + 0.84, w2 - 0.60, 0.80,
         [{"t": "零到三，四个档位，只增不减。", "sz": 20, "c": INK, "ls": 1.26}])
    for j, (t, col) in enumerate([("0", "D7E8DC"), ("1", "A9CBB4"), ("2", "D9A96A"), ("3", "C0725C")]):
        chip(s, x2 + 0.30 + j * 1.10, BT + 1.80, 0.95, 0.66, t, fill=col, color=NAVY, sz=22, radius=0.15)
    text(s, x2 + 0.30, BT + 2.58, w2 - 0.60, 0.50,
         [{"t": "每一档对应不同的处置动作。", "sz": 20, "c": MUTED}])
    rect(s, LM, 5.60, CW, 0.92, SOFT2)
    text(s, LM + 0.30, 5.64, CW - 0.60, 0.84,
         [{"t": "多出来的不是一句新功能，而是一个带着历史的状态。", "sz": 21, "b": True, "c": NAVY}],
         anchor=MSO_ANCHOR.MIDDLE)
    foot(s, 34, 35, "附录 · 疫苗瓶")

    # ------------------------------------------------ 35 迁回 + 三关
    s = new_slide(prs); S.append(s)
    head(s, "附录 · 迁回", "别搬那张贴纸，搬那个结构")
    card(s, LM, BT, w2, 2.55, fill=SOFT)
    text(s, LM + 0.30, BT + 0.18, w2 - 0.60, 0.46, [{"t": "不要搬", "sz": 22, "b": True, "c": MUTED}])
    text(s, LM + 0.30, BT + 0.76, w2 - 0.60, 1.60,
         [{"t": "那张会变色的贴纸。", "sz": 20, "c": INK, "ls": 1.26},
          {"t": "它是化学显色，换到软件里没有对应物。", "sz": 20, "c": INK, "ls": 1.26, "sp": 6}])
    card(s, x2, BT, w2, 2.55, fill=SOFT2, line=AMBER, lw=1.4)
    text(s, x2 + 0.30, BT + 0.18, w2 - 0.60, 0.46, [{"t": "要搬", "sz": 22, "b": True, "c": AMBER}])
    text(s, x2 + 0.30, BT + 0.76, w2 - 0.60, 1.60,
         [{"t": "不存一个布尔值，改存零到三的单调档位。", "sz": 20, "c": INK, "ls": 1.26},
          {"t": "三档只能核销，不能再新增。", "sz": 20, "c": INK, "ls": 1.26, "sp": 6}])
    rect(s, LM, 4.82, CW, 1.55, SOFT2)
    text(s, LM + 0.30, 4.86, CW - 0.60, 1.44,
         [{"t": "过三关还是同样的问法：能不能写成结构，前提还在不在，有没有新的可测量。",
           "sz": 21, "b": True, "c": NAVY, "ls": 1.24},
          {"t": "档位本身，就是不联网也能看的审计证据。", "sz": 20, "c": INK, "ls": 1.24, "sp": 6}],
         anchor=MSO_ANCHOR.MIDDLE)
    foot(s, 35, 35, "附录 · 疫苗瓶")

    assert len(S) == 35, len(S)
    for i, sl in enumerate(S):
        if i < len(NOTES):
            sl.notes_slide.notes_text_frame.text = NOTES[i]
    prs.save(str(PPTX))
    print("PAGES", len(S))
    print("NOTES_CHARS", sum(len(n) for n in NOTES))
    print("SAVED", PPTX)


def main():
    build()


if __name__ == "__main__":
    main()
