# -*- coding: utf-8 -*-
from pathlib import Path

f = Path(r"D:\doc\ZGCA\Lectures\TRIZ\FOS_科普_15min_优化版_20261022\build_deck_v3.py")
s = f.read_text(encoding="utf-8")

pairs = [
    # 1) header boxes: give 20pt single lines room
    ('    rect(s, LM, 0.42, 0.055, 0.34, AMBER)\n'
     '    text(s, LM + 0.20, 0.40, CW - 0.2, 0.36,\n'
     '         [{"t": kicker, "sz": 20, "c": MUTED, "spc": 1.0}], name="kicker")\n'
     '    text(s, LM, 0.82, CW, 0.58,\n'
     '         [{"t": title, "sz": 30, "b": True, "c": NAVY, "ls": 1.06}], name="title")\n'
     '    if sub:\n'
     '        text(s, LM, 1.46, CW, 0.42,\n'
     '             [{"t": sub, "sz": 20, "c": MUTED, "ls": 1.1}], name="sub")',
     '    rect(s, LM, 0.40, 0.055, 0.40, AMBER)\n'
     '    text(s, LM + 0.20, 0.37, CW - 0.2, 0.44,\n'
     '         [{"t": kicker, "sz": 20, "c": MUTED, "spc": 1.0}], name="kicker")\n'
     '    text(s, LM, 0.80, CW, 0.66,\n'
     '         [{"t": title, "sz": 30, "b": True, "c": NAVY, "ls": 1.06}], name="title")\n'
     '    if sub:\n'
     '        text(s, LM, 1.44, CW, 0.46,\n'
     '             [{"t": sub, "sz": 20, "c": MUTED, "ls": 1.1}], name="sub")'),
    # 2) footer boxes
    ('    text(s, LM, FOOT_Y + 0.10, 9.0, 0.34, [{"t": section, "sz": 20, "c": GREY}])\n'
     '    text(s, LM + CW - 2.6, FOOT_Y + 0.10, 2.6, 0.34,',
     '    text(s, LM, FOOT_Y + 0.07, 9.0, 0.42, [{"t": section, "sz": 20, "c": GREY}])\n'
     '    text(s, LM + CW - 2.6, FOOT_Y + 0.07, 2.6, 0.42,'),
    # 3) principle grid geometry: taller cells, wider text
    ('    y0, rows, cols = 1.98, 5, 4\n'
     '    rowh = (6.74 - y0 - 4 * 0.08) / rows\n'
     '    colw = (CW - 3 * 0.14) / cols',
     '    y0, rows, cols = 1.62, 5, 4\n'
     '    rowh = (6.76 - y0 - 4 * 0.08) / rows\n'
     '    colw = (CW - 3 * 0.12) / cols'),
    ('        x = LM + c * (colw + 0.14)', '        x = LM + c * (colw + 0.12)'),
    ('        text(s, x + 0.14, y, colw - 0.24, rowh,',
     '        text(s, x + 0.10, y, colw - 0.16, rowh,'),
    # 4) pages 4/5: no subtitle, so the grid can start higher
    ('    head(s, "01 · 工具箱", "四十条发明原理（一）", "不是四十个公式，是四十个提问的角度")',
     '    head(s, "01 · 工具箱 · 四十个提问的角度", "四十条发明原理（一）")'),
    ('    head(s, "01 · 工具箱", "四十条发明原理（二）", "同一条原理，在机械、化工、软件里都成立")',
     '    head(s, "01 · 工具箱 · 同一条原理跨行业都成立", "四十条发明原理（二）")'),
    # 5) page 6: replace the badly wrapping text with a chip
    ('    text(s, LM + w2 + 0.10, 4.10, CW - 2 * w2 - 0.20, 0.60,\n'
     '         [{"t": "同一串\\n时间数字", "sz": 20, "b": True, "c": CORAL}], align=PP_ALIGN.CENTER)',
     '    ov = box(s, LM + w2 + 0.19, 3.99, 0.92, 0.92, fill=CORAL, line=None,\n'
     '             shape=MSO_SHAPE.OVAL, radius=0.5)\n'
     '    label(ov, [{"t": "矛盾", "sz": 20, "b": True, "c": WHITE}], margins=(0.02, 0.02, 0.01, 0.01))'),
    # 6) page 7: shorten two bodies that ran wider than the card
    ('"把日志聚合成一个根"', '"日志聚合成一个根"'),
    ('"把时刻从点改成区间"', '"时刻从点改成区间"'),
    ('"闰时日历提前注入演练"', '"闰时日历提前注入"'),
]

for old, new in pairs:
    n = s.count(old)
    assert n == 1, (n, old[:50])
    s = s.replace(old, new)

f.write_text(s, encoding="utf-8")
print("patched", len(pairs))
