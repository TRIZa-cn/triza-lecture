# -*- coding: utf-8 -*-
from pathlib import Path

f = Path(r"D:\doc\ZGCA\Lectures\TRIZ\FOS_科普_15min_优化版_20261022\build_deck_v3.py")
s = f.read_text(encoding="utf-8")

old = ('        chip(s, LM + 0.22, y + 0.20, 1.95, 0.55, lv, fill=NAVY, color=WHITE, sz=20)\n'
       '        text(s, LM + 2.40, y + 0.10, CW - 2.56, 0.75, [{"t": txt, "sz": 20, "c": INK}],')
new = ('        chip(s, LM + 0.22, y + 0.20, 2.42, 0.55, lv, fill=NAVY, color=WHITE, sz=20)\n'
       '        text(s, LM + 2.88, y + 0.10, CW - 3.04, 0.75, [{"t": txt, "sz": 20, "c": INK}],')
assert s.count(old) == 3, s.count(old)
s = s.replace(old, new)

old2 = '        text(s, x + 0.22, BT + 0.68, bw - 0.44, 0.70,'
new2 = '        text(s, x + 0.22, BT + 0.66, bw - 0.44, 0.80,'
assert s.count(old2) == 1, s.count(old2)
s = s.replace(old2, new2)

old3 = '    text(s, LM, 4.80, 8.0, 0.95, [{"t": "谢谢", "sz": 52, "b": True, "c": NAVY}])'
new3 = '    text(s, LM, 4.78, 8.0, 1.10, [{"t": "谢谢", "sz": 52, "b": True, "c": NAVY}])'
assert s.count(old3) == 1, s.count(old3)
s = s.replace(old3, new3)

f.write_text(s, encoding="utf-8")
print("patched b")
