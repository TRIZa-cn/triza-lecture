# -*- coding: utf-8 -*-
import sys
from pathlib import Path
D = Path(r"D:\doc\ZGCA\Lectures\TRIZ\TRIZa_V8_15min_20261012")
p = D / "build_v8b.py"
s = p.read_text(encoding="utf-8")
a = "bullets(s, x + 0.26, 2.62, 5.30, 3.20, items, mc=col, gap=12)"
b = "bullets(s, x + 0.26, 2.62, 5.30, 3.30, items, mc=col, gap=8)"
assert a in s
p.write_text(s.replace(a, b), encoding="utf-8", newline="\n")
print("FIXED")
