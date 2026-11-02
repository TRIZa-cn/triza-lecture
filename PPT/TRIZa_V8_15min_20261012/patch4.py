# -*- coding: utf-8 -*-
import io, sys
from pathlib import Path
D = Path(r"D:\doc\ZGCA\Lectures\TRIZ\TRIZa_V8_15min_20261012")
bp = D / "build_v8.py"
src = bp.read_text(encoding="utf-8").lstrip("\ufeff")
a = '''        text(s, x + 0.28, 2.74, 3.20, 0.70,
             [{"t": h1, "sz": 22, "b": True, "c": NAVY, "ls": 1.1}])
        text(s, x + 0.28, 3.50, 3.22, 1.94,
             [{"t": b1, "sz": 20, "c": INK, "ls": 1.18}])'''
b = '''        text(s, x + 0.28, 2.74, 3.20, 0.82,
             [{"t": h1, "sz": 22, "b": True, "c": NAVY, "ls": 1.1}])
        text(s, x + 0.28, 3.58, 3.22, 1.86,
             [{"t": b1, "sz": 20, "c": INK, "ls": 1.18}])'''
assert a in src
src = src.replace(a, b)
c = '"L_cert = highest level whose validation requirements and hard gates were passed;  -1 if none.",'
e = '"L_cert = the highest level whose validation and hard gates were passed; -1 if none.",'
assert c in src
src = src.replace(c, e)
bp.write_text(src, encoding="utf-8", newline="\n")
print("PATCHED2")
