# -*- coding: utf-8 -*-
import re
p = r"D:\doc\ZGCA\Lectures\TRIZ\PPT\build_w2.py"
src = open(p, encoding="utf-8").read().splitlines()
want = ("def page(", "def footer(", "def band(", "def table(", "def cover(", "def bullet(", "def para(", "def rect(")
for i, l in enumerate(src):
    if any(l.startswith(w) for w in want):
        print("  ".join([l.strip()] + [src[i+k].strip() for k in (1, 2, 3) if i + k < len(src) and src[i+k].strip()]))
        print()