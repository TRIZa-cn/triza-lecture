# -*- coding: utf-8 -*-
p = r"D:\doc\ZGCA\Lectures\TRIZ\PPT\TRIZa-Lecture-SKILL.md"
lines = open(p, encoding="utf-8").read().splitlines()
for i, l in enumerate(lines, 1):
    if "16" in l or "字号" in l or "pt" in l.lower():
        print(f"{i:4d}| {l[:150]}")