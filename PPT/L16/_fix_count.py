# -*- coding: utf-8 -*-
import re, os
DOC = r"D:\doc\ZGCLab\triza\project\TRIZa_网安参数与矛盾库_设计稿.md"
t = open(DOC, encoding="utf-8").read()
orig = t
pairs = [
    ("为准：24 个参数", "为准：25 个参数"),
    ("**24**（P01~P18 + P19~P24 补缺）", "**25**（P01~P18 + P19~P24 补缺 + P25 操作可逆性）"),
]
for a, b in pairs:
    n = t.count(a)
    t = t.replace(a, b)
    print(f"替换 {n} 处: {a[:28]}")
if t != orig:
    import shutil
    shutil.copyfile(DOC, DOC + ".bak-20261029")
    open(DOC, "w", encoding="utf-8").write(t)
    print("saved")
left = [l.strip()[:100] for l in t.splitlines() if "24" in l and ("参数" in l or "P24" in l)]
print("\n仍含 24 的相关行（人工确认）:")
for l in left: print("  ", l)