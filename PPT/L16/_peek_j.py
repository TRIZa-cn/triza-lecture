# -*- coding: utf-8 -*-
p = r"D:\doc\ZGCLab\triza\project\TRIZa_网安参数与矛盾库_设计稿.md"
lines = open(p, encoding="utf-8").read().splitlines()
for i, l in enumerate(lines, 1):
    if l.startswith("## ") or l.startswith("### "):
        if l.startswith("## J") or l.startswith("## K") or "落库" in l:
            print(f"{i:4d}: {l}")
print("---- J 段附近 ----")
start = next(i for i, l in enumerate(lines) if l.startswith("## J"))
end = next((i for i, l in enumerate(lines) if i > start and l.startswith("## ")), len(lines))
for i in range(start, min(end, start + 45)):
    print(f"{i+1:4d}| {lines[i]}")