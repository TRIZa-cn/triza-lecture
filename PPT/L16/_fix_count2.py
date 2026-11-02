# -*- coding: utf-8 -*-
import os, sys
DOC = r"D:\doc\ZGCLab\triza\project\TRIZa_网安参数与矛盾库_设计稿.md"
t = open(DOC, encoding="utf-8").read()
pairs = [
 ("安全参数种子（24 个：P01~P18 + P19~P24 补缺） | **24** |",
  "安全参数种子（25 个：P01~P18 + P19~P24 补缺 + P25 操作可逆性） | **25** |"),
 ("| 参数总数 | **24**（P01~P18 + P19~P24） |",
  "| 参数总数 | **25**（P01~P18 + P19~P24 + P25） |"),
]
for a, b in pairs:
    n = t.count(a); t = t.replace(a, b); print(f"{n} 处 <- {a[:34]}")
open(DOC, "w", encoding="utf-8").write(t)
left = [l.strip()[:90] for l in t.splitlines() if ("24 个" in l or "**24**" in l)]
print("仍写 24 的行:", left or "无")