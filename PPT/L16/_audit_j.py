# -*- coding: utf-8 -*-
import os, re
p = r"D:\doc\ZGCLab\triza\project\TRIZa_网安参数与矛盾库_设计稿.md"
t = open(p, encoding="utf-8").read()
if "**实现者必读的三条约束**" in t:
    t = t.replace("**实现者必读的三条约束**", "**实现者必读的四条约束**")
    open(p, "w", encoding="utf-8").write(t)
    print("标题改为四条")

# 一致性核对：目录里实际有哪些产物，表里登记了哪些
D = r"D:\doc\ZGCLab\triza\project\design\akbank"
actual = []
for r, _, fs in os.walk(D):
    for f in fs:
        if f.endswith((".json", ".md")) and ".bak-" not in f:
            actual.append(os.path.relpath(os.path.join(r, f), r"D:\doc\ZGCLab\triza\project").replace("\\", "/"))
registered = set(re.findall(r"`(design/akbank/[^`]+)`", t))
print("\n目录实际产物 %d 个：" % len(actual))
for a in sorted(actual):
    print(("  [已登记] " if a in registered else "  [未登记] ") + a)
print("\n表里提到但目录不存在：", [x for x in registered if x not in actual] or "无")