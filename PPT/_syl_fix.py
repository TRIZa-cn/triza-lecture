# -*- coding: utf-8 -*-
from docx import Document
from docx.shared import Pt
DOC = r"D:\doc\ZGCA\Lectures\TRIZ\广东工业大学教学大纲【2026TRIZ创新构思法研讨班@刘文印】.docx"
d = Document(DOC)

# 1) 去掉重复标签：核心概念：核心概念： / AI协同点：AI协同点：
fixed = 0
for p in d.paragraphs:
    for label in ("核心概念：", "AI协同点："):
        dbl = label + label
        if p.text.startswith(dbl):
            # 第二个 run 以 label 开头
            for r in p.runs[1:]:
                if r.text.startswith(label):
                    r.text = r.text[len(label):]
                    fixed += 1
                    break
print("labels fixed:", fixed)

# 2) 合并重复的“以上各章的方法与工具”
TOTAL = "以上十六章合计 16 学时"
DUP = "以上各章的方法与工具，均在 TRIZa 的发明工作台上配套实现。"
total_p = None
dup_p = None
for p in d.paragraphs:
    if p.text.startswith(TOTAL):
        total_p = p
    elif p.text.startswith(DUP):
        dup_p = p
assert total_p is not None and dup_p is not None
r = total_p.add_run(dup_p.text.strip())
r.font.size = Pt(12)
dup_p._element.getparent().remove(dup_p._element)
print("merged dup sentence")

d.save(DOC)

# ---------- 校验 ----------
d2 = Document(DOC)
texts = [p.text.strip() for p in d2.paragraphs]
print("paras:", len(texts))
print("dup label left:", sum(1 for t in texts if t.startswith("核心概念：核心概念：") or t.startswith("AI协同点：AI协同点：")))
print("dup sentence left:", sum(1 for t in texts if t.startswith(DUP)))
chaps = [t for t in texts if t.startswith("第") and ("学时" in t) and ("核心概念" not in t)]
print("chapter lines:", len(chaps))
for t in chaps: print("   ", t)
total = sum(int("".join(c for c in t.split("（")[-1] if c.isdigit()) or 0) for t in chaps if "（" in t)
print("学时合计（从标题解析）:", total)