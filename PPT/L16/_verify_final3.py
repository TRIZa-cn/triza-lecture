# -*- coding: utf-8 -*-
import re
from pptx import Presentation
from docx import Document
P = r"D:\doc\ZGCA\Lectures\TRIZ\PPT\TRIZ_W2A_矛盾解剖_15min_优化版.pptx"
prs = Presentation(P)
s2, s4, s5 = prs.slides[1], prs.slides[3], prs.slides[4]
t4 = [sh for sh in s4.shapes if sh.name == "TextBox 2"][0].text_frame.text
p5 = [sh for sh in s5.shapes if sh.name == "TextBox 5"][0].text_frame.paragraphs[0].text
n2, n4, n5 = (s.notes_slide.notes_text_frame.text for s in (s2, s4, s5))

checks = {
 "s4_title": t4 == "约束有三种来源，对应三种矛盾",
 "s5_half": "把约束追到对象上" in p5,
 "s2_ask": "你手上那个约束，是谁给的" in n2,
 "s4_map": "约束落在三个不同的地方" in n4,
 "s5_trace": "把约束追到了对象上" in n5,
}
print("slide checks:", checks)

# 文档同步：把 deck notes 原文去查 docx
vd = "\n".join(p.text for p in Document(r"D:\doc\ZGCA\Lectures\TRIZ\TRIZa-声音模型训练-录音讲稿.docx").paragraphs)
gd = "\n".join(p.text for p in Document(r"D:\doc\ZGCA\Lectures\TRIZ\TRIZa-16讲讲稿集-v1.docx").paragraphs)
for name, n in (("S02", n2), ("S04", n4), ("S05", n5)):
    print(name, "voice:", n in vd, "| 讲稿集:", n in gd, "| len", len(n))

tot = sum(len(s.notes_slide.notes_text_frame.text) for s in prs.slides)
print("A notes:", tot, "→", round(tot/270, 2), "min @270 /", round(tot/258, 2), "min @258")
print("s2+s4+s5 :", len(n2), len(n4), len(n5))