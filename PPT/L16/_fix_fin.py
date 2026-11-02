# -*- coding: utf-8 -*-
from pptx import Presentation
P = r"D:\doc\ZGCA\Lectures\TRIZ\PPT\TRIZ_W2A_矛盾解剖_15min_优化版.pptx"
prs = Presentation(P)

# 定位所有出现
print("=== 改前 ===")
for i, s in enumerate(prs.slides, 1):
    for sh in s.shapes:
        if sh.has_text_frame:
            for p in sh.text_frame.paragraphs:
                for r in p.runs:
                    if "鳍片" in r.text:
                        print(f"  slide {i} [{sh.name}] run: {r.text}")
    n = s.notes_slide.notes_text_frame.text if s.has_notes_slide else ""
    if "鳍片" in n:
        print(f"  slide {i} NOTES 含 {n.count('鳍片')} 处")

# 1) 讲稿：连"鳍片就是散热片"那句一起去掉冗余
s7 = prs.slides[6]
n7 = s7.notes_slide.notes_text_frame.text
n7 = n7.replace("机身边缘的鳍片上再吹走。鳍片就是散热片，很多片薄金属",
                "机身边缘的散热片上再吹走。散热片是很多片薄金属")
s7.notes_slide.notes_text_frame.text = n7
print("notes7 剩余鳍片:", n7.count("鳍片"), "| 字数:", len(n7))

# 2) 页面文字
cnt = 0
for s in prs.slides:
    for sh in s.shapes:
        if sh.has_text_frame:
            for p in sh.text_frame.paragraphs:
                for r in p.runs:
                    if "鳍片" in r.text:
                        r.text = r.text.replace("鳍片", "散热片"); cnt += 1
print("页面 run 替换:", cnt)
prs.save(P)

prs2 = Presentation(P)
left = 0
for s in prs2.slides:
    for sh in s.shapes:
        if sh.has_text_frame and "鳍片" in sh.text_frame.text: left += 1
    if s.has_notes_slide and "鳍片" in s.notes_slide.notes_text_frame.text: left += 1
print("=== 改后残留:", left)
print("slide7 正文:", [sh.text_frame.text for sh in prs2.slides[6].shapes
                     if sh.has_text_frame and "散热片" in sh.text_frame.text])