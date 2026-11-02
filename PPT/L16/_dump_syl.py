# -*- coding: utf-8 -*-
from docx import Document
p=r"D:\doc\ZGCA\Lectures\TRIZ\广东工业大学教学大纲【2026TRIZ创新构思法研讨班@刘文印】.docx"
d=Document(p)
print("=== 段落 ===")
for para in d.paragraphs:
    t=para.text.strip()
    if t: print(t)
print("\n=== 表格 ===")
for ti,tb in enumerate(d.tables,1):
    print("--- 表%d (%dx%d) ---"%(ti,len(tb.rows),len(tb.columns)))
    for r in tb.rows:
        cells=[c.text.strip().replace("\n"," / ") for c in r.cells]
        print(" | ".join(cells))