# -*- coding: utf-8 -*-
from docx import Document
DOC = r"D:\doc\ZGCA\Lectures\TRIZ\广东工业大学教学大纲【2026TRIZ创新构思法研讨班@刘文印】.docx"
d = Document(DOC)
out=[]
for i,p in enumerate(d.paragraphs):
    b = "".join(("*" if r.bold else "") + r.text for r in p.runs) if p.runs else ""
    out.append(f"[{i:03d}] {p.text[:110]}")
open(r"D:\doc\ZGCA\Lectures\TRIZ\PPT\_syl_after.txt","w",encoding="utf-8").write("\n".join(out))
print("total", len(d.paragraphs))