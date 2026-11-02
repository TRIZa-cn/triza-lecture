# -*- coding: utf-8 -*-
from docx import Document
V = r"D:\doc\ZGCA\Lectures\TRIZ\TRIZa-声音模型训练-录音讲稿.docx"
G = r"D:\doc\ZGCA\Lectures\TRIZ\TRIZa-16讲讲稿集-v1.docx"
KEY = ["你手上那个约束，是谁给的", "约束落在三个不同的地方", "把约束追到对象上", "把约束往下追"]
for name, path in (("录音稿", V), ("讲稿集", G)):
    t = "\n".join(p.text for p in Document(path).paragraphs)
    print(f"{name}: {len(t)} 字 | " + " ".join(f"{k[:6]}={'✓' if k in t else '✗'}" for k in KEY))
vd = Document(V)
segs = [p.text for p in vd.paragraphs if p.text.startswith("S") and "　" in p.text]
print("录音稿段数:", len(segs), "| 首末:", segs[0][:14], "...", segs[-1][:14])