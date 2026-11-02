# -*- coding: utf-8 -*-
import glob, os, re
from docx import Document
from pptx import Presentation

DOC = r"D:\doc\ZGCA\Lectures\TRIZ\广东工业大学教学大纲【2026TRIZ创新构思法研讨班@刘文印】.docx"
out = []
d = Document(DOC)
out.append(f"PARAS {len(d.paragraphs)} TABLES {len(d.tables)}")
for i, p in enumerate(d.paragraphs):
    t = p.text
    szs = sorted({r.font.size.pt for r in p.runs if r.font.size})
    b = any(r.font.bold for r in p.runs)
    out.append(f"[{i:03d}][{p.style.name}]{szs}{' B' if b else ''} {t}")
open(r"D:\doc\ZGCA\Lectures\TRIZ\PPT\_syllabus_now.txt", "w", encoding="utf-8").write("\n".join(out))

# PPT 扫描
out2 = []
for f in sorted(glob.glob(r"D:\doc\ZGCA\Lectures\TRIZ\**\*.pptx", recursive=True)):
    if "_raw" in f or "_voices" in f or "_kang" in f or "_male" in f or "_fos" in f or "_edge" in f:
        continue
    try:
        prs = Presentation(f)
    except Exception as e:
        out2.append(f"!! {os.path.basename(f)} {e}")
        continue
    hits = set()
    for s in prs.slides:
        for sh in s.shapes:
            if sh.has_text_frame:
                for para in sh.text_frame.paragraphs:
                    tx = para.text
                    if "创新构思法" in tx:
                        hits.add(tx.strip())
            if sh.has_table:
                for row in sh.table.rows:
                    for c in row.cells:
                        if "创新构思法" in c.text_frame.text:
                            hits.add(c.text_frame.text.strip())
    out2.append(f"### {os.path.basename(f)}  slides={len(prs.slides)}  rel={os.path.relpath(f, r'D:\doc\ZGCA\Lectures\TRIZ')}")
    for h in sorted(hits):
        out2.append("    | " + h)
open(r"D:\doc\ZGCA\Lectures\TRIZ\PPT\_deck_names.txt", "w", encoding="utf-8").write("\n".join(out2))
print("done", len(out), len(out2))