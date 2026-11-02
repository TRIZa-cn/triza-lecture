# -*- coding: utf-8 -*-
from docx import Document
from docx.shared import Pt
import sys
p = r"D:\doc\ZGCA\Lectures\TRIZ\北京中关村学院教学大纲【2026TRIZ创新构思法研讨班@刘文印】.docx"
d = Document(p)
print("PARAS", len(d.paragraphs), "TABLES", len(d.tables))
sec = d.sections[0]
print("PAGE", sec.page_width.cm, sec.page_height.cm, "margins", sec.left_margin.cm, sec.right_margin.cm, sec.top_margin.cm, sec.bottom_margin.cm)
print("---")
body = d.element.body
from docx.table import Table
from docx.text.paragraph import Paragraph
idx_p = 0; idx_t = 0
for child in body.iterchildren():
    tag = child.tag.split('}')[-1]
    if tag == 'p':
        para = Paragraph(child, d)
        t = para.text.strip()
        st = para.style.name
        sz = None
        for r in para.runs:
            if r.font.size: sz = r.font.size.pt; break
        print(f"[P{idx_p:03d}][{st}][{sz}] {t[:200]}")
        idx_p += 1
    elif tag == 'tbl':
        tb = Table(child, d)
        print(f"[T{idx_t:02d}] rows={len(tb.rows)} cols={len(tb.columns)}")
        for i, row in enumerate(tb.rows):
            cells = [c.text.strip().replace('\n',' / ') for c in row.cells]
            print("   R%d: %s" % (i, " | ".join(cells)[:220]))
        idx_t += 1