# -*- coding: utf-8 -*-
import sys, docx
from pathlib import Path
D = Path(r"D:\doc\ZGCA\Lectures\TRIZ")
for name in ["00_矛盾论阶段约束与补救规划FOS研究报告.docx",
             "北京中关村学院教学大纲【2026TRIZ创新构思法研讨班@刘文印】.docx",
             "TRIZ发明等级.docx"]:
    p = D / name
    print("=" * 70)
    print("FILE:", name, "exists:", p.exists())
    if not p.exists():
        continue
    d = docx.Document(str(p))
    print("paras:", len(d.paragraphs), "tables:", len(d.tables))
    n = 0
    for para in d.paragraphs:
        t = para.text.strip()
        if t:
            n += 1
            print(f"[{para.style.name}] {t}")
    print("-- non-empty:", n, "chars:", sum(len(x.text) for x in d.paragraphs))
    for ti, tb in enumerate(d.tables):
        print(f"-- table {ti}: {len(tb.rows)}x{len(tb.columns)}")
        for r in tb.rows[:6]:
            print("   | " + " | ".join(c.text.strip().replace("\n", " ")[:40] for c in r.cells))
