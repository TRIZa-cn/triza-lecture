# -*- coding: utf-8 -*-
import glob, os, re
from pptx import Presentation
print(f"{'deck':<40}{'页':>3}{'正文最小':>8}{'最多字':>7}{'最多要点':>8}{'最多表格行':>10}{'讲稿':>6}{'分钟':>6}")
bad = []
for f in sorted(glob.glob(r"D:\doc\ZGCA\Lectures\TRIZ\PPT\L16\*.pptx")):
    prs = Presentation(f)
    min_sz, maxchars, maxbul, maxtbl, notes = 99, 0, 0, 0, 0
    for s in prs.slides:
        chars = 0; bul = 0
        for sh in s.shapes:
            if sh.has_text_frame and sh.text_frame.text.strip():
                t = sh.text_frame.text.strip()
                if (sh.top or 0) / 914400 > 6.9:
                    continue
                chars += len(re.sub(r"\s", "", t))
                bul += sum(1 for p in sh.text_frame.paragraphs if p.text.strip())
                for p in sh.text_frame.paragraphs:
                    for r in p.runs:
                        if r.font.size and r.text.strip():
                            min_sz = min(min_sz, r.font.size.pt)
            if sh.has_table:
                maxtbl = max(maxtbl, len(sh.table.rows))
        maxchars = max(maxchars, chars); maxbul = max(maxbul, bul)
        notes += len(s.notes_slide.notes_text_frame.text)
        for p in (s.notes_slide.notes_text_frame.paragraphs if s.has_notes_slide else []):
            pass
    name = os.path.basename(f)[:38]
    print(f"{name:<40}{len(prs.slides):>3}{min_sz:>8.0f}{maxchars:>7}{maxbul:>8}{maxtbl:>10}{notes:>6}{notes/270:>6.1f}")
    if maxchars > 250 or maxtbl > 8 or min_sz < 16:
        bad.append((name, maxchars, maxtbl, min_sz))
print()
print("超限页:", bad if bad else "无")