from pathlib import Path
from pptx import Presentation
OUT = Path(r"D:\doc\ZGCA\Lectures\TRIZ\FOS_科普_15min_优化版_20261022")
prs = Presentation(str(OUT / "FOS_科普_15min_优化版.pptx"))
lines = ["# 功能导向检索 · 十五分钟逐字稿", "",
         "> 每页一段，可直接照读。全文 " , ""]
body = []
total = 0
for i, s in enumerate(prs.slides, 1):
    nt = s.notes_slide.notes_text_frame.text.strip()
    total += len(nt)
    body.append("## 第 %02d 页\n\n%s\n" % (i, nt))
lines[2] = "> 每页一段，可直接照读。全文约 %d 字，正常语速约 15–16 分钟。" % total
out = "\n".join(lines + body)
(OUT / "讲稿_逐字稿.md").write_text(out, encoding="utf-8")
print("MD", len(out))
