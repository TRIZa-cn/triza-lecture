# -*- coding: utf-8 -*-
import zipfile, re, os
src = r"D:\doc\ZGCA\Lectures\TRIZ\PPT\TRIZ_W1_矛盾15m.pptx"
z = zipfile.ZipFile(src)
media = [(n, z.getinfo(n).file_size) for n in z.namelist() if n.startswith("ppt/media/")]
print("源文件(录音版)自带媒体:", len(media))
for n, s in media: print("   %-28s %8.1f KB" % (n, s/1024))

# 每个 slide 引用了哪些 media
for n, s in media:
    base = os.path.basename(n)
    for i in range(1, 19):
        rel = "ppt/slides/_rels/slide%d.xml.rels" % i
        if rel in z.namelist() and base in z.read(rel).decode("utf-8", "ignore"):
            print("   → 出现在第 %d 页" % i)

from pptx import Presentation
prs = Presentation(src)
print("\n各页图片数:")
for i, s in enumerate(prs.slides, 1):
    pics = [sh for sh in s.shapes if sh.shape_type == 13]
    if pics:
        print("   第 %2d 页: %d 张 %s" % (i, len(pics),
              [(round(sh.width/914400,2), round(sh.height/914400,2)) for sh in pics]))
print("\nnotes 总字数:", sum(len(s.notes_slide.notes_text_frame.text) for s in prs.slides),
      "| 页数:", len(prs.slides))