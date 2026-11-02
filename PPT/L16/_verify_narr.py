# -*- coding: utf-8 -*-
import zipfile, re, os, glob
P = r"D:\doc\ZGCA\Lectures\TRIZ\PPT\TRIZ_W1_矛盾15m_演播版.pptx"
z = zipfile.ZipFile(P)
media = [n for n in z.namelist() if n.startswith("ppt/media/")]
mp3 = [n for n in media if n.lower().endswith(".mp3")]
print("包内 media 文件:", len(media), "| 其中 mp3:", len(mp3))
print("总大小 %.2f MB" % (sum(z.getinfo(n).file_size for n in mp3)/1048576))

src = {os.path.basename(p): os.path.getsize(p)
       for p in glob.glob(r"D:\doc\ZGCA\Lectures\TRIZ\PPT\LWY录音20261030\_嵌入MP3\*.mp3")}

slides = sorted([n for n in z.namelist() if re.match(r"ppt/slides/slide\d+\.xml$", n)],
                key=lambda s: int(re.search(r"(\d+)", s).group()))
print("slide 部件:", len(slides))
ok_adv = ok_audio = 0
print("\n页  advTm   音频元素  隐藏位置(off-slide)  media大小")
for i, s in enumerate(slides, 1):
    xml = z.read(s).decode("utf-8", "ignore")
    m = re.search(r'<p:transition[^>]*advTm="(\d+)"', xml)
    adv = int(m.group(1))/1000 if m else None
    has_audio = ("audioFile" in xml) or ("<a:audioFile" in xml) or ('r:link' in xml and "media" in xml)
    # 媒体形状是否在画布外：找 a:off 最大坐标
    offs = [(int(a), int(b)) for a, b in re.findall(r'<a:off x="(-?\d+)" y="(-?\d+)"', xml)]
    offslide = any(x >= 914400 or y >= 480000 for x, y in offs)
    if adv is not None: ok_adv += 1
    if has_audio: ok_audio += 1
    print("%2d  %-7s %-9s %-19s" % (i, ("%.2fs" % adv) if adv else "无", "有" if has_audio else "无", "是" if offslide else "否"))

print("\n有 advTm 的页: %d/18 | 有音频元素: %d/18" % (ok_adv, ok_audio))
print("媒体文件名:", sorted(os.path.basename(n) for n in mp3)[:3], "...")

# 幻灯片内容未丢：页数与备注
from pptx import Presentation
prs = Presentation(P)
notes = sum(len(s.notes_slide.notes_text_frame.text) for s in prs.slides if s.has_notes_slide)
print("python-pptx 复核: %d 页, 备注 %d 字, 画布 %.0fx%.0f" % (
    len(prs.slides), notes, prs.slide_width/914400*72, prs.slide_height/914400*72))
srcfile = r"D:\doc\ZGCA\Lectures\TRIZ\PPT\TRIZ_W1_矛盾15m.pptx"
print("源文件未改动:", os.path.getsize(srcfile), "B",
      __import__("datetime").datetime.fromtimestamp(os.path.getmtime(srcfile)).strftime("%m-%d %H:%M:%S"))