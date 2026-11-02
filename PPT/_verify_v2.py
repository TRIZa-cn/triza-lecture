# -*- coding: utf-8 -*-
import json, os, re, zipfile
from docx import Document
from pptx import Presentation

D = r"D:\doc\ZGCA\Lectures\TRIZ\PPT"
NAR = os.path.join(D, "TRIZ_W2A_矛盾解剖_15min_优化版_自动解说.pptx")
meta = json.load(open(os.path.join(D, "_narration_opt.json"), encoding="utf-8"))

z = zipfile.ZipFile(NAR)
slides = sorted([n for n in z.namelist() if re.match(r"ppt/slides/slide\d+\.xml$", n)],
                key=lambda s: int(re.search(r"(\d+)\.xml", s).group(1)))
adv = []; media = 0
for n in slides:
    x = z.read(n).decode("utf-8")
    m = re.findall(r'advTm="(\d+)"', x)
    adv.append(int(m[0])/1000 if m else None)
    media += int("<p:audio" in x or 'audioFile' in x)
z.close()

dur = [r["dur"] for r in meta["rows"]]
adv_ok = all(a is not None and abs(a - (d + 0.6)) < 0.2 for a, d in zip(adv, dur))

syl = Document(r"D:\doc\ZGCA\Lectures\TRIZ\广东工业大学教学大纲【2026TRIZ创新构思法研讨班@刘文印】.docx")
st = [p.text.strip() for p in syl.paragraphs]
chaps = [t for t in st if re.match(r"^第[一二三四五六七八九十]+章", t)]
hours = sum(int(m.group(1)) for t in chaps for m in [re.search(r"（(\d+)学时", t)] if m)
voice = r"D:\doc\ZGCA\Lectures\TRIZ\TRIZa-声音模型训练-录音讲稿.docx"
vd = Document(voice)
vsegs = [p.text for p in vd.paragraphs if re.match(r"^S\d\d　", p.text)]
opt = Presentation(os.path.join(D, "TRIZ_W2A_矛盾解剖_15min_优化版.pptx"))
n10 = opt.slides[9].notes_slide.notes_text_frame.text

print(json.dumps({
 "a_slides": len(slides), "a_audio": media, "a_adv": sum(1 for a in adv if a),
 "adv_ok": adv_ok, "adv_sec": adv[:3], "dur_sec": dur[:3],
 "cpm": meta["chars_per_min"], "min": round(meta["total_sec"]/60, 2),
 "s10_zhuti": ("主体" in n10 and "承担主体" in "".join(
     sh.text_frame.text for sh in opt.slides[9].shapes if sh.name == "Rectangle 5")),
 "s10_no_ariz": "ARIZ" not in n10,
 "syl_chapters": len([c for c in chaps if "附" not in c]), "syl_hours": hours,
 "syl_16h": any(t.startswith("学时/学分：16学时") for t in st),
 "voice_segs": len(vsegs), "voice_mb": round(os.path.getsize(voice)/1024/1024, 2),
}, ensure_ascii=False))