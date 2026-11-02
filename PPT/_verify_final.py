# -*- coding: utf-8 -*-
import json, os, re
from docx import Document
from pptx import Presentation

D = r"D:\doc\ZGCA\Lectures\TRIZ\PPT"
# --- 更新术语账（按实际改动）---
tj = json.load(open(os.path.join(D, "_terms_A.json"), encoding="utf-8"))
FIX = {
 "通用工程参数": "第6页：参数就是可测量的量（重量/温度/速度/功耗/噪声）",
 "分割": "第7页：分割＝拆成互不影响的块",
 "抽取": "第7页：抽取＝把起作用的部件单独拿走",
 "动态化": "第7页：动态化＝让部件能变能适应",
 "中介物": "第7页：中介物＝中间加个传递载体",
 "鳍片": "第7页讲稿：鳍片就是散热片，多片薄金属",
}
for t in tj["terms"]:
    if t["name"] in FIX:
        t["explained"] = True
        t["where"] = FIX[t["name"]]
json.dump(tj, open(os.path.join(D, "_terms_A.json"), "w", encoding="utf-8"), ensure_ascii=False, indent=1)

# --- 复核 ---
OPT = os.path.join(D, "TRIZ_W2A_矛盾解剖_15min_优化版.pptx")
NAR = os.path.join(D, "TRIZ_W2A_矛盾解剖_15min_优化版_自动解说.pptx")
meta = json.load(open(os.path.join(D, "_narration_opt.json"), encoding="utf-8"))

nar = Presentation(NAR)
audio = adv = 0
for s in nar.slides:
    has_media = any(sh.shape_type == 16 for sh in s.shapes)  # MEDIA
    audio += int(has_media)
    adv += int(bool(s._element.find('.//{http://schemas.openxmlformats.org/presentationml/2006/main}transition')))

def footer(p):
    pr = Presentation(p); out = set()
    for s in pr.slides:
        for sh in s.shapes:
            if sh.has_text_frame and "创新构思法" in sh.text_frame.text:
                out.add(sh.text_frame.text.strip()); break
    return out

f_nar = footer(NAR)
decks = ["TRIZ_W2A_矛盾解剖_15min.pptx", "TRIZ_W2A_矛盾解剖_15min_优化版.pptx",
         "TRIZ_W2A_矛盾解剖_15min_自动解说.pptx", "TRIZ_W2B_盒子与矛盾论_15min.pptx",
         "TRIZ_W2B_盒子与矛盾论_15min_自动解说.pptx"]
footers_ok = all(all("实战研讨班" in t for t in footer(os.path.join(D, n))) for n in decks) \
             and all("实战研讨班" in t for t in f_nar)

opt = Presentation(OPT)
total_notes = sum(len(s.notes_slide.notes_text_frame.text) for s in opt.slides)
s10_notes = opt.slides[9].notes_slide.notes_text_frame.text
syl = Document(r"D:\doc\ZGCA\Lectures\TRIZ\广东工业大学教学大纲【2026TRIZ创新构思法研讨班@刘文印】.docx")
st = [p.text.strip() for p in syl.paragraphs]
chaps = [t for t in st if re.match(r"^第[一二三四五六七八九十]+章", t)]
hours = sum(int(m.group(1)) for t in chaps for m in [re.search(r"（(\d+)学时", t)] if m)
voice = r"D:\doc\ZGCA\Lectures\TRIZ\TRIZa-声音模型训练-录音讲稿.docx"
vd = Document(voice)
vsegs = [p.text for p in vd.paragraphs if re.match(r"^S\d\d　", p.text)]

res = {
 "audio": audio, "adv": adv, "slides": len(nar.slides),
 "cpm": meta["chars_per_min"], "min": round(meta["total_sec"]/60, 2), "chars": meta["total_chars"],
 "footers_ok": footers_ok,
 "s10_has_zhuti": "承担主体" in s10_notes,
 "s10_no_ariz": "ARIZ" not in s10_notes,
 "syl_chapters": len([c for c in chaps if "附" not in c]),
 "syl_hours": hours, "syl_16h_line": any(t.startswith("学时/学分：16学时") for t in st),
 "syl_total_line": any(t.startswith("以上十六章合计 16 学时") for t in st),
 "voice_docx": os.path.isfile(voice), "voice_segs": len(vsegs),
 "voice_size": os.path.getsize(voice) if os.path.isfile(voice) else 0,
}
print(json.dumps(res, ensure_ascii=False))