import json, io
from pptx import Presentation
p = r"D:\doc\ZGCA\Lectures\TRIZ\FOS_科普_20min_20261022\FOS_科普_20min_20261022\FOS_科普_15min.pptx"
prs = Presentation(p)
out = []
out.append(f"SIZE {prs.slide_width/914400:.3f} x {prs.slide_height/914400:.3f} in, N={len(prs.slides)}")
for i, s in enumerate(prs.slides, 1):
    d = {"slide": i, "layout": s.slide_layout.name, "shapes": []}
    for sh in s.shapes:
        e = {"name": sh.name, "type": str(sh.shape_type), "l": round(sh.left/914400,2) if sh.left is not None else None,
             "t": round(sh.top/914400,2) if sh.top is not None else None,
             "w": round(sh.width/914400,2) if sh.width is not None else None,
             "h": round(sh.height/914400,2) if sh.height is not None else None}
        if sh.has_text_frame:
            e["text"] = sh.text_frame.text
            runs = []
            for para in sh.text_frame.paragraphs:
                for r in para.runs:
                    runs.append({"t": r.text, "sz": r.font.size.pt if r.font.size else None,
                                 "b": r.font.bold, "font": r.font.name,
                                 "color": str(r.font.color.rgb) if r.font.color and r.font.color.type is not None and str(r.font.color.type)=="MSO_THEME_COLOR.NOT_THEME_COLOR" else None})
            e["runs"] = runs
        d["shapes"].append(e)
    if s.has_notes_slide:
        d["notes"] = s.notes_slide.notes_text_frame.text
    out.append(json.dumps(d, ensure_ascii=False, indent=1))
with open(r"D:\doc\ZGCA\Lectures\TRIZ\_dump15.json","w",encoding="utf-8") as f:
    f.write("\n".join(out))
print("ok")
