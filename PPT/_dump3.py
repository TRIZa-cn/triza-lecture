from pptx import Presentation
p = r"D:\doc\ZGCA\Lectures\TRIZ\FOS_科普_20min_20261022\FOS_科普_20min_20261022\FOS_科普_15min.pptx"
prs = Presentation(p)
L=[]
for i,s in enumerate(prs.slides,1):
    L.append("="*80)
    L.append(f"SLIDE {i}")
    for sh in s.shapes:
        t=""
        if sh.has_text_frame:
            t=sh.text_frame.text.replace("\n"," ⏎ ")
        if t.strip()=="" and not sh.has_text_frame:
            t="<shape>"
        L.append(f"  {sh.shape_type} | {round(sh.left/914400,2)},{round(sh.top/914400,2)} {round(sh.width/914400,2)}x{round(sh.height/914400,2)} | {t}")
    if s.has_notes_slide:
        L.append("  ---NOTES---")
        L.append(s.notes_slide.notes_text_frame.text)
open(r"D:\doc\ZGCA\Lectures\TRIZ\_deck15.txt","w",encoding="utf-8").write("\n".join(L))
print(len(L))
