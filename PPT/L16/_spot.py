from pptx import Presentation
p = r"D:\doc\ZGCA\Lectures\TRIZ\PPT\L16\12_ARIZ求解流程.pptx"
prs = Presentation(p)
for i, s in enumerate(prs.slides[:3], 1):
    print(f"-- s{i}")
    for sh in s.shapes:
        if sh.has_text_frame and sh.text_frame.text.strip():
            print("   ", sh.text_frame.text.strip().replace("\n", " / ")[:100])
print("notes1:", prs.slides[0].notes_slide.notes_text_frame.text[:70])