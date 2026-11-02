# -*- coding: utf-8 -*-
from pptx import Presentation
from pptx.util import Pt
P = r"D:\doc\ZGCA\Lectures\TRIZ\PPT\TRIZ_W1_矛盾15m.pptx"
prs = Presentation(P)

def set_band(slide, new_text):
    # 找该页最下方那条色带（T 在 5.3~6.3 之间的 AUTO_SHAPE）
    best = None
    for sh in slide.shapes:
        if sh.has_text_frame and sh.height is not None and sh.top is not None:
            t = sh.top / 914400; h = sh.height / 914400
            if 5.3 < t < 6.4 and h > 0.5 and sh.width/914400 > 8:
                if best is None or t > best.top/914400: best = sh
    if best is None:
        print("未找到色带"); return False
    tf = best.text_frame
    tf.clear()
    p = tf.paragraphs[0]
    p.line_spacing = 1.0
    import sys
    sys.path.insert(0, r"D:\doc\ZGCA\Lectures\TRIZ\PPT")
    import build_w2 as B
    B.rich(p, new_text, 20, B.INK)
    return True

print("p7 :", set_band(prs.slides[6], "约束就是矛盾——所以没有矛盾时，你可以亲手造一个。"))
print("p11:", set_band(prs.slides[10], "约束不是牢笼，是脚手架——盒子得你自己搭。"))
prs.save(P)

chk = Presentation(P)
for i in (7, 11):
    s = chk.slides[i-1]
    print(f"\n第{i}页 色带文字:")
    for sh in s.shapes:
        if sh.has_text_frame and sh.top and 5.3 < sh.top/914400 < 6.5 and sh.width/914400 > 8:
            print("   " + sh.text_frame.text.strip())
print("\n最终页数:", len(chk.slides), "| notes:", sum(len(x.notes_slide.notes_text_frame.text) for x in chk.slides), "字")