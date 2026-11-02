# -*- coding: utf-8 -*-
import sys, os, json
sys.path.insert(0, r"D:\doc\ZGCA\Lectures\TRIZ\PPT\L16")
import _l16_render as R
import _l16_spec_a as A
import _l16_spec_a2 as A2
import _l16_spec_b as B
import _l16_spec_c as C

ALL = A.SPECS + A2.SPECS + B.SPECS + C.SPECS
print("lectures:", len(ALL), [s["no"] for s in ALL])
out = R.build_all(ALL)
for r in out:
    print(r)
print("decks:", len(out), "pages:", sum(r["pages"] for r in out),
      "notes:", sum(r["notes"] for r in out))

# ---------- A 讲加搭桥句 ----------
from pptx import Presentation
OPT = r"D:\doc\ZGCA\Lectures\TRIZ\PPT\TRIZ_W2A_矛盾解剖_15min_优化版.pptx"
prs = Presentation(OPT)
s2 = prs.slides[1]
n = s2.notes_slide.notes_text_frame.text
BRIDGE = "电池这个例子是经典的 TRIZ 案例，虽然离软件的场景远，但问题结构是一样的——你看，做漏洞扫描，追得深就跑得慢，对不对？"
if BRIDGE in n:
    print("bridge already there")
else:
    anchor = "还不能起火。"
    assert anchor in n, n[:80]
    n2 = n.replace(anchor, anchor + BRIDGE, 1)
    s2.notes_slide.notes_text_frame.text = n2
    prs.save(OPT)
    print("bridge inserted, notes2:", len(n2))