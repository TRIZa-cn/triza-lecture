# -*- coding: utf-8 -*-
import json, os
from pptx import Presentation
A = r"D:\doc\ZGCA\Lectures\TRIZ\PPT\TRIZ_W2A_矛盾解剖_15min_优化版.pptx"
KEY = "被别人反驳"
prs = Presentation(A)
for i, s in enumerate(prs.slides, 1):
    n = s.notes_slide.notes_text_frame.text if s.has_notes_slide else ""
    if KEY in n:
        print(f"[命中] 第 {i} 页 讲稿，共 {len(n)} 字")
        print("---- 全文 ----")
        print(n)
    for sh in s.shapes:
        if sh.has_text_frame and (KEY in sh.text_frame.text or "口号" in sh.text_frame.text):
            print(f"[命中] 第 {i} 页 形状 {sh.name}: {sh.text_frame.text[:120]}")

print("\n==== contradiction.schema.json ====")
p = r"D:\doc\ZGCLab\triza\project\design\akbank\schema\contradiction.schema.json"
print(json.dumps(json.load(open(p, encoding="utf-8")), ensure_ascii=False, indent=1))