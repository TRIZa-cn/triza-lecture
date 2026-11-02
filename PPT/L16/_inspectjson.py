# -*- coding: utf-8 -*-
import json
d = json.load(open(r"D:\doc\ZGCA\Lectures\TRIZ\PPT\assets\contradiction_matrix_raw.json", encoding="utf-8"))
print("== 说明 ==")
print(d["$schema_description"][:400])
print("\n== meta ==")
m = d["meta"]
print(json.dumps(m, ensure_ascii=False)[:900])
c = d["cells"]
print("\n== cells ==")
print("类型:", type(c).__name__, "| 长度:", len(c))
if isinstance(c, list):
    print("前3项:", json.dumps(c[:3], ensure_ascii=False))
elif isinstance(c, dict):
    kk = list(c)[:3]
    for k in kk: print("  %s -> %s" % (k, json.dumps(c[k], ensure_ascii=False)[:120]))