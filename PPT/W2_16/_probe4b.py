# -*- coding: utf-8 -*-
import io, re, json
p = r"D:\doc\ZGCA\Lectures\TRIZ\PPT\W2_16\w_s0204.py"
t = io.open(p, encoding="utf-8").read()
i = t.find('"n": 4,')
seg = t[i:]
# 截到下一个 session 起点（若有）
j = seg.find('"n": 5,')
if j > 0:
    seg = seg[:j]
out = []
for m in re.finditer(r'"([a-z_0-9]+)":\s*(\[.*?\]|"(?:[^"\\]|\\.)*")', seg, re.S):
    k, v = m.group(1), m.group(2)
    if v.startswith('['):
        try:
            v = json.loads(v)
        except Exception:
            continue
        v = " / ".join(str(x) for x in v)
    else:
        try:
            v = json.loads(v)
        except Exception:
            continue
    if isinstance(v, str) and ("原理" in v or "条" in v):
        out.append(k + " :: " + v)
io.open(r"D:\doc\ZGCA\Lectures\TRIZ\PPT\W2_16\_s4_dump.txt", "w", encoding="utf-8").write("\n".join(out))
print("lines", len(out))