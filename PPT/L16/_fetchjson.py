# -*- coding: utf-8 -*-
import json, urllib.request, os
URL = "https://raw.githubusercontent.com/Antropocosmist/triz-engineering-solver/main/resources/contradiction_matrix.json"
req = urllib.request.Request(URL, headers={"User-Agent": "Mozilla/5.0"})
raw = urllib.request.urlopen(req, timeout=30).read()
print("下载字节:", len(raw))
d = json.loads(raw)
print("顶层类型:", type(d).__name__)
if isinstance(d, dict):
    ks = list(d.keys())
    print("顶层键数:", len(ks), "| 前5:", ks[:5])
    k0 = ks[0]; v0 = d[k0]
    print("示例键 %r -> 类型 %s" % (k0, type(v0).__name__))
    if isinstance(v0, dict):
        print("  该键下子键数:", len(v0), "| 前5:", list(v0.keys())[:5])
        print("  第一个子项:", list(v0.items())[0])
    elif isinstance(v0, list):
        print("  列表长度:", len(v0), "| 前3:", v0[:3])
        print("  是否列表的列表:", isinstance(v0[0], list), "| 内层长度:", len(v0[0]) if isinstance(v0[0], list) else "-")
out = r"D:\doc\ZGCA\Lectures\TRIZ\PPT\assets\contradiction_matrix_raw.json"
os.makedirs(os.path.dirname(out), exist_ok=True)
open(out, "wb").write(raw)
print("已存:", out)