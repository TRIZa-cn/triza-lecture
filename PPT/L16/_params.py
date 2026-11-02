# -*- coding: utf-8 -*-
import json, urllib.request
B = "https://raw.githubusercontent.com/Antropocosmist/triz-engineering-solver/main/resources/"
d = json.load(open(r"D:\doc\ZGCA\Lectures\TRIZ\PPT\assets\contradiction_matrix_raw.json", encoding="utf-8"))
print("== meta 全文 ==")
for k, v in d["meta"].items():
    s = json.dumps(v, ensure_ascii=False)
    print(" %s: %s" % (k, s[:600]))

for name in ("39_parameters.md",):
    try:
        req = urllib.request.Request(B + name, headers={"User-Agent": "Mozilla/5.0"})
        txt = urllib.request.urlopen(req, timeout=30).read().decode("utf-8")
        open(r"D:\doc\ZGCA\Lectures\TRIZ\PPT\assets\39_parameters.md", "w", encoding="utf-8").write(txt)
        print("\n== %s (%d 字) ==" % (name, len(txt)))
        print(txt[:1400])
    except Exception as e:
        print(name, "FAIL", e)