# -*- coding: utf-8 -*-
import sys, json, pathlib, importlib
ROOT = pathlib.Path(r"D:\doc\ZGCA\Lectures\TRIZ\PPT\W2_16")
sys.path.insert(0, str(ROOT))
import w_kit

OUT = ROOT / "out"
OUT.mkdir(exist_ok=True)

EXT = {}
for mn in ["w_notes_ext", "w_notes_ext_b", "w_notes_ext_c", "w_notes_ext_d", "w_notes_ext_e", "w_notes_ext_f", "w_notes_ext_g", "w_notes_ext_h", "w_notes_ext_i", "w_notes_ext_j"]:
    try:
        m = importlib.import_module(mn)
    except ModuleNotFoundError:
        continue
    for k, v in getattr(m, "EXT", {}).items():
        EXT.setdefault(k, {}).update(v)

try:
    from w_notes_a import NOTES_A
except Exception:
    NOTES_A = {}

mods = sys.argv[1:] or ["w_s0204", "w_s0506", "w_s0708", "w_s0910",
                        "w_s1112", "w_s1314", "w_s1516"]
sess = []
for m in mods:
    sess += importlib.import_module(m).SESSIONS

for s in sess:
    n = s["n"]
    if n in NOTES_A:
        s["notes"] = dict(NOTES_A[n])
    for k, v in EXT.get(n, {}).items():
        s["notes"][k] = s["notes"].get(k, "") + v

rows = []
for s in sorted(sess, key=lambda x: x["n"]):
    prs = w_kit.build_session(s)
    name = "\u7b2c%02d\u6b21_%s.pptx" % (s["n"], s.get("slug", "course"))
    f = OUT / name
    prs.save(str(f))
    a = w_kit.audit(f)
    rows.append(dict(n=s["n"], file=name, slides=a["slides"],
                     min_font=a["min_font"], notes=a["notes"],
                     est_min=round(a["est_min"], 1), warns=a["warns"],
                     kb=round(f.stat().st_size / 1024)))
    print("  n=%2d pages=%2d minf=%.0f notes=%4d est=%4.1f warns=%d"
          % (s["n"], a["slides"], a["min_font"], a["notes"], a["est_min"], a["warns"]))
print(json.dumps(rows, ensure_ascii=False))