# -*- coding: utf-8 -*-
import re
from pathlib import Path
D = Path(r"D:\doc\ZGCA\Lectures\TRIZ\TRIZa_V8_15min_20261012")
p = D / "build_v8b.py"
src = p.read_text(encoding="utf-8")

def sub_fn(src, name, body):
    pat = re.compile(r"def %s\(prs\):.*?(?=\ndef )" % name, re.S)
    m = pat.search(src)
    assert m, name
    return src[:m.start()] + body.rstrip() + "\n" + src[m.end():]

P02 = '''
def p02(prs):
    s = new_slide(prs)
    head(s, "01 \\u00b7 THE GAP", "Prompting yields code. Principles yield evidence.")
    left = [
        "The raw problem goes straight into the prompt.",
        "A few rounds may still produce runnable code.",
        "Nothing forces the model to name the trade-off.",
        "Different wording next run: not reproducible, no ablation.",
    ]
    right = [
        "A contradiction is written down before generation.",
        "Each candidate carries one principle and one parameter pair.",
        "Ablatable: remove the operator and the effect should vanish.",
        "Comparable: the protocol is frozen first, so runs line up.",
    ]
    for x, title, items, col in ((LM, "PROMPT ONLY", left, CORAL),
                                 (LM + 6.09, "THROUGH TRIZ", right, GREEN)):
        card(s, x, 1.78, 5.80, 4.14, fill=SOFT, accent=col)
        text(s, x + 0.26, 1.94, 5.30, 0.62,
             [{"t": title, "sz": 20, "b": True, "c": col}])
        bullets(s, x + 0.26, 2.62, 5.30, 3.20, items, mc=col, gap=12)
    band(s, 6.06, [("Ablatable, comparable, reproducible ", {"c": INK, "b": True}),
                   ("\\u2014 or none of the three.", {"c": INK})], h=0.58)
    foot(s, "The gap", "02")
    return s
'''

P10 = '''
def p10(prs):
    s = new_slide(prs)
    head(s, "09 \\u00b7 THE DECISION", "Improved on three measures. Not promoted.")
    text(s, LM, 1.74, 6.5, 1.02,
         [{"runs": [("0.83256", {"c": MUTED, "b": True, "sz": 32}),
                    ("  \\u2192  ", {"c": AMBER, "b": True, "sz": 32}),
                    ("0.84839", {"c": NAVY, "b": True, "sz": 32})]}])
    text(s, LM, 2.80, 6.5, 0.62, [{"t": "macro-F1, exact endpoints", "sz": 20, "c": MUTED}])
    text(s, LM, 3.46, 6.5, 0.72,
         [{"runs": [("\\u0394 = +0.01584", {"c": INK, "b": True, "sz": 24}),
                    ("     1.90% relative", {"c": MUTED, "sz": 20})]}])
    stats = [("43 / 45", "pairs positive"),
             ("5 / 5", "seed blocks positive"),
             ("+0.01802", "recall"),
             ("\\u22120.00168", "false positive rate")]
    paras = []
    for i, (val, lab) in enumerate(stats):
        paras.append({"runs": [(val + "   ", {"c": GREEN, "b": True, "sz": 22}),
                               (lab, {"c": INK, "sz": 20})],
                      "sz": 20, "sb": 0 if i == 0 else 8})
    text(s, 7.30, 1.76, 5.30, 1.94, paras)
    card(s, LM, 4.22, CW, 1.30, fill=SOFT2, accent=CORAL)
    text(s, LM + 0.26, 4.30, CW - 0.52, 1.14, [
        {"runs": [("Our own gates: ", {"c": INK, "b": True, "sz": 20}),
                  ("relative F1 gain \\u2265 3.2277%   and   FPR reduction \\u2265 0.005",
                   {"c": INK, "sz": 20})], "sz": 20},
        {"runs": [("Both unmet. ", {"c": CORAL, "b": True, "sz": 20}),
                  ("A completed evaluation is not an authorisation.", {"c": INK, "sz": 20})],
         "sz": 20, "sb": 3},
        {"t": "Deltas are computed before rounding; the rounded pair alone would read +0.01583.",
         "sz": 20, "c": MUTED, "sb": 3}])
    band(s, 5.64, [("Complete evidence, insufficient authorisation.",
                    {"c": NAVY, "b": True})], h=0.66)
    foot(s, "The decision", "10")
    return s
'''

P13 = '''
def p13(prs):
    s = new_slide(prs)
    head(s, "12 \\u00b7 SCOPE", "What this does not show")
    items = [
        ("Not causal", "No isolation of TRIZ, FI or the controller: a workflow, not three "
                       "separable effects."),
        ("No matched search", "No matched-budget independent search, and no sensitivity study "
                              "over reference populations."),
        ("Not independent", "Related candidates and reused caches are not independent "
                            "confirmations."),
        ("Reproduction limited", "Sample-level predictions missing, one specification original "
                                 "missing."),
    ]
    for i, (name, body) in enumerate(items):
        x = LM + (i % 2) * 6.09
        y = 1.82 + (i // 2) * 1.94
        card(s, x, y, 5.80, 1.72, fill=SOFT, accent=CORAL)
        text(s, x + 0.26, y + 0.16, 5.28, 0.34,
             [{"t": name, "sz": 22, "b": True, "c": NAVY}])
        text(s, x + 0.26, y + 0.60, 5.28, 1.04,
             [{"t": body, "sz": 20, "c": INK, "ls": 1.04}])
    band(s, 5.78, [("Stated in the paper, ", {"c": INK, "b": True}),
                   ("not discovered in review.", {"c": INK})])
    foot(s, "Scope", "13")
    return s
'''

P14 = '''
def p14(prs):
    s = new_slide(prs)
    head(s, "13 \\u00b7 TAKE-AWAYS", "Three things worth keeping")
    items = [
        "Type the direction, not only the result: contradiction, principle and contract "
        "make the search auditable before the number exists.",
        "Treat functional information as an instrument tied to a fixed reference population, "
        "not as a universal score.",
        "Keep promotion separate from improvement: our most interesting candidate improved "
        "three measures and was still not promoted.",
    ]
    y = 1.84
    for i, it in enumerate(items):
        text(s, LM, y, 0.55, 0.72, [{"t": str(i + 1), "sz": 26, "b": True, "c": AMBER}])
        text(s, LM + 0.62, y + 0.02, 11.2, 0.92, [{"t": it, "sz": 22, "c": INK, "ls": 1.05}])
        y += 1.06
    shape(s, LM, 5.14, CW, 0.014, fill=LINE)
    text(s, LM, 5.26, 6.0, 0.72, [{"t": "Thank you.", "sz": 34, "b": True, "c": NAVY}])
    text(s, 6.90, 5.30, 5.70, 1.14, [
        {"t": "TRIZa: Guiding Agent Evolution", "sz": 20, "b": True, "c": INK},
        {"t": "Paper, Appendix A\\u2013H and the decision records travel with the proceedings.",
         "sz": 20, "c": MUTED, "ls": 1.04, "sb": 4}])
    foot(s, "Take-aways", "14")
    return s
'''

for name, body in (("p02", P02), ("p10", P10), ("p13", P13), ("p14", P14)):
    src = sub_fn(src, name, body)

FIX = [
 ("text(s, x + 0.26, y + 0.62, 5.28, 0.94,", "text(s, x + 0.26, y + 0.60, 5.28, 1.06,"),
 ("text(s, LM + 5.45, 3.18, 6.2, 0.34,", "text(s, LM + 5.45, 3.10, 6.2, 0.70,"),
 ("text(s, 8.05, 1.92, 4.55, 0.94,", "text(s, 8.05, 1.90, 4.55, 1.02,"),
 ("text(s, 8.05, 2.90, 4.55, 0.94,", "text(s, 8.05, 2.88, 4.55, 1.02,"),
 ("text(s, 8.05, 3.94, 4.55, 1.10,", "text(s, 8.05, 3.92, 4.55, 1.06,"),
]
for a, b in FIX:
    assert a in src, a
    src = src.replace(a, b)

p.write_text(src, encoding="utf-8", newline="\n")
print("PATCHED", len(src))
