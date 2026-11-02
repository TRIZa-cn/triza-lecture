# -*- coding: utf-8 -*-
import re
from pathlib import Path

D = Path(r"D:\doc\ZGCA\Lectures\TRIZ\TRIZa_V8_15min_20261012")
bp = D / "build_v8.py"
src = bp.read_text(encoding="utf-8").lstrip("\ufeff")

BODIES = {}

BODIES["p01"] = '''def p01(s):
    shape(s, 0, 0, 13.3333, 0.16, NAVY)
    shape(s, LM, 1.06, 0.06, 0.98, AMBER)
    text(s, LM + 0.26, 1.02, CW - 0.3, 0.46,
         [{"t": "AI4Science 2026  \\u00b7  Shenzhen, China  \\u00b7  October 23-25", "sz": 20, "c": MUTED, "spc": 0.6}])
    text(s, LM, 1.58, 11.3, 1.50, [
        {"t": "TRIZ^a: Guiding Agent Evolution", "sz": 36, "b": True, "c": NAVY, "ls": 1.06},
        {"t": "from Pattern Recognition to Solution Invention", "sz": 36, "b": True, "c": NAVY, "ls": 1.06},
    ])
    shape(s, LM, 3.22, 2.20, 0.05, AMBER)
    text(s, LM, 3.42, 10.6, 0.90, [
        {"t": "Wenyin Liu (GDUT)   \\u00b7   Yiheng Huang (BUPT)   \\u00b7   Kai Wang (Beijing Denglu Technology)",
         "sz": 20, "c": INK, "ls": 1.2},
    ])
    text(s, LM, 4.02, 10.6, 0.90, [
        {"t": "Evaluation feedback says whether a candidate worked.", "sz": 22, "c": BLUE, "ls": 1.2},
        {"t": "TRIZ says which edit to try next.", "sz": 22, "c": BLUE, "ls": 1.2},
    ])
    card(s, LM, 5.40, CW, 1.24, fill=SOFT, line=None, radius=0.08)
    text(s, LM + 0.30, 5.40, CW - 0.6, 1.24, [
        {"t": "Paper + supplementary appendix A-H + contracts, code, measurements and decision ledgers:",
         "sz": 20, "c": MUTED, "ls": 1.15},
        {"t": "triza.cn/triza-paper-exp-data2026", "sz": 20, "b": True, "c": NAVY, "ls": 1.15},
    ], anchor=MSO_ANCHOR.MIDDLE)
    foot(s, "TRIZa V8  \\u00b7  AI4Science 2026", 1)'''

BODIES["p03"] = '''def p03(s):
    head(s, "03 \\u00b7 THIS PAPER", "Three commitments, one workflow")
    items = [
        ("01", "A named direction", "Contradiction, principle and typed transformation, fixed before the candidate."),
        ("02", "A measurement of harm", "Functional information and Ideality, against a fixed reference population."),
        ("03", "An authorisation boundary", "A better metric is not permission to replace the running system."),
    ]
    for i, (n, h1, b1) in enumerate(items):
        x = LM + i * 4.07
        card(s, x, 1.90, 3.75, 3.58)
        chip(s, x + 0.28, 2.06, 0.78, 0.58, n, fill=NAVY)
        text(s, x + 0.28, 2.74, 3.20, 0.70,
             [{"t": h1, "sz": 22, "b": True, "c": NAVY, "ls": 1.1}])
        text(s, x + 0.28, 3.50, 3.22, 1.94,
             [{"t": b1, "sz": 20, "c": INK, "ls": 1.18}])
    band(s, 5.68, "Everything after this slide is evidence about how those three behave.",
         color=BLUE, fill=SOFT2)
    foot(s, "03 \\u00b7 This paper", 3)'''

BODIES["p04"] = '''def p04(s):
    head(s, "04 \\u00b7 BACKGROUND", "TRIZ gives names, not recipes")
    cols = [
        ("Contradiction", "Improving A costs you B. The pair is stated, not assumed."),
        ("40 principles", "Named directions: segmentation, local quality, prior action, parameter change."),
        ("Ideality", "Useful function over harm and cost. A ratio, not a score."),
    ]
    for i, (h1, b1) in enumerate(cols):
        x = LM + i * 4.18
        card(s, x, 2.00, 3.30, 2.38, fill=SOFT if i == 1 else WHITE)
        text(s, x + 0.24, 2.14, 2.85, 0.46, [{"t": h1, "sz": 20, "b": True, "c": NAVY}])
        text(s, x + 0.24, 2.68, 2.86, 1.64,
             [{"t": b1, "sz": 20, "c": INK, "ls": 1.16}])
    text(s, LM + 3.34, 3.00, 0.80, 0.44, [{"t": ">", "sz": 24, "b": True, "c": AMBER}],
         align=PP_ALIGN.CENTER)
    text(s, LM + 7.52, 3.00, 0.80, 0.44, [{"t": ">", "sz": 24, "b": True, "c": AMBER}],
         align=PP_ALIGN.CENTER)
    band(s, 4.70, "A principle does not specify an estimator or a value. The LLM picks it; the frozen protocol tests it.")
    text(s, LM, 5.70, CW, 0.90, [
        {"t": "Our contribution is not the theory. It is making the theory executable inside a loop that can run code and measure results.",
         "sz": 20, "c": MUTED, "ls": 1.2},
    ])
    foot(s, "04 \\u00b7 Background", 4)'''

BODIES["p06"] = '''def p06(s):
    head(s, "06 \\u00b7 STAGE ONE", "Fix the question before the answer")
    items = [
        ("HIL brief", "Target function and constraints that cannot move."),
        ("Contradiction", "Improving and worsening property, with testable criteria."),
        ("Frozen protocol", "Split, metric and procedure fixed before generation."),
        ("AKBank", "Source-tagged mappings, plus prior outcomes."),
    ]
    for i, (h1, b1) in enumerate(items):
        x = LM + (i % 2) * 5.83
        y = 1.88 + (i // 2) * 1.70
        card(s, x, y, 5.55, 1.56)
        chip(s, x + 0.24, y + 0.22, 0.58, 0.46, str(i + 1), fill=AMBER, sz=20)
        text(s, x + 0.96, y + 0.16, 4.40, 0.42,
             [{"t": h1, "sz": 20, "b": True, "c": NAVY}])
        text(s, x + 0.96, y + 0.62, 4.42, 0.88,
             [{"t": b1, "sz": 20, "c": INK, "ls": 1.14}])
    band(s, 5.40, "If the measurement is decided after the result, no later claim is safe.")
    foot(s, "06 \\u00b7 Stage one", 6)'''

BODIES["p10"] = '''def p10(s):
    head(s, "10 \\u00b7 MEASUREMENT", "Functional information: rarity, not quality")
    cards = [
        ("Useful FI", "I_U = -log2[ (1 + h_U) / (1 + N) ]",
         "h_U counts references at or above the candidate's useful level."),
        ("Harm FI", "I_H = -log2[ (1 + h_H) / (1 + N) ]",
         "h_H counts the same over false positives, recall loss, leakage, resources."),
    ]
    for i, (h1, f1, b1) in enumerate(cards):
        x = LM + i * 6.14
        card(s, x, 1.86, 5.75, 2.10)
        text(s, x + 0.26, 1.98, 5.2, 0.40, [{"t": h1, "sz": 20, "b": True, "c": MUTED}])
        text(s, x + 0.26, 2.40, 5.2, 0.46, [{"t": f1, "sz": 20, "b": True, "c": NAVY}])
        text(s, x + 0.26, 2.94, 5.24, 0.94, [{"t": b1, "sz": 20, "c": INK, "ls": 1.14}])
    band(s, 4.12, "FI is relative to a declared reference population and protocol.", color=AMBER)
    text(s, LM, 4.98, CW, 1.72, [
        {"t": "Change the population, the qualification rule or the protocol, and the comparison changes with it.",
         "sz": 20, "c": INK, "ls": 1.2},
        {"t": "The add-one correction keeps the estimate finite when nothing matches. It is not a confidence bound.",
         "sz": 20, "c": INK, "ls": 1.2, "sp": 10},
    ])
    foot(s, "10 \\u00b7 Measurement", 10)'''

BODIES["p11"] = '''def p11(s):
    head(s, "11 \\u00b7 MEASUREMENT", "Ideality, and a separate certificate")
    card(s, LM, 1.84, CW, 1.10, fill=SOFT, radius=0.06)
    text(s, LM + 0.30, 1.84, CW - 0.6, 1.10, [
        {"t": "J_l = I_U^LB  /  max( \\u03b5 , I_H^UB + \\u03bb_C \\u00b7 C_path )",
         "sz": 24, "b": True, "c": NAVY},
    ], anchor=MSO_ANCHOR.MIDDLE)
    card(s, LM, 3.14, 5.75, 1.72)
    text(s, LM + 0.26, 3.24, 5.2, 0.40, [{"t": "Certification level", "sz": 20, "b": True, "c": MUTED}])
    text(s, LM + 0.26, 3.66, 5.26, 1.12, [
        {"t": "L_cert = highest level whose validation requirements and hard gates were passed;  -1 if none.",
         "sz": 20, "c": INK, "ls": 1.14},
    ])
    card(s, 6.86, 3.14, 5.75, 1.72)
    text(s, 7.12, 3.24, 5.2, 0.40, [{"t": "What a ratio cannot buy", "sz": 20, "b": True, "c": MUTED}])
    text(s, 7.12, 3.66, 5.26, 1.12, [
        {"t": "Integrity, safety, replication and resource gates are mandatory. A larger J does not compensate.",
         "sz": 20, "c": INK, "ls": 1.14},
    ])
    band(s, 5.04, "A method specification, not a measured quantity.", color=CORAL, fill=C("FBEEEA"), h=0.66, sz=20)
    text(s, LM, 5.86, CW, 0.90, [
        {"t": "Historical runs use harm point estimates; the campaign uses construction steps over six as a cost proxy, not measured runtime or energy.",
         "sz": 20, "c": MUTED, "ls": 1.18},
    ])
    foot(s, "11 \\u00b7 Measurement", 11)'''

BODIES["p13"] = '''def p13(s):
    head(s, "13 \\u00b7 EVIDENCE", "Four groups, four different questions")
    cols = [(LM, 2.55, "Task"), (3.27, 2.95, "Comparison"), (6.22, 3.35, "Observed"), (9.57, 3.04, "Status")]
    for x, w, t in cols:
        text(s, x, 1.90, w, 0.42, [{"t": t, "sz": 20, "b": True, "c": MUTED}])
    shape(s, LM, 2.36, CW, 0.014, LINE)
    rows = [
        ("PowerDuck GOOSE", "paired weighted-F1", "0.96943 \\u2192 0.99817   +2.87 pp",
         "complete; no principle attribution", GREEN),
        ("CICIoT2023 legacy", "macro-F1, 9 families", "+4.23 pp",
         "complete; no run-level identity chain", GREEN),
        ("CIC-DDoS2019", "weighted-F1, 10 folds", "+0.15 pp",
         "complete; full registered scope", GREEN),
        ("Formal V2", "macro-F1, 45 pairs", "0.83256 \\u2192 0.84839   +0.01584",
         "complete; NOT promoted", CORAL),
        ("V17 / V18", "multi-metric sequence", "V17 passes a separate champion rule",
         "neither meets strict V2", CORAL),
    ]
    for i, (a, b1, c1, d1, col) in enumerate(rows):
        y = 2.42 + i * 0.86
        text(s, LM, y, 2.55, 0.82, [{"t": a, "sz": 20, "b": True, "c": NAVY, "ls": 1.12}],
             anchor=MSO_ANCHOR.MIDDLE)
        text(s, 3.27, y, 2.95, 0.82, [{"t": b1, "sz": 20, "c": INK, "ls": 1.12}],
             anchor=MSO_ANCHOR.MIDDLE)
        text(s, 6.22, y, 3.35, 0.82, [{"t": c1, "sz": 20, "c": INK, "ls": 1.12}],
             anchor=MSO_ANCHOR.MIDDLE)
        text(s, 9.57, y, 3.04, 0.82, [{"t": d1, "sz": 20, "b": True, "c": col, "ls": 1.12}],
             anchor=MSO_ANCHOR.MIDDLE)
        if i < 4:
            shape(s, LM, y + 0.86, CW, 0.010, LINE)
    band(s, 6.14, "\\u201ccomplete\\u201d means the stated evaluation finished, not that the candidate was promoted.",
         color=AMBER, fill=SOFT, h=0.72, sz=20)
    foot(s, "13 \\u00b7 Evidence", 13)'''

BODIES["p16"] = '''def p16(s):
    head(s, "16 \\u00b7 CONTROL", "Policy is a trade-off, not a winner",
         "510 overlapping subsets of nine held-out families, 2,040 decisions, four candidates each")
    px0, px1, py0, py1 = 2.30, 11.30, 1.98, 5.66
    xmin, xmax, ymin, ymax = 0.86, 1.005, 0.30, 1.00

    def X(v):
        return px0 + (v - xmin) / (xmax - xmin) * (px1 - px0)

    def Y(v):
        return py1 - (v - ymin) / (ymax - ymin) * (py1 - py0)

    for v in (0.4, 0.6, 0.8, 1.0):
        shape(s, px0, Y(v), px1 - px0, 0.008, LINE)
        text(s, 1.45, Y(v) - 0.19, 0.72, 0.38, [{"t": "%.1f" % v, "sz": 20, "c": GREY}],
             align=PP_ALIGN.RIGHT)
    for v in (0.88, 0.92, 0.96, 1.00):
        text(s, X(v) - 0.55, py1 + 0.06, 1.10, 0.42,
             [{"t": "%.2f" % v, "sz": 20, "c": GREY}], align=PP_ALIGN.CENTER)
    shape(s, px0, py0, 0.014, py1 - py0, GREY)
    shape(s, px0, py1, px1 - px0, 0.014, GREY)
    pts = [(0.963, 0.508, "1"), (0.998, 0.409, "2"), (0.989, 0.811, "3"), (0.883, 0.955, "4")]
    for sx, sy, n in pts:
        cx, cy = X(sx), Y(sy)
        sp = oval(s, cx - 0.22, cy - 0.22, 0.44, 0.44, NAVY)
        label(sp, [{"t": n, "sz": 20, "b": True, "c": WHITE}], margins=(0, 0, 0, 0))
    text(s, LM, 6.20, CW, 0.72, [
        {"t": "1 fixed 5%   \\u00b7   2 safety-first   \\u00b7   3 balanced V2   \\u00b7   4 calibrated V1        x: specificity     y: qualification F1",
         "sz": 20, "c": MUTED},
    ])
    foot(s, "16 \\u00b7 Control", 16)'''

BODIES["p18"] = '''def p18(s):
    head(s, "18 \\u00b7 SCOPE", "What this paper does not show")
    items = [
        ("No component-level causation", "TRIZ, FI and the controller are not separated by an ablation."),
        ("No matched-budget search", "Related candidates and reused caches are not independent."),
        ("No population sensitivity", "We do not vary the reference set or its qualification rule."),
        ("Incomplete provenance", "Sample-level predictions and one specification original are missing."),
    ]
    for i, (h1, b1) in enumerate(items):
        x = LM + (i % 2) * 5.83
        y = 1.82 + (i // 2) * 1.54
        card(s, x, y, 5.55, 1.40)
        text(s, x + 0.26, y + 0.12, 5.0, 0.40,
             [{"t": h1, "sz": 20, "b": True, "c": NAVY}])
        text(s, x + 0.26, y + 0.56, 5.06, 0.82,
             [{"t": b1, "sz": 20, "c": INK, "ls": 1.14}])
    band(s, 4.92, "These limits are stated in the paper, not discovered in review.",
         color=BLUE, fill=SOFT2, h=0.70)
    text(s, LM, 5.82, CW, 0.80, [
        {"t": "Reproduction rests on released records, a 108-file manifest and a no-training replay of 2,040 retained decisions.",
         "sz": 20, "c": MUTED, "ls": 1.16},
    ])
    foot(s, "18 \\u00b7 Scope", 18)'''

BODIES["p19"] = '''def p19(s):
    head(s, "19 \\u00b7 TAKE-AWAYS", "Three things to take away")
    rows = [
        ("1", "Type the direction, not only the result.",
         "A contradiction, a principle and a contract make the search auditable before the number exists."),
        ("2", "Functional information is an instrument, not a score.",
         "It is defined against a fixed reference population and a frozen protocol."),
        ("3", "Promotion is a separate decision from improvement.",
         "Our gates ask for more than a better metric, and they refuse candidates that improve one."),
    ]
    for i, (n, h1, b1) in enumerate(rows):
        y = 1.84 + i * 1.42
        text(s, LM, y - 0.04, 0.80, 0.72, [{"t": n, "sz": 40, "b": True, "c": AMBER}])
        text(s, LM + 0.96, y, 10.90, 0.48, [{"t": h1, "sz": 22, "b": True, "c": NAVY}])
        text(s, LM + 0.96, y + 0.50, 10.90, 0.78, [{"t": b1, "sz": 20, "c": INK, "ls": 1.14}])
    band(s, 6.06, "The most interesting result here is a candidate that improved three measures and was still not promoted.",
         color=CORAL, fill=C("FBEEEA"), h=0.66, sz=20)
    foot(s, "19 \\u00b7 Take-aways", 19)'''

for name, body in BODIES.items():
    pat = re.compile(r"def %s\(s\):.*?(?=\n\n# -{10,}|\n\ndef main\()" % name, re.S)
    m = pat.search(src)
    assert m, name
    src = src[:m.start()] + body + src[m.end():]

bp.write_text(src, encoding="utf-8", newline="\n")
print("PATCHED", len(BODIES))
