# -*- coding: utf-8 -*-
import io, re, sys
from pathlib import Path

D = Path(r"D:\doc\ZGCA\Lectures\TRIZ\TRIZa_V8_15min_20261012")
bp = D / "build_v8.py"
src = bp.read_text(encoding="utf-8")

old = re.search(r'    steps = \["propose a candidate".*?anchor=MSO_ANCHOR\.MIDDLE\)\n(?=    card\(s, 6\.55)', src, re.S)
assert old, "p02 block not found"
new = '''    steps = ["propose a candidate", "execute it", "evaluate it", "propose again"]
    for i, t in enumerate(steps):
        y = 2.46 + i * 0.56
        sp = oval(s, LM + 0.30, y + 0.06, 0.42, 0.42, SOFT2)
        label(sp, [{"t": str(i + 1), "sz": 20, "b": True, "c": BLUE}], margins=(0, 0, 0, 0))
        text(s, LM + 0.86, y, 4.3, 0.54, [{"t": t, "sz": 20, "c": INK}],
             anchor=MSO_ANCHOR.MIDDLE)
    text(s, LM + 0.30, 4.74, 5.0, 0.42,
         [{"t": "and each step is a guess", "sz": 20, "b": True, "c": CORAL}],
         anchor=MSO_ANCHOR.MIDDLE)
'''
src = src[:old.start()] + new + src[old.end():]
bp.write_text(src, encoding="utf-8")

sys.path.insert(0, str(D))
import notes_v8
n = list(notes_v8.NOTES)
n[11] = ("Traceability is what makes the claim checkable rather than rhetorical. The schema links "
         "parent and candidate identity, the contradiction, the principle, the operator arguments, "
         "the edit, the protocol fingerprint, the metrics, the gates, and the decision. A record "
         "missing any of those links is not treated as evidence. Where no matrix hit exists, the "
         "recommendation is recorded as a fallback, and it is not later rewritten as a matrix match. "
         "Stage labels also stay at their own stage. That discipline is what makes a record usable: "
         "a reader can see which decision was made, on which evidence, and whether the comparison "
         "was against the paired baseline or against the incumbent.")
with io.open(D / "notes_v8.py", "w", encoding="utf-8", newline="\n") as f:
    f.write('# -*- coding: utf-8 -*-\n"""Verbatim read-aloud notes for the TRIZa V8 15-minute conference talk."""\n\nNOTES = [\n')
    for i, t in enumerate(n):
        f.write("    # %02d\n    %r,\n" % (i + 1, t))
    f.write("]\n")

bad = sorted({ch for ch in src if ord(ch) > 127})
print("BUILD_NONASCII", [(ch, hex(ord(ch))) for ch in bad])
print("NOTES", len(n), "words", sum(len(x.split()) for x in n))
