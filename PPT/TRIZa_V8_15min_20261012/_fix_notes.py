# -*- coding: utf-8 -*-
import sys, pprint
from pathlib import Path
D = Path(r"D:\doc\ZGCA\Lectures\TRIZ\TRIZa_V8_15min_20261012")
sys.path.insert(0, str(D))
import notes_v8b
N = list(notes_v8b.NOTES)

REP = [
 ("The claim of the talk is on the slide, and it is one sentence:",
  "The claim of this talk is one sentence:"),
 ("This is the slide I would ask you to remember.",
  "This is the result I would ask you to remember."),
 ("and it is worth one slide.", "and it deserves a minute."),
]
for a, b in REP:
    N = [x.replace(a, b) for x in N]

ADD = {
2: "That distinction is the reason the rest of the talk is about records rather than about "
   "prompts. A principle is a commitment you can check afterwards, and a prompt is not.",
4: "In practice this is where most agent projects fail, not in the model. They begin with a "
   "measurement they can move, and discover the constraint only after the numbers arrive.",
6: "None of this is specific to network intrusion detection. The same four steps applied to a "
   "retrieval system would name a different contradiction and reach a different set of principles.",
12: "We would rather publish that list than let a reader discover it. It also tells you what the "
    "next experiment has to be, which is a matched-budget search against the same protocol.",
}
for i, t in ADD.items():
    N[i] = N[i] + " " + t

src = ("# -*- coding: utf-8 -*-\n"
       "\"\"\"Verbatim read-aloud notes for the TRIZa 15-minute talk (14 slides).\"\"\"\n\n"
       "NOTES = " + pprint.pformat(N, width=100) + "\n")
(D / "notes_v8b.py").write_text(src, encoding="utf-8", newline="\n")
w = sum(len(x.split()) for x in N)
print("slides", len(N), "words", w, "min@135", round(w / 135.0, 1))
print("meta", [i for i, x in enumerate(N) if "slide" in x.lower() or "speaker" in x.lower()])
print("nonascii", sorted({c for x in N for c in x if ord(c) > 127}))
