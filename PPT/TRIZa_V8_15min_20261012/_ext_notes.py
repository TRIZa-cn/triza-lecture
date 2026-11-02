# -*- coding: utf-8 -*-
import sys, pprint
from pathlib import Path
D = Path(r"D:\doc\ZGCA\Lectures\TRIZ\TRIZa_V8_15min_20261012")
sys.path.insert(0, str(D))
import notes_v8b
N = list(notes_v8b.NOTES)

EX = {
1: "The difference is not the quality of the code. It is auditability. Two people can take the "
   "same candidate, read the same contract, and disagree about the evidence instead of about taste.",
3: "Nothing in this diagram is a new model. The stages are ordinary engineering; what is unusual is "
   "that each stage leaves a record the next one is forced to consume.",
5: "This is why we call the object typed. An untyped edit has no operator, so there is nothing to "
   "ablate and nothing to compare; you can only report that it worked.",
7: "Both bounds come from the same reference population, and the protocol identifier travels with "
   "them, so two numbers are only comparable when that identifier matches.",
8: "Read the fourth row differently from the first three. It is the only row where the candidate was "
   "evaluated against gates that were written before the run, and it is the only row that failed them.",
10: "The honest reading is that we have not shown which policy is right. We have shown that the "
    "choice changes the outcome, and that it can be stated before the experiment instead of after.",
11: "For a review this is the more useful result, because a rejected candidate with a complete record "
    "tells you where the method stops, and a promoted one rarely does.",
13: "If you take one operational rule from this talk, take the ordering: contract first, then the "
    "number. Everything else in the method can be renegotiated; that order cannot.",
}
for i, t in EX.items():
    N[i] = N[i] + " " + t
src = "# -*- coding: utf-8 -*-\n\"\"\"Verbatim read-aloud notes for the TRIZa 15-minute talk (14 slides).\"\"\"\n\nNOTES = " + pprint.pformat(N, width=100) + "\n"
(D / "notes_v8b.py").write_text(src, encoding="utf-8", newline="\n")
import importlib
importlib.reload(notes_v8b)
w = sum(len(x.split()) for x in N)
print("slides", len(N), "words", w, "min@135", round(w / 135.0, 1))
bad = [i for i, x in enumerate(N) if "slide" in x.lower() or "speaker" in x.lower()]
print("meta", bad)
