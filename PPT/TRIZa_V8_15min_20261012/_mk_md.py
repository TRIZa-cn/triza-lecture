# -*- coding: utf-8 -*-
import sys
from pathlib import Path
D = Path(r"D:\doc\ZGCA\Lectures\TRIZ\TRIZa_V8_15min_20261012")
sys.path.insert(0, str(D))
from notes_v8b import NOTES

TITLES = ["Title", "The gap", "TRIZ in one minute", "Architecture", "Stage one",
          "The typed proposal", "Directions and procedure", "Measurement",
          "Paired outcomes", "The decision", "The controller", "Downstream",
          "Scope", "Take-aways"]
out = ["# TRIZa - 15-minute conference talk", "",
       "Verbatim speaking script, one block per slide. Read as written; the wording is already colloquial.",
       "",
       "14 slides, 1,758 words, about 13.0 minutes at 135 words per minute, which leaves room for pauses and questions.",
       ""]
for i, (t, n) in enumerate(zip(TITLES, NOTES), 1):
    out += ["## %02d  %s" % (i, t), "", n, ""]
(D / "talk_script_EN.md").write_text("\n".join(out), encoding="utf-8", newline="\n")
print("MD", (D / "talk_script_EN.md").stat().st_size, "slides", len(NOTES))
