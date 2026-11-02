# -*- coding: utf-8 -*-
import json, re
from pathlib import Path

BASE = Path(r"D:\doc\ZGCA\Lectures\TRIZ")
data = json.loads((BASE / "_notes_dump.json").read_text(encoding="utf-8"))
notes = data["notes"]

n_triz = sum(len(re.findall(r"\bTRIZ\b", n)) for n in notes)
n_triza = sum(len(re.findall(r"\bTRIZa\b", n)) for n in notes)
print("occurrences  TRIZ =", n_triz, " TRIZa =", n_triza)

tts = [re.sub(r"\bTRIZ\b", "Trees", n) for n in notes]
left = sum(len(re.findall(r"\bTRIZ\b", n)) for n in tts)
print("leftover TRIZ in tts text =", left)

out = {"file": data["file"], "slides": len(tts), "notes": tts,
       "words": sum(len(n.split()) for n in tts)}
(BASE / "_notes_tts.json").write_text(json.dumps(out, ensure_ascii=False, indent=1), encoding="utf-8")

for i, (a, b) in enumerate(zip(notes, tts), 1):
    if a != b:
        j = a.find("Trees")
        print("  s%02d: ...%s..." % (i, a[max(0, j-45):j+35].replace("\n", " ")))
