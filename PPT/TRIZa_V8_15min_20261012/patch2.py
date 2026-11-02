# -*- coding: utf-8 -*-
import io, sys
from pathlib import Path
D = Path(r"D:\doc\ZGCA\Lectures\TRIZ\TRIZa_V8_15min_20261012")
sys.path.insert(0, str(D))
import notes_v8
n = list(notes_v8.NOTES)
n[14] = n[14].replace("This is the slide I would ask you to remember.", "This is the result I would ask you to remember.")
n[19] = n[19].replace("is at the URL on the slide, in English and in Chinese", "is public, in English and in Chinese, at the address you see here")
with io.open(D / "notes_v8.py", "w", encoding="utf-8", newline="\n") as f:
    f.write('# -*- coding: utf-8 -*-\n"""Verbatim read-aloud notes for the TRIZa V8 15-minute conference talk."""\n\nNOTES = [\n')
    for i, t in enumerate(n):
        f.write("    # %02d\n    %r,\n" % (i + 1, t))
    f.write("]\n")
print("PATCHED_NOTES", len(n), sum(len(x.split()) for x in n))
