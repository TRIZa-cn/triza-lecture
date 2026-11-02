# -*- coding: utf-8 -*-
import json, wave, hashlib, zipfile
from pathlib import Path
D = Path(r"D:\doc\ZGCA\Lectures\TRIZ\PPT")
z = D / "TRIZa_15min_handoff.zip"
p = D / "TRIZa_AI4Science_15min_v2.pptx"
f = D / "TRIZa_AI4Science_15min_v2.pdf"
w = D / "TRIZa_AI4Science_15min_v2_Andrew_8pct_22k_mono.wav"

def sha(x):
    h = hashlib.sha256()
    with open(x, "rb") as fh:
        for b in iter(lambda: fh.read(1 << 20), b""):
            h.update(b)
    return h.hexdigest()

zf = zipfile.ZipFile(z)
names = zf.namelist()
inner = {n.split("/")[-1]: n for n in names}
src_hash = sha(p)
pkg_hash = hashlib.sha256(zf.read(inner["TRIZa_AI4Science_15min_v2.pptx"])).hexdigest()
with wave.open(str(w)) as h:
    sec = h.getnframes() / h.getframerate()

print(json.dumps({
  "zip_exists": z.exists(),
  "zip_mb": round(z.stat().st_size / 1048576, 1),
  "entries": len(names),
  "entry_names": names,
  "pptx_hash_match": src_hash == pkg_hash,
  "wav_sec": round(sec, 2),
  "sources_unchanged": p.stat().st_mtime < z.stat().st_mtime and f.stat().st_mtime < z.stat().st_mtime and w.stat().st_mtime < z.stat().st_mtime,
}, ensure_ascii=False))
