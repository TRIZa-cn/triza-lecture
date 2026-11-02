# -*- coding: utf-8 -*-
import asyncio, json, re, wave
from pathlib import Path
from pptx import Presentation
import edge_tts, miniaudio

D = Path(r"D:\doc\ZGCA\Lectures\TRIZ\PPT")
V2 = D / "TRIZa_AI4Science_15min_v2.pptx"

prs = Presentation(str(V2))
s2 = prs.slides[1]
t = s2.notes_slide.notes_text_frame.text
assert "Latin acronym" in t, "marker missing"
s2.notes_slide.notes_text_frame.text = t.replace("It is a Latin acronym for", "It is a Russian acronym for", 1)
prs.save(str(V2))

prs = Presentation(str(V2))
notes = [s.notes_slide.notes_text_frame.text.strip() for s in prs.slides]
print("russian_fix", "Russian acronym" in notes[1], "| latin_left", "Latin acronym" in notes[1])
print("slides", len(notes))

SUBS = [
    (r"\bTRIZa\b", "Treesa"), (r"\bTRIZ\b", "Trees"),
    (r"\bCICIoT\b", "CIC IoT"), (r"\bCIC-DDoS\b", "CIC DDoS"),
    (r"\bweighted-F1\b", "weighted F one"), (r"\bmacro-F1\b", "macro F one"),
    (r"\bF1\b", "F one"), (r"\bFPR\b", "F P R"),
    (r"\bV1\b", "V one"), (r"\bV2\b", "V two"), (r"\bFI\b", "F I"),
]
tts = []
for n in notes:
    x = n
    for pat, rep in SUBS:
        x = re.sub(pat, rep, x)
    tts.append(x)

(D / "_notes_v2.json").write_text(json.dumps(
    {"file": str(V2), "slides": len(notes), "words": sum(len(re.findall(r"[A-Za-z][A-Za-z'-]*", n)) for n in notes), "notes": notes},
    ensure_ascii=False, indent=1), encoding="utf-8")
(D / "_notes_tts_v2.json").write_text(json.dumps(
    {"file": str(V2), "slides": len(tts), "words": 0, "notes": tts}, ensure_ascii=False, indent=1), encoding="utf-8")

OUT = D / "_edge_v2"
asyncio.run(edge_tts.Communicate(tts[1], "en-US-AndrewNeural", rate="-8%").save(str(OUT / "s02.mp3")))
dec = miniaudio.decode_file(str(OUT / "s02.mp3"), output_format=miniaudio.SampleFormat.SIGNED16)
with wave.open(str(OUT / "s02.wav"), "wb") as w:
    w.setnchannels(dec.nchannels); w.setsampwidth(2); w.setframerate(dec.sample_rate)
    w.writeframes(dec.samples.tobytes())
(OUT / "s02.mp3").unlink()

durs = []
for i in range(1, len(notes) + 1):
    with wave.open(str(OUT / ("s%02d.wav" % i))) as w:
        durs.append(round(w.getnframes() / w.getframerate(), 2))
tot = sum(durs)
print("s02_sec", durs[1], "| total_sec", round(tot, 1), "| total_min", round(tot / 60, 2))
(D / "_edge_v2_durs.json").write_text(json.dumps(
    {"durations": durs, "total": round(tot, 1), "voice": "en-US-AndrewNeural", "rate": "-8%"}), encoding="utf-8")
