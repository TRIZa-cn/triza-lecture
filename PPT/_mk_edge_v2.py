# -*- coding: utf-8 -*-
import asyncio, json, re, wave
from pathlib import Path
import edge_tts, miniaudio

D = Path(r"D:\doc\ZGCA\Lectures\TRIZ\PPT")
OUT = D / "_edge_v2"
OUT.mkdir(exist_ok=True)
for f in list(OUT.glob("s*.wav")) + list(OUT.glob("*.mp3")):
    f.unlink()

d = json.loads((D / "_notes_v2.json").read_text(encoding="utf-8"))
notes = d["notes"]

SUBS = [
    (r"\bTRIZa\b", "Treesa"),
    (r"\bTRIZ\b", "Trees"),
    (r"\bCICIoT\b", "CIC IoT"),
    (r"\bCIC-DDoS\b", "CIC DDoS"),
    (r"\bweighted-F1\b", "weighted F one"),
    (r"\bmacro-F1\b", "macro F one"),
    (r"\bF1\b", "F one"),
    (r"\bFPR\b", "F P R"),
    (r"\bV1\b", "V one"),
    (r"\bV2\b", "V two"),
    (r"\bFI\b", "F I"),
]
tts, counts = [], {}
for n in notes:
    t = n
    for pat, rep in SUBS:
        t, k = re.subn(pat, rep, t)
        if k:
            counts[pat] = counts.get(pat, 0) + k
    tts.append(t)
print("substitutions", counts)

(D / "_notes_tts_v2.json").write_text(json.dumps(
    {"file": str(D / "TRIZa_AI4Science_15min_v2.pptx"), "slides": len(tts),
     "words": d["words"], "notes": tts}, ensure_ascii=False, indent=1), encoding="utf-8")

VOICE = "en-US-AndrewNeural"
RATE = "-8%"

async def main():
    for i, t in enumerate(tts, 1):
        await edge_tts.Communicate(t, VOICE, rate=RATE).save(str(OUT / ("s%02d.mp3" % i)))

asyncio.run(main())

durs = []
for i in range(1, len(tts) + 1):
    mp3 = OUT / ("s%02d.mp3" % i)
    dec = miniaudio.decode_file(str(mp3), output_format=miniaudio.SampleFormat.SIGNED16)
    with wave.open(str(OUT / ("s%02d.wav" % i)), "wb") as w:
        w.setnchannels(dec.nchannels); w.setsampwidth(2); w.setframerate(dec.sample_rate)
        w.writeframes(dec.samples.tobytes())
    durs.append(round(dec.num_frames / dec.sample_rate, 2))
    mp3.unlink()

tot = sum(durs)
print("voice", VOICE, "rate", RATE)
print("durations", durs)
print("total_sec", round(tot, 1), "total_min", round(tot / 60, 2), "wpm", round(d["words"] / (tot / 60), 1))
(D / "_edge_v2_durs.json").write_text(json.dumps({"durations": durs, "total": round(tot, 1), "voice": VOICE, "rate": RATE}), encoding="utf-8")
