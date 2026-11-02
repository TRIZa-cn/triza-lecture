# -*- coding: utf-8 -*-
import asyncio, json, re, wave
from pathlib import Path
import edge_tts, miniaudio

BASE = Path(r"D:\doc\ZGCA\Lectures\TRIZ")
OUT = BASE / "_edge"
for f in list(OUT.glob("s*.wav")) + list(OUT.glob("*.mp3")):
    f.unlink()

d = json.loads((BASE / "_notes_tts.json").read_text(encoding="utf-8"))
notes = [re.sub(r"\bTRIZa\b", "Treesa", n) for n in d["notes"]]
VOICE = "en-US-AndrewNeural"
RATE = "-8%"

async def main():
    for i, t in enumerate(notes, 1):
        await edge_tts.Communicate(t, VOICE, rate=RATE).save(str(OUT / ("s%02d.mp3" % i)))

asyncio.run(main())

durs = []
for i in range(1, len(notes) + 1):
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
(BASE / "_edge_durs.json").write_text(json.dumps({"durations": durs, "total": round(tot, 1), "voice": VOICE, "rate": RATE}), encoding="utf-8")
