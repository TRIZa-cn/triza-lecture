# -*- coding: utf-8 -*-
import asyncio, json, re
from pathlib import Path
import edge_tts, miniaudio
from pptx import Presentation

D = Path(r"D:\doc\ZGCA\Lectures\TRIZ\PPT")
VOICE = "zh-CN-YunyangNeural"
RATE = "-25%"
DECKS = [("A", D / "TRIZ_W2A_矛盾解剖_15min.pptx"),
         ("B", D / "TRIZ_W2B_盒子与矛盾论_15min.pptx")]
OUT = D / "narration"


def clean(t):
    t = t.replace("**", "").replace("　", " ")
    t = re.sub(r"[ \t]+", " ", t)
    return "\n".join(x.strip() for x in t.splitlines() if x.strip())


def dur_of(p):
    d = miniaudio.decode_file(str(p))
    return d.num_frames / float(d.sample_rate)


async def main():
    info = {}
    total_all = 0.0
    for tag, deck in DECKS:
        prs = Presentation(str(deck))
        od = OUT / tag
        od.mkdir(parents=True, exist_ok=True)
        rows, tot = [], 0.0
        for i, s in enumerate(prs.slides, 1):
            txt = clean(s.notes_slide.notes_text_frame.text) if s.has_notes_slide else ""
            if not txt:
                rows.append({"slide": i, "chars": 0, "dur": 0.0, "file": None})
                continue
            f = od / ("s%02d.mp3" % i)
            if not f.exists() or f.stat().st_size < 1000:
                await edge_tts.Communicate(txt, VOICE, rate=RATE).save(str(f))
            d = dur_of(f)
            tot += d
            rows.append({"slide": i, "chars": len(txt), "dur": round(d, 2), "file": str(f)})
        info[tag] = {"deck": str(deck), "voice": VOICE, "rate": RATE,
                     "rows": rows, "total": round(tot, 1)}
        total_all += tot
        print("%s: slides=%d audio=%d chars=%d total=%.1fs (%.1f min)"
              % (tag, len(rows), sum(1 for r in rows if r["file"]),
                 sum(r["chars"] for r in rows), tot, tot / 60))
    (D / "_narration.json").write_text(json.dumps(info, ensure_ascii=False, indent=1),
                                       encoding="utf-8")
    print("all total %.1f min" % (total_all / 60))

asyncio.run(main())