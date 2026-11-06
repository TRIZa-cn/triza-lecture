# -*- coding: utf-8 -*-
"""按读音表替换 notes 里的字词，产出可直接喂 TTS 的文本 + 一份报告。

覆盖两类问题，同一套机制：
  - 中文多音字（polyphone.json）：调低 → 条低、调一次 → 调整一次
  - 英文缩写（abbreviations.json）：TRIZ → trees、P06 → 第 6 页

原则（见 polyphone.md / abbreviations.md）：
  - 不改 deck、不改 notes 原文，只产出"合成用的那一份文本"
  - 精确片段匹配，不用位置、不用正则
  - **长串先替换**：TRIZ 先跑会把 TRIZa 吃成 treesa
  - missed 非空 -> 退出码 1，不许合成

用法：
    python apply_polyphone.py --deck <pptx> \
        --table polyphone.json --table abbreviations.json \
        --out-speech speech.json --out-report report.json
"""
import argparse
import json
import os
import sys


def load_notes(deck_path):
    from pptx import Presentation
    prs = Presentation(deck_path)
    out = []
    for i, s in enumerate(prs.slides, start=1):
        t = ""
        if s.has_notes_slide:
            t = s.notes_slide.notes_text_frame.text
        out.append({"page": i, "text": t})
    return out


def load_tables(paths):
    entries, per_table = [], {}
    for p in paths:
        tab = json.load(open(p, encoding="utf-8"))
        items = tab.get("entries", tab if isinstance(tab, list) else [])
        per_table[os.path.basename(p)] = len(items)
        for e in items:
            e = dict(e)
            e["_table"] = os.path.basename(p)
            entries.append(e)
    # 长的 before 先替换：否则 TRIZ 会把 TRIZa 吃成 treesa
    entries.sort(key=lambda e: -len(e.get("before", "")))
    return entries, per_table


def apply(entries, notes):
    applied, missed, skipped = [], [], []
    for e in entries:
        if not e.get("speech", True):
            skipped.append({"id": e.get("id"), "before": e.get("before"),
                            "why": "speech=false：不进配音"})
            continue
        hits = 0
        for n in notes:
            if e["before"] in n["text"]:
                n["text"] = n["text"].replace(e["before"], e["after"])
                hits += 1
        rec = {"id": e.get("id"), "table": e.get("_table"), "kind": e.get("kind", "char"),
               "before": e["before"], "after": e["after"], "hits": hits,
               "status": e.get("status", "")}
        (applied if hits else missed).append(rec)
    return applied, missed, skipped


def leftover_watch(words, notes):
    """替换后仍出现的多音字，提示人工复核（只提示，不报错）。

    「调整」里的「调」虽然还在，但读音已经对了，跳过。
    不跳过的话每次都会误报，提示变噪音，人就不看了。
    """
    SAFE_NEXT = {"调": "整"}
    out = []
    for w in words:
        for n in notes:
            idx = 0
            while True:
                pos = n["text"].find(w, idx)
                if pos < 0:
                    break
                idx = pos + 1
                if n["text"][pos + 1:pos + 2] in SAFE_NEXT.get(w, ""):
                    continue
                lo, hi = max(0, pos - 12), min(len(n["text"]), pos + 12)
                out.append({"char": w, "page": n["page"],
                            "ctx": n["text"][lo:hi].replace("\n", " ")})
    return out


def abbr_watch(notes):
    """替换后 notes 里还剩哪些连续大写串（可能是没进表的缩写）。"""
    import re
    PAT = re.compile(r"TRIZa|[A-Z][A-Z0-9]{1,7}")
    out = []
    for n in notes:
        for m in PAT.finditer(n["text"]):
            lo, hi = max(0, m.start() - 12), min(len(n["text"]), m.end() + 12)
            out.append({"token": m.group(0), "page": n["page"],
                        "ctx": n["text"][lo:hi].replace("\n", " ")})
    return out


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--deck", required=True)
    ap.add_argument("--table", required=True, action="append",
                    help="可重复；多张表合并，长串优先")
    ap.add_argument("--out-speech", default="speech.json")
    ap.add_argument("--out-report", default="report.json")
    ap.add_argument("--watch", default="调", help="替换后还要盯的中文字，逗号分隔")
    a = ap.parse_args()

    entries, per_table = load_tables(a.table)
    notes = load_notes(a.deck)
    applied, missed, skipped = apply(entries, notes)
    watch = [w for w in a.watch.split(",") if w]
    leftover = leftover_watch(watch, notes)
    left_abbr = abbr_watch(notes)

    report = {
        "deck": os.path.basename(a.deck),
        "tables": per_table,
        "checked": len(entries),
        "applied": len(applied),
        "missed": missed,
        "skipped": skipped,
        "leftover_char": leftover,
        "leftover_abbr": left_abbr,
        "detail_applied": applied,
        "speech_chars": sum(len(n["text"]) for n in notes),
        "verdict": "OK" if not missed else "STOP",
    }
    json.dump(report, open(a.out_report, "w", encoding="utf-8"),
              ensure_ascii=False, indent=1)
    json.dump(notes, open(a.out_speech, "w", encoding="utf-8"),
              ensure_ascii=False, indent=1)

    print(json.dumps({k: report[k] for k in
                      ("deck", "tables", "checked", "applied", "missed", "skipped",
                       "leftover_char", "leftover_abbr", "verdict")},
                     ensure_ascii=False, indent=1))
    if missed:
        print("\n[STOP] 有 %d 条没匹配到：讲稿改过了，表过期了。修表再合成。"
              % len(missed), file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
