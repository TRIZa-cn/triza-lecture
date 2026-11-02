# -*- coding: utf-8 -*-
import json, shutil, re
from pathlib import Path
from pptx import Presentation

D = Path(r"D:\doc\ZGCA\Lectures\TRIZ\PPT")
SRC = D / "TRIZa_AI4Science_15min.pptx"
DST = D / "TRIZa_AI4Science_15min_v2.pptx"

NEW = {
2: "So what is TRIZ? It is a Latin acronym for the Theory of Inventive Problem Solving, developed by Genrich Altshuller, based on analyses of a large patent corpus. He found that inventive moves recur, and they cluster into a small set of principles. The entry point is a contradiction of a pair of parameters. You want more of A, and that costs you B. You cannot wish for only one side; you get a contradicting pair of parameters. In our setting, the pair is concrete: raise attack recall without raising false positives. Then come forty named principles, which are directions rather than recipes. Segmentation says split the decision; Local Quality says make behaviour depend on a declared context. And Ideality says judge the result by useful function over harm and cost. TRIZ never tells you which estimator to use. It tells you which class of change to try. That distinction is the reason the rest of the talk is about records rather than about prompts. A principle is a commitment you can check afterwards, and a prompt is not.",
3: "Then, why TRIZ? Let's start with the alternative: ChatGPT or other LLM plus Agent, because they are the default today. If you do not use TRIZ, the raw problem statement goes into the prompt, and a few rounds later you may well have runnable code. That is a real result. But three things are missing. The model's understanding of the problem stays shallow, because nothing forced it to name the trade-off. The solution is not reproducible, because the next run starts from a different phrasing. And you cannot say why this solution rather than another, so you cannot ablate it. But with TRIZ, every candidate traces to one principle and one pair of parameters, which makes it interpretable, ablatable, comparable and reproducible. The difference is not the quality of the code. It is auditability.",
7: "We illustrate our methods with four exemplar directions guided by four principles, and the procedure that keeps them honest. Segmentation partitions decisions behind a stable interface. Local Quality assigns context-specific behaviour. Prior Action prepares reusable state before the critical path. Parameter Change moves a threshold or a weight. The procedure is four steps. 1. State the property or parameter to improve, the one that may worsen, and what is protected. 2. Retrieve principles with their source, and label a fallback as a fallback; that is, if no matrix hit, just say so. State how the operation helps and how it harms. Record the principle and the arguments before evaluation, because a missing rationale cannot be supplied afterwards. None of this is specific to network intrusion detection. The same four steps applied to a retrieval system would name a different contradiction and reach a different set of principles.",
9: "We used three cyber attack detection tasks over three datasets to evaluate how our method works. After many rounds of evolutions, all three tasks showed an increase in the performance metrics, some larger, some smaller. Next we explain these comparisons in more detail. Here are the paired outcomes. Each row is a baseline and an observed candidate inside its own protocol: PowerDuck weighted-F1, plus two point eight seven points; CICIoT legacy macro-F1, plus four point two three; CIC-DDoS, plus zero point one five over all ten folds; and the formal candidate, plus zero point zero one five eight. Two cautions. The sizes are not a ranking across tasks, because the metrics differ. And PowerDuck's gain is against its paired baseline, not against the incumbent it had to beat, where the margin was zero point two two. Read the fourth row differently from the first three. It is the only row where the candidate was evaluated against gates that were written before the run, and it is the only row that failed them.",
11: "The controller has to decide which candidates are functionally qualified, and that turns out to be a trade-off rather than a winner. On the right axis, specificity is the rejection rate for unqualified candidates. Safety-first sits far right and rejects almost everything, including qualified candidates. Calibrated V1 sits at the top left: it accepts far more, and its recall is much higher. Balanced V2 sits between them. We report this as a policy choice with stated consequences, and we also say that the two thousand and forty decisions come from overlapping subsets, so they are not independent replications. We have not shown which policy is right; we have shown that the choice changes the outcome, and that it can be stated before the experiment instead of after.",
12: "So what does this paper not show? It does not isolate the causal contribution of TRIZ, of functional information, or of the controller. The cases show an executable workflow, not three separable effects. There is no matched-budget independent search and no reference-population sensitivity study, which are our future work.",
13: "Our main contributions are: TRIZ-guided solution generation, and Operational Ideality based on functional information. Next, three conclusions: 1. Type the direction, not only the result: a contradiction, a principle and a contract make the search auditable before the number exists. 2. Treat functional information as an instrument tied to a fixed reference population, not as a universal score. 3. Keep promotion separate from improvement, because in our runs the most interesting result is a candidate that improved three measures and was still not promoted. We also put the appendix and the decision records on our website. That's all of my presentation. I am happy to take questions. Thanks.",
}

src_mtime = SRC.stat().st_mtime
shutil.copy2(SRC, DST)
prs = Presentation(str(DST))
old = []
for i, s in enumerate(prs.slides, 1):
    old.append(s.notes_slide.notes_text_frame.text.strip() if s.has_notes_slide else "")

for i, s in enumerate(prs.slides, 1):
    if i in NEW:
        s.notes_slide.notes_text_frame.text = NEW[i]
prs.save(str(DST))

prs2 = Presentation(str(DST))
new = []
for s in prs2.slides:
    new.append(s.notes_slide.notes_text_frame.text.strip())

changed = [i + 1 for i in range(len(new)) if new[i] != old[i]]
words = sum(len(re.findall(r"[A-Za-z][A-Za-z'-]*", t)) for t in new)

KEYS = ["two point eight seven", "four point two three", "zero point one five",
        "zero point zero one five eight", "forty-five", "one point nine percent",
        "three point two percent", "zero point zero zero five", "two thousand and forty"]
found = sum(1 for k in KEYS if any(k in t for t in new))

print("v2_slides", len(new), "words", words, "changed", changed)
print("key_numbers_found", found, "of", len(KEYS))
for i in changed:
    print("--- slide %d ---" % i)
    print("OLD:", old[i-1][:400])
    print("NEW:", new[i-1][:400])

(D / "_notes_v2.json").write_text(json.dumps(
    {"file": str(DST), "slides": len(new), "words": words, "notes": new}, ensure_ascii=False, indent=1), encoding="utf-8")
print("src_mtime_preserved", abs(SRC.stat().st_mtime - src_mtime) < 0.001)
