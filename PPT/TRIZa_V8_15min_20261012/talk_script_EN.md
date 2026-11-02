# TRIZa - 15-minute conference talk

Verbatim speaking script, one block per slide. Read as written; the wording is already colloquial.

14 slides, 1,758 words, about 13.0 minutes at 135 words per minute, which leaves room for pauses and questions.

## 01  Title

Good morning. I am Wenyin Liu, from Guangdong University of Technology, and this is joint work with Yiheng Huang and Kai Wang. The claim of this talk is one sentence: TRIZ parameterisation, which supplies reasoning and direction, multiplied by LLM agents that solve automatically, which supplies execution, equals a self-evolving loop. In fifteen minutes I want to show what that multiplication actually looks like when it runs, and what it does not buy you.

## 02  The gap

Start with the alternative, because it is the default today. If you do not use TRIZ, the raw problem statement goes into the prompt, and a few rounds later you may well have runnable code. That is a real result. But three things are missing. The model's understanding of the problem stays shallow, because nothing forced it to name the trade-off. The solution is not reproducible, because the next run starts from a different phrasing. And you cannot say why this solution rather than another, so you cannot ablate it. Through TRIZ, every candidate traces to one principle and one pair of parameters, which makes it ablatable, comparable and reproducible. The difference is not the quality of the code. It is auditability. Two people can take the same candidate, read the same contract, and disagree about the evidence instead of about taste.

## 03  TRIZ in one minute

TRIZ in one minute, for those who have not used it. It comes from the analysis of a large patent corpus: inventive moves recur, and they cluster into a small set of principles. The entry point is a contradiction. You want more of A, and that costs you B. Not a wish, a pair. In our setting the pair is concrete: raise attack recall without raising false positives. Then come forty named principles, which are directions rather than recipes. Segmentation says split the decision; Local Quality says make behaviour depend on a declared context. And Ideality says judge the result by useful function over harm and cost. TRIZ never tells you which estimator to use. It tells you which class of change to try. That distinction is the reason the rest of the talk is about records rather than about prompts. A principle is a commitment you can check afterwards, and a prompt is not.

## 04  Architecture

Here is the loop: four stages in three layers. Layer one formulates the problem, a human brief, the contradiction, and a protocol that is frozen before anything is generated. Layer two is the core. The agent selects a principle, instantiates its contract, and produces a candidate. Layer three controls: promote, request revision, redirect to another principle, or stop. The knowledge bank keeps source-tagged mappings and prior outcomes, so a recommendation can say where it came from. And high-risk actions need a human operator with the required authority, which is a design decision, not a formality. Nothing in this diagram is a new model. The stages are ordinary engineering; what is unusual is that each stage leaves a record the next one is forced to consume.

## 05  Stage one

Stage one is about fixing the question before the answer. The brief names the target function and the constraints that cannot move. The contradiction extractor writes down the property we want to improve, the property that may get worse, and a testable criterion for each. Then the protocol is frozen: the split, the metric implementation, the evaluation procedure. That ordering is the whole point. If the measurement is decided after you see the result, no later claim is safe, no matter how good the number looks. In practice this is where most agent projects fail, not in the model. They begin with a measurement they can move, and discover the constraint only after the numbers arrive.

## 06  The typed proposal

The central object is the typed proposal. It is a tuple: the parent artifact, the contradiction, the selected principle, the declared operator, its arguments, and the frozen protocol. The candidate is that operator applied to the parent. The contract states four things: input and output types, permitted operations, protected conditions, and required checks. One rule matters more than the rest. A favourable result cannot retrospectively make an untyped edit TRIZ-guided. You commit to the direction before you see the number, and the record keeps every link, so someone else can re-run the decision. This is why we call the object typed. An untyped edit has no operator, so there is nothing to ablate and nothing to compare; you can only report that it worked.

## 07  Directions and procedure

Four directions we actually use, and the procedure that keeps them honest. Segmentation partitions decisions behind a stable interface. Local Quality assigns context-specific behaviour. Prior Action prepares reusable state before the critical path. Parameter Change moves a threshold or a weight. The procedure is four steps. State the property to improve, the one that may worsen, and what is protected. Retrieve principles with their source, and label a fallback as a fallback. State how the operation helps and how it harms. Record the principle and the arguments before evaluation, because a missing rationale cannot be supplied afterwards. None of this is specific to network intrusion detection. The same four steps applied to a retrieval system would name a different contradiction and reach a different set of principles.

## 08  Measurement

How a candidate is scored. Functional information asks how rare a configuration is among references that pass the same protocol at the same evidence level. Useful FI is minus log base two of the fraction reaching the candidate's useful level; harm FI is the same construction over false positives, recall loss, leakage and resource violations. The add-one correction keeps the estimate finite when nothing matches. It is not a confidence bound. Ideality then divides a lower bound on useful FI by an upper bound on harm FI plus a path cost. And certification is a separate quantity: the highest evidence level whose gates were actually passed, or minus one. A larger ratio cannot buy its way past a safety or replication gate. Both bounds come from the same reference population, and the protocol identifier travels with them, so two numbers are only comparable when that identifier matches.

## 09  Paired outcomes

Here are the paired outcomes, and I want to be careful about how to read them. Each row is a baseline and an observed candidate inside its own protocol: PowerDuck weighted-F1, plus two point eight seven points; CICIoT legacy macro-F1, plus four point two three; CIC-DDoS, plus zero point one five over all ten folds; and the formal candidate, plus zero point zero one five eight. Two cautions. The sizes are not a ranking across tasks, because the metrics differ. And PowerDuck's gain is against its paired baseline, not against the incumbent it had to beat, where the margin was zero point two two. Read the fourth row differently from the first three. It is the only row where the candidate was evaluated against gates that were written before the run, and it is the only row that failed them.

## 10  The decision

This is the result I would ask you to remember. The formal campaign completed forty-five family and seed pairs. Macro-F1 rose from zero point eight three two five six to zero point eight four eight three nine, a difference of zero point zero one five eight four, or one point nine percent relative. Recall rose, false positive rate fell, forty-three of forty-five pairs were positive, and all five seed blocks moved the same way. And the candidate was not promoted, because our own gates ask for three point two three percent relative gain and a zero point zero zero five FPR reduction. Complete evidence, insufficient authorisation.

## 11  The controller

The controller has to decide which candidates are functionally qualified, and that turns out to be a trade-off rather than a winner. On the right axis, specificity is the rejection rate for unqualified candidates. Safety-first sits far right and rejects almost everything, including qualified candidates. Calibrated V1 sits at the top left: it accepts far more, and its recall is much higher. Balanced V2 sits between them. We report this as a policy choice with stated consequences, and we also say that the two thousand and forty decisions come from overlapping subsets, so they are not independent replications. The honest reading is that we have not shown which policy is right. We have shown that the choice changes the outcome, and that it can be stated before the experiment instead of after.

## 12  Downstream

Downstream is where the loop said no, and it deserves a minute. V14 lifts macro-F1 from zero point eight three three to zero point eight seven five. V17 adds a little more, but its detector false positive rate nearly halves, which is the wrong direction. It passes its own champion rule, not strict balanced V2, and its recall stays two point four three points below the RF100 baseline. V18 was rejected under its retained contract. And notice the MCC panel: the value is higher than the same-fit V17 core and lower than the older cached V17. The number did not change. The reference did. For a review this is the more useful result, because a rejected candidate with a complete record tells you where the method stops, and a promoted one rarely does.

## 13  Scope

So what does this paper not show? It does not isolate the causal contribution of TRIZ, of functional information, or of the controller. The cases show an executable workflow, not three separable effects. There is no matched-budget independent search and no reference-population sensitivity study. Related candidates and reused caches are not independent confirmations. And reproduction is still limited: sample-level predictions are missing, one specification original is missing, and some historical provenance is incomplete. These are stated in the paper, not discovered in review. We would rather publish that list than let a reader discover it. It also tells you what the next experiment has to be, which is a matched-budget search against the same protocol.

## 14  Take-aways

Three things to take away. Type the direction, not only the result: a contradiction, a principle and a contract make the search auditable before the number exists. Treat functional information as an instrument tied to a fixed reference population, not as a universal score. And keep promotion separate from improvement, because in our runs the most interesting result is a candidate that improved three measures and was still not promoted. The paper, the appendix and the decision records are at the address on the right. Thank you, and I am happy to take questions. If you take one operational rule from this talk, take the ordering: contract first, then the number. Everything else in the method can be renegotiated; that order cannot.
