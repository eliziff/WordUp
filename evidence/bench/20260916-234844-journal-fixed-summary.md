# Benchmark 20260916-234844-journal-fixed

model gpt-5.6-luna effort xhigh; exe build/wordup.exe; timeout 1800s per run

| task | arm | runs | passed | median wall s | median commands | median input tok | median output tok | median suite ms |
|---|---|---:|---:|---:|---:|---:|---:|---:|
| journal-mcgill-lj-health | components | 2 | 2 | 976.3 | 51.0 | 4778577.0 | 47155.5 | 3027.2 |
| journal-mcgill-lj-health | core | 2 | 2 | 1352.7 | 61.0 | 6653857.0 | 62855.0 | 2844.8 |
| journal-osgoode-hall-lj | components | 2 | 2 | 1444.4 | 66.0 | 2141512.0 | 25236.5 | 3024.9 |
| journal-osgoode-hall-lj | core | 2 | 2 | 953.7 | 44.5 | 5054538.5 | 44815.5 | 2783.9 |
| journal-ubc-l-rev | components | 2 | 2 | 1641.2 | 74.0 | 3682322.0 | 26905.0 | 2832.2 |
| journal-ubc-l-rev | core | 2 | 1 | 1004.6 | 51.5 | 7014626.5 | 46724.5 | 3171.0 |

## Runs

- journal-mcgill-lj-health / components / run-2: PASS; wall 925.0 s; commands 43; timed_out False; components []
- journal-mcgill-lj-health / components / run-1: PASS; wall 1027.6 s; commands 59; timed_out False; components ['document.stories']
- journal-mcgill-lj-health / core / run-2: PASS; wall 1030.6 s; commands 49; timed_out False; components []
- journal-mcgill-lj-health / core / run-1: PASS; wall 1674.8 s; commands 73; timed_out False; components []
- journal-osgoode-hall-lj / core / run-2: PASS; wall 832.1 s; commands 41; timed_out False; components []
- journal-osgoode-hall-lj / core / run-1: PASS; wall 1075.2 s; commands 48; timed_out False; components []
- journal-osgoode-hall-lj / components / run-2: PASS; wall 1088.7 s; commands 60; timed_out False; components ['operation.safe-edit']
- journal-ubc-l-rev / core / run-1: PASS; wall 945.3 s; commands 44; timed_out False; components []
- journal-osgoode-hall-lj / components / run-1: PASS; wall 1800.1 s; commands 72; timed_out True; components None
- journal-ubc-l-rev / core / run-2: FAIL hidden native suite; wall 1063.9 s; commands 59; timed_out False; components []
- journal-ubc-l-rev / components / run-1: PASS; wall 1482.3 s; commands 57; timed_out False; components []
- journal-ubc-l-rev / components / run-2: PASS; wall 1800.0 s; commands 91; timed_out True; components None

## Reading (2026-09-17)

Context. The earlier journal re-run `20260916-230047-journal-rerun` (same executable, six workers) scored 0/12 on both arms; every failure was a grading defect, not an agent failure: the hidden suites read the body paragraph's whole range (which also holds the footnote reference mark, so Word reports a mixed font) and asserted `UpdateStylesOnOpen` although the prompt never asked for it. `tools/bench/make_journal_tasks.py` now reads the paragraph's first word and the note's first text word and states the requirement in the prompt; the three regenerated positive-control templates pass the corrected suites. This run uses the corrected tasks with workspaces kept.

Result. Hidden native pass rate: components 6/6, core 5/6 (the one core failure is genuine: the agent's macro curled the opening quote as a closing one). Median input tokens per task are lower on the components arm for all three tasks (4.78M vs 6.65M, 2.14M vs 5.05M, 3.68M vs 7.01M). Wall time is not better: two components runs hit the 1800 s budget (their partial work still passed grading) and the arm was slower on two of three tasks.

Decision rule (AGENTS.md): a component survives while the components arm beats core on hidden native pass rate or cuts tokens or wall time meaningfully. Both the pass-rate and the token conditions hold, so the library stays. The evidence is thin: N=2 per cell, and only two of the six components runs called a component at all (`document.stories` once, `operation.safe-edit` once), so the token saving is more plausibly from the arm's manifest and documentation shaping than from component calls. Before pruning or expanding the library, run N=5 with the corrected tasks and compare per-component usage against outcome.
