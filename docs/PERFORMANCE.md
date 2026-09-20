# Native feedback measurements

A fresh artifact-bound run on 2026-09-15 against Microsoft Word 16.0 passed 22
assertions in 5,169.8 ms (Word host handshake: 2,116.1 ms) on the private
desktop; `word_executed=true`, `vba_compiled=true`, and no user Word process was
attached. Warm native-call medians from the same artifact were 36.5 ms for
selection normalization, 9.0 ms for the form/style/footnote/table acceptance,
5.8 ms for building-block insertion, 11.3 ms for hotkey registration and
execution, and 0.53 ms for an application property read. These are measured
primitive timings, not a claim about a complete journal Setup workflow. The
raw acceptance report is `reports/word-launch-guard/reports/acceptance.json`.

Measured in real Microsoft Word on Windows x64, September 12, 2026. Processes run at Below Normal priority, with the owned job capped at 50% CPU and 2048 MB. These are measurements on the development computer, not a one-second guarantee or a test on low-end hardware.

The signed Studio acceptance suite freshly executes 17 assertions: actual whole-project compilation, native signature presence, persisted nested UserForm, Ribbon initialization and button action, scratch VBA, native page rendering, form callback, template hotkey, context-menu callback, document contents, inherited content control, saved building block, and selection/format preservation. It also captures the real Word and UserForm windows. It does not substitute a cached pass for Word execution.

## Real-manuscript macro timings

These are separate from the primitive and generated-template measurements above.
Each row ran the actual VBA against a supplied manuscript copy in hidden
Microsoft Word 16.0, with the source and result checked through Word's exact
`WordOpenXML`. The macro time is the operation's own timer; the native-step time
also includes the surrounding assertions and Word calls. It excludes Word host
startup unless noted.

| Operation and input | Document facts | Macro time | Native operation step | Result |
|---|---|---:|---:|---|
| `UBC.FormatManuscript` on Dyson | 61,775 characters; 156 footnotes; 82,335-byte DOCX | 3.297 s | 3.773 s | Passed; body, notes, fields/bookmarks and exact source content preserved |
| `UBC.FormatManuscript` on Rizzuto | 43,151 characters; 146 footnotes; 84,485-byte DOCX | 3.969 s | 4.456 s | Passed; table, two sections, notes and exact source content preserved |
| ALR full Setup on Dyson | 190 paragraphs; 156 footnotes; 684,955-byte Word XML | 4.49 s (median of 3, quiet machine) | 10.4 s native step | Passed; 20 selected tasks; full XML equal to the original under the explicit policy |
| ALR full Setup on Rizzuto | 169 paragraphs; 146 footnotes; one table; 692,344-byte Word XML | 7.66 s (median of 3, quiet machine) | 16.3 s native step | Passed; 20 selected tasks; text identical to the original, XML equal except one duplicate review comment the original adds |

The complete fresh Word runs were 8.274 s and 8.370 s for the two UBC rows,
including startup, compilation, document open, XML snapshots and cleanup. The
ALR Dyson quiet run was 18.602 s end to end; the ALR Rizzuto state/undo run was
36.609 s because it additionally measured a 7.029-second full undo/state
restoration check. Those wall times are harness workflows, not the macro
throughput numbers.

The ALR rows are macro stage time from `ALR_BatchReport` on the 2026-09-16 build, measured three times each on a quiet machine (scratch script `parity_and_timing.py`; Dyson 4.43-4.96 s, Rizzuto 7.65-10.14 s with one outlier). The original July 22 macro takes 53.96 s on Dyson and 22.18 s on Rizzuto on the same machine, so the refactor is about 12x and 2.9x faster. The gains came from a one-shot paragraph index shared by the style, indent, review and emphasis stages, run-level instead of per-character formatting writes, gating the quote passes on a one-time text read per story, visiting only bold paragraphs when removing bold inside brackets, and writing page setup values only when they differ. Every stage was re-checked against the original's output with the stage-by-stage XML snapshot suite (`make_bisect.py`): the paragraph-offset regression that deleted characters from the middle of words (empty-paragraph removal by string offsets) was found and removed that way, and the remaining run-boundary differences were traced to writes the original never made. The slowest current stages are the curly-quote passes (about 0.3 s Dyson, 0.45 s Rizzuto), style normalization (0.75 s / 1.3 s, kept per paragraph for parity), TOC refresh (0.8 s / 1.2 s, Word's own) and page setup on the two-section Rizzuto paper (1.1 s).

| Run | Wall time | Owned-job CPU time |
|---|---:|---:|
| Earlier warm suite before UI Automation optimization | 4.70 s | 5.33 s |
| Native UI Automation lookup, warm | 2.66 s | 3.33 s |
| Completion notifications and disposable-document reuse, warm, experimental redraw suppression | 2.70 s | 2.50 s |
| Same-session comparison, redraw suppression | 4.03 s | 4.19 s |
| Same-session comparison, ordinary redraw | 3.51 s | 3.08 s |
| Final suite with explicit UserForm readiness before capture, warm | 2.62 s | 2.81 s |

Redraw suppression was removed: the comparison did not establish a benefit. Word scheduling and machine load produce substantial variability. The one-second full-suite target has **not** been achieved.

Visual review found that the UI Automation action may return before its UserForm appears. Earlier fast runs could capture Word before the form was visible. The final suite explicitly waits for the form's Close control before capturing windows; its actual form, Ribbon and page images were visually reviewed. This corrected run took 4.58 seconds including Word startup and 2.62 seconds warm, with approximately 256 MiB peak owned-job memory. The earlier rows are optimization experiments, not equivalent visual-completeness evidence.

The 2.70-second observation spent 793 ms opening the signed template, 308 ms compiling, 278 ms creating a document from the template, 246 ms running isolated scratch VBA, and 222 ms capturing windows. Selection/format verification took 42 ms after reusing the disposable document instead of creating another one (previously 317 ms). These individual figures are observations, not promised latency ceilings.

Completion notifications replace repeated task polling on Windows. The independent wall-clock watchdog remains active if the caller stops waiting. An opt-in native regression test checks successful return, runtime error number and line, a caller deadline that leaves the diagnostic lane alive, and termination of an infinite macro without polling. No changes to user Word processes or global process priorities are needed.

For a warm workflow, start one authorized local session and reuse it:

```powershell
.\wordup.exe -w WORKSPACE --execute session start
.\wordup.exe -w WORKSPACE rpc test '@test-request.json'
.\wordup.exe -w WORKSPACE rpc test '@test-request.json'
.\wordup.exe -w WORKSPACE session stop
```

The request contains `path` (the template's absolute path) and `suite` (the suite JSON object). Each response contains per-operation timings, assertion results, artifact and suite hashes, owned-job CPU time, peak job memory, and visual evidence paths. Use a fresh process when testing startup behavior or process isolation. Repeated `selftest` calls through a session now also reuse its Word host unless `fresh` is requested.

Local raw evidence: `build/lightning-ready-cold.json` and `build/lightning-ready-warm.json` contain the final results. Earlier experiments: `build/uia-final-warm.json`, `build/lightning-warm.json`, `build/lightning-ab-redraw.json`, and `build/lightning-ab-plain.json`. Build outputs are excluded from Git; timings above are a summary of those measurements.

## Offline writer loop

`tools/writer-compare/go` separates direct package mutation, complete workspace
rebuilds, and resident-cache checks. The latest local run on the complex Atelier
template measured these Go-only lanes:

| Operation | Forced rebuild warm median | Same-output cache warm median |
|---|---:|---:|
| unchanged workspace | 19.11 ms | 1.03 ms |
| module edit | 23.24 ms | 1.05 ms |
| class add | 24.46 ms | 1.30 ms |
| control caption | 21.67 ms | 2.10 ms |
| control font | 21.73 ms | 1.55 ms |

The direct Go package lane measured 4.52 ms for a module edit and 5.12 ms for a
class add. These are writer measurements, not Word compilation or rendering.
The resident cache retains parsed state only inside the existing idle-expiring
agent session; external source, evidence, or artifact changes invalidate it and
force a real build. On Windows it uses the filesystem change stamp for an
unchanged artifact and falls back to the recorded SHA-256 when that stamp is
unavailable; full source, package-part, and CFB verification still happens on
the rebuild path.

A 2026-09-14 run on the supplied complex macro fixture measured 27.72 ms for
unchanged forced rebuilds, 54.04 ms for a module edit, 58.90 ms for a class add,
and 58.93 ms for a form-caption edit; same-output cache warm medians were
2.07, 2.12, 2.92, and 1.99 ms respectively. The direct Go package lane was
39.35 ms for a module edit and 38.43 ms for a class add. The form-font lane
varied from 1.6 to 628.0 ms under the same run, so it remains a machine-load
observation rather than a gate. No language or cutover conclusion follows from
one run; the measurements identify which path deserves further profiling.

The older pyOpenVBA comparison is historical evidence, not a current gate: this
checkout does not contain a separate repository at the pinned revision, so no
current cross-language result is claimed. When that oracle is available, run
`tools/writer-compare/compare.py` against the exact pinned checkout before using
language or cutover conclusions.

The same warm session measured a second `check` request at 23.3 ms after a first
request that included session startup (520.8 ms). A direct process check took
704.6 ms, which is startup overhead rather than the warm agent loop.

## Shared XML core and agent manifest (2026-09-15)

Go-only measurements on Linux amd64, Go 1.23.2, Intel Xeon Platinum 8370C,
GOMAXPROCS=5. Baseline is production source at
`6a9d431d43275b4b1118e9fc3d4a7c76759f6743`; the same benchmark functions and
portable wrapper were copied into its checkout without changing production
code. Five 200 ms benchmark samples per lane; values below are medians.
The 100-style baseline needs about 500 ms for one operation, so each sample
contains one iteration. These are local observations, not latency guarantees
or Word/VBA compile, UI, render, signing, or end-to-end workflow measurements.

| Workload | Baseline | Shared core | Speedup |
|---|---:|---:|---:|
| Create 10 styles, 5 run + 5 paragraph properties per style | 9.010 ms | 0.903 ms | 9.98x |
| Create 100 styles, same properties | 501.724 ms | 10.484 ms | 47.85x |
| Build and JSON-encode the complete tool manifest | 4.598 ms | 0.418 ms | 11.01x |

For the 100-style workload, allocated bytes per operation fell from
139,632,016 to 3,619,768 (97.4% less), and allocations from 2,042,826 to 31,592.
The encoded tool manifest fell from 309,238 bytes for 50 tools to 26,692 bytes
for 51 tools (91.4% less), including the new `styles.apply` operation. Neither
measurement depends on response caching or skipping the requested edits.

`xml_edit.go` provides the shared original-offset splice plan. `properties.go`
defines typed property mappings once; styles, numbering and composition reuse
them. Styles and numbering are indexed once per batch instead of reparsing
and copying the growing whole part per style and per property. Only changed
fragments are rebuilt, with retained namespace scopes and bounded edit plans.

Reproduce with the repository wrapper (use `tools/go.ps1` on Windows):

```sh
WORDUP_GO=/path/to/go ./tools/go.sh test -run '^$' -bench 'Benchmark(StyleBatch|ToolManifest)$' -benchtime=200ms -count=5 -benchmem ./internal/office ./internal/agent
```

Benchmarks run only with `-bench`; normal tests do not execute benchmark loops.
CI retains bounded observations rather than enforcing noisy time thresholds.
Its offline suite and race checks are separate from native Word acceptance.
Static checks retain all analyzers except ANTLR's known generated unreachable
gotos, while the unreachable analyzer remains enabled on handwritten code.
Raw local samples are retained in `docs/evidence/shared-core-before.txt` and
`docs/evidence/shared-core-after.txt`.

## ALR Suggester on a real manuscript (2026-09-17)

The 8.1 s figure below was measured on the Dyson demo, which is small and holds no
fields. On the real submission corpus the same build took two to eight minutes, and
one paper did not finish at all. Measured on `0_63-2 The Pore Space Race`
(278 paragraphs, 261 notes), macro stage time from `ALR_BatchReport`:

| Stage (Suggester, tracked) | Before | After |
|---|---:|---:|
| Spelling List | 21.6 s | 4.4 s |
| Wording | 48.0 s | 5.4 s |
| Honorifics | 4.1 s | 0.0 s |
| Latin italics | 21.9 s | 0.2 s |
| Latin roman | 15.3 s | 0.2 s |
| Audit | 6.0 s | 0.3 s |
| Numbers | did not finish in 600 s | 0.1 s |
| Dates | never reached | 1.9 s |
| Bracketed paragraphs | never reached | 3.7 s |
| All 13 rule stages | never finished | 17.7 s |

Three causes, each measured before it was fixed:

1. A story holding fields fell back to one Word `Find` per quotation mark, 5.4 s for
   the main story and 11.6 s for the notes, on every call. The field map answers the
   same question with string work.
2. The quote-span cache held one story while a run walks about 17 stories per rule
   family, so it never hit. It is now a 24-slot ring keyed on story and length.
3. `Numbers` was the last rule driving Word's wildcard `Find` (percent, thousands,
   year spans, ratios). All four are string scans now, which is the 600 s to 0.1 s line.

Knock-on effects: `suggester-probe` 64 s to 8.4 s, `setup-rules-probe` 15 s to 8.6 s,
the Dyson proposal report 240 s to 14 s, Rizzuto 950 s to 19 s.

Red ink stays small: across the judge, section, journal and Ibid proposals in
`judges-journals-probe`, 68 revisions cost 276 characters inserted and 195 deleted,
about four in and three out per proposal. The Ibid converter alone went from a
200-character delete-and-retype to inserting `Ibid` and deleting 58 characters.

## ALR Suggester and Setup after the two-GUI reorganization (2026-09-16, night)

Measured on the Dyson submission (190 paragraphs, 156 footnotes) while a six-worker benchmark ran on the same machine; macro stage time from `ALR_BatchReport`. Superseded by the corpus measurements above: this document is small and holds no fields, so it never exercised the slow paths.

| Stage (Suggester, tracked) | Before fixes | After |
|---|---:|---:|
| Spelling List (135 rules) | 8.5 s | 0.8 s |
| Wording | 5.3 s | 0.7 s |
| Section words (legacy, 24 wildcard Finds, regex-gated) | 3.6 s | 3.2 s |
| Numbers | never finished (infinite empty-match loop inside the TOC field) | 1.0 s |
| Dates | — | 0.8 s |
| Bracketed paragraphs | 10.6 s (legacy, two Finds per footnote) | 0.1 s |
| Whole Suggester (12 concepts; 13 since Latin italics moved here on 2026-09-17) | — | 8.1 s |

What changed: rules reach Word's Find only when an in-memory whole-word scan of the story finds their text; quote spans come from one string scan (or one Find pass when the story holds hidden content) and are shifted arithmetically after each write instead of rebuilt; every Find loop goes through `FindNext`, which skips empty or non-advancing matches and any match touching a field. Full Setup with all 24 concepts through the form on Dyson (23 since 2026-09-17, Latin italics now a Suggester proposal): about 6 s (original stage set 4.5 s median, unchanged). Red ink on Dyson: 85 tracked proposals, 78 inserted and 39 deleted characters.
