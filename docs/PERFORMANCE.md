# Native feedback measurements

Measured in real Microsoft Word on Windows x64, September 12, 2026. Processes run at Below Normal priority, with the owned job capped at 50% CPU and 2048 MB. These are measurements on the development computer, not a one-second guarantee or a test on low-end hardware.

The signed Studio acceptance suite freshly executes 17 assertions: actual whole-project compilation, native signature presence, persisted nested UserForm, Ribbon initialization and button action, scratch VBA, native page rendering, form callback, template hotkey, context-menu callback, document contents, inherited content control, saved building block, and selection/format preservation. It also captures the real Word and UserForm windows. It does not substitute a cached pass for Word execution.

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
