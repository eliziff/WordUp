# Wordwright Local 0.2.0 — actual validation record

**Native Windows/Word and Mac/Word execution: NOT RUN.** These are cross-compiled executables with a tested offline core, not a complete zero-intervention Office compatibility certification. See `docs/LIMITS.md` for unmet requirements.

## Executed

- **27 named top-level tests passed**, zero failed and zero skipped tests with the private ALR specimen supplied. Four packages currently have no automated tests. `tests.jsonl` is the raw Go test event record; `test-summary.json` names the actual tests.
- `go test -race ./...` passed on Linux. This does not execute platform-specific native Windows/Mac code.
- `go vet ./...` passed on Linux. Windows-target `go vet -unsafeptr=false` passed for the native/IPC/deployment packages; unsafe-pointer analysis was explicitly disabled for those raw-ABI checks. No claim of an unsafe-code proof.
- Compression fuzzing: **80,043 executions** in the recorded short run. CFB fuzzing: **2,242 executions**. Both passed. These bounded runs are not comprehensive fuzzing.
- Actual CLI import, builds, inspection, generated example, warm Unix-socket session and repeated RPC calls ran successfully.
- Linux `selftest` returned **not_run / nonzero**, rather than a fake pass, because its Word suite does not target Linux. An explicit `native.start` invocation also returned the unavailable-host error. See both raw records.
- Source build-cache tampering, stale source writes, path traversal, ZIP symlink entries, unknown properties, empty assertion suites, artifact-hash mismatch and concurrent installation changes were exercised.
- Deployment policy/backup tests use **explicit synthetic reports**, not native Word evidence. Windows and Mac activation behavior has not run.

## Private ALR specimen

The 708,117-byte supplied template was imported and rebuilt unchanged with **byte-for-byte identity**. The original was not modified.

The controlled VBA edit retained all 15 components; the other 14 component sources, all 26 UserForm storage streams, and `word/document.xml` remained unchanged in the preservation test. Separate binary-form tests read and edit all four native ALR forms, including the MultiPage form. This establishes the tested source/storage properties, not successful form instantiation in Word.

The private proof bundle includes a baseline-identical DOTM and an edited `NOT-WORD-VERIFIED` candidate. Do not install that candidate over a working template until relevant native acceptance passes. Its artifact hash is in the private `proof.json`.

## Integrated example

`wordwright example OUTPUT` builds a native 10,106-byte template with four VBA components, one persistent nested MultiPage form, RibbonX, named styles, a content control, real footnote, table, header/footer and reusable saved part. The source fixture includes actual VBA checks for form persistence, styles, selection-boundary/format preservation, Ribbon initialization and saved-block insertion. The native checks have **not run** here.

No screenshot or LibreOffice rendering was used as evidence for this native rewrite. Page rendering is implemented through Word/GDI, but no native rendered page is supplied in this release.

## Measured offline performance

Twenty full CLI process launches per scenario, warm OS disk cache, on INTEL(R) XEON(R) PLATINUM 8573C in a Linux container. Includes process launch and command work. Does not measure Windows performance, Word startup, COM latency, native macro execution, compilation or rendering.

| Scenario | Median | p95 |
|---|---:|---:|
| alr_cached_build | 58.082 ms | 63.397 ms |
| alr_changed_source_build | 144.885 ms | 158.831 ms |
| alr_inspection | 87.087 ms | 104.339 ms |
| studio_cached_build | 2.973 ms | 3.608 ms |

The separate Go algorithm benchmark records roughly 143–153 MB/s VBA compression and about 0.80–0.83 ms for its new-project serialization fixture. Algorithm microbenchmarks are not end-to-end template/runtime timings. Raw samples and benchmark outputs are retained.

## Build facts

Windows x64 executable: **4,131,840 bytes**. No executable packer, dynamic runtime downloader, Python, Node, browser engine, VSTO or third-party Go modules. The Go runtime/standard library are linked into the binaries; OS APIs and installed Office are not bundled.

Production Go/assembly source including comments: **12,160 lines**. Including tests/comments: **13,032 lines**. Line count is a size observation, not evidence of correctness.

Exact binary hashes and architecture/build information: `evidence/release.json`. The executables are not Authenticode-signed or notarized. Mac full native UI/render/compile coverage and guaranteed activation across a pending OS reboot are not implemented. These remain open release gates, not accepted substitutes for the requested end product.
