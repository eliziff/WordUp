# Native serial evidence

Native suites can declare `"serial_trace": "trace.json"`. Their macro must
export the trace to `$output/trace.json` after its editing action and custom
undo record have finished. WordUp validates and hashes it before acceptance
can pass and writes a readable `trace.md` beside it. Missing, truncated,
malformed or unfinished-write evidence fails acceptance.

`test.compare` compares the exact ordered decoded events and reports the first
difference. `test.freeze` retains the raw trace and readable report; relocated
bundles verify their hashes. Existing XML policies do not relax trace equality.
Different JSON whitespace is not an event difference; raw file hashes remain
available. Neither event equality nor an XML comparison means DOCX byte identity.

## Instrumenting a project

The standalone module is `internal/verify/vba/WordUp_Trace.bas`, also embedded
as `verify.TraceRecorderSource` for the native regression fixture. Copy that
module into the project's `vba/` directory and build normally. WordUp does not
automatically intercept arbitrary VBA. Do not import through VBE.

1. Call `WordUp_TraceStart` with a JSON array naming the paths observed.
2. Use `WordUp_TraceEvent(kind, stage, range, dataMembers)` for `read`, `plan`
   and `skip` observations. `dataMembers` is an escaped JSON object body; use
   `WordUp_TraceQuote` for strings. Do not duplicate the recorder's reserved
   range fields (`start`, `end`, `story_type`, `story_length`, `text`, `touches`).
3. Immediately before a write, obtain its operation ID from
   `WordUp_TraceWrite(stage, range, property, requestedJSON)`.
4. Execute the original write without changing its error semantics. Immediately
   record its actual result with `WordUp_TraceResult(operation, range, error)`.
5. Close the undo record, then call `WordUp_TraceFinish(path)`. Observation
   failures are raised here, outside the editing action. Finish releases the
   buffer and retained objects on success and failure; `WordUp_TraceReset`
   explicitly abandons an interrupted observation.

The recorder captures live range text and coordinates, document and
story-instance IDs, text/formatting requests and outcomes, tracking settings,
and error numbers. IDs are assigned by encounter order, never by volatile COM
handles. Separate documents and story instances do not share position spaces.

After successful writes, Word maintains duplicate range anchors. Subsequent
read/plan/write events report cross-stage touches using those anchors' current
positions, including boundary touches and collapsed anchors. These are
diagnostic candidates, not a complete dependency graph or proof of causality.
Non-overlapping logical dependencies must be instrumented by the project.

The format is schema 1 with `complete`, `coverage`, and `events`. Every event
has a consecutive `sequence`, `kind`, `stage`, `document`, `story`, and `data`.
Write/result events share a unique `operation`. Event data includes exact range
coordinates/text; results require an integer `error`. Project-specific data
stays in `data` and is compared without exclusions. Coverage is a declaration,
not an automatically verified claim that all macro paths are observed.

## Evidence and limits

The non-legal native regression compiles the module, performs tracked text and
formatting edits serially, verifies a later read touches the earlier write,
compares repeated traces exactly, and checks one-step undo. Run with
`WORDUP_NATIVE_TEST=1` through `tools/go.ps1 test ./internal/verify -run
TestNativeSerialTrace -count=1 -v`.

Projects supply their own workflows, coverage declarations and assertions.
Run the production stages in their production order, retaining pending revisions
between stages. Tracing is optional diagnostic evidence, not a requirement to
instrument every read before improving a macro. It is not a Word-free engine.

Warm test input filenames now include a per-run suffix. This prevents a prior
SaveAs from blocking a later run that uses the same logical input name. Tests
of filename-dependent behaviour must explicitly account for disposable names.
Input bytes and frozen baseline names remain unchanged. Operation-token
expansion uses fresh storage so `$output` and `$input` never mutate the saved
suite or its hash.

The native host also clears the original staged-file identity when the original
open handle is unloaded after SaveAs. A native regression reopens changed bytes
at the same source filename, independently of the runner's unique filenames.
