# Running a project corpus

Corpus execution is one tool, not a required verification ladder. Individual
native operations, profiling, suites, corpus runs, replay, XML comparison and
optional serial tracing can be composed in whatever order the project needs.
WordUp supplies execution and evidence; the project supplies acceptance policy.

`test.corpus` runs an existing native suite sequentially with different declared
inputs. It does not supply editorial rules, a corpus, or an alternative VBA engine.
Use `call test.corpus @request.json` (or the same RPC/MCP operation) with execution
authority and these parameters:

```json
{
  "path": "dist/Example.dotm",
  "output": "reports/corpus",
  "suite": {
    "schema": 1,
    "name": "Project assertions",
    "inputs": {"document": "C:/fixtures/first.docx"},
    "steps": [
      {"name": "Open", "operation": {"op": "open", "file": "$input:document$", "as": "doc"}},
      {"name": "Assert", "operation": {"op": "get", "target": "doc", "member": "ReadOnly"}, "assert": [{"kind": "equals", "expected": false}]},
      {"name": "Close", "operation": {"op": "unload", "target": "doc"}}
    ]
  },
  "cases": [
    {"id": "First", "inputs": {"document": "fixtures/first.docx"}},
    {"id": "Second", "inputs": {"document": "fixtures/second.docx"}}
  ]
}
```

Replace the illustrative assertion with the project's actual workflow and checks.
Each case replaces exactly the suite's declared input names. Case IDs are simple
letters/digits/underscores, start with a letter, and are case-insensitively unique.
Case paths are workspace-relative or absolute; suite input paths are absolute.
Suites must close documents and release handles they open. Cases share a warm
owned Word host; an execution failure closes it before the next case.

Every invocation allocates a new `run-*` directory containing `corpus.json` and
`index.md`. Reports retain exact artifact/input hashes, assertions and timings.
Passing cases also freeze all retained evidence. Originals are not edited: the
existing suite runner supplies disposable working copies through `$input:name$`.
Direct paths hidden in VBA or operations are not isolated or made hermetic.

To resume, add `reference` pointing to a previous `corpus.json`. Only matching
artifact, runner executable hash, suite and input bytes can reuse a case, and every frozen
file is hash-checked. Missing/damaged evidence reruns that case. Failed cases rerun;
earlier attempts remain intact. Any failed or unfinished case prevents aggregate
success. A running/interrupted summary is not a pass. Resume does not detect
changes to Word installations, settings, external services or undeclared inputs;
omit `reference` when those change or fresh execution is needed.

Planner replay uses ordinary suite operations and assertions. Optional serial
traces and existing `test.compare` can provide additional diagnostics. Their
coverage is explicit; neither asserts universal fidelity or DOCX byte parity.

## Interpreting evidence

An expected protected-span skip is a normal project outcome, not an execution
failure. Assert both preservation of the protected content and completion of
eligible edits. Do not turn arbitrary write errors into successful skips.

Keep assertion mismatches separate from operation errors, task deadlines, worker
exits and verified Word exits. A lost connection alone is not proof of a Word
crash. Native acceptance reports retain the suite operation, the failing
observation, timings and available host/window diagnostics. Asynchronous deadline
causes are propagated to waiting calls rather than replaced by `session_closed`.
Explicit reruns retain separate attempts; they do not erase earlier failures.

For ordinary output review, suites can capture `xml.snapshot` before and after
the editing action (outside its undo record) and inspect those captures without
Word. Keep tracked deletions/insertions, field structure and formatting visible;
accepted text alone cannot establish preservation or serial correctness. Detailed
serial instrumentation is optional, useful when intermediate behavior matters.
Execution success and human/editorial output approval remain distinct claims.

`call xml.review @request.json` provides a bounded before/after view without
starting Word:

```json
{"reference":"reports/before.xml","path":"reports/after.xml","offset":0,"limit":20}
```

Use `part` for an explicitly selected DOCX/OPC part. Flat OPC snapshots can be
reviewed directly. The result includes changed paragraph rows with displayed
text, revision/field/control markup, direct formatting, source hashes and byte
offsets. Increase `offset` while `more` is true; previews and evidence arrays
explicitly indicate truncation. Changed paragraphs are detected from their raw
XML, so serialization-only differences can also appear. Paragraphs are aligned
by ordinal within each story, not by inferred editorial identity; paragraph
insertions can shift later rows. Use `xml.compare` for namespace-aware equality
and raw byte identity. Review does not resolve inherited formatting or cover
every package feature, and never marks output editorially approved.

Offline comparison proves only the compared observations. Pure planner replay
checks known inputs; it does not reproduce Word's state transitions. Native
`test.replay` actually reruns Word. Neither should be described as a Word-free
proof of native serial behavior or DOCX byte identity.
