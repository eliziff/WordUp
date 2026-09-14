# Expected XML

The user describes the result. If an expected artifact will help test it,
mechanically copy the starting XML captured from Word and edit only the intended
differences. That edited XML is model-authored silver.
For a smaller change, an ordinary assertion may be enough.

Run the feature on the same starting document and capture the resulting XML
from Word using the existing native XML capture operations. Compare like with
like: the same range, story, or complete Flat OPC export on both sides.
Preserve untouched formatting, notes, fields, bookmarks and other markup in the
copied expectation. For heading styling, edit the actual paragraph properties,
not a separate list of labels.

Use `xml.verify` with two file paths. No fixture manifest is required:

```json
{"reference":"tests/expected.xml","path":"reports/actual.xml"}
```

```powershell
wordup -w WORKSPACE call xml.verify '@comparison.json'
```

Exact bytes are the default. A mismatch returns an error with hashes and a
zero-based first-different-byte offset, plus namespace-aware difference context
when applicable. `comparison: "semantic"` explicitly allows namespace-prefix
and equivalent serialization differences; it does not ignore content or
formatting. `xml.compare` remains available for diagnostic comparisons with
explicit policies. Do not relax comparison merely to obtain a pass.

Try a plausible defect, such as lost italics, to check that verification rejects
it. Test reusable features on varied documents. XML comparison is deterministic;
the model-authored expectation can still be wrong. Native behavior and visual
fidelity require actual Word execution and rendering respectively.

Native suites can capture and verify directly. When the expected or actual
path is a DOCX/DOTM, `named.part` selects the same XML part in memory:

```json
[
  {"name":"Capture output","operation":{"op":"xml.snapshot","target":"doc","file":"$output/actual.xml"}},
  {"name":"Check expected XML","operation":{"op":"xml.verify","target":"C:/Work/tests/expected.xml","file":"$output/actual.docx","named":{"part":"word/document.xml"}},"assert":[{"path":"/matches_expected","kind":"equals","expected":true}]}
]
```

Here `target` is the expected file for `xml.verify`, and `named.comparison` can
explicitly select `semantic`. Neither comparison operation counts as Word
execution. Use `xml.snapshot` outside the editing action's undo record.

Word can regenerate editing-history identifiers or add document IDs during an
edit. Exact comparison reports these differences too. When they are irrelevant
to the particular test, use the existing suite `xml.compare` with an explicit
`named.xml_policy` and assert `/equal`. Keep the policy local and test defects
under that same policy. Do not discard IDs referenced by comments or other
document features. There is no automatic ignore list.

## Repeating a native test

Each test saves a distinct `saved_report` and retains the tested template bytes;
`reports/acceptance.json` is only the latest report. To preserve a manuscript or
expected XML file across iterations, declare an input in the suite:

```json
"inputs": {"manuscript":"C:/Work/input.docx","expected":"C:/Work/expected.xml"}
```

Use `$input:manuscript$` or `$input:expected$` in operation paths. Each run reads
the declared files once, records their hashes and creates separate baseline and
working copies. `test.replay` uses the captured inputs and resets working copies;
an explicit artifact `path` lets it test new code against the same inputs.
Without `path`, it uses the retained tested artifact. Changed snapshot hashes
fail instead of silently reading current source files. `test.compare` requires
matching captured input hashes. Paths accessed by macro code outside these
declared inputs are not frozen or isolated. Old reports cannot retroactively
prove which external files were used; capture a new baseline.

`test.freeze` takes a passing report `reference` and a new `output` directory.
It copies the retained template, declared inputs, report and evidence without
rewriting XML. Move that directory as a unit. Pass its `bundle.json` as the
reference to `test.replay` or `test.compare`; loading verifies the copied hashes
and resolves evidence paths while preserving the original suite hash. Existing
destinations are refused. Bundles contain local evidence and may contain private
documents or paths; freezing does not publish or sanitize them.
