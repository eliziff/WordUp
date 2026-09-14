# Agent API

All commands are local. Use `wordup -w WORKSPACE call METHOD @parameters.json` for one-shot operations and `rpc METHOD @parameters.json` for a previously started warm session. `serve` accepts newline-delimited `{id,method,params}` JSON. `mcp` is stdio JSON-RPC with initialization and tools/call. Stdout contains protocol data; errors are structured and command failure is nonzero.

## Exact-source XPath evidence

`call xml.query @query.json` queries a UTF-8 XML file, or a selected DOCX/DOTM
part, without starting Word. Prefixes bind to namespace URIs independently
of the source document's prefixes:

```json
{"path":"manuscript.docx","part":"word/styles.xml","query":"//w:style[@w:styleId='Heading1']/w:rPr","namespaces":{"w":"http://schemas.openxmlformats.org/wordprocessingml/2006/main"},"limit":20}
```

Responses retain package/part hashes and original UTF-8 element byte offsets
(`element_start`, `element_end`, exclusive end). Attribute/text matches locate
their containing element. Previews above 1024 bytes become prefix/length/hash
objects, not complete XML. `truncated` means more matches exist, not an exact
total. Use `count(...)`, `boolean(...)` or `string(...)` for scalar evidence.
No matches is an empty query, not a formatting pass. Properties are explicit
XML, not resolved style, theme, numbering or Word layout properties.

Native suites can query an earlier saved snapshot:

```json
{"name":"No explicit 24-point runs","operation":{"op":"xml.query","file":"$output/after.xml","member":"count(//w:rPr/w:sz[@w:val='48'])","named":{"namespaces":{"w":"http://schemas.openxmlformats.org/wordprocessingml/2006/main"}}},"assert":[{"kind":"equals","path":"/value","expected":0}]}
```

Query steps do not count as Word execution. `named` also accepts `part` and
`limit`. Large assertion error values carry JSON byte counts and hashes;
full actual values remain in observations and expected values in the suite.
Full reports are still detailed; general bounded tool summaries are not yet
implemented.

`read` can retrieve large report or evidence files in bounded chunks with
`offset` and `limit`. The response always includes the full-file SHA-256,
total byte count, returned range, and `next_offset` when more bytes remain;
offsets are raw bytes (a chunk that is not valid UTF-8 is returned as base64),
and the default with neither field remains a complete read.

`xml.patch` applies a small direct byte-range edit to an XML source already in
the workspace (for example `package/word/styles.xml`). Use the exact offsets
and `source_sha256` returned by `xml.query` as `offset`, `length`, and
`expected_sha256`; provide replacement `text` or `base64`. The result is
checked as UTF-8 and strict XML before an atomic guarded write. It preserves
all bytes outside the range and intentionally does not rewrite a DOCX/DOTM
ZIP; import the artifact first and edit its ordinary package source.

## Direct native object access (Windows)

Create `operation.json`:

```json
{"operation":{"op":"new","as":"doc"}}
```

Then use `rpc native.call @operation.json`. Subsequent operations retain object handles:

```json
{"operation":{"op":"get","target":"doc","member":"Content","as":"body"}}
```

```json
{"operation":{"op":"put","target":"body","member":"Text","value":"An actual Word document.\r"}}
```

```json
{"operation":{"op":"get","target":"body","member":"Text"}}
```

Methods accept positional `args` and named `named` parameters. Pass a retained object using `{"object":"body"}` and an omitted optional COM parameter using `{"missing":true}`. Object-returning calls require an `as` handle. Default application handle is `app`. Prefer document/Range objects to changing the user's application-wide defaults. `batch` contains an ordered `steps` array and avoids a process launch for every member access. Before opening a changed revision of a staged filename, close the prior document with `unload`, passing both its `target` handle and its original `file` path; this explicitly releases the staged-version guard. Keep filenames unique for simultaneously loaded revisions.

## Scratch native VBA

```json
{"operation":{"op":"eval","value":"Dim d As Document\nSet d = Documents.Add\nd.Content.Text = \"Native VBA wrote this.\"\nEvaluate = d.Paragraphs.Count\nd.Close SaveChanges:=wdDoNotSaveChanges"}}
```

This generates a separate temporary DOTM containing a real `Public Function Evaluate() As Variant`, runs it through Word, and returns the result plus generated source/hash. It does not modify the target template. Use a normal source component for persistent behavior. Native errors remain errors; the harness does not evaluate the body itself.

For a currently paused VBA frame, use `vba.immediate` with `paused:true` and
`text:"? expression"`; optional `symbols` requests best-effort parameter/local/
module-variable values. Each unavailable symbol remains an explicit result row.
This path uses the owned VBE Immediate Window and returns the buffer text Word
exposed; it does not substitute the scratch project or claim a typed value.

## Modal forms and UI

```json
{"operation":{"op":"begin","steps":[{"op":"run","macro":"ProjectName.ModuleName.ShowModalForm"}]}}
```

Use `poll` with `value` equal to the returned task token. While the macro is blocked, `ui.windows` and `ui.tree` can inspect owned windows, `ui.invoke` can invoke an accessible control, and `ui.set_value` can fill a field. Use actual HWND and accessibility child identifiers from that inspection. No guessed controls or global keyboard input. `ui.capture` captures an owned window. When finished, use `forget` to release a completed task record. A process has a bounded task budget.

## Native render and compile

```json
{"operation":{"op":"render","target":"doc","file":"C:\\Work\\MyTemplate\\reports\\pages","value":144}}
```

`child` optionally selects a 1-based page. Native PNG rendering is Windows-only. `format:"pdf"` requests Word's own PDF exporter. For actual compile observation, pass `op:"compile"`, a document `target`, and `member` with the project name for menu-target verification if VBProject access is unavailable. Use `begin` for a call that can block on compiler diagnostics and inspect the owned VBE accessibility tree. A disabled command is not a witnessed compiler pass.

## Acceptance suites

```json
{
  "schema":1,
  "name":"Citation behavior",
  "platforms":["windows"],
  "steps":[
    {"name":"Load exact candidate","operation":{"op":"open","file":"$artifact","as":"candidate"},"assert":[{"kind":"equals","path":"/open_and_repair","expected":false}]},
    {"name":"Execute native acceptance","operation":{"op":"run","macro":"$project.Tests.CheckCitation"},"assert":[{"kind":"equals","expected":"CITATION_OK"}]}
  ]
}
```

`$artifact`, `$project`, `$filename` are resolved from the exact staged candidate. Tests must include assertions. `path` uses an RFC 6901 JSON pointer into the actual returned value; empty path means the entire result. Kinds: `equals`, `not_equals`, `contains`, `matches`, `near` with tolerance, and `greater_than`. A suite can require a separately observed compile via `require_compile:true`. A raw native call can execute arbitrary code; an acceptance suite should state observable outcomes rather than trusting the generated code's claims alone.

The CLI `test` command uses a fresh process. Via agent `test`, set `fresh:true` for release acceptance; warm runs are useful while developing but cannot authorize deployment.

## Deployment

`deployment.json`:

```json
{"path":"dist/MyTemplate.dotm","reference":"reports/acceptance.json","output":"C:\\Templates\\MyTemplate.dotm"}
```

Use authorized `call deploy '@deployment.json'`. The app checks the exact artifact/suite/observations, stages the bytes, backs up the old file and either activates or waits locally for Word to exit. Windows registers a current-user login command to resume pending activation, and removes it on completion. `call activate` with `path` equal to the returned plan retries it explicitly. `call restore` with that same `path` restores its previous template, refusing changed backups or newer target edits. Both require `--execute` and Word to be closed; neither closes user Word processes.

## Compile and sign

Authorized `call compile '{}'` builds the current workspace, compiles the complete project in Word, and signs the output. Optional `output` overrides its default path. It automatically creates or reuses a WordUp local certificate. `sign` instead takes an existing unsigned `path` and a new `output`. Advanced `signing` fields are `thumbprint`, `store_location`, `signtool` and `timestamp_url`; all are optional. Microsoft SignTool and the registered Office SIP remain prerequisites.

The signing report distinguishes `signed`, `digest_verified` and certificate trust (`verified`, `trust_error`). A locally self-signed certificate does not automatically become trusted on someone else's computer. Compilation records its native observations in `reports/compile.json`; when candidate creation or signing fails, `retained_stage` identifies the preserved staging directory. On failure, staged input and window diagnostics remain available at the report's paths; the prior output is preserved. Unchanged compile results are explicitly labeled `cached`; fresh acceptance is still required for deployment.

## Images and style references

`check` performs the offline source/XML checks, validates the assembled OPC
package graph (content types and relationship targets), and also returns a compact
`package_validation` object (`checked`, `valid`, and an error when invalid) plus
`inventory`: public procedures, UserForm event handlers and their declared
controls, Ribbon callback bindings (including conservative declaration-shape
warnings), VBA hotkey/context-menu registrations, and
the recorded component files with their current hash state. Persistent form
design JSON is parsed and its filename/name wiring is checked; an unpaired
design is a warning. Missing or modified wiring is reported in the same
diagnostics list. This is an
inventory and static wiring check, not proof that Word compiled or dispatched
an event; use native `compile` and acceptance steps for that.

`reference.document` accepts a DOCX/DOTM path and returns the exact `word/theme/theme1.xml`, `word/styles.xml`, `word/numbering.xml`, and (when present) `word/fontTable.xml` parts with hashes, plus a compact `part_inventory` of every OPC part (path, byte count, and SHA-256) so opaque custom XML, extension, and asset parts are discoverable without a normalized model. `content_types` preserves the package's declared defaults/overrides, and `relationship_inventory` preserves each `.rels` source, declared target/type/mode, and safely resolved internal target part; this exposes the package graph without guessing Word semantics. It also returns theme font mappings, paragraph/direct-formatting observations, and source-located `section_observations` for exact page geometry, columns, section breaks and header/footer references, plus document/settings/relationship hashes. The main story is returned as `paragraph_observations`; headers, footers, notes, comments, and glossary entries are returned separately as part-qualified `story_observations` when present. Paragraphs with tracked edits carry compact `revision_evidence` (kind, source bounds, author/date/id when present, and UTF-16 units); content-control evidence records source bounds, alias/tag/id/lock/type and data bindings; field, hyperlink, and bookmark evidence is source-located with its identifiers, instructions, and relationship targets, and artwork evidence records drawing bounds, dimensions, alternative text, and image relationships. It is not a complete resolved cascading-style engine; use the returned parts or `xml.query` for exact properties. `reference.image` supplies actual image pixels to MCP/vision clients. An ordinary CLI agent can also read that same local image file. `image.compare` takes `reference`, `path`, optional `tolerance` and optional `output` for a difference PNG. Different dimensions fail: there is no hidden alignment or rescaling that could create a misleading pass.

## Mac native prototype

Use `dictionary` to inspect the installed Word SDEF and the implemented operation surface. Native `get`, `put`, `invoke`, `run` and `ae.send` use actual Apple Events. Event/property codes are discovered from the installed dictionary, not guessed from Windows names. Mac UI/render/full-compile operations are unsupported and must not be silently rerouted to Windows or emulated. `compat` is advisory source review; only actual Mac Word evidence can verify Mac behavior.

## Reusable structure evidence

`call structure.source '{}'` returns the editable standalone `structure.detect` component and its contract. `call structure.inspect '@parameters.json'` accepts a document `path` and reads package evidence without Word: source-hashed paragraph locations, stable source IDs, style ancestry, direct versus inherited outline levels, table/textbox containment and raw numbering definitions. It also returns the exact `styles_xml`, `numbering_xml`, `theme_xml` and `font_table_xml` inputs (when present) with their hashes, so properties outside the generic detector can be copied or compared directly without a guessed formatting recipe. Duplicate style IDs are returned as compact source spans; conflicting definitions are marked as paragraph-level ambiguity and the deterministic source-order fallback is explicit. Unique Word paragraph IDs remain compact source IDs; repeated IDs fall back to their XML paths and carry `paragraph_id_ambiguous=true`, so malformed input cannot silently overwrite a row. Explicit `numId=0` is retained as disabled-numbering evidence instead of being discarded. `structure.resolve` adds generic candidate scoring, hierarchy, contradictions, ambiguity and stable parent IDs while keeping the factual evidence intact. `structure.compare` accepts a silver XML `reference`, resolves every named source document, and reports bounded role, level, parent, and any explicitly requested `ambiguous`/`ambiguity` or `contradiction` mismatches. Boolean ambiguity labels compare the resolved flag; generic labels such as `direct-format` and `roman-marker` compare against the corresponding direct-formatting and marker evidence. Contradiction labels remain model-authored metadata; comparison checks whether detector evidence is present rather than inventing a second taxonomy. Mismatch examples retain source XML spans and the relevant style, outline, numbering, formatting and contradiction evidence. Publication-specific role and style mapping remains separate. See [the detection boundary and reviewed upstream mechanisms](STRUCTURE-DETECTION.md).

The standalone `WU_DetectStructure(document)` returns the versioned 1.2.1 `result(paragraphIndex, WU_*)` table. Its public column constants keep candidate score and confidence separate, then cover role, level, 1-based parent row, ambiguity, evidence, source properties, Word UTF-16 range, marker/list evidence, context, stable range identity, stable parent identity, provenance, contradictions and alternatives.

`component.list`, `component.get`, `component.add`, `component.status`, and `component.diff` expose editable components. Pass `component` for a built-in, or `path` for a local directory containing `component.json` and its listed files. Bundle files are copied as bytes, so images, form assets, and other binary sources retain their exact contents and hashes; text files remain editable source. Binary files may live in `assets/` or directly under `package/`; VBA, form, XML, JSON, and other editable source paths must remain text. `component.add` and `component.diff` accept declared string `parameters`; `module_prefix` is currently the typed adaptation seam and must be a valid VBA identifier. Undeclared parameters and manifest parameters without a typed adapter are rejected. Installation writes ordinary source and records its initial SHA-256 hashes, provenance, supported platforms, applied parameters, and any declared Ribbon fragment source-to-part mappings in `.wordwright/components.json`; Ribbon mappings are collision-checked before any source is written and composed by the normal package build. Repeating the same unchanged installation is idempotent, while different parameters, adapted source, or a colliding file are never overwritten. `component.status` reports `compatibility` as `compatible`, `unsupported`, or `unknown` in addition to file state, and reports `version_state` (`current`, `outdated`, or `unknown`) when a bundled starting version is available. `component.diff` returns only per-file hashes, sizes, state, and first differing byte; call `component.get` when source text is needed.

`xml.verify` compares expected XML `reference` directly with actual XML `path`.
Set `part` (for example `word/document.xml`) when either path is a DOCX/DOTM
or Flat OPC export; the named package part is selected in memory, so no
extraction fixture is needed. Copy starting Word XML and edit only intended
differences to create the expectation. Comparison defaults to exact bytes;
`comparison: "semantic"`
explicitly selects namespace-aware equality without ignored content. Mismatch
returns an error plus hashes and difference locations. This command reads XML;
it does not execute Word.
