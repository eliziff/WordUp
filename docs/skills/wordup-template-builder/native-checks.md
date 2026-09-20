# Use existing WordUp checks

Use `wordup -w WORKSPACE call METHOD '@parameters.json'` for offline operations; add `--execute` for native execution. A warm `session start` plus `rpc` avoids repeated host startup. Paths inside parameters resolve against the workspace unless absolute; the CLI's `@parameters.json` is read from the current working directory.

- `structure.inspect` with a document `path` reads style ancestry, direct/inherited outline evidence and containment without launching Word. Follow it with `structure.resolve` when the agent needs generic candidate roles, levels, parents and ambiguity before conversion.
- `reference.document` with `{"path":"dist/Template.dotm"}` inspects saved package styles, theme fonts, numbering, content types, the relationship graph, a hash inventory of every OPC part, and paragraph/story observations. Inspect both source and output; this is not a complete resolved style engine.
- `check` checks source syntax/XML/Ribbon diagnostics. Use a native `compile` step separately; parser success is not compiler success.
- `test` with `path` and a `suite` executes native assertions; choose `fresh:true` when an independent host is needed. Keep requirement-specific suites in the workspace and reuse `test.replay`; use `test.compare` for compatible baseline/candidate suites.
- `native.call` supports `get`, `put`, `invoke`, `eval`, `xml`, `render`, and UI operations. Use retained objects and batches rather than repeated process launches. See the API for asynchronous forms and captures.

## Saved style assertion pattern

In a suite, open `$artifact` as `template`. Resolve a required style with the native operations below, then assert actual properties. Replace the example style name and expected values with reference-derived requirements. Missing styles must fail rather than be created by the test.

```json
[
  {"name":"Resolve saved heading", "operation":{"op":"get","target":"template","member":"Styles","args":["House Heading 1"],"as":"heading"}},
  {"name":"Resolve heading paragraphs", "operation":{"op":"get","target":"heading","member":"ParagraphFormat","as":"headingParagraph"}},
  {"name":"Heading has semantic level one", "operation":{"op":"get","target":"headingParagraph","member":"OutlineLevel"},"assert":[{"kind":"equals","expected":1}]}
]
```

Apply the same pattern to `Font`, `BaseStyle`, `NextParagraphStyle` and required paragraph/numbering properties. This snippet is a pattern, not a complete acceptance suite. Also create a document from `$artifact` and verify inherited properties; after conversion, inspect actual paragraph styles. These are three different observations.

For exact XML, use a suite operation `{"op":"xml","target":"document"}` after the edit and after its undo record closes. Retain the raw snapshot and declare any comparison exclusions. Counts alone cannot prove preserved content or formatting. For native layout, `{"op":"render","target":"document","value":144}` captures pages; optional `child` chooses a 1-based page. Inspect the returned images, not merely file existence.

For UI, locate/invoke actual controls using the native UI operations, then assert the document result after completion. A returned invocation can precede completion. Use read-only eventual observations where appropriate; never retry a modifying action just because its result is delayed.

## Requested user preview

```json
{"path":"dist/Template.dotm","document":"inputs/manuscript.docx"}
```

Pass this to `--execute call preview`. The current implementation opens an isolated copy with the template attached, enables automatic style updates, applies them immediately, saves and reopens the copy, and verifies the attachment and setting. Acceptance is not a prerequisite; an optional reference reports its status and whether the artifact hash matches. Omit `document` to start a new document from the template. This does not run the conversion workflow and does not replace conversion acceptance.
