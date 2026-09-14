# Agent diagnostics

Native sessions use an owned private desktop by default. Visible preview or a
visible test session requires an explicit user request; debugging is not an
implicit request to display Word.

Use `native.call` with `{"operation":{"op":"xml.snapshot","target":"body",
"file":"reports/after.xml"}}` to capture a document or range handle. Obtain a
range with `get document.Content as body`. Omitting `file` returns the XML inline.
The result includes SHA-256, byte count and package part names. The UTF-8 file
preserves the live `WordOpenXML` string without XML normalization or document
saving. It is Word's Flat OPC representation, not the original ZIP bytes.
Capture before and after, then assert namespace-aware paragraph, run, style,
field and relationship structure. Preserve the raw evidence even when comparing
selected semantic properties; Word can change incidental metadata.

`vba.immediate` accepts `{"text":"? ActiveDocument.Paragraphs.Count"}` and
returns a typed result, generated source and scratch artifact hash. It uses a
separate native VBA project. Set `paused:true` to evaluate against a currently
paused VBA frame through the owned VBE Immediate Window; optional `symbols`
return best-effort parameter/local/module-variable rows and explicitly list
unavailable symbols. The paused path returns buffer text, not a typed value.
For statements use `native.call`
with `{"operation":{"op":"eval","value":"Evaluate = 6 * 7"}}`.
Both require execute authority. Evaluation has a default 30-second deadline.

Scratch runtime errors include number, description, source, `Erl` and generated
source. `Erl` is meaningful only for code with VBA line numbers. These wrappers
do not automatically instrument arbitrary existing macros, collect a VBA stack,
or catch native Word access violations. The Windows host monitors its owned Word
process and reports its PID and exit code when it exits; host information states
whether that monitor was established. UI diagnostics remain a separate channel.
Worker stream failures also retain the worker PID, bounded log tail, and exit
code when available; a still-running worker is reported explicitly. The opt-in
`TestNativeOwnedWordExitDiagnostics` and `TestNativeOwnedWorkerExitDiagnostics`
prove both paths using controlled termination of newly owned processes, not a
reproduction of the original demo crash.

A successful accessibility action is not evidence that an asynchronous form
callback completed. Assert the resulting document and form state. Bounded
`eventually_ms` on read-only `get` or `ui.find` steps can observe completion;
mutating actions are never automatically retried.

The reported small-demo Word crash remains unresolved. Passing scratch or XML
tests does not establish that the original form interaction is safe.

XML comparisons retain the global verdict and raw hashes, and now include the first policy-aware token difference in up to eight changed Flat OPC parts. Locations are byte offsets in the original snapshot, with bounded nearby XML. This helps expose later-part changes after an earlier metadata mismatch; it does not resolve inherited formatting or enumerate every mismatch within one part.
Use xml_policy.context_parts (an array of exact part names, such as /word/footnotes.xml) to select up to eight part contexts. This selects diagnostic detail only: all parts still contribute to the verdict and hash evidence. Changed parts without detail carry context_omitted=true.

`test.compare` also compares each matching `xml.snapshot` step automatically, validating captured bytes against their recorded SHA-256 first. It accepts `xml_policy` and returns `xml_snapshots` alongside asserted-result differences and timings. Missing or modified captures fail comparison. Raw acceptance reports and saved CLI `{result: ...}` envelopes work directly with comparison, replay, preview and deployment.

XML policy remains explicit: `generated_toc_bookmarks` maps generated `_Toc<number>` names by bookmark order, including hyperlink anchors and REF/PAGEREF targets; it does not discard bookmark positions or text. `attribute_equivalences` declares exact listed values for an element/attribute pair, useful for known temporary paths without ignoring every relationship target. These policies do not establish byte identity. Bounded token previews show the differing area; `canonical_token_byte` is an offset in the diagnostic token, while source XML offsets remain separate.

`generated_comment_ids` maps modern comment paragraph and durable IDs through the original comment identity in complete Flat OPC snapshots, preserving reply-parent links, text and resolved state; unresolved links fail, and timestamps remain differences unless explicitly excluded.

For expected modal prompts, `ui.find`/`ui.invoke` accept `named: {"scope":"dialog","name":"OK","role":43,"message_pattern":"^Expected operation completed$","wait_ms":3000}`: only an owned dialog whose message matches is selected, and its observed message is retained in the result.

Each acceptance run writes `progress.jsonl` in its evidence directory as steps start and finish; agents can read complete lines during execution to locate a stalled step without waiting for the final report. `eventually_ms` retries missing read-only `ui.find` selectors, but never ambiguous selectors or mutating actions.

Acceptance `compile` steps automatically capture and acknowledge recognized English VBE compile-error dialogs. The error retains the dialog text and native module/selection/source line. When source localization is available, failure capture targets the VBE screenshot plus remaining modal dialogs and XML, avoiding unrelated document/IME trees. Unknown/localized dialogs retain the timeout/diagnostic path; this is not a general automatic dialog clicker. Native proof: `workspaces/compile-proof/reports/automatic-compiler-focused.json` rejects an undefined type, reports BadCode line 2, and captures the VBE image; total 2.441 s, failure capture 125.907 ms. VBE line numbers exclude hidden Attribute lines in exported source.
