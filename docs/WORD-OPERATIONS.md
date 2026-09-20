# Word operations

Preserve character-level semantics when optimizing character-level edits, and verify equivalent output on representative documents.

## Bulk editing and undo

Read unchanged text once and use retained ranges or batches where possible. Repeated collection indexing, full-story reads and evidence capture can dominate editing time. Use For Each when edits do not invalidate enumeration; destructive edits may require reverse order or stable ranges. Measure setup, editing and observation separately.

Normally use one Application.UndoRecord custom record per user-facing edit. Save Application.ScreenUpdating before disabling it for bulk work; restore its prior value and close any opened record on both success and failure. Read-only actions need no undo record.

Native checks have shown that WordOpenXML export inside an undo record can split grouping. Capture outside it. A main-story ReplaceAll after a header edit, and a footnote insertion immediately followed by ReplaceAll, also split records in reproduced sequences. Where applicable, put header writes last and avoid that insertion/ReplaceAll adjacency; verify the actual restored state. Document.Undo can return False for a record spanning header stories even when it restores the document.

## Revisions and notes

Document.Revisions does not necessarily include footnote/endnote edits; inspect the relevant story or range. Range.Text can include tracked deletions. Use revision-aware XML, or accept revisions only in a disposable copy when checking final accepted text. That check does not replace testing the production sequence with pending revisions intact.

Text offsets are not automatically Word range offsets, especially around fields and revisions. Preserve fields and verify the target range before writing. Do not blindly trim the last character of Footnote.Range; remove a trailing paragraph mark only when present.

## UI and saved output

A returned UI action may precede completion. Observe the resulting document or form state; bounded read-only waits are available without repeating the modifying action. See [diagnostics](DIAGNOSTICS.md) and the [API](API.md) for selectors, asynchronous calls and native captures.

Attaching a template and updating its styles does not assign newly named styles to existing paragraphs. Verify conversion separately when it is required. Native renders establish appearance; XML and object-model checks establish the properties they actually inspect.
