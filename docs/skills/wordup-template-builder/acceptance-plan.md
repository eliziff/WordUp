# Template acceptance plan

Optional checklist for substantial template work. Use only rows relevant to the request; no copied plan or separate document is required. This is a reference, not a coverage certificate. Keep unknowns visible.

Record source/reference paths and hashes, target artifact hash, platform, and the intended workflow (new documents, conversion, or both).

| Requirement | Expected behavior / source evidence | Automated observation | Visual evidence | Status / gap |
|---|---|---|---|---|
| Each document role and heading level | Destination style, semantic outline level, numbering; reference page/section | Saved style properties AND converted paragraph style/outline/numbering | Representative occurrence of every required level | Unresolved |
| Style system | Font, size, emphasis, base/next style, spacing, indents, keep-with-next/page breaks, gallery visibility as needed | Reopen DOTM; inspect named styles and property values | Style gallery plus rendered specimens | Unresolved |
| Page/section design | Geometry, margins, first/odd/even headers, page numbering | Native section/header/footer/field properties | First, interior and section-boundary pages | Unresolved |
| Existing-document conversion | Mapping rules, permitted text/format changes, ambiguity behavior | Before/after XML; paragraphs actually assigned destination styles; ambiguous input case | Representative converted manuscript | Unresolved |
| New document | Correct styles, starting content and workflow without prior manuscript state | New from exact DOTM; native property/content assertions | New document and initial UI | Unresolved |
| Reusable content, if required | Building blocks, fields, tables and insertion context | Insert actual saved content and verify editable structure | Inserted result | Unresolved |
| Ribbon/forms | Required actions, states, keyboard flow and appearance | Real interaction followed by completion and result assertions | Actual ribbon/forms, including relevant states | Unresolved |
| Recovery/preservation | Undo, repeat semantics, protected content, notes/tables/fields/bookmarks as applicable | One-step undo; repeat run; meaningful error case; exact XML with declared policy | Review affected boundaries where needed | Unresolved |
| Performance | Representative input size and useful latency target | Per-operation and full workflow timings; partial diagnostics on timeout | Not applicable unless responsiveness is a requirement | Unresolved |

For each evidence entry, identify the report, suite step, artifact hash and image/page where applicable. A test of one heading level cannot close the hierarchy row. Reference ambiguity is a gap to resolve or disclose, not permission to manufacture a convention. Scope visual claims to the pages and UI states reviewed.
