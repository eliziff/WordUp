# Native verification implementation status

WordUp owns execution, input isolation, optional evidence recording, comparison,
profiling and corpus orchestration. Projects own workflows, documents, domain
assertions and actual findings. Shipped tests use generated fixtures.

## Implemented and checked

- Opt-in VBA serial recorder with live document/story identity, ordered
  write/result pairs, range observations and cross-stage live-anchor touches.
- Trace validation, readable reports, exact ordered comparison and hash-checked
  frozen evidence. Coverage remains explicitly declared, not inferred.
- Native generated regression covering tracked serial text/format edits,
  downstream overlap, repeated events, one-step undo and distinct stories.
- Disposable input isolation and immutable suite placeholder expansion.
- Native SaveAs/unload regression reopening changed bytes at the original
  filename, without relying on the test runner's per-run filename suffix.
- Generic `test.corpus`: sequential warm execution, fresh input copies,
  immutable attempts, executable/artifact/input hashes, verified frozen-evidence
  resume, failure aggregation and a readable index. Generated native fixtures
  verify repeat execution and resume without changing original documents.
- File-backed WordOpenXML capture streams the native BSTR to disk, preserving
  UTF-8 bytes and hashing the saved file without the inline response-size cap.
  Inline responses retain their existing limit. Tests cover large exports,
  surrogate boundaries and preservation of an existing file on malformed XML.
- Timestamped Word startup milestones are retained on successful launches as
  well as failures, including the accessibility connection and settings reads.

The combined Go suite and fresh Windows serial/corpus/isolation/startup checks
passed on 19 September 2026. The native integrated self-test passed 22 assertions
with actual compilation in `build/acceptance-final-core02`. Its context-menu
check now runs in the document created from the attached template, alongside
the real Ribbon/form checks; the custom menu was absent in the reopened
template-editor window. This is evidence for the attached-document workflow,
not a claim that the template-editor context-menu behavior has been repaired.

No accepted-revision simulation, stage reordering or in-undo XML export is used.
Neither trace equality nor normalized XML equality establishes DOCX byte identity.
Exhaustive read logging and universal dependency graphs are not prerequisites.

See [corpus usage](CORPUS.md). Project-specific corpus content and acceptance
findings stay in projects. Replay uses ordinary suite operations and assertions.

Windows native checks establish only their recorded assertions. Mac execution
remains unverified. No deployment is implied by this status document.
