# Future features

## Optional human feedback inside Word

Let a user point at a document range, Ribbon control or form element in their real
Word session and attach a comment that the agent can resolve to that element.
This is an optional addition to the headless agent verification loop, not a
replacement with a conventional chat sidebar or enterprise add-in workflow.

## Remaining authoring and setup work

Broaden form/control coverage and Ribbon callback/type/idMso diagnostics; add
richer optional runtime tracing; automate the Microsoft signing prerequisites
where supported. Native Mac execution and UI/render/compile parity remain a
stretch goal. Current boundaries are recorded in LIMITS.md.

## ALR heading detection using document structure and Word XML

Requested September 12, 2026. Future work; not part of the current native-runtime fixes.

Combine the existing techniques in `legal-structure-parser` with richer Microsoft
Word XML information to improve heading detection in the ALR template. The current
macro maps outline levels to their ALR heading-level equivalents, but that mapping
was not consistently clean in the reference implementation.

Private reference: `ALR Macro [July 22 2026].dotm`, supplied beside the repository
source. The template is excluded from public Git history.

Investigate paragraph styles and their inheritance, direct and inherited outline
levels, numbering definitions and numbering levels, run/paragraph formatting,
heading text patterns, and surrounding document structure. Compare those signals
with the parser's existing techniques before choosing a combined approach.

Acceptance should cover ambiguous or missing outline levels, manually formatted
headings, multilevel numbering, and ordinary paragraphs that resemble headings.
Measure detection accuracy and correct ALR style assignment on representative
documents while preserving text, numbering, and unrelated formatting. Low-confidence
cases should remain reviewable rather than silently receive the wrong heading level.
