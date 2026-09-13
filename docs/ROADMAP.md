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

## Heading detection using document structure and Word XML

Combine the existing techniques in `legal-structure-parser` with richer Microsoft
Word XML information to improve heading detection across documents. Separate
the source hierarchy from the target template's styles; stored outline levels
are evidence, not necessarily the intended structure.

Investigate paragraph styles and their inheritance, direct and inherited outline
levels, numbering definitions and numbering levels, run/paragraph formatting,
heading text patterns, and surrounding document structure. Compare those signals
with the parser's existing techniques before choosing a combined approach.

Acceptance should cover ambiguous or missing outline levels, manually formatted
headings, multilevel numbering, and ordinary paragraphs that resemble headings.
Measure detection accuracy and correct target-style assignment on representative
documents while preserving text, numbering, and unrelated formatting. Low-confidence
cases should remain reviewable rather than silently receive the wrong heading level.
