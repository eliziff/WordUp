# Word structure detection: reuse boundary

The standalone VBA observer returned by `structure.source` is currently a native evidence collector plus a small marker recognizer. It is not the completed heading detector. Keep it out of the per-paragraph editing loop: redundant COM reads can dominate conversion time; collect evidence once and reuse it.

## Local source review

Reviewed local legal-structure revision `f73bcd5d66575c0504e97fc2c3b95bf5589377e5` and legal-pdf-parser revision `e1b060bdf9b92b9f6295bfc1e7933bd5a446af22`, reviewed as developer-local upstream sources.

- `legal-structure/src/candidates.rs::detect_structure_candidate_runs`: candidate runs retain parent relationships, source ranges, grammar families and sequence validity. A candidate is separate from a resolved primary document role.
- `legal-pdf-support/src/pairing_support.rs::enumerator_interpretations`: retains multiple Roman/alphabetic interpretations; numeric hierarchies include depth. Do not collapse ambiguous I/V/X into one family prematurely.
- `parse_heading_ladder` in the same file: prefers sequence increments, opens new families, records illegal restarts, forward gaps, midcounter openings and violations; family consistency and numeric footnote suspicion remain observable.
- `legal-pdf-structure/src/structure.rs`: excludes contents and ineligible regions before classification; uses character-weighted bold share, body-flow continuity, heading plausibility, repeated family/level agreement, and wrapped heading geometry. Style evidence corroborates grammar and can demote false positives.
- `legal-pdf-structure/src/structure/graph.rs`: keeps region/source evidence separate from resolved hierarchy and its parent stack.

These mechanisms should be reused before introducing a new scoring system. PDF-specific geometry thresholds are not Word paragraph rules.

## Word adapter requirements

Collect source facts once, before editing. Retain source hash and exact locations, and distinguish Word UTF-16 story positions from package XML byte offsets. Resolve paragraph style inheritance and direct overrides; preserve native outline, list identity/level/start/restart, paragraph indentation/spacing, font family/size, character-weighted emphasis, fields/TOC and table/textbox/story containment. Store conflicting signals rather than overwriting them.

Feed eligible candidates and style-family evidence into the grammar/sequence resolver. Keep native outline evidence distinguishable from inferred levels; do not assume it is editorially correct. Unnumbered headings need style-family and neighboring-structure evidence. Quotations and ordinary numbered lists must remain subordinate, even when their local numbering is well formed. Map the resolved source hierarchy into the requested house style only after detection.

## Baseline finding

Manuscripts can contain multiple levels of unnumbered headings and a separate TOC heading. A numbering-only detector cannot recover that hierarchy. Test malformed and missing outlines as well as correctly styled inputs; native outline evidence alone does not prove general recovery.

A useful acceptance corpus includes damaged outline levels, equivalent direct-formatted headings, ambiguous Roman/letter sequences, numbering restarts, quoted instruments, ordinary lists, tables, fields/TOC, mixed inline emphasis and non-ASCII text. Require expected hierarchy and source spans; benchmark analysis separately from editing. Passing a prefix fixture is not acceptance.

## Implemented package adapter

`structure.inspect` now resolves stored paragraph-style ancestry and direct outline overrides, retains original property XML and numbering XML, and marks table/textbox containment. Style definitions, document-default run properties and the theme font scheme are emitted once; paragraphs reference their style chain by ID instead of repeating the same XML. Runs retain every direct font slot and theme token, resolved theme typefaces, exact source bounds and property XML, size, emphasis and UTF-16 text weight. Paragraph-mark run properties are reported separately when a paragraph carries direct bold, italic, caps or small-caps formatting, so heading evidence is not lost when Word stores it on `w:pPr/w:rPr` instead of each run. Direct and style-inherited numbering is joined to its concrete/abstract definitions, yielding level, family, label pattern, start, restart and start-override evidence. It reconstructs displayed labels while retaining an uncertainty flag for unsupported formats or missing ancestor counters. Resolution uses paragraph-level weighted bold/caps evidence, so a bold phrase inside ordinary prose is not promoted as if the whole paragraph were emphasized; non-bullet native numbering can corroborate a heading style without promoting an ordinary list by itself. Generic named prefixes such as Part, Chapter, Theme, Section and Appendix retain separate ladder families, so nested labels do not collapse into one sequence. Native outline claims remain evidence, not an inferred hierarchy.
