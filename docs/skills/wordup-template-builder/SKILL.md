---
name: wordup-template-builder
description: Build or refactor reusable Word DOTM templates from manuscripts, existing macros, and visual style references using WordUp. Covers saved styles, heading conversion, reusable content, UI, and native acceptance; use for template workflows rather than one-off document formatting.
---

# Build a reusable Word template

Deliver a reusable `.dotm` and evidence for its intended workflow. A formatted manuscript, successful compile, or working button alone is insufficient. Use WordUp's source workspace and native acceptance suites; do not make agents recreate build, launch, capture, or comparison infrastructure in one-off scripts.

Read [the acceptance plan](acceptance-plan.md) when defining requirements and [native checks](native-checks.md) when implementing them. WordUp's [API](../../API.md), [source format](../../../README.md#source-format), and [reuse guidance](../../REUSE.md) describe the existing primitives.

## Minimum workflow

1. **Establish the target.** Inspect the input's saved styles and structure with `reference.document`; inspect the visual reference itself. Record each required document role and its destination style, heading/outline level, numbering behavior, and reference location. Distinguish measured conventions, user requirements, and unresolved inferences. Do not transplant another publication's hierarchy merely because its macro is the starting point.
2. **Author the template itself.** Save the style system in the DOTM: fonts, paragraph spacing/indents, inheritance, following style, outline levels, numbering links, and relevant pagination properties. Define page/section geometry, running matter and reusable content where required. Use `styles/recipe.json`, package XML, and native operations as appropriate; inspect the rebuilt artifact, not just recipe intent. Do not add a content recipe merely to edit styles: it replaces the body.
3. **Implement conversion separately from attachment.** Define how existing paragraphs map into the destination styles, including direct formatting and inconsistent source headings. Use native structure, style metadata, numbering and text evidence; expose ambiguity rather than silently inventing roles. Attaching a template and updating its styles does not apply newly named styles to existing paragraphs. Test new-from-template and conversion of unrelated documents as separate workflows.
4. **Build the whole interaction.** Give users a discoverable entry point and coherent controls for the promised workflow. Exercise actual Ribbon callbacks and form events, keyboard/focus behavior, enabled/disabled states and completion/error feedback. Review actual appearance at usable sizes. An accessible label or successful dispatch does not prove attractive layout or completed editing.
5. **Make editing recoverable and bounded.** Group each editing action into one undo record; suppress screen updates during bulk work and restore prior state on success/failure. Verify undo, repeat behavior, selection boundaries and content preservation. Define progress/time limits for expensive operations; collect WordUp runtime diagnostics when they fail.
6. **Prove requirements on the saved artifact.** Run source checks, observed native compilation and artifact-bound assertions. Cover required heading levels individually, stored style definitions, conversion results, reusable content, real UI actions and representative manuscripts. Capture exact XML outside undo records. Compare only under explicit, justified policies. Inspect native page renders and UI captures against the reference; log pages/controls actually reviewed. Never count a missing case as passing.
7. **Hand off honestly.** Report verified scope and unresolved gaps separately. Use `preview` only when the user requests visible Word, with a fresh copied manuscript and automatic style updates. Leave execution to the user when they ask to run it themselves. Preserve original documents/templates.

For source-structure detection, read [the shared-core boundary](../../STRUCTURE-DETECTION.md). Reuse existing grammar and document-wide evidence; do not put a rich per-paragraph COM observer into an editing hot loop.

## Acceptance rule

Every required row in the acceptance plan needs an observable check and retained evidence, or an explicit unresolved status. A reference-driven template is not ready merely because the test suite passes: missing hierarchy/style/UI coverage remains unfinished. Mark irrelevant features not applicable with a reason; do not add features just to fill a checklist.

Use warm native sessions for focused iterations and fresh acceptance for the final artifact. Keep tests proportional: one meaningful boundary case is better than repeated existence checks. Do not promise full fidelity, parity, or performance beyond the actual inputs and observations.
