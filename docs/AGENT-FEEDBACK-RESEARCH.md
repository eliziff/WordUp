# Agent feedback: evidence and acceptance requirements

The product target remains complete, locally authored and shareable Word macro templates: VBA, UserForms, RibbonX, hotkeys, context menus, styles, building blocks, and conversion of existing manuscripts. A narrow demo is not product acceptance. Human preview is optional; routine verification runs in owned hidden Word.

## Sources and implications

- [SWE-agent, 2024](https://arxiv.org/abs/2405.15793) experimentally studies agent-computer interfaces and reports improved software repair performance with a purpose-built interface. This supports investing in usable operations and feedback; it does not establish that any particular WordUp diagnostic improves agent performance.
- [Writing effective tools for agents](https://www.anthropic.com/engineering/writing-tools-for-agents) recommends realistic tasks, held-out evaluations, meaningful compact responses, and measuring tool errors, calls, runtime and token use. Apply this to complete template-building and repair tasks; do not equate a growing tool count with broader usable capability.
- [Demystifying evals for agents](https://www.anthropic.com/engineering/demystifying-evals-for-ai-agents) distinguishes agent trajectories from actual outcomes and discusses complementary evaluation layers. WordUp needs both reproducible execution records and independent document/UI acceptance. Agent narration is not an oracle.
- [Playwright best practices](https://playwright.dev/docs/best-practices) uses retrying assertions and failure traces. WordUp should poll observable state within a deadline, without replaying editing actions. Retain evidence from the original failed attempt; retry success cannot explain or erase it.
- [Playwright trace viewer](https://playwright.dev/docs/trace-viewer) connects action history to snapshots, logs, errors and timing. Transfer the pattern to Word: action identity, document/selection state, accessibility target, before/after evidence and failure diagnostics in one navigable report. Playwright itself is not a native Word driver.
- [Microsoft VBA Err object](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/err-object) documents runtime error properties. Preserve number, source and description before cleanup can overwrite them. A returned error string alone is insufficient.
- [Microsoft Office automation diagnostics](https://support.microsoft.com/en-us/office/you-receive-run-time-error-429-when-you-automate-office-applications) recommends Erl for locating failures. Erl needs numbered VBA statements; an uninstrumented zero is not a source location. Instrumentation must preserve continuation lines, labels and conditional compilation, and retain a source map.

These imply the following WordUp acceptance requirements, rather than blanket claims about model improvement:

1. Every failure retains its structured cause through RPC, asynchronous execution, suite step wrapping and report serialization. Source locations must distinguish original and generated code. Missing diagnostics must be explicit.
2. Compact failure summaries link to full raw evidence. Large accessibility trees must not obscure the failing statement or assertion. Evidence should be captured once and reused.
3. Assertions check document outcomes, selection boundaries, fields, bookmarks, footnotes, formatting and one-step undo. Native UI actions must execute the actual callback and produce those outcomes.
4. Exact XML is retained alongside targeted semantic comparisons and native renders. Serialization changes are investigated with no-operation controls; broad exclusions cannot manufacture parity.
5. Authoring, exercising, inspecting and comparing are normal reusable app operations. Required basic workflows must not depend on bespoke orchestration scripts.
6. Diagnostic quality is tested by seeded failures: compile errors, wrong callbacks, malformed forms, runtime errors, stalled operations and incorrect output. Tests must demonstrate detection and useful localization, not merely that logging code ran.
7. Measure cold/warm latency, CPU and peak/idle memory after correctness. Keep expensive evidence failure-triggered where possible; preserve explicit visual acceptance. Faster verification cannot drop coverage silently.

## Completion gates and diagnostic evaluation

The following are acceptance targets, not a declaration that they are implemented. The execution harness can be tested without a model subscription. Measuring improvement in model repair success additionally requires a fixed external agent/model and comparable runs; no such controlled model comparison has been completed.

| Surface | Required evidence | Current gap to close |
| --- | --- | --- |
| VBA compilation | Seeded syntax/type/reference errors, module and source location, actual native compiler result | Scratch Erl diagnostics do not prove localization of all artifact compile failures |
| Runtime and crashes | Original error, execution identity, last known action, bounded timeout, retained source and owned-process diagnostics | Full paused locals/call stack and the original quick-demo crash remain unresolved |
| UserForms and RibbonX | Package validation, load, actual event invocation, state transitions, resulting document change, native image | Form-load and screenshot tests do not cover every callback or arbitrary control |
| Hotkeys and context menus | Installed binding, actual invocation, target selection, outcome, removal/restore | End-to-end coverage remains incomplete |
| Template conversion | Arbitrary existing input, styles, sections, notes, fields, building blocks, metadata and native pagination | Complete publication-specific parity and the full editorial workflow remain incomplete |
| Mutation safety | Protected-region rejection, one-step undo, state restoration after success/error, repeated invocation | Must be exercised per editing feature, not inferred from helper existence |
| Agent usability | Fresh workspace to finished template and seeded repair through documented operations | Record required ad hoc scripting, ambiguous errors, failed discovery and missing operations as product defects |
| Performance | Same accepted tests and artifacts, cold/warm timings, phase breakdown, CPU and idle/peak memory | Individual fast runs are observations; controlled distributions and resource baselines remain necessary |

For each seeded defect, preserve the defect input, exact artifact hash, failing assertion, raw evidence, corrected artifact and unchanged acceptance check. A useful diagnostic must detect the defect and narrow the repair location or violated invariant. Include missing/wrong callbacks, wrong selection boundaries, lost formatting, unavailable controls, runtime faults and bounded runaway cases. Compare diagnostic variants on held-out tasks using repair success, false-pass rate, tool calls and elapsed time. Do not allow a repair to weaken its acceptance criteria unnoticed.

Report-level assertion failures now retain structured `assertion_failed` details through contextual wrapping, including expected/actual values, assertion index and JSON pointer. `actual_available` and `pointer_error` distinguish a missing observation from an observed null. The existing flat step-error fields remain available. A focused regression checks both cases; this is evidence of diagnostic transport correctness, not a measured improvement in model success rate.

## Current concrete gap

An unrelated-manuscript test compiled but failed at runtime with error 438 and Erl 0. Detailed scratch source was present in the step observation, but the report-level error lost the structured cause when wrapped with a step name. The shared report conversion now unwraps native faults while preserving the contextual message; a regression test checks code and line retention. The runtime cause was that an opened template was not installed as an add-in for the active manuscript. Adding the existing `addin` operation fixed the test without changing the macro.

Scratch evaluation now numbers eligible physical statement lines automatically in both direct and prepared evaluation. Runtime faults include `body_line`, `body_line_text`, `line_numbering`, and `location_kind`; complete generated source remains available. Existing caller numbering is preserved, as are continuation lines, comments, named labels and compiler directives. Source bodies exceeding VBA's line-number range remain unnumbered with an explicit mode. `TestNativeScratchLocations` exercises both executors in actual hidden Word, including continuation syntax and conditional branches. This provides the last numbered statement reported by Erl, not a call stack or automatic instrumentation of the tested template's own code.

Fresh conversion/undo acceptance passed in 2.496 seconds after this change. A separate large manuscript test passed with 156 footnotes and 61,775 body characters; conversion plus its preservation checks measured 496 ms, and the full run with exact input/output XML took 4.362 seconds. These prove the current page/footnote subset, not completion of a complete editorial workflow.

## Real manuscript corpus

Read-only discovery found private manuscripts with both letter-sized and compact
page geometries, notes, tables and inconsistent editorial structure. These are
local inputs, not redistributable fixtures. Match documents by content before
asserting before/after parity; filenames and editing-stage labels alone are
insufficient. Development examples must not silently become generic style
requirements.

## Targeted native visual diagnostics

`ui.diagnostics` accepts `named.window` (exact owned window title) or `hwnd` to inspect a particular visible owned window. Owned modal dialogs remain included; no selector retains the original full capture. A missing requested target returns `diagnostic_window_not_found`, exercised in the native test. This makes routine form inspection reusable without collecting every document window. In a representative UI suite, targeted form capture took 148 ms versus 3,076 ms for the prior full capture; full-suite elapsed time was 4.107 seconds versus 7.802 seconds. A later run after changing the form status copy took 4.929 seconds. These are individual observations, not a guaranteed latency or a resolution of the separate UIA expand/collapse delay.
## Source-grounded manuscript roles

`reference.document` paragraph observations now include source part, exact UTF-8 byte range, namespace-qualified XPath, available Word paragraph ID, and inline footnote/endnote/comment references. Root metadata supplies the XPath namespace binding and explains that these are package XML locations, not native Word paragraph indexes. Nested textbox paragraphs are observed separately instead of contaminating the containing paragraph's text/style. Foreign-namespace text is excluded; inline tabs/breaks are retained without mistaking paragraph tab-stop settings for document text. A regression test covers those cases. A large manuscript inspection took 18.247 ms and identified title paragraphs at XML positions 3/4 and the author at position 6 with footnote ID 1. Inferring editorial roles still belongs to the agent; these factual locations do not claim automatic classification or native layout validation.

## Further feedback-loop research: actionable evidence

[Anthropic tool engineering](https://www.anthropic.com/engineering/writing-tools-for-agents) emphasizes explicit tool contracts, useful error responses, and evaluation against realistic tasks. [Its agent evaluation guidance](https://www.anthropic.com/engineering/demystifying-evals-for-ai-agents) separates task outcomes from the agent trajectory and recommends multiple evaluation layers. [Playwright traces](https://playwright.dev/docs/trace-viewer) associate actions with snapshots, logs and timing; [auto-retrying assertions](https://playwright.dev/docs/best-practices) observe readiness instead of assuming fixed delays.

For WordUp, the concrete targets are artifact-bound outcomes alongside the exact operation transcript, expected/actual assertion values, compiler source locations and observed dialogs, runtime source information where available, native XML and rendered/UI evidence, and independently reported operation latency. Existing polling must retry observations rather than replay editing actions. Missing evidence must remain distinguishable from a negative result. These are engineering applications of the sources, not proof that WordUp improves model success rates. That requires repeated agent task completion and diagnosis evaluations against a fixed corpus, including seeded failures; it remains outstanding.
