# Agent papercuts (WordUp loop)

Friction an agent hit while doing real template work with the toolchain,
recorded as it happened so we can decide what to fold into the agent loop.
Each entry: what happened, what it cost, and a candidate fix. Newest first.

## 2026-09-16 evening, ALR two-GUI reorganization (Claude Fable 5.1)

1. **Scratch `eval` cannot see template procedures, and a reference to one hangs
   for the full deadline.** An `eval` body that called `WU_BatchReport()` (a
   procedure that lives in the template project, not in the generated scratch
   add-in) compiled into Word's modal "Compile error in hidden module" dialog.
   `Application.Run` blocked until the 300 s `native_deadline`; the report said
   only `native_deadline` with an empty `worker_log_tail`. The previous agent
   burned two 300 s runs and an evening of tracing on what looked like a macro
   hang. Candidate fixes, in order of value: (a) a dialog watchdog during `run`
   and `eval` that notices an owned `#32770` window, captures its text, closes
   it and returns `macro_dialog_blocked` with the message (this also converts
   every stray `MsgBox` hang into an instant diagnosis); (b) compile the scratch
   project before running it and fail with `scratch_vba_compile_failed` plus the
   generated source; (c) state in `help` for `eval` that template procedures
   are reached only through `Application.Run "Name"`.

2. **Path resolution differs by verb.** `build OUTPUT` resolves the output
   against the workspace (`build dist/X.dotm` wanted; `build workspaces/ws/dist/X.dotm`
   silently created `workspaces/ws/workspaces/ws/dist/`). `test ARTIFACT SUITE`
   resolves the artifact against the workspace but the suite against the
   current directory. Operation `file` values inside a `native.call` JSON are
   repo-relative. Three failed invocations before the pattern was clear.
   Candidate fix: resolve every user-supplied path the same way (try the
   workspace, then the cwd, and say which one matched in the result), or
   document the rule in one line of `help` per verb.

3. **A failing step reports only its first failed assertion.** The verify runner
   stops at the first assertion that fails inside a step, so a step with eleven
   assertions needed a second round to learn that the later ones also held.
   Candidate fix: evaluate every assertion of the step and list all failures;
   keep stopping at the step boundary.

4. **Screenshot paths are buried.** `ui.diagnostics` with `file` writes
   `window-<hwnd>.png` under `dist/acceptance-assets/<hash>/run-<n>/…`, but the
   `test` result only carries the path inside `result.windows[].screenshot`.
   Candidate fix: surface captured files in a top-level `artifacts` list of the
   step observation and print them in the CLI summary.

5. **MultiPage tabs are driven by an undocumented role.** Switching a form's
   tab needs `ui.invoke` with `role: 37`; the only place that number appears is
   an older suite. Candidate fix: name the common MSAA roles in `help` for
   `ui.find`/`ui.invoke` (button 43, checkbox 44, page tab 37, window 16).

6. **The bench's ALR fixture is a frozen snapshot.** `tools/bench/fixtures/alr-refactor-seeded`
   is a copy of the private ALR workspace with seeded defects; after the
   template refactor it still carries the retired Style Pass form and macros,
   so the `alr-*` tasks measure an old template. Candidate fix: a
   `make_alr_tasks.py` that re-copies the workspace and re-applies the seeded
   defects from a small patch list, like `make_journal_tasks.py` does for
   journals.

7. **No live view of an owned Word from the same CLI.** While a `run` was
   blocked, a second `wordup … call native.call` starts its own Word, so the
   hung process could not be inspected (`ui.diagnostics` needs the owning
   host). Candidate fix: `doctor --owned` that lists the toolchain's live Word
   processes with their visible windows and dialog text.

Things that worked well and should stay: the `compile` op's pinpointed VBE
error (module, procedure, line text, source line) turned a compile failure into
a one-round fix; `mode: replace` form designs rendered correctly on the first
build; `begin`/`ui.invoke`/`poll` drove both redesigned forms unattended.

## Added later the same night

8. **`poll` never waits.** A `poll` step returns `{"status":"running"}` at
   once, and the next object-model step then blocks behind the still-running
   macro until its own timeout ("Native operation did not finish"). The wait
   exists but is spelled `eventually_ms` on the step with an assertion on
   `/status`; nothing in the failure pointed there. Candidate fix: make `poll`
   accept `wait_ms` directly, or have the runner suggest `eventually_ms` when
   a `poll` returns running and the next step times out.

9. **Two suite file shapes.** Half the workspace suites were wrapped as
   `{"path","fresh","suite":{…}}` (the shape another verb consumes) and `test`
   rejects them with `json: unknown field "path"`. Candidate fix: accept the
   wrapper in `test` (read `suite`, treat `path` as the artifact) or reject it
   with a message that names the other shape.

10. **A batch gives no timing until it ends, and one slow step kills it.**
    `batch` returns only after every step, so a 12-stage run gave no signal
    for eleven minutes; `profile` returns per-step `duration_ms` and keeps
    partial evidence, but nothing suggested it. Candidate fix: mention
    `profile` in the `batch` help text and in the `native_step_timeout` fault.

11. **Wildcard Find inside fields.** Word's wildcard Find over a story with a
    TOC field returns thousands of matches inside the field code whose `Text`
    reads back empty; a template author has to discover this the hard way.
    Candidate fix: a note in `docs/WORD-OPERATIONS.md` and a bundled
    field-blind Find helper (the `FindNext` pattern in `ALR_Rules.bas`).

12. **Killing a hung CLI by command-line pattern is guesswork.** With six
    benchmark workers each spawning `wordup.exe __host` and `WINWORD.EXE`, the
    only way to find the Word behind one CLI call was walking parent PIDs by
    hand. Candidate fix: `session list`/`doctor --owned` printing CLI PID,
    host PID, Word PID and the operation in flight.

13. **A hidden suite graded a requirement the prompt never stated, and read a
    mixed range.** All twelve journal runs of `20260916-230047-journal-rerun`
    failed grading: four because the suite asserted `UpdateStylesOnOpen` while
    the task prompt never mentioned it, eight because the body-font check read
    `Paragraphs(4).Range.Font`, a range that also holds the footnote reference
    mark, so Word reported a mixed font (empty name, size 9999999). The agents'
    own reports said "native acceptance passed", and they were right about what
    they were asked. Fixed in `make_journal_tasks.py` (first word of the body
    paragraph, first text word of the note, requirement added to the prompt).
    Candidate fix for the loop: a bench self-check that runs each hidden suite
    against its positive-control template AND lists every graded property the
    prompt does not mention, before any agent run is paid for.

## 2026-09-17, ALR corpus overhaul (five Opus agents in parallel)

14. **The harness reports a killed run as a pass.** `wordup … --execute test`
    writes its result as JSON on stdout; when the process dies (session limit,
    kill, host crash) or when its output is piped, stdout comes back empty and
    the exit code is not distinguishable by a caller that only parses JSON. A
    runner that reads `result.error` therefore prints "passed" for a suite that
    never ran. One false green reached a phase report before it was caught.
    Candidate fix: always emit a JSON envelope, with a fault code, on the way
    out, including on signal; and make a non-zero exit code carry a machine
    readable reason. Consumers should treat empty stdout as a failure, which
    `workspaces/alr-refactor/tools/run_suites.py` now does.

15. **Default step deadlines do not survive a busy machine.** `open` and `eval`
    default to 30 s. With four agents each holding a Word instance, opening a
    staged template routinely exceeded it, and a Suggester run over a real
    manuscript exceeded the 900 s step budget. The shared workspace suites
    carry no `timeout_ms`, so every agent hit this and each one discovered it
    separately. Candidate fix: scale the default with the number of live hosts
    the toolchain owns, or let a run set a global multiplier, and say in the
    fault which deadline was hit and what it defaults to.

16. **An agent could not write its own report file.** The corpus audit was
    asked to produce `reports/qa/SUMMARY.md` and its harness refused the write,
    so a 4,000 word analysis had to be relayed through the conversation and
    written to disk by the parent session. Candidate fix: allow a subagent to
    write under a declared output directory, or state the restriction up front
    so the agent plans a different deliverable.

17. **Two agents writing the same suite output path collide.** Two phases ran
    the shared `corpus-proposals-dyson.json`, which writes a fixed report path,
    and overwrote each other's evidence mid-analysis. Candidate fix: suites
    that write files should namespace the path by artifact hash or run id, the
    way `dist/acceptance-assets/<hash>/run-<n>/` already does for screenshots.

18. **Freezing the artifact under test is the pattern that made the audit
    possible.** Copying the build to `dist/qa/ALR_qa.dotm` let the audit measure
    a stable target for five hours while three other agents rebuilt the same
    source. Worth recommending in the docs as the way to audit a moving tree,
    together with the rule that every observation names the frozen hash and
    every root cause names a procedure rather than a line number.

19. **Word re-orients a curly quote that a replacement inserts.** With
    `Options.AutoFormatAsYouTypeReplaceQuotes` on, replacing `'9` with `’9`
    literally yields `‘9`, because the inserted mark follows a space and Word
    applies its own contextual rule to a programmatic write. Any pass that
    cares about quote direction has to switch the option off around its writes
    and back afterwards. Candidate fix: a note in `docs/WORD-OPERATIONS.md` and
    a literal-replace helper in the bundled recipes.

20. **Word's speller cannot identify an acronym.** It accepts `ai`, `us`, `it`,
    `sec`, `act` and `ghg` as words, so "is this a dictionary word" does not
    separate AI from AND when title-casing a heading typed in capitals. What
    worked instead was harvesting the capital runs the author used inside
    ordinary mixed-case text elsewhere in the same document. Candidate fix:
    record the finding; a speller check is a tempting and wrong first instinct.

21. **Seeded probes certify the wrong thing.** Every defect in the ranked list
    the audit produced, including one that silently overwrote words in a
    footnote for 72% of real manuscripts, passed the seeded acceptance suites
    for weeks. The suites proved the rules do what they should on sentences
    written to exercise them; nothing proved the rules leave alone what they
    should. The fix that worked was mechanical: run the real corpus, dump every
    proposal with surrounding context to a text file, and read it. Candidate
    fix for the loop: make "dump every change with context and read it" a
    first-class verb, so a rule change cannot be reported as verified until
    someone has looked at what it did to real documents.
