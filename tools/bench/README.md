# Agent benchmark

Measures whether WordUp's model-facing abstractions (bundled VBA components, journal generation, structure detection) actually help a coding agent build better templates than the core surface alone. This is the gate that decides which abstractions stay: an abstraction survives only when the `components` arm beats `core` on hidden native pass rate, or cuts tokens or wall time meaningfully on the tasks it applies to.

## What runs

For each task, arm and run the runner:

1. copies the fixture workspace (or generates the Studio example) into a disposable run directory;
2. stages `tool\wordup.exe`, a `tool\wordup.cmd` wrapper that pins `WORDUP_TOOL_PROFILE` (`core` hides `component.*`, `journal.*`, `structure.*` from the manifest and from dispatch) and a run-local TEMP, plus the user docs;
3. writes an `AGENTS.md` with the task prompt and the arm's note, then runs `codex exec` (model `gpt-5.6-luna`, reasoning `xhigh`, `--json`, `--output-last-message`, `--output-schema schemas/result.json`) with the prompt on stdin;
4. grades the result with the task's hidden native suite (`wordup --execute test`), offline `check` diagnostics and file-content checks, and records tokens, wall time, command count, timed-out flag and leftover Word processes.

Arms: `none` (control: grade the untouched fixture, which must fail for a seeded defect or feature task), `core`, `components`. Codex runs with `danger-full-access` by default because Word needs its own profile directories; the run directory is disposable and the fixture originals are never touched.

Tokens are recorded, not priced: the Codex path is subscription-backed. Do not reset or probe account usage counters from here.

## Commands

    python tools/bench/run_bench.py --arms none --runs 1          # validate suites: everything should FAIL
    python tools/bench/run_bench.py --arms core,components --runs 5
    python tools/bench/run_bench.py --task alr-* --exe dist/older/wordup.exe --label baseline

Results land in `tools/bench/results/<stamp>/` (gitignored): per-run `summary.json`, `events.jsonl`, `last.md`, `grade-test-report.json`, and an aggregate `summary.md` with per-task medians.

## Tasks

Journal-generation tasks (`journal-*`) are produced by `python tools/bench/make_journal_tasks.py --exe build/wordup.exe`: each fixture is a blank workspace holding the journal's inferred profile (`journal/profile.json`), the prompt asks for a Public Sub `WU_Typeset` that applies it unattended, and the hidden suite asserts the concrete measured values (page size, margins, fonts, heading case, running head or page number, footnote tab, curly quotes) plus body-text preservation and one-step undo. The `none` arm must fail them; a template produced by `journal.create`, whose generated `WU_Typeset` follows the same contract, is the positive control.

`tasks/*.json` declare `id`, `fixture` (`{"type":"workspace","path":...}` relative to the repo, or `{"type":"example"}`), `prompt`, and `grade` (`suite` path relative to this directory, optional `check_forbid` substrings that must not appear in `check` output, optional `file_contains` globs). Suites live in `suites/` and use the ordinary acceptance-suite schema; the agent never sees them. Workspaces under `workspaces/` are private and gitignored, so the ALR tasks only run on a machine that has them.
