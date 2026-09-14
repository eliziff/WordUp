# Agent operating contract

THIS PROJECT HAS ZERO USERS: NEVER PRESERVE A LEGACY SHIM, EXISTING WORKSPACE FORMAT, OR BACKWARD-COMPATIBILITY PATH.

ALWAYS RUN GO THROUGH `tools/go.ps1`; NEVER USE `go`, `go run`, OR DEFAULT GO CACHE/TEMP PATHS DIRECTLY.

THE INNER EDIT/BUILD/CHECK LOOP MUST STAY FAST. Reuse a warm Word session and run the narrowest relevant check while iterating; reserve fresh Word startup, rendering, signing, and broad corpus/acceptance runs for explicit outer-loop verification. Measure named operations separately. If routine work takes seconds or any command is unexpectedly slow, stop narrating or merely increasing its timeout: identify the blocking boundary, profile it, and fix or remove it. A deadline contains a runaway; it does not excuse latency or replace diagnosis.

Choose proportionate evidence of the user's intended outcome. When an expected artifact is useful, mechanically copy the starting Word XML, edit only the intended differences, and compare it directly with XML captured after the operation in Word. Ordinary assertions may suffice for smaller changes. Test reusable behavior beyond its development example and investigate mismatches rather than automatically updating expectations. See [expected XML](docs/EXPECTED-RESULTS.md).

You are controlling a local executable called WordUp. Complete the user's Word template work yourself; do not ask them to import modules, open VBE, configure a VM, or run/debug your tests manually.

1. Locate the supplied executable, run `version`, `help`, and `doctor`. Read README.md and docs/LIMITS.md. Native execution requires already-installed Word. Do not assert native success from these probes.
2. Run the app's `--execute selftest` in a disposable output directory when execution is authorized. It generates its own native template and assertions. Inspect the actual failure report and owned-window diagnostics if it fails. Do not hide failure behind passing offline checks.
3. Import the user's original artifact into a separate workspace; do not overwrite the original. Use a warm local session or stdio MCP while iterating. Explicit `--execute` authority permits arbitrary code with the user's OS authority; a private desktop is not a malware sandbox.
4. Edit ordinary UTF-8 VBA, persistent form data, RibbonX, and native OOXML. Unknown properties must not be discarded. Use raw package XML or native Word access. Do not flatten legal documents or turn editable Word features into images.
5. Use `read` hashes for guarded writes, and check the build report's exact output hash. A successful package build is not a VBA compile or Word runtime pass.
6. Test exact intended behavior and counterexamples on disposable documents. Assert selection boundaries, formatting preservation, document structure, repeated-operation behavior, form event results, Ribbon callbacks, and generated content as appropriate. Use actual Word renders for layout acceptance. Reference-image measurements are estimates until checked against a native render.
7. Explore unfamiliar functionality through native object-model access or scratch native VBA. Do not replace complex Word behavior with a mock implementation and call it verified.
8. Run fresh, artifact-bound native acceptance before deployment. A self-test validates only its own recorded assertions. Observe compiler results separately; never mark compilation true just because one macro ran. Record Windows and Mac evidence separately.
9. Deploy through the app only after relevant tests pass. Keep backups and stale-target guards. Windows pending activation registers automatic login recovery; the recovery command has been tested, but an actual reboot has not. Use restore with an installed activation plan to recover its previous template without overwriting newer user edits.
10. Report real gaps. Current Windows evidence is in docs/VALIDATION-WINDOWS.md; Mac execution remains unverified. Native-host diagnostics are work for the agent, not a reason to invent a pass. A passing fixture establishes its tested behaviors, not every possible template or third-party control.

Authoring default: wrap each user-facing editing action in one `Application.UndoRecord` custom record. Disable `Application.ScreenUpdating` for bulk edits, remembering its previous value. Shared success/error cleanup must close any opened undo record and restore the previous screen-updating value. Read-only actions need no undo record. Verify one-step undo and state restoration after failures. Capture WordOpenXML outside the editing action and its undo record: a native opening-layout test demonstrated that exporting it during the record disrupted undo grouping.

Quick workflow (replace paths and use @JSON files to avoid shell escaping):

    wordup import ORIGINAL.dotm WORKSPACE
    wordup -w WORKSPACE --execute session start
    wordup -w WORKSPACE rpc help {}
    wordup -w WORKSPACE rpc build {}
    wordup -w WORKSPACE rpc native.call @operation.json
    wordup -w WORKSPACE --execute test ARTIFACT.dotm SUITE.json
    wordup -w WORKSPACE --execute call deploy @deployment.json
    wordup -w WORKSPACE session stop

No user-provided API keys or external service accounts are needed by this harness. A coding/vision agent is external to the harness, and Microsoft Word is not redistributed inside it.

For footnote/endnote revision tests, inspect the relevant story or note Range.Revisions. Do not assume Document.Revisions includes those edits: an actual Supra regression showed it did not. Range.Text can include tracked deletions; verify final text by accepting revisions in the disposable story (and undo that test-only acceptance separately), or inspect the revision-aware XML. Do not blindly remove the last character of a Footnote.Range: remove a trailing paragraph mark only when one is actually present.
