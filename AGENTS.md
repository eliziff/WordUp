# Agent operating contract

You are controlling a local executable called Wordwright. Complete the user's Word template work yourself; do not ask them to import modules, open VBE, configure a VM, or run/debug your tests manually.

1. Locate the supplied executable, run `version`, `help`, and `doctor`. Read README.md and docs/LIMITS.md. Native execution requires already-installed Word. Do not assert native success from these probes.
2. Run the app's `--execute selftest` in a disposable output directory when execution is authorized. It generates its own native template and assertions. Inspect the actual failure report and owned-window diagnostics if it fails. Do not hide failure behind passing offline checks.
3. Import the user's original artifact into a separate workspace; do not overwrite the original. Use a warm local session or stdio MCP while iterating. Explicit `--execute` authority permits arbitrary code with the user's OS authority; a private desktop is not a malware sandbox.
4. Edit ordinary UTF-8 VBA, persistent form JSON, RibbonX, native OOXML, style/content/Quick Part recipes. Unknown properties must not be discarded. Use raw package XML or native Word access where convenience recipes are incomplete. Do not flatten legal documents or turn editable Word features into images.
5. Use `read` hashes for guarded writes, and check the build report's exact output hash. A successful package build is not a VBA compile or Word runtime pass.
6. Test exact intended behavior and counterexamples on disposable documents. Assert selection boundaries, formatting preservation, document structure, repeated-operation behavior, form event results, Ribbon callbacks, and generated content as appropriate. Use actual Word renders for layout acceptance. Reference-image measurements are estimates until checked against a native render.
7. Explore unfamiliar functionality through native object-model access or scratch native VBA. Do not replace complex Word behavior with a mock implementation and call it verified.
8. Run fresh, artifact-bound native acceptance before deployment. A self-test validates only its own recorded assertions. Observe compiler results separately; never mark compilation true just because one macro ran. Record Windows and Mac evidence separately.
9. Deploy through the app only after relevant tests pass. Keep backups and stale-target guards. Pending activation does not survive an OS reboot automatically in this preview; resume it through the app, not by asking the user to move files.
10. Report real gaps. These supplied Windows/Mac builds were cross-compiled, not executed in Microsoft Word by their author. Native-host diagnostics are work for the agent, not a reason to invent a pass.

Quick workflow (replace paths and use @JSON files to avoid shell escaping):

    wordwright import ORIGINAL.dotm WORKSPACE
    wordwright -w WORKSPACE --execute session start
    wordwright -w WORKSPACE rpc help {}
    wordwright -w WORKSPACE rpc build {}
    wordwright -w WORKSPACE rpc native.call @operation.json
    wordwright -w WORKSPACE --execute test ARTIFACT.dotm SUITE.json
    wordwright -w WORKSPACE --execute call deploy @deployment.json
    wordwright -w WORKSPACE session stop

No user-provided API keys or external service accounts are needed by this harness. A coding/vision agent is external to the harness, and Microsoft Word is not redistributed inside it.
