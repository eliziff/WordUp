# Historical Windows validation

This page preserves the recorded 0.3.0 native evidence. Current 0.4.0
acceptance is indexed in [CURRENT-ACCEPTANCE.md](CURRENT-ACCEPTANCE.md); the
commands below use the pinned Windows toolchain wrapper.

Real Microsoft Word was exercised on Windows x64 on September 12, 2026. All processes launched for this work ran Below Normal. The supplied Mac backend was not executed.

## Release acceptance

The generated, unsigned `WordUpStudio.dotm` passed **16 native assertions** in a fresh owned Word process using the 0.3.0 executable. Its SHA-256 was `3d61a7b267c4a7a79fd699e3d5366a493fe5ece45b5329ecf12829bf478aca19`. The suite SHA-256 was `829d60898f521e72b9751d89093e9e25e3ec221a4e75d2c08b8e8aa5c7033a8a`. This release run took 11.38 seconds including startup and cleanup; it was not a warm benchmark.

Coverage includes whole-project compilation; persisted nested MultiPage forms; Ribbon initialization and a real button action; scratch VBA; actual Word page rendering; form event callbacks; a template hotkey; a real context-menu action; document text; inherited content controls; saved building-block insertion; and selection boundaries with formatting preservation. An explicit UserForm readiness check precedes window capture. See the [native form](images/studio-form.png) and [Word-rendered page](images/studio-page.png).

The separate signed-template suite passed **17 assertions**, adding Word's `VBASigned` check. Its corrected warm run recorded 2.62 seconds, with actual form, Ribbon and page images inspected. See [performance measurements and their limits](PERFORMANCE.md). A passed fixture does not certify arbitrary templates, all controls or all Office builds.

## Failure and containment checks

- Scratch VBA returned the expected value, and a deliberate failure returned error number 5 and source line 100.
- Completion notifications were tested independently of polling. A short caller deadline left the diagnostic lane alive; the independent task watchdog terminated an infinite macro.
- Earlier native isolation tests verified that another Word instance survived termination of the owned runaway job and that source artifact bytes stayed unchanged.
- The 0.3.0 compile command successfully compiled and signed a generated project. A deliberately malformed declaration then failed native compilation, retained the candidate and VBE screenshot, and preserved the previous signed output byte-for-byte. The observed failure took 10.94 seconds.

## Delivery and recovery

`TestNativeLoginRecovery` generated a template, obtained 16 fresh native assertions, registered the current-user login command, executed its actual Windows Script Host script, verified installation of the exact tested bytes, checked removal of the login entry, and restored the previous file. It passed in 10.30 seconds. This exercised the resume mechanism, **not an actual OS reboot**.

Policy tests also cover interrupted replacement before the installed-state write, backup integrity, rejection of newer target edits, and repeatable restore. Signing tests separately cover automatic certificate creation/reuse and tampered VBA rejection. Local certificate integrity and trusted-publisher status are different report fields; no personal certificate is hardcoded.

## Reproduce

```powershell
.\tools\go.ps1 test -p 2 ./internal/... ./cmd/...
$env:WORDUP_NATIVE_TEST = '1'
.\tools\go.ps1 test -p 2 ./internal/verify -run '^TestNativeCompletion$' -count=1 -v
.\tools\go.ps1 test -p 2 ./internal/deploy -run '^TestNativeLoginRecovery$' -count=1 -v
```

Native tests require usable installed Word. `WORDUP_SIGNING_TEST=1` enables the separate signing test when SignTool and Office SIP are installed. `TestNativeFeedback` is the broader deliberate-error/isolation suite; it is opt-in and substantially longer than ordinary acceptance.

Local detailed records are excluded from public Git because they contain machine paths and signing identity metadata. Relevant records are `build/release-final-acceptance.json`, `build/release-compile-pass.json`, `build/release-compile-fail.json`, `build/lightning-ready-warm.json`, and `build/release-tests.txt`. The public evidence summary preserves the distinction between fresh execution, warm execution, policy fixtures, and historical offline results.
