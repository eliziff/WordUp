# Capability and evidence boundaries

These are release gaps, not redefinitions of the requested product.

`xml.query` accepts UTF-8 XML up to 16 MiB and 200,000 elements, with the
existing XML depth/DTD restrictions. XPath expressions are limited to 4096
bytes, namespace maps to 64 entries, and returned matches to 100 (default 20).
Cancellation is checked around evaluation and between matches; upstream XPath
evaluation is synchronous and has no hard deadline for pathological predicates.
Do not treat it as a sandbox for hostile queries. Queries never rewrite source
XML or establish effective formatting/layout fidelity.

| Requirement | Current evidence and limits |
|---|---|
| Native Windows authoring and execution | Actual Word compile, VBA runtime, nested persisted forms, Ribbon callbacks, context menu, template hotkey, document contents and saved parts passed. These fixtures do not prove every possible macro. |
| Forms and custom ActiveX | Common controls, Frame, MultiPage and Page are implemented. Unknown binary data is preserved; unsupported edits fail. Third-party controls and dependencies are not universally supported or bundled. |
| Ribbon verification | Embedded 2007/2010 XSD validation on Windows, duplicate IDs and lexical callback checks, plus actual tested callbacks. Built-in idMso existence and all callback type signatures are not fully checked. |
| Native page/UI appearance | Actual Word page renders and owned-window captures have been visually reviewed. Agents must inspect their own outputs; screenshot generation alone does not establish good layout. |
| UI automation | Named selectors, bounded waits, MSAA/UI Automation and separate async macro/UI lanes. No universal owner-drawn gesture, drag or file-picker support. |
| Error diagnostics | Scratch error locations, automatic paused module/line capture and reset for recognized English VBA runtime-error dialogs, compiler observations, retained failed candidates and owned-window/XML evidence. Native Call Stack frames are captured for recognized English runtime-error dialogs; local-variable tracing is not implemented. |
| Resource containment | Owned job limits and independent asynchronous watchdog tested, including survival of another Word instance. Limits do not apply when a recipient opens the DOTM normally. |
| Fast iteration | Complete signed suite recorded 2.62 s warm with 17 assertions. Timing varies; the one-second target remains unmet. See PERFORMANCE.md. |
| Signing | Automatic local certificate creation/reuse and strongest V3 digest verification tested. SignTool and registered Office SIP remain prerequisites; automatic prerequisite installation and automatic recipient trust are not provided. |
| Delivery and restore | Exact fresh-artifact/suite checks, backups, stale-target protection, interrupted-replacement recovery and guarded restore tested. |
| Login recovery | Current-user Windows login command was registered, executed and removed in a native delivery test. No actual OS reboot was performed. Windows Script Host must be available; policy may prevent execution. |
| Document/image references | Raw XML plus document/image inspection. No embedded model, OCR service, universal resolved-style engine or font-identification guarantee. |
| Mac stretch | Offline code and Apple Events prototype exist. Actual Mac Word execution and native UI/render/compile parity remain unverified/incomplete. |

## Security and environment

Native VBA can do what the current user can do, including files, network access and external programs. A private desktop separates visible UI; it does NOT isolate malicious macros, OS globals, clipboard, printer settings or external COM applications launched by a macro. The caller authorizes execution explicitly. The app does not attach to or intentionally close the user's existing Word process. Raw native APIs remain powerful and must be used with the user's permission.

Microsoft Word must be installed, activated and able to run normally. Word is not an open-source runtime bundled with this program. Office Group Policy, enterprise application controls, downloaded-file restrictions and OS permissions remain effective; the app does not bypass them or silently lower the user's Trust Center settings. macOS Automation prompts may require actual human consent. Executables are not Authenticode-signed or notarized in this delivery.

## Binary preservation versus interpretation

Unedited data stays untouched where possible. Supported form edits preserve unmodeled binary tails, fonts and pictures; unsupported form properties/controls are not translated to approximate substitutes. Codepage handling is explicit; Windows can use installed Windows codepages, while the standalone cross-platform core directly supports its implemented encodings. Unrepresentable or unsupported encodings error rather than silently replacing characters.

The writer implements published interoperable source streams and invalidates old execution caches. It is not a standalone VBA compiler. Compression/CFB fuzzing and reparse checks are useful but not equivalent to an independent Office conformance certification. Bounded input limits protect memory: generally 256 MB source/package budgets, 64 million image pixels, bounded IPC messages and response trees. The Windows native ABI builds are 64-bit (x64 and ARM64); no 32-bit host executable is supplied.

Source assertions, deterministic byte round trips, native runtime assertions, compiler observations, layout checks, and cross-platform verification are distinct. Reports preserve that distinction. Synthetic hosts used in harness tests are explicitly labeled and never counted as actual Word evidence. Passing a native test only establishes the tested scenario on the recorded Office build.

`toc.inspect` on Windows inspects hyperlink-backed TOC entries with decimal page labels against native bookmark-range adjusted page numbers. It reports no-TOC, unlinked TOCs, non-decimal labels and missing destinations as incomplete/unresolved, not success. It does not update fields or infer arbitrary custom numbering formats; named `repaginate: true` explicitly refreshes Word pagination.
