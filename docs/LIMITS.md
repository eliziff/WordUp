# Capability and evidence boundaries

These are release gaps, not redefinitions of the requested product.

| Requirement | Implementation | Evidence in this release |
|---|---|---|
| One local executable, no separately installed runtime | Go native binary; same image runs worker/session | Real PE/Mach-O/ELF binaries built. Linux executable exercised. Windows/Mac binaries not run. |
| No separate configured desktop/VM | Windows private Win32 desktop and owned Word | Source and cross-build only; actual coexistence/cleanup not run. |
| Editable VBA including new/deleted components | Direct MS-CFB/MS-OVBA writer and cache invalidation | Actual new projects, source edits and ALR round trips; not Word compilation. |
| Persistent native forms | Binary MS-OFORMS common controls, Frame, MultiPage, Page, metadata | New forms plus read/edit checks on all four ALR forms; not instantiated in Word here. |
| Arbitrary forms/custom ActiveX | Preserve unknown bytes; explicit unsupported-property errors; native object model remains available | NOT a universal binary designer for every third-party control. Registered third-party dependencies are not bundled. |
| Ribbon | Preserve/edit native XML, relationships and callback lexical checks | Native parts built; actual Ribbon load test supplied but not executed. No complete Ribbon XSD/idMso/type checker. |
| Native content/styles/saved parts | Typed recipes plus full raw XML access | Package/structure tests; no native render performed. Recipe coverage is not all OOXML. |
| DOCX/picture to template | DOCX style/property observations, image delivery and pixel comparison | Model must infer design; no bundled model, OCR service or promise of exact font identification from pixels. |
| Arbitrary native Word behavior | Raw Automation access and scratch native VBA | Real implementation, not an emulator; native behavior unverified. |
| Native UI testing | Owned-window accessibility, actions, captures; async macro/UI lanes | Unexecuted. Not universal support for every owner-drawn control, drag gesture or file picker. No hidden global SendInput fallback. |
| Native page rendering | Word Page.EnhMetaFileBits and Windows GDI | Compiled; not run. Not replaced with a web/LibreOffice renderer. |
| Full VBA compile | Targeted native VBE Compile command, observed enabled-to-disabled transition | Not run. Policy/VBE restrictions can fail. Already-disabled command does not prove a fresh compile. |
| Fast feedback | Native binary edits, artifact/source-hash cache, warm process, batches, local IPC | Offline CLI measurements only. No native Word latency data. |
| Safe deployment | Exact artifact and suite binding, backup, old-target guard, deferred local activation | File/policy tests on Linux use labeled synthetic proof. Actual Windows install/locks not run. |
| Next Word launch after OS reboot | Durable pending activation record | NO automatic login/startup resume in this preview. Not a fulfilled no-intervention reboot guarantee. |
| Mac portable authoring | Native offline binary, conditional checks and runtime probe VBA | Checks tested; not proof of portable behavior. Some Mac compile constants are conservatively unknown without the runtime probe. |
| Actual Mac execution | Direct Apple Events and installed SDEF dictionary | Cross-compiled prototype; not run. No native Mac UI/render/compile parity. |

## Security and environment

Native VBA can do what the current user can do, including files, network access and external programs. A private desktop separates visible UI; it does NOT isolate malicious macros, OS globals, clipboard, printer settings or external COM applications launched by a macro. The caller authorizes execution explicitly. The app does not attach to or intentionally close the user's existing Word process. Raw native APIs remain powerful and must be used with the user's permission.

Microsoft Word must be installed, activated and able to run normally. Word is not an open-source runtime bundled with this program. Office Group Policy, enterprise application controls, downloaded-file restrictions and OS permissions remain effective; the app does not bypass them or silently lower the user's Trust Center settings. macOS Automation prompts may require actual human consent. Executables are not Authenticode-signed or notarized in this delivery.

## Binary preservation versus interpretation

Unedited data stays untouched where possible. Supported form edits preserve unmodeled binary tails, fonts and pictures; unsupported form properties/controls are not translated to approximate substitutes. Codepage handling is explicit; Windows can use installed Windows codepages, while the standalone cross-platform core directly supports its implemented encodings. Unrepresentable or unsupported encodings error rather than silently replacing characters.

The writer implements published interoperable source streams and invalidates old execution caches. It is not a standalone VBA compiler. Compression/CFB fuzzing and reparse checks are useful but not equivalent to an independent Office conformance certification. Bounded input limits protect memory: generally 256 MB source/package budgets, 64 million image pixels, bounded IPC messages and response trees. The Windows native ABI builds are 64-bit (x64 and ARM64); no 32-bit host executable is supplied.

Source assertions, deterministic byte round trips, native runtime assertions, compiler observations, layout checks, and cross-platform verification are distinct. Reports preserve that distinction. Synthetic hosts used in harness tests are explicitly labeled and never counted as actual Word evidence. Passing a native test only establishes the tested scenario on the recorded Office build.
