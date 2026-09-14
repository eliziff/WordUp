# WordUp 0.4.0

Development builds now incorporate Rubberduck's GPL-3.0-or-later VBA grammar.
See [infrastructure reuse and licensing](docs/INFRASTRUCTURE-REUSE.md).

Local reference: [upstream Word/VBA and dependency documentation](docs/upstream/README.md),
plus [reproduced Word behavior and documentation gaps](docs/WORD-OPERATIONS.md).

A native, source-first Word template toolchain for coding agents. One executable; no Python, Node, browser runtime, cloud service, VM, module-import ritual, or separately configured test desktop.

**Windows engineering preview, verified in actual Microsoft Word.** Native acceptance covers whole-project VBA compilation, saved nested UserForms, Ribbon and context-menu callbacks, template hotkeys, editable document content, Quick Parts and real page rendering. Signed-template acceptance passed 17 assertions. Failure, runaway containment and login recovery have separate native tests. Mac execution remains unverified. See [current evidence](docs/VALIDATION-WINDOWS.md) and [remaining limits](docs/LIMITS.md); historical `VALIDATION.md` describes the original offline preview.

## Start with an agent

See [agent diagnostics](docs/DIAGNOSTICS.md) for live XML evidence, Immediate-style VBA evaluation, runtime errors, and the boundaries of those checks. Verification runs on an owned private desktop by default; displaying Word requires an explicit user request.

Unpack the matching binary and give a local coding agent its path, this README, and the template or document you want changed. The agent can use an ordinary terminal; MCP is optional. Microsoft Word must already be installed and usable for actual VBA execution and Word rendering. The harness does not bundle Microsoft Office or change its activation, enterprise policy, or the user's Trust Center registry settings.

A useful instruction to the agent is in `AGENTS.md`. The app's `selftest` creates its own integrated example and tests it—no user-imported modules or hand-created fixtures required:

```powershell
.\wordup.exe --execute selftest
```

That command deliberately returns a nonzero exit code and structured evidence on failure. It is the first native gate, not a proof of every future macro. The agent should run it and investigate the resulting logs itself rather than asking the user to do VBE work.

For a real project:

```powershell
.\wordup.exe import "C:\Templates\MyTemplate.dotm" "C:\Work\MyTemplate"
.\wordup.exe -w "C:\Work\MyTemplate" --execute session start
.\wordup.exe -w "C:\Work\MyTemplate" rpc help '{}'
.\wordup.exe -w "C:\Work\MyTemplate" rpc build '{}'
```

`session start` creates a local app process from this same executable. The native Word host is lazy-started and stays warm. Windows uses a user-restricted named pipe; Mac/Linux use a local socket. No TCP listener, third-party service, or system service. `session stop` closes the owned worker, not the user's Word. An idle session exits after ten minutes.

For reliable shell quoting, pass operation parameters as UTF-8 JSON files:

```powershell
.\wordup.exe -w "C:\Work\MyTemplate" rpc native.call @operation.json
```

An MCP client can instead launch:

```text
wordup.exe --workspace C:\Work\MyTemplate --execute mcp
```

There is no global agent configuration installer. Clients that support stdio MCP can use that command; any terminal-capable coding agent can use the CLI and local session. Protocol revisions 2025-03-26, 2025-06-18 and 2025-11-25 are explicitly negotiated; newer capabilities are not silently claimed.

For reference-driven template work, use the [WordUp template-builder agent skill](docs/skills/wordup-template-builder/SKILL.md). It defines the minimum style, conversion, UI and native-evidence workflow, with a reusable acceptance plan.

On Windows, run the pinned toolchain through `tools/go.ps1`. It keeps Go's build cache, module cache, and temporary work inside the repository tool area, disables redundant VCS probing, and keeps normal iterations offline. On a fresh checkout, fetch dependencies explicitly with `tools/go.ps1 -Online mod download`.

## Source format

For legal-document projects, agents can reuse [Legal Structure Parser](https://github.com/eliziff/legal-structure-parser) and [Legal PDF Parser](https://github.com/eliziff/legal-pdf-parser). WordUp's generated agent instructions link to both. They are optional development tools; the executable and delivered templates do not require them. See [reuse guidance](docs/REUSE.md).

```text
project.json                   VBA project name and explicit build choices
.wordwright/base.opc           Immutable, hash-checked original package
.wordwright/index.json         Baseline source hashes
vba/*.bas                      Standard VBA modules
vba/*.cls                      Classes and the document module
vba/*.vba                      UserForm event code
forms/<name>.json              Persistent native MSForms design, in points
package/                       Full native OPC parts, including RibbonX and assets
assets/                        Images and other editable package inputs
references/                    Agent reference material
reports/                       Machine-readable build and acceptance results
tests/suite.json               Native observable assertions
```

The original file is not used as a scratchpad. The workspace preserves untouched package parts. A no-change import/build of a development specimen was byte-identical. Source changes rebuild the native CFB/VBA project and compressed streams; they do not automate typing into VBE. Ordinary module/form source builds therefore do not require enabling VBIDE trust access.

Edit document content, styles, numbering and saved parts in the package XML or through native Word operations on a disposable copy. The source tree records the actual document, not a second partial description of it.

## Native Windows execution design

The same executable creates a private Win32 desktop on the user's own interactive window station, then launches an owned Word process there. It never switches the user's visible desktop. It binds through the owned process's `_WwG` window and `AccessibleObjectFromWindow`, verifies the resulting Word HWND/PID, and uses raw COM `IDispatch`, `VARIANT`, `SAFEARRAY`, and BSTR calls. There is no Python COM bridge, VSTO runtime, injected DLL, or emulator.

The `/a` startup switch avoids automatically loading the user's normal template and add-ins. An owned job object provides process-tree cleanup. Default file inspection disables macros. Authorized execution uses the owned process's documented AutomationSecurity property; no Trust Center or execution-policy registry changes are made. Group Policy and OS security remain authoritative.

Native Windows tests exercised compilation, modal compiler diagnostics, independent Word-process survival and runaway cleanup. The owned job runs Below Normal with default limits of 50% CPU, 2048 MB memory, 16 processes and 30 seconds per operation. These limits apply to the WordUp-owned job, not to a shared template opened normally by its recipient.

## Broad capability without a fixed operation menu

Native calls can access any member exposed by Word's object model through `get`, `invoke`, `put`, and `putref`, with named/optional parameters and retained object references. A scratch `eval` operation builds a separate temporary native VBA project and executes a function body in Word. It returns the generated source and artifact hash with the observed result. This is native VBA, not a Python/JavaScript approximation.

Native UI operations inspect and operate on owned windows through accessibility APIs. Asynchronous `begin`/`poll` lets an agent inspect or interact with a modal form while the Word call remains blocked. Completed task records can be released with `forget`. Unsupported accessibility actions fail; there is no invisible global mouse/keyboard fallback.

The renderer obtains actual Word page `EnhMetaFileBits`, rasterizes it through Windows GDI, and writes PNGs. It does not use a web imitation or LibreOffice to claim Word layout correctness. `reference.image` and `image.compare` support a vision-capable agent's iterative design loop. The harness contains no embedded inference model.

## Tests and installation

`test` uses an exact staged copy of an artifact and records its SHA-256, suite, results, assertion count, host metadata, and timing. A fresh test starts a separate owned Word host. It does not edit the source artifact. Empty assertions and unavailable runtimes cannot pass.

```powershell
.\wordup.exe -w "C:\Work\MyTemplate" --execute test "C:\Work\MyTemplate\dist\MyTemplate.dotm" "C:\Work\MyTemplate\tests\suite.json"
```

`deploy` requires a passing, fresh, artifact-bound native acceptance record. It verifies the old installed file's hash, makes a backup, and atomically installs only when Word is closed. If Word is open, an app-owned local process waits. Windows registers a current-user login command for pending activation and removes it on completion. The login script has been executed in a native integration test; an actual OS reboot was not performed. `call restore` with the activation plan's `path` restores the previous template while refusing newer user edits.

`compile` builds, compiles in Word, and signs the output. It automatically creates a local non-exportable signing certificate or reuses an existing WordUp certificate. No personal certificate identity is built in. Windows SDK SignTool and Microsoft's registered Office SIP are currently prerequisites; automatic prerequisite installation is not implemented. A self-signed signature's integrity and publisher trust are reported separately. Failed compilation retains the candidate, compiler diagnostics and screenshots; it does not replace the working output.

## Mac stretch implementation

The same offline builder and conservative conditional-compilation/portability checks run in native Apple Silicon and Intel binaries. A native Apple Events backend reads the installed Word scripting dictionary and sends events directly through system frameworks; there is no `osascript` subprocess. It is an unexecuted prototype, not Windows parity: Mac UI capture, page rendering and full-project compile verification are not implemented. Windows test passes are never reclassified as Mac passes. Normal macOS Automation consent and signing/notarization requirements are not bypassed.

## Measurements

The complete signed Windows suite recorded 2.62 seconds warm and 4.58 seconds including startup, with 17 fresh assertions and visually checked Word, Ribbon and UserForm captures. Timing varies with machine load; a one-second full-suite target has not been achieved. See [measurement details](docs/PERFORMANCE.md). `evidence/benchmarks.json` retains the original offline Linux measurements.

## Developing the harness

End users do not need Go. Building the source requires Go 1.23+:

```text
go test ./internal/... ./cmd/...
go test -race ./internal/... ./cmd/...
go build -trimpath -ldflags="-s -w" -o wordup ./cmd/wordup
```

The core uses the pinned ANTLR Go runtime and its transitive dependencies in `go.mod`/`go.sum`. `scripts/build.sh` builds the standalone platform executables; it does not package optional Windows helpers. Open XML SDK validation and FlaUI patterns require the `office-tools` bundle built by `tools/office-bridge/build.ps1`; independent binary inspection requires the `wordup-oletools` bundle built by `tools/oletools/build.ps1`. Keep those folders beside the executable in a complete Windows distribution. End users do not need Go or Python to use the packaged builds. Private development specimens are excluded from the generic source package; setting `WORDUP_SPECIMEN` to a specimen path enables the optional specimen tests.

Read `docs/API.md`, `docs/LIMITS.md`, `VALIDATION.md`, and the source. This preview is not signed or notarized. Do not circumvent OS security warnings or deploy an unverified candidate over a working template.
