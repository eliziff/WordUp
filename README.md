# Wordwright Local 0.2.0

A native, source-first Word template toolchain for coding agents. One executable; no Python, Node, browser runtime, cloud service, VM, module-import ritual, or separately configured test desktop.

**Engineering preview.** The source/package/form core has been executed and tested on Linux against the supplied ALR template. Windows and Mac executables are real cross-compiled binaries. **The Windows Word backend and Mac Apple Events backend have NOT been executed against Microsoft Word in this build environment.** See `VALIDATION.md` and `docs/LIMITS.md`. Compiling the executable is not native Word acceptance. No edited sample is labeled Word-verified.

## Start with an agent

Unpack the matching binary and give a local coding agent its path, this README, and the template or document you want changed. The agent can use an ordinary terminal; MCP is optional. Microsoft Word must already be installed and usable for actual VBA execution and Word rendering. The harness does not bundle Microsoft Office or change its activation, enterprise policy, or the user's Trust Center registry settings.

A useful instruction to the agent is in `AGENTS.md`. The app's `selftest` creates its own integrated example and tests it—no user-imported modules or hand-created fixtures required:

```powershell
.\wordwright.exe --execute selftest
```

That command deliberately returns a nonzero exit code and structured evidence on failure. It is the first native gate, not a proof of every future macro. The agent should run it and investigate the resulting logs itself rather than asking the user to do VBE work.

For a real project:

```powershell
.\wordwright.exe import "C:\Templates\MyTemplate.dotm" "C:\Work\MyTemplate"
.\wordwright.exe -w "C:\Work\MyTemplate" --execute session start
.\wordwright.exe -w "C:\Work\MyTemplate" rpc help '{}'
.\wordwright.exe -w "C:\Work\MyTemplate" rpc build '{}'
```

`session start` creates a local app process from this same executable. The native Word host is lazy-started and stays warm. Windows uses a user-restricted named pipe; Mac/Linux use a local socket. No TCP listener, third-party service, or system service. `session stop` closes the owned worker, not the user's Word. An idle session exits after ten minutes.

For reliable shell quoting, pass operation parameters as UTF-8 JSON files:

```powershell
.\wordwright.exe -w "C:\Work\MyTemplate" rpc native.call @operation.json
```

An MCP client can instead launch:

```text
wordwright.exe --workspace C:\Work\MyTemplate --execute mcp
```

There is no global agent configuration installer. Clients that support stdio MCP can use that command; any terminal-capable coding agent can use the CLI and local session. Protocol revisions 2025-03-26, 2025-06-18 and 2025-11-25 are explicitly negotiated; newer capabilities are not silently claimed.

## Source format

```text
project.json                   Project metadata, component kinds, references
.wordwright/base.opc           Immutable, hash-checked original package
.wordwright/index.json         Baseline source hashes
vba/*.bas                      Standard VBA modules
vba/*.cls                      Classes and the document module
vba/*.vba                      UserForm event code
forms/<name>.json              Persistent native MSForms design, in points
package/                       Full native OPC parts, including RibbonX and assets
styles/recipe.json             Named style and numbering changes
content/recipe.json            Native document-body composition
building_blocks/recipe.json    Reusable native saved parts
assets/                        Images used by recipes
references/                    Agent reference material
reports/                       Machine-readable build and acceptance results
tests/suite.json               Native observable assertions
```

The original file is not used as a scratchpad. The workspace preserves untouched package parts. A no-change ALR import/build was byte-identical. Source changes rebuild the native CFB/VBA project and compressed streams; they do not automate typing into VBE. Ordinary module/form source builds therefore do not require enabling VBIDE trust access.

**`content/recipe.json` replaces the document body.** Do not introduce it when only styles or macros should change. For existing complex documents, edit their original XML surgically or use native Word operations on a test copy. A convenience recipe is not a full replacement for Word's object model.

## Native Windows execution design

The same executable creates a private Win32 desktop on the user's own interactive window station, then launches an owned Word process there. It never switches the user's visible desktop. It binds through the owned process's `_WwG` window and `AccessibleObjectFromWindow`, verifies the resulting Word HWND/PID, and uses raw COM `IDispatch`, `VARIANT`, `SAFEARRAY`, and BSTR calls. There is no Python COM bridge, VSTO runtime, injected DLL, or emulator.

The `/a` startup switch avoids automatically loading the user's normal template and add-ins. An owned job object provides process-tree cleanup. Default file inspection disables macros. Authorized execution uses the owned process's documented AutomationSecurity property; no Trust Center or execution-policy registry changes are made. Group Policy and OS security remain authoritative.

This removes the previous separate-machine/VM requirement from the implementation. **Coexistence with a live user Word session, modal-dialog handling, compiler selection, security behavior, and cleanup still need actual Windows acceptance.** The source includes the mechanisms; the Linux test results do not prove them.

## Broad capability without a fixed operation menu

Native calls can access any member exposed by Word's object model through `get`, `invoke`, `put`, and `putref`, with named/optional parameters and retained object references. A scratch `eval` operation builds a separate temporary native VBA project and executes a function body in Word. It returns the generated source and artifact hash with the observed result. This is native VBA, not a Python/JavaScript approximation.

Native UI operations inspect and operate on owned windows through accessibility APIs. Asynchronous `begin`/`poll` lets an agent inspect or interact with a modal form while the Word call remains blocked. Completed task records can be released with `forget`. Unsupported accessibility actions fail; there is no invisible global mouse/keyboard fallback.

The renderer obtains actual Word page `EnhMetaFileBits`, rasterizes it through Windows GDI, and writes PNGs. It does not use a web imitation or LibreOffice to claim Word layout correctness. `reference.image` and `image.compare` support a vision-capable agent's iterative design loop. The harness contains no embedded inference model.

## Tests and installation

`test` uses an exact staged copy of an artifact and records its SHA-256, suite, results, assertion count, host metadata, and timing. A fresh test starts a separate owned Word host. It does not edit the source artifact. Empty assertions and unavailable runtimes cannot pass.

```powershell
.\wordwright.exe -w "C:\Work\MyTemplate" --execute test "C:\Work\MyTemplate\dist\MyTemplate.dotm" "C:\Work\MyTemplate\tests\suite.json"
```

`deploy` requires a passing, fresh, artifact-bound native acceptance record. It verifies the old installed file's hash, makes a backup, and atomically installs only when Word is closed. If Word is open, an app-owned local process waits for it to exit. It does not close the user's Word. **This preview does not register a startup task: if the OS shuts down while activation is pending, the agent must resume the durable plan through the app.** Do not claim guaranteed next-boot activation across an OS reboot.

## Mac stretch implementation

The same offline builder and conservative conditional-compilation/portability checks run in native Apple Silicon and Intel binaries. A native Apple Events backend reads the installed Word scripting dictionary and sends events directly through system frameworks; there is no `osascript` subprocess. It is an unexecuted prototype, not Windows parity: Mac UI capture, page rendering and full-project compile verification are not implemented. Windows test passes are never reclassified as Mac passes. Normal macOS Automation consent and signing/notarization requirements are not bypassed.

## Measurements

See `evidence/benchmarks.json` for twenty whole-process trials per scenario. In this Linux container, the 708 KB ALR template's changed-source build had a median of about 145 ms; cached ALR builds about 58 ms; the integrated example's cached build about 3 ms. These include launching the CLI. They are **not Word startup, execution or Windows benchmarks**. Native Word latency has not been measured.

## Developing the harness

End users do not need Go. Building the source requires Go 1.23+:

```text
go test ./...
go test -race ./...
go build -trimpath -ldflags="-s -w" -o wordwright ./cmd/wordwright
```

There are no external Go modules. `go list -m all` contains only this module. `scripts/build.sh` reproduces the platform binaries. The private ALR specimen is excluded from the generic source package; setting `WORDWRIGHT_SPECIMEN` to its path enables the optional specimen tests.

Read `docs/API.md`, `docs/LIMITS.md`, `VALIDATION.md`, and the source. This preview is not signed or notarized. Do not circumvent OS security warnings or deploy an unverified candidate over a working template.
