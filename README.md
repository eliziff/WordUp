# WordUp 0.4.0

WordUp is a source-first toolkit for building, testing, and deploying Microsoft Word templates with coding agents.

It turns a `.docm` or `.dotm` file into an editable workspace containing its VBA, UserForms, Ribbon configuration, document content, styles, assets, and tests. An agent can change those sources with ordinary development tools, rebuild the native Word file, and verify the result in Microsoft Word.

WordUp is designed for work such as:

- developing and maintaining macro-enabled Word templates;
- editing VBA modules, classes, document code, and UserForms;
- creating Ribbon controls, context menus, keyboard shortcuts, and Quick Parts;
- changing document content, styles, numbering, and package assets;
- compiling VBA and exercising macros in an owned Word session;
- rendering real Word pages for visual review;
- recording repeatable acceptance checks; and
- installing a verified template with backup and restore support.

WordUp is currently a Windows engineering preview. Its Windows workflow has been tested in Microsoft Word, including VBA compilation, UserForms, Ribbon callbacks, hotkeys, document editing, Quick Parts, signing, rendering, failure containment, and deployment recovery. See [Windows validation](docs/VALIDATION-WINDOWS.md) and [current limits](docs/LIMITS.md).

## How it works

A typical WordUp project follows this loop:

1. Import an existing Word document or template into a source workspace.
2. Edit its VBA, forms, package content, assets, and acceptance tests.
3. Build a new native `.docm` or `.dotm` artifact.
4. Compile, run, inspect, and render it in Microsoft Word.
5. Test the finished artifact against observable expectations.
6. Deploy the verified result with a recoverable backup.

WordUp preserves the imported package as a hash-checked baseline, so edits happen in source files and builds produce new artifacts rather than modifying the original in place.

## Quick start

Download and unpack the Windows build, then confirm that WordUp can create and test its included example:

```powershell
.\wordup.exe --execute selftest
```

Import a real template and start a reusable local session:

```powershell
.\wordup.exe import "C:\Templates\MyTemplate.dotm" "C:\Work\MyTemplate"
.\wordup.exe -w "C:\Work\MyTemplate" --execute session start
.\wordup.exe -w "C:\Work\MyTemplate" rpc build '{}'
```

For parameters that are awkward to quote in a shell, place the JSON in a UTF-8 file:

```powershell
.\wordup.exe -w "C:\Work\MyTemplate" rpc native.call @operation.json
```

Run `doctor` to diagnose the local environment and `rpc help '{}'` to browse available operations. Microsoft Word must be installed for native compilation, execution, and rendering.

## Working with coding agents

Give an agent the `wordup.exe` path, this README, and the document or template task. Imported workspaces include agent instructions, and the [template-builder skill](docs/skills/wordup-template-builder/SKILL.md) provides guidance for styles, conversions, UI work, and native verification.

Agents can drive WordUp through its command-line interface or through stdio MCP:

```text
wordup.exe --workspace C:\Work\MyTemplate --execute mcp
```

The [diagnostics guide](docs/DIAGNOSTICS.md) covers live XML inspection, VBA evaluation, runtime errors, and useful evidence to collect when a document behaves unexpectedly.

## Project structure

An imported workspace keeps the editable parts of a Word project in familiar files:

```text
project.json                   Project identity and build choices
vba/*.bas                      Standard VBA modules
vba/*.cls                      Class and document modules
vba/*.vba                      UserForm event code
forms/<name>.json              UserForm designs
package/                       Word package parts, RibbonX, and embedded content
assets/                        Images and other editable inputs
references/                    Reference material for agents
tests/suite.json               Native acceptance assertions
reports/                       Build and test results
.wordwright/                   Imported baseline and source hashes
```

The source tree represents the complete document project. VBA, form definitions, package XML, and assets can be reviewed and versioned together.

## Build, test, and deploy

Build the workspace through the local session:

```powershell
.\wordup.exe -w "C:\Work\MyTemplate" rpc build '{}'
```

Run an acceptance suite against the exact staged artifact:

```powershell
.\wordup.exe -w "C:\Work\MyTemplate" --execute test `
  "C:\Work\MyTemplate\dist\MyTemplate.dotm" `
  "C:\Work\MyTemplate\tests\suite.json"
```

WordUp records the tested artifact hash, assertions, host metadata, timing, and results. Deployment requires a fresh passing record for that artifact, retains a backup, and supports restoring the previous template.

The `compile` workflow builds the project, compiles it in Word, and signs the output with a local WordUp certificate. The [API reference](docs/API.md) documents build, native automation, rendering, testing, signing, and deployment operations.

## Learn more

Contributor guidance lives in [AGENTS.md](AGENTS.md). Additional references include:

- [API and operations](docs/API.md)
- [Known limits](docs/LIMITS.md)
- [Windows validation](docs/VALIDATION-WINDOWS.md)
- [Word behavior notes](docs/WORD-OPERATIONS.md)
- [Performance measurements](docs/PERFORMANCE.md)
- [Infrastructure reuse and licensing](docs/INFRASTRUCTURE-REUSE.md)
