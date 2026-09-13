# Headless Rubberduck integration

Pinned source: https://github.com/rubberduck-vba/Rubberduck at
`fae50adab188126a5e7d2a1cefc3328cc18af482` (GPL-3.0-or-later).
The local source checkout lives under `workspaces/reuse/rubberduck-engine`.
Build checks reject a different revision or tracked source changes.

```powershell
./tools/rubberduck/build.ps1 -Proof
```

Contributor prerequisites: Visual Studio 2022 MSBuild and .NET SDK 8.0.425.
Paths are parameters; the SDK may be installed locally. The build restores
upstream dependencies from NuGet and compiles the production Parsing,
VBEditor, Resources, SettingsProvider and InternalApi assemblies. It leaves
upstream tracked source untouched. `Headless.targets` skips add-in deployment
and repository analyzers and uses Roslyn for satellite assemblies.

The semantic proof links the upstream in-memory editor test fixtures and runs
the real preprocessing/declaration/reference pipeline. It is not native Word
execution, an installed add-in, or a production agent endpoint. Test-only
Moq/NUnit dependencies must not be mistaken for shipped application dependencies.

Remaining integration: production workspace/reference adapters, source/attribute
coordinate mapping, installed type-library loading, inspection selection,
incremental invalidation, bounded protocol results and packaging/license audit.
Do not report deep `check` or full Rubberduck integration from this proof alone.

Local SDK bootstrap used Microsoft's 8.0.425 win-x64 archive, verified against
the release metadata SHA-512:
`f0b6f15bf6f1a0507205c0cb102ab99e1dee875c4682c8ed94665be1d580186a06b21455e83b3a01a0ff7f4cd887b67420f2e2fe09ed985534a4cea488ae1af9`.
The SDK is a contributor tool, not a template runtime requirement.
