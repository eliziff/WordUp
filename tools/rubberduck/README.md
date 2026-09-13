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

`WordUp.Analysis.csproj` builds the workspace adapter separately from test code.
It supplies read-only module snapshots to the real preprocessing, declaration
and reference pipeline. Hidden attributes retain their separate parser view and
diagnostics map back to physical source lines. Unsupported editor operations
throw; no editor or Word instance is created. Standard/class modules are supported;
document/form metadata remains unfinished and is rejected explicitly.

The build also includes upstream CodeAnalysis and its project dependencies.
`Headless.targets` references the pinned SDK's .NET Framework-compatible
`netstandard` facade for the Immutable dependency, without importing SDK task
internals or changing upstream source. Keep this dependency in the analysis
toolchain, not in generated templates.

The proof exercises both this adapter and upstream test-only editor fixtures,
comparing declarations and reference locations for cross-module calls, invalid
members, conditional code and property accessors. Moq/NUnit belong only to the
proof project, not the analysis assembly. This is not native Word execution or
an agent endpoint yet.

`WordUp.Analysis.exe serve` now accepts the existing helper protocol: one JSON
string array `["analyze", requestJson]` per line. Requests contain `Modules`
(`Name`, `Kind`, `Path`, `Source`) and `Inspections`. Only standard/class modules
and the two verified inspections are accepted. Input lines are bounded to 4 Mi
characters; results identify source-only scope, original source hashes and mapped
file locations. No source files are opened or modified by this worker.
The proof starts the real executable and checks errors followed by successful
analysis in the same process. Main application wiring, process-job containment,
retrievable large results and release packaging remain unfinished.

The proof runs the actual upstream unused-variable inspection on a defect,
its correction, and the defect again, asserting the variable, module, line and
resource description. This establishes that inspection path, not all inspections
or incremental cache invalidation.

Installed type-library reflection uses upstream `ComLibraryProvider` (no
registration) and `ComProject`. References require matching GUID, version and
name; unavailable references are retained as errors. Library bytes are hashed.
Fresh resolver declarations prevent stale reference bindings between analyses.
To exercise an installed Word library as well, set `WORDUP_TEST_WORD_TYPELIB`
to its absolute `MSWORD.OLB` path before running the proof. That check resolves
`Word.Document` and `Document.Name` twice and rejects a mismatched library identity.

Remaining integration: workspace manifest/host/designer metadata, transitive
library coverage, inspection selection, incremental invalidation, bounded protocol
results and packaging/license audit. Type reflection inherits upstream's
best-effort handling of individual library entries; it is not native compilation.
Do not report deep `check` or full Rubberduck integration from this proof alone.

Local SDK bootstrap used Microsoft's 8.0.425 win-x64 archive, verified against
the release metadata SHA-512:
`f0b6f15bf6f1a0507205c0cb102ab99e1dee875c4682c8ed94665be1d580186a06b21455e83b3a01a0ff7f4cd887b67420f2e2fe09ed985534a4cea488ae1af9`.
The SDK is a contributor tool, not a template runtime requirement.
