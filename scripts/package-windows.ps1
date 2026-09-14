param(
 [ValidateSet('windows/amd64','windows/arm64')]
 [string]$Platform = 'windows/amd64',
 [string]$Executable = '',
 [string]$Destination = ''
)
$ErrorActionPreference = 'Stop'
[Diagnostics.Process]::GetCurrentProcess().PriorityClass = 'BelowNormal'
$root = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
$architecture = if ($Platform -eq 'windows/arm64') { 'arm64' } else { 'x64' }
if ([string]::IsNullOrWhiteSpace($Executable)) {
 $Executable = Join-Path $root ('dist/wordup-windows-' + $architecture + '.exe')
}
if ([string]::IsNullOrWhiteSpace($Destination)) {
 $Destination = Join-Path $root ('dist/windows-' + $architecture)
}
$destinationRoot = [IO.Path]::GetFullPath($Destination)
$executablePath = (Resolve-Path -LiteralPath $Executable).Path
$executableBytes = [IO.File]::ReadAllBytes($executablePath)
if ($executableBytes.Length -lt 64 -or $executableBytes[0] -ne 0x4d -or $executableBytes[1] -ne 0x5a) {
 throw "Executable is not a Windows PE file: $executablePath"
}
$peOffset = [BitConverter]::ToInt32($executableBytes, 0x3c)
if ($peOffset -lt 0 -or $peOffset + 6 -gt $executableBytes.Length -or $executableBytes[$peOffset] -ne 0x50 -or $executableBytes[$peOffset + 1] -ne 0x45) {
 throw "Executable has an invalid PE header: $executablePath"
}
$machine = [BitConverter]::ToUInt16($executableBytes, $peOffset + 4)
$expectedMachine = if ($Platform -eq 'windows/arm64') { 0xaa64 } else { 0x8664 }
if ($machine -ne $expectedMachine) {
 throw ("Executable machine 0x{0:x4} does not match {1}: {2}" -f $machine, $Platform, $executablePath)
}
$required = @('office-tools/WordUp.OfficeTools.exe','office-tools/OPENXML-LICENSE','wordup-oletools/wordup-oletools.exe','wordup-oletools/DEPENDENCIES.txt','wordup-oletools/licenses')
foreach ($relative in $required) {
 if (!(Test-Path -LiteralPath (Join-Path $root ('bin/' + $relative)))) { throw "Missing $relative; build tools/office-bridge/build.ps1 and tools/oletools/build.ps1 first." }
}
$analysisSource = Join-Path $root 'tools/rubberduck/bin/analysis/Release/net462'
if (!(Test-Path -LiteralPath (Join-Path $analysisSource 'WordUp.Analysis.exe'))) { throw 'Missing analysis/WordUp.Analysis.exe; build tools/rubberduck/build.ps1 first.' }
$runtimeNotices = @('internal/vbaparse/ANTLR-LICENSE','internal/vbaparse/X-EXP-LICENSE','internal/structure/LICENSE','internal/structure/PDF-LICENSE','internal/office/XPATH-NOTICES','docs/upstream/rubberduck/LICENSE')
foreach ($relative in $runtimeNotices) {
 if (!(Test-Path -LiteralPath (Join-Path $root $relative))) { throw "Missing compiled dependency notice: $relative" }
}
$pythonNotices = Join-Path $root 'bin/wordup-oletools/licenses'
$noticeInventory = Get-Content -LiteralPath (Join-Path $pythonNotices 'inventory.json') -Raw | ConvertFrom-Json
if (!$noticeInventory.distributions.Count) { throw 'Python bundle license inventory is empty.' }
foreach ($dependency in $noticeInventory.distributions) {
 if ($dependency.license_files_missing -or !$dependency.license_files.Count) { throw "Missing dependency notice: $($dependency.name)" }
 foreach ($notice in $dependency.license_files) {
  $noticePath = Join-Path $pythonNotices $notice.file
  if ((Get-FileHash -LiteralPath $noticePath -Algorithm SHA256).Hash -ne $notice.sha256) { throw "Dependency notice hash mismatch: $($dependency.name)/$($notice.file)" }
 }
}
if (Test-Path -LiteralPath $destinationRoot) { throw "Destination already exists: $destinationRoot. Choose a new staging directory." }
New-Item -ItemType Directory -Path $destinationRoot | Out-Null
Copy-Item -LiteralPath $executablePath -Destination (Join-Path $destinationRoot 'wordup.exe')
foreach ($folder in @('office-tools','wordup-oletools')) {
 Copy-Item -LiteralPath (Join-Path $root ('bin/' + $folder)) -Destination $destinationRoot -Recurse
}
Copy-Item -LiteralPath $analysisSource -Destination (Join-Path $destinationRoot 'analysis') -Recurse
foreach ($file in @('LICENSE','LICENSE-MIT-ORIGINAL','THIRD-PARTY-NOTICES.md')) {
 Copy-Item -LiteralPath (Join-Path $root $file) -Destination $destinationRoot
}
$licenseDirectory = Join-Path $destinationRoot 'licenses'
New-Item -ItemType Directory -Path $licenseDirectory | Out-Null
foreach ($relative in $runtimeNotices) {
 Copy-Item -LiteralPath (Join-Path $root $relative) -Destination $licenseDirectory
}
@"
WordUp Windows $architecture development distribution

Keep wordup.exe, office-tools and wordup-oletools together. Native template execution requires installed Microsoft Word. Go and Python are not required on the end user's computer. OfficeTools uses Windows .NET Framework 4.8. VBA signing additionally requires Windows SDK SignTool and Office SIP.

The main executable targets Windows $architecture ($Platform). The bundled oletools helper is a
portable x64 Windows payload and relies on Windows x64 emulation on ARM64.

Run wordup.exe help and wordup.exe doctor first. Native execution requires --execute. This package is not a claim that every template or workflow has passed verification.

Source, documentation and dependency notices: https://github.com/eliziff/WordUp
The source corresponding to a published binary must accompany its release; do not publish this development stage independently of that source.
"@ | Set-Content -Encoding UTF8 -LiteralPath (Join-Path $destinationRoot 'README.txt')
$files = @(Get-ChildItem -LiteralPath $destinationRoot -File -Recurse | Sort-Object FullName | ForEach-Object {
 [ordered]@{path=$_.FullName.Substring($destinationRoot.Length + 1).Replace('\','/'); bytes=$_.Length; sha256=(Get-FileHash -LiteralPath $_.FullName -Algorithm SHA256).Hash.ToLowerInvariant()}
})
[ordered]@{schema=1; platform=$Platform; files=$files} | ConvertTo-Json -Depth 5 | Set-Content -Encoding UTF8 -LiteralPath (Join-Path $destinationRoot 'manifest.json')
Write-Output $destinationRoot
