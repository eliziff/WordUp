param(
 [string]$Source = (Join-Path $PSScriptRoot '../../workspaces/reuse/rubberduck-engine'),
 [string]$MSBuild = 'C:\Program Files (x86)\Microsoft Visual Studio\2022\BuildTools\MSBuild\Current\Bin\MSBuild.exe',
 [string]$DotnetRoot = (Join-Path $PSScriptRoot '../../../.tools/dotnet-8.0.425'),
 [string]$Packages = (Join-Path $PSScriptRoot '../../../.tools/nuget'),
 [switch]$Proof
)
$ErrorActionPreference = 'Stop'
[Diagnostics.Process]::GetCurrentProcess().PriorityClass = 'BelowNormal'
$sourceRoot = (Resolve-Path -LiteralPath $Source).Path
$revision = & git -C $sourceRoot rev-parse HEAD
if ($LASTEXITCODE -ne 0 -or $revision -ne 'fae50adab188126a5e7d2a1cefc3328cc18af482') {
 throw 'Expected pinned Rubberduck revision fae50adab188126a5e7d2a1cefc3328cc18af482.'
}
$changes = & git -C $sourceRoot status --porcelain --untracked-files=no
if ($LASTEXITCODE -ne 0 -or $changes) { throw 'Upstream tracked source is modified; retain adaptations outside the checkout.' }
$env:DOTNET_ROOT = [IO.Path]::GetFullPath($DotnetRoot)
$env:DOTNET_MSBUILD_SDK_RESOLVER_CLI_DIR = $env:DOTNET_ROOT
$env:DOTNET_CLI_TELEMETRY_OPTOUT = '1'
$env:DOTNET_SKIP_FIRST_TIME_EXPERIENCE = '1'
$env:MSBuildEnableWorkloadResolver = 'false'
$env:MSBuildSDKsPath = Join-Path $env:DOTNET_ROOT 'sdk/8.0.425/Sdks'
if (!(Test-Path -LiteralPath $env:MSBuildSDKsPath)) { throw 'Install the pinned .NET SDK 8.0.425 or supply DotnetRoot.' }
$overrides = Join-Path $PSScriptRoot 'Headless.targets'
& $MSBuild (Join-Path $sourceRoot 'Rubberduck.Parsing/Rubberduck.Parsing.csproj') /restore /t:Build /p:Configuration=Release /p:RestoreSources=https://api.nuget.org/v3/index.json "/p:SolutionDir=$sourceRoot\" "/p:RestorePackagesPath=$([IO.Path]::GetFullPath($Packages))" "/p:CustomAfterMicrosoftCommonTargets=$overrides" /nologo /verbosity:minimal /warnasmessage:MSB4011
if ($LASTEXITCODE -ne 0) { throw 'Upstream Rubberduck parser build failed.' }
if ($Proof) {
 & $MSBuild (Join-Path $PSScriptRoot 'SemanticProof.csproj') /restore /t:Build /p:Configuration=Release /p:RestoreSources=https://api.nuget.org/v3/index.json "/p:SolutionDir=$sourceRoot\" "/p:RubberduckSource=$sourceRoot" "/p:RestorePackagesPath=$([IO.Path]::GetFullPath($Packages))" "/p:CustomAfterMicrosoftCommonTargets=$overrides" /nologo /verbosity:minimal /warnasmessage:MSB4011
 if ($LASTEXITCODE -ne 0) { throw 'Semantic proof build failed.' }
 & (Join-Path $PSScriptRoot 'bin/Release/net462/SemanticProof.exe')
 if ($LASTEXITCODE -ne 0) { throw 'Semantic proof failed.' }
}
