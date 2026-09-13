param([string]$Go='go')
$ErrorActionPreference='Stop'
[Diagnostics.Process]::GetCurrentProcess().PriorityClass='BelowNormal'
$root=[IO.Path]::GetFullPath((Join-Path $PSScriptRoot '../..'))
$previous=$env:WORDUP_OFFICE_TOOLS_TEST
Push-Location $root
try {
    & $Go test -p 2 -c -o bin/office-tools.test.exe ./internal/native
    if($LASTEXITCODE -ne 0){throw 'Integration test compilation failed'}
    $env:WORDUP_OFFICE_TOOLS_TEST='1'
    & ./bin/office-tools.test.exe '-test.run' '^TestOfficeToolsIntegration$' '-test.v' '-test.timeout' '30s'
    if($LASTEXITCODE -ne 0){throw 'Office helper integration test failed'}
} finally {
    $env:WORDUP_OFFICE_TOOLS_TEST=$previous
    Pop-Location
}
