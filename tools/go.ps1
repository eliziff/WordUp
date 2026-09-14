$ErrorActionPreference = 'Stop'

$online = $args.Count -gt 0 -and $args[0] -eq '-Online'
if ($online) { $args = @($args | Select-Object -Skip 1) }

$repo = Split-Path -Parent $PSScriptRoot
$tools = Join-Path (Split-Path -Parent $repo) '.tools'
$go = Join-Path $tools 'go\bin\go.exe'
if (-not (Test-Path -LiteralPath $go -PathType Leaf)) {
    throw "Pinned Go toolchain not found: $go"
}

$env:GOCACHE = Join-Path $tools 'gocache'
$env:GOMODCACHE = Join-Path $tools 'gomodcache'
$env:GOTMPDIR = Join-Path $tools 'gotmp'
$env:GOFLAGS = (($env:GOFLAGS -replace '(^|\s)-buildvcs=\S+', '') + ' -buildvcs=false').Trim()
$env:GOTOOLCHAIN = 'local'
if (-not $online) { $env:GOPROXY = 'off' }
New-Item -ItemType Directory -Force $env:GOCACHE, $env:GOMODCACHE, $env:GOTMPDIR | Out-Null

$processPath = [Environment]::GetEnvironmentVariable('Path', 'Process')
[Environment]::SetEnvironmentVariable('PATH', $null, 'Process')
[Environment]::SetEnvironmentVariable('Path', $processPath, 'Process')
$process = Start-Process -FilePath $go -ArgumentList $args -WorkingDirectory $repo -NoNewWindow -PassThru
$process.PriorityClass = 'BelowNormal'
$process.WaitForExit()
exit $process.ExitCode
