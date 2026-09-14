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

# Windows environment names are case-insensitive, but PowerShell can leave
# both `Path` and `PATH` after a caller assigns one spelling. Start-Process
# rejects that duplicate block. Normalize to one entry before launching Go.
$processPath = [Environment]::GetEnvironmentVariable('Path', 'Process')
Remove-Item -LiteralPath Env:\PATH -Force -ErrorAction SilentlyContinue
Remove-Item -LiteralPath Env:\Path -Force -ErrorAction SilentlyContinue
$env:Path = $processPath
# Start-Process joins an argument array into one Windows command line without
# adding quotes. Quote arguments containing whitespace so repository paths such
# as "reference/Style Guide [Fall].dotm" remain one argument.
$argumentList = @($args | ForEach-Object {
    $value = [string]$_
    if ($value -match '[\s"]') {
        '"' + $value.Replace('"', '\"') + '"'
    } else {
        $value
    }
})
$process = Start-Process -FilePath $go -ArgumentList $argumentList -WorkingDirectory $repo -NoNewWindow -PassThru
$process.PriorityClass = 'BelowNormal'
$process.WaitForExit()
exit $process.ExitCode
