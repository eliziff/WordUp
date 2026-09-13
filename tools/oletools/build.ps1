$ErrorActionPreference='Stop'
[Diagnostics.Process]::GetCurrentProcess().PriorityClass='BelowNormal'
$python=Join-Path $PSScriptRoot '.venv/Scripts/python.exe'
if(!(Test-Path $python)){ & python -m venv (Join-Path $PSScriptRoot '.venv') }
& $python -m pip install --disable-pip-version-check -r (Join-Path $PSScriptRoot 'requirements.lock')
if($LASTEXITCODE -ne 0){throw 'oletools dependency installation failed'}
$root=[IO.Path]::GetFullPath((Join-Path $PSScriptRoot '../..'))
& $python -m PyInstaller --noconfirm --noupx --onedir --name wordup-oletools --distpath (Join-Path $root 'bin') --workpath (Join-Path $root 'build/oletools') --specpath (Join-Path $root 'build') --copy-metadata oletools --copy-metadata olefile (Join-Path $PSScriptRoot 'inspect_binary.py')
if($LASTEXITCODE -ne 0){throw 'oletools bundle build failed'}
& $python -m pip freeze | Out-File -Encoding utf8 (Join-Path $root 'bin/wordup-oletools/DEPENDENCIES.txt')

& $python (Join-Path $PSScriptRoot 'collect_licenses.py') (Join-Path $root 'bin/wordup-oletools/licenses')
if($LASTEXITCODE -ne 0){throw 'oletools license inventory failed'}
Copy-Item -LiteralPath (Join-Path $PSScriptRoot 'requirements.lock') -Destination (Join-Path $root 'bin/wordup-oletools/requirements.lock')