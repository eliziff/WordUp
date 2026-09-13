$ErrorActionPreference='Stop'
[Diagnostics.Process]::GetCurrentProcess().PriorityClass='BelowNormal'
$packages=Join-Path $PSScriptRoot 'packages'
$output=[IO.Path]::GetFullPath((Join-Path $PSScriptRoot '../../bin/office-tools'))
New-Item -ItemType Directory -Force $packages,$output | Out-Null
Copy-Item -LiteralPath (Join-Path $PSScriptRoot 'OPENXML-LICENSE') -Destination (Join-Path $output 'OPENXML-LICENSE')
$dependencies=@(@('documentformat.openxml','3.3.0','net46'),@('documentformat.openxml.framework','3.3.0','net46'),@('flaui.core','4.0.0','net48'),@('flaui.uia3','4.0.0','net48'),@('interop.uiautomationclient','10.19041.0','net45'),@('system.management','5.0.0','net45'))
foreach($p in $dependencies){
 $id=$p[0];$version=$p[1];$target=Join-Path $packages $id
 if(!(Test-Path $target)){
  $zip=Join-Path $packages ($id+'.zip')
  Invoke-WebRequest -UseBasicParsing ('https://api.nuget.org/v3-flatcontainer/'+$id+'/'+$version+'/'+$id+'.'+$version+'.nupkg') -OutFile $zip -TimeoutSec 30
  Expand-Archive -LiteralPath $zip -DestinationPath $target
 }
 $lib=Join-Path $target ('lib/'+$p[2])
 if(Test-Path $lib){Get-ChildItem -LiteralPath $lib -Filter '*.dll' | Copy-Item -Destination $output}
 Get-ChildItem -LiteralPath $target -File | Where-Object Name -Match 'license|notice|nuspec' | ForEach-Object {Copy-Item -LiteralPath $_.FullName -Destination (Join-Path $output ($id+'-'+$_.Name))}
}
$refs=@('/r:System.Web.Extensions.dll','/r:System.Xml.dll','/r:System.Core.dll','/r:System.Drawing.dll','/r:System.Management.dll')
$refs+=('/r:'+(Join-Path $env:WINDIR 'Microsoft.NET/Framework64/v4.0.30319/WPF/WindowsBase.dll'))
$refs+=Get-ChildItem -LiteralPath $output -Filter '*.dll' | ForEach-Object {'/r:'+ $_.FullName}
$compiler=Join-Path $env:WINDIR 'Microsoft.NET/Framework64/v4.0.30319/csc.exe'
& $compiler /nologo /target:exe /platform:x64 ('/out:'+(Join-Path $output 'WordUp.OfficeTools.exe')) $refs (Join-Path $PSScriptRoot 'Program.cs')
if($LASTEXITCODE -ne 0){throw 'Office helper compilation failed'}
