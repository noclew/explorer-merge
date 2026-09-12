$ErrorActionPreference = 'Stop'
$compiler = Join-Path $env:WINDIR 'Microsoft.NET\Framework64\v4.0.30319\csc.exe'
if (-not (Test-Path -LiteralPath $compiler)) { throw 'The Windows .NET Framework C# compiler was not found.' }
$outputDirectory = Join-Path $PSScriptRoot 'dist'
New-Item -ItemType Directory -Force -Path $outputDirectory | Out-Null
$sources = @(Get-ChildItem -LiteralPath (Join-Path $PSScriptRoot 'src') -Filter '*.cs' | ForEach-Object FullName)
& $compiler /nologo /target:winexe /platform:anycpu /optimize+ /reference:System.Windows.Forms.dll /reference:System.Drawing.dll /reference:Microsoft.CSharp.dll /reference:System.Xml.Linq.dll "/win32icon:$PSScriptRoot\assets\ExplorerMerge-character.ico" "/out:$outputDirectory\ExplorerMerge.exe" @sources
if ($LASTEXITCODE -ne 0) { throw 'Build failed.' }
Write-Output "Built: $outputDirectory\ExplorerMerge.exe"
