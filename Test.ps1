param([switch]$Integration)
$ErrorActionPreference = 'Stop'
$compiler = Join-Path $env:WINDIR 'Microsoft.NET\Framework64\v4.0.30319\csc.exe'
$outputDirectory = Join-Path $PSScriptRoot 'test-output'
New-Item -ItemType Directory -Force -Path $outputDirectory | Out-Null
$sources = @(Get-ChildItem -LiteralPath (Join-Path $PSScriptRoot 'src') -Filter '*.cs' | ForEach-Object FullName)
$sources += Join-Path $PSScriptRoot 'tests\Tests.cs'
& $compiler /nologo /target:exe /main:Tests /reference:System.Windows.Forms.dll /reference:System.Drawing.dll /reference:Microsoft.CSharp.dll /reference:System.Xml.Linq.dll "/out:$outputDirectory\Tests.exe" @sources
if ($LASTEXITCODE -ne 0) { throw 'Test build failed.' }
$testArguments = @()
if ($Integration) { $testArguments += '--integration' }
& "$outputDirectory\Tests.exe" @testArguments
if ($LASTEXITCODE -ne 0) { throw 'Tests failed.' }
