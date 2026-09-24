<#
    build.ps1 — gradi RomsCatiaUI.exe pomocu .NET Framework vbc.exe (bez Visual Studija).
    Pokretanje:  cd src\02_vb ;  .\build.ps1
#>
$ErrorActionPreference = "Stop"
$root = $PSScriptRoot

$vbc = Join-Path $env:WINDIR "Microsoft.NET\Framework64\v4.0.30319\vbc.exe"
if (-not (Test-Path $vbc)) {
    $vbc = Join-Path $env:WINDIR "Microsoft.NET\Framework\v4.0.30319\vbc.exe"
}
if (-not (Test-Path $vbc)) { throw "vbc.exe (.NET Framework 4.x) nije pronaden." }

$outDir = Join-Path $root "build"
New-Item -ItemType Directory -Force -Path $outDir | Out-Null
$outExe = Join-Path $outDir "RomsCatiaUI.exe"

$sources = Get-ChildItem -Path (Join-Path $root "src") -Recurse -Filter *.vb | ForEach-Object { $_.FullName }
if (-not $sources) { throw "Nema .vb izvornih datoteka u src\." }

$refs = @(
    "System.dll",
    "System.Core.dll",
    "System.Drawing.dll",
    "System.Windows.Forms.dll",
    "System.Configuration.dll",
    "Microsoft.VisualBasic.dll"
)

$vbcArgs = @(
    "/nologo",
    "/target:winexe",
    "/optionexplicit+",
    "/optionstrict-",
    "/rootnamespace:RomsCatiaUI",
    "/out:$outExe",
    "/imports:Microsoft.VisualBasic,System,System.Collections.Generic,System.Windows.Forms,System.Drawing"
)
$vbcArgs += ($refs | ForEach-Object { "/r:$_" })
$vbcArgs += $sources

Write-Host "Gradim $outExe ..." -ForegroundColor Cyan
& $vbc @vbcArgs
if ($LASTEXITCODE -ne 0) { throw "vbc build nije uspio (exit $LASTEXITCODE)." }

# Kopiraj App.config -> <exe>.config
$appConfig = Join-Path $root "App.config"
if (Test-Path $appConfig) {
    Copy-Item $appConfig (Join-Path $outDir "RomsCatiaUI.exe.config") -Force
}

Write-Host "Gotovo: $outExe" -ForegroundColor Green
