# -*- mode: powershell -*-
<#
.SYNOPSIS
    Build FeeMan: sync deps -> PyInstaller -> Inno Setup installer.

.DESCRIPTION
    Requires: uv (https://docs.astral.sh/uv/), Inno Setup 6 (optional, for installer).
    Notes are kept in ASCII so the script parses correctly under both
    Windows PowerShell 5.1 and PowerShell 7 regardless of console codepage.

.PARAMETER SkipInstaller
    Only build the exe, skip the installer.

.PARAMETER Clean
    Remove build/ and dist/ before building.

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File build.ps1
    powershell -ExecutionPolicy Bypass -File build.ps1 -Clean
    powershell -ExecutionPolicy Bypass -File build.ps1 -SkipInstaller
#>
[CmdletBinding()]
param(
    [switch]$SkipInstaller,
    [switch]$Clean
)

$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'
$root = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $root

function Write-Step([string]$msg) {
    Write-Host ""
    Write-Host "==> $msg" -ForegroundColor Cyan
}

function Fail([string]$msg) {
    Write-Host "ERROR: $msg" -ForegroundColor Red
    exit 1
}

# ---------------------------------------------------------------- preflight
Write-Step "Checking toolchain"

if (-not (Get-Command uv -ErrorAction SilentlyContinue)) {
    Fail "uv not found. Install: powershell -c ""irm https://astral.sh/uv/install.ps1 | iex"""
}
Write-Host ("uv   : " + ((uv --version) -join " "))

$iscc = $null
foreach ($candidate in @(
        "C:\Program Files (x86)\Inno Setup 6\ISCC.exe",
        "C:\Program Files\Inno Setup 6\ISCC.exe",
        (Join-Path $env:LOCALAPPDATA "Programs\Inno Setup 6\ISCC.exe"))) {
    if (Test-Path $candidate) { $iscc = $candidate; break }
}

if (-not $SkipInstaller -and -not $iscc) {
    Write-Host "WARNING: Inno Setup 6 (ISCC.exe) not found - skipping installer." -ForegroundColor Yellow
    Write-Host "         Install with: winget install JRSoftware.InnoSetup" -ForegroundColor Yellow
    $SkipInstaller = $true
}
if ($iscc) { Write-Host ("iscc : " + $iscc) }

# ---------------------------------------------------------------- clean
if ($Clean) {
    Write-Step "Cleaning build/ and dist/"
    foreach ($d in @("build", "dist")) {
        if (Test-Path $d) { Remove-Item $d -Recurse -Force }
    }
}

# ---------------------------------------------------------------- deps
Write-Step "Syncing dependencies (uv sync)"
uv sync
if ($LASTEXITCODE -ne 0) { Fail "uv sync failed" }

# ---------------------------------------------------------------- self-check
Write-Step "Self-check"
uv run python -c "import fee_core, fee_app, openpyxl, tkinter; print('imports OK')"
if ($LASTEXITCODE -ne 0) { Fail "self-check failed" }

# ---------------------------------------------------------------- package
Write-Step "Packaging with PyInstaller"
uv run pyinstaller fee_app.spec --noconfirm --clean --distpath dist --workpath build
if ($LASTEXITCODE -ne 0) { Fail "PyInstaller failed" }

$exe = Join-Path $root "dist\FeeMan\FeeMan.exe"
if (-not (Test-Path $exe)) { Fail "missing $exe" }
$exeKB = (Get-Item $exe).Length / 1KB
$totalMB = ((Get-ChildItem (Join-Path $root "dist\FeeMan") -Recurse |
             Measure-Object Length -Sum).Sum) / 1MB
Write-Host ("exe  : dist\FeeMan\FeeMan.exe ({0:N0} KB, folder {1:N1} MB)" -f $exeKB, $totalMB)

# ---------------------------------------------------------------- installer
if (-not $SkipInstaller) {
    Write-Step "Building installer with Inno Setup"
    & $iscc "installer.iss"
    if ($LASTEXITCODE -ne 0) { Fail "Inno Setup failed" }

    $setup = Get-ChildItem "dist\FeeMan-Setup-*.exe" -ErrorAction SilentlyContinue |
             Sort-Object LastWriteTime -Descending | Select-Object -First 1
    if (-not $setup) { Fail "installer not produced" }
    Write-Host ("setup: dist\{0} ({1:N1} MB)" -f $setup.Name, ($setup.Length / 1MB))
}

Write-Step "Done"
Write-Host "App       : dist\FeeMan\FeeMan.exe"
if (-not $SkipInstaller) {
    Write-Host "Installer : dist\FeeMan-Setup-1.0.0.exe"
}
