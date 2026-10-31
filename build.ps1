# -*- mode: powershell -*-
<#
.SYNOPSIS
    Build FeeMan: compile the C++ launcher, stage the embedded-runtime layout,
    then optionally build the Inno Setup installer.

.DESCRIPTION
    The installed app does NOT bundle Python. FeeMan.exe checks runtime\python.exe
    and openpyxl on start; if missing it runs bootstrap_runtime.ps1, which downloads
    the Python 3.11.9 embeddable runtime and installs openpyxl. Tcl/Tk and tkinter
    are shipped under runtime\ (the official embed zip does not contain them).

    Requirements:
      * MinGW-w64 (g++ and windres), or a Visual Studio developer prompt (cl and rc)
      * Inno Setup 6 (optional, only for the installer)

.PARAMETER SkipInstaller
    Stage dist\FeeMan only; do not build the setup exe.

.PARAMETER Clean
    Remove build\ and dist\ before building.

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

if ($Clean) {
    Write-Step "Cleaning build\ and dist\"
    foreach ($d in @("build", "dist")) {
        if (Test-Path $d) { Remove-Item $d -Recurse -Force }
    }
}

$buildDir = Join-Path $root "build"
$stage = Join-Path $root "dist\FeeMan"
if (Test-Path $stage) { Remove-Item $stage -Recurse -Force }
New-Item -ItemType Directory -Force -Path $buildDir, $stage | Out-Null

# ---------------------------------------------------------------- launcher
Write-Step "Compiling launcher"

$exe = Join-Path $stage "FeeMan.exe"
$gpp = Get-Command g++ -ErrorAction SilentlyContinue
$windres = Get-Command windres -ErrorAction SilentlyContinue
$cl = Get-Command cl -ErrorAction SilentlyContinue
$rc = Get-Command rc -ErrorAction SilentlyContinue

if ($gpp -and $windres) {
    $res = Join-Path $buildDir "launcher.res.o"
    & $windres -O coff -i "launcher.rc" -o $res
    if ($LASTEXITCODE -ne 0) { Fail "windres failed" }
    & $gpp -std=c++17 -O2 -Wall -Wextra -DUNICODE -D_UNICODE -o $exe `
        "launcher.cpp" $res -mwindows -luser32 -static -static-libgcc -static-libstdc++
    if ($LASTEXITCODE -ne 0) { Fail "g++ failed" }
}
elseif ($cl -and $rc) {
    $res = Join-Path $buildDir "launcher.res"
    & $rc /nologo /fo $res "launcher.rc"
    if ($LASTEXITCODE -ne 0) { Fail "rc failed" }
    $msvcArgs = @(
        "/nologo", "/utf-8", "/std:c++17", "/EHsc", "/O2",
        "/DUNICODE", "/D_UNICODE",
        "/Fo:$buildDir\", "/Fe:$exe",
        "launcher.cpp", $res, "/link", "/SUBSYSTEM:WINDOWS", "user32.lib"
    )
    & $cl @msvcArgs
    if ($LASTEXITCODE -ne 0) { Fail "cl failed" }
}
else {
    Fail "No C++ toolchain found. Install MinGW-w64 (g++/windres) or run from a Visual Studio Developer prompt."
}

# ---------------------------------------------------------------- stage
Write-Step "Staging files into dist\FeeMan"

foreach ($f in @("fee_app.py", "fee_core.py", "bootstrap_runtime.ps1", "bootstrap_runtime.bat", "README.md", "LICENSE")) {
    Copy-Item -LiteralPath $f -Destination $stage -Force
}
Copy-Item -LiteralPath "assets\app.ico" -Destination $stage -Force

$runtimeSrc = Join-Path $root "runtime"
$runtimeDst = Join-Path $stage "runtime"
New-Item -ItemType Directory -Force -Path $runtimeDst | Out-Null
foreach ($f in @("tcl86t.dll", "tk86t.dll", "_tkinter.pyd")) {
    Copy-Item -LiteralPath (Join-Path $runtimeSrc $f) -Destination $runtimeDst -Force
}
Copy-Item -LiteralPath (Join-Path $runtimeSrc "tcl") -Destination (Join-Path $runtimeDst "tcl") -Recurse -Force
Copy-Item -LiteralPath (Join-Path $runtimeSrc "tkinter") -Destination (Join-Path $runtimeDst "tkinter") -Recurse -Force
Get-ChildItem -LiteralPath $runtimeDst -Recurse -Directory -Filter "__pycache__" |
    Remove-Item -Recurse -Force

$required = @(
    "FeeMan.exe",
    "fee_app.py",
    "fee_core.py",
    "bootstrap_runtime.ps1",
    "bootstrap_runtime.bat",
    "runtime\_tkinter.pyd",
    "runtime\tcl\tcl8.6\init.tcl",
    "runtime\tcl\tk8.6\tk.tcl",
    "runtime\tkinter\__init__.py"
)
foreach ($item in $required) {
    if (-not (Test-Path (Join-Path $stage $item))) { Fail "missing staged file: $item" }
}
if (Test-Path (Join-Path $stage "runtime\python.exe")) {
    Fail "runtime\python.exe must not be staged (it is downloaded on first run)"
}

$exeKB = (Get-Item $exe).Length / 1KB
$stageMB = ((Get-ChildItem $stage -Recurse | Measure-Object Length -Sum).Sum) / 1MB
Write-Host ("launcher : {0:N0} KB" -f $exeKB)
Write-Host ("stage    : {0:N1} MB" -f $stageMB)

# ---------------------------------------------------------------- installer
if (-not $SkipInstaller) {
    $iscc = $null
    foreach ($candidate in @(
            "C:\Program Files (x86)\Inno Setup 6\ISCC.exe",
            "C:\Program Files\Inno Setup 6\ISCC.exe",
            (Join-Path $env:LOCALAPPDATA "Programs\Inno Setup 6\ISCC.exe"))) {
        if (Test-Path $candidate) { $iscc = $candidate; break }
    }

    if (-not $iscc) {
        Write-Host "WARNING: Inno Setup 6 (ISCC.exe) not found; skipping installer." -ForegroundColor Yellow
        Write-Host "         Install with: winget install JRSoftware.InnoSetup" -ForegroundColor Yellow
        $SkipInstaller = $true
    } else {
        Write-Step "Building installer with Inno Setup"
        & $iscc "installer.iss"
        if ($LASTEXITCODE -ne 0) { Fail "Inno Setup failed" }
        $setup = Get-ChildItem "dist\FeeMan-Setup-*.exe" -ErrorAction SilentlyContinue |
                 Sort-Object LastWriteTime -Descending | Select-Object -First 1
        if (-not $setup) { Fail "installer not produced" }
        Write-Host ("setup    : dist\{0} ({1:N1} MB)" -f $setup.Name, ($setup.Length / 1MB))
    }
}

Write-Step "Done"
Write-Host "Portable : dist\FeeMan\FeeMan.exe (first run downloads Python runtime)"
if (-not $SkipInstaller) {
    Write-Host "Installer: dist\FeeMan-Setup-1.0.0.exe"
}