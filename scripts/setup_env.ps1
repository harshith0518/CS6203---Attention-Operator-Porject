#requires -Version 5.1
<#
.SYNOPSIS
    CS-6023 project preflight + cuTile environment setup.

.DESCRIPTION
    Audits the machine against the hard prerequisites for cuTile Python
    (GPU compute capability 8.x+, driver r580+, CUDA Toolkit 13.2+),
    creates a local .venv, installs cuda-tile[tileiras], and runs a smoke test.

    It never modifies drivers or system software. Anything it cannot fix, it
    reports and stops.

.NOTES
    Written for Windows PowerShell 5.1 (no pwsh 7 dependency).
    If script execution is blocked, run:
        powershell -ExecutionPolicy Bypass -File .\scripts\setup_env.ps1

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File .\scripts\setup_env.ps1
    powershell -ExecutionPolicy Bypass -File .\scripts\setup_env.ps1 -SkipInstall
#>
param(
    [switch]$SkipInstall,
    [string]$PythonExe = ""
)

$ErrorActionPreference = 'Continue'
$root = Split-Path -Parent $PSScriptRoot
$venv = Join-Path $root '.venv'
$script:fail = @()
$script:warn = @()

function Section($t) { Write-Host "`n=== $t ===" -ForegroundColor Cyan }
function Ok($t)      { Write-Host "  [OK]   $t" -ForegroundColor Green }
function Warn2($t)   { Write-Host "  [WARN] $t" -ForegroundColor Yellow; $script:warn += $t }
function Bad($t)     { Write-Host "  [FAIL] $t" -ForegroundColor Red;    $script:fail += $t }

# Runs an interpreter given as @('exe', 'arg1', ...) plus extra args.
function Run-Py {
    param([string[]]$Argv, [string[]]$ArgList)
    if ($Argv.Count -gt 1) {
        & $Argv[0] $Argv[1..($Argv.Count - 1)] @ArgList
    } else {
        & $Argv[0] @ArgList
    }
}

Write-Host "CS-6023 cuTile environment check" -ForegroundColor White
Write-Host "Workspace: $root"

# ---------------------------------------------------------------- GPU / driver
Section "GPU and driver"
if (Get-Command nvidia-smi -ErrorAction SilentlyContinue) {
    $rows = & nvidia-smi --query-gpu=name,driver_version,memory.total,compute_cap --format=csv,noheader
    if ($LASTEXITCODE -ne 0 -or -not $rows) {
        Bad "nvidia-smi present but returned nothing (is the driver loaded?)"
    }
    foreach ($row in ($rows -split "`r?`n" | Where-Object { $_.Trim() })) {
        $f = $row.Split(',') | ForEach-Object { $_.Trim() }
        if ($f.Count -lt 4) { continue }
        $name = $f[0]; $drv = $f[1]; $mem = $f[2]; $cc = $f[3]
        Ok "GPU          : $name"
        Ok "VRAM         : $mem"
        Ok "Compute cap. : $cc"

        if ($cc -match '^(\d+)') {
            $ccMajor = [int]$Matches[1]
            if ($ccMajor -lt 8) {
                Bad "compute capability $cc is below 8.x -- cuTile Python will not run on this GPU"
            } elseif ($ccMajor -eq 8) {
                Ok  "compute capability 8.x is supported (Ampere/Ada support landed in cuTile 1.2.0, CTK 13.2)"
                Warn2 "CC 8.x has no TMA, no thread-block clusters and no FP8 MMA, so the cuTile knobs allow_tma / num_ctas / mma_scaled are INERT here. Plan the tunability experiment around tile shapes instead -- this is itself a report finding."
            } else {
                Ok "compute capability $cc is supported"
            }
        }

        $drvMajor = 0
        if ($drv -match '^(\d+)') { $drvMajor = [int]$Matches[1] }
        Ok "Driver       : $drv"
        if ($drvMajor -lt 580) {
            Bad "driver $drv is older than r580 -- cuTile Python CANNOT run. Install driver >= 580 (>= r580.126.09 / r582.16 if you also want detailed ncu profiling of cuTile kernels on Windows), reboot, and re-run this script."
        } elseif ($drvMajor -lt 582) {
            Warn2 "driver $drv runs cuTile fine, but detailed Nsight Compute profiling of cuTile kernels on Windows needs >= r582.16."
        } else {
            Ok "driver $drv satisfies both the runtime (>=580) and cuTile profiling (>=582.16) requirements"
        }
    }
} else {
    Bad "nvidia-smi not found -- no NVIDIA driver, or not on PATH"
}

# ---------------------------------------------------------------- CUDA toolkit
Section "CUDA Toolkit (needed only for the hand-written CUDA C++ half)"
if (Get-Command nvcc -ErrorAction SilentlyContinue) {
    $nvOut = (& nvcc --version) -join "`n"
    if ($nvOut -match 'release (\d+)\.(\d+)') {
        $maj = [int]$Matches[1]; $min = [int]$Matches[2]
        Ok "nvcc         : CUDA $maj.$min"
        if ($maj -lt 13 -or ($maj -eq 13 -and $min -lt 2)) {
            Bad "CUDA $maj.$min is older than 13.2 -- install CUDA Toolkit 13.2+ (Ampere cuTile support starts there)"
        } else {
            Ok "CUDA $maj.$min satisfies the 13.2+ requirement"
        }
    } else {
        Warn2 "nvcc present but the version string could not be parsed"
    }
} else {
    Warn2 "nvcc not on PATH. The cuTile half can still run via 'pip install cuda-tile[tileiras]', but the CUDA C++ half needs CUDA Toolkit 13.2+ installed."
}

$vswhere = "${env:ProgramFiles(x86)}\Microsoft Visual Studio\Installer\vswhere.exe"
if (Test-Path $vswhere) {
    $vspath = & $vswhere -latest -products * -requires Microsoft.VisualStudio.Component.VC.Tools.x86.x64 -property installationPath
    if ($vspath) { Ok "MSVC         : $vspath" }
    else         { Warn2 "Visual Studio present but the 'Desktop development with C++' workload is not installed; nvcc needs an MSVC host compiler." }
} else {
    Warn2 "No Visual Studio / Build Tools found; nvcc on Windows needs an MSVC host compiler."
}

# ---------------------------------------------------------------- disk
Section "Disk space (CUDA Toolkit needs ~8 GB)"
try {
    $vol = Get-Volume -DriveLetter C -ErrorAction Stop
    $freeGb = [math]::Round($vol.SizeRemaining / 1GB, 1)
    if ($freeGb -lt 12) { Warn2 "only $freeGb GB free on C:" } else { Ok "$freeGb GB free on C:" }
} catch {
    Warn2 "could not read disk space (permissions); check manually"
}

# ---------------------------------------------------------------- python
Section "Python 3.10-3.14"
# Prefer 3.12: widest wheel coverage for cupy-cuda13x / torch / cuda-tile.
$candidates = @()
if ($PythonExe) { $candidates += ,@($PythonExe) }
$candidates += ,@('py', '-3.12')
$candidates += ,@('py', '-3.13')
$candidates += ,@('py', '-3.14')
$candidates += ,@('py', '-3')
$candidates += ,@('python')

$pyArgv = $null
$pyver = $null
foreach ($cand in $candidates) {
    if (-not (Get-Command $cand[0] -ErrorAction SilentlyContinue)) { continue }
    try {
        $v = (Run-Py -Argv $cand -ArgList @('-c', "import sys;print('%d.%d'%sys.version_info[:2])") 2>$null)
        if ($LASTEXITCODE -eq 0 -and $v) {
            $pyArgv = $cand
            $pyver = ($v | Select-Object -First 1).Trim()
            break
        }
    } catch { }
}

if (-not $pyArgv) {
    Bad "no usable python found; pass -PythonExe 'C:\path\to\python.exe'"
} else {
    $maj = [int]$pyver.Split('.')[0]; $min = [int]$pyver.Split('.')[1]
    if ($maj -ne 3 -or $min -lt 10 -or $min -gt 14) {
        Bad "python $pyver is outside the 3.10-3.14 range supported by cuTile Python"
    } else {
        Ok "python $pyver via '$($pyArgv -join ' ')'"
        if ($min -ge 14) {
            Warn2 "python $pyver works for cuTile, but binary wheels for cupy-cuda13x / torch may lag on 3.14. If an install below fails, re-run with -PythonExe pointing at a 3.12 interpreter."
        }
    }
}

# ---------------------------------------------------------------- summary
Section "Preflight summary"
if ($script:fail.Count -eq 0) {
    Write-Host "  No blockers found." -ForegroundColor Green
} else {
    Write-Host "  $($script:fail.Count) blocker(s):" -ForegroundColor Red
    $script:fail | ForEach-Object { Write-Host "    - $_" -ForegroundColor Red }
}
if ($script:warn.Count -gt 0) {
    Write-Host "  $($script:warn.Count) warning(s):" -ForegroundColor Yellow
    $script:warn | ForEach-Object { Write-Host "    - $_" -ForegroundColor Yellow }
}

if ($script:fail.Count -gt 0) {
    Write-Host "`nStopping: fix the blockers above and re-run. Do not start writing code yet." -ForegroundColor Red
    exit 1
}
if ($SkipInstall) {
    Write-Host "`n-SkipInstall given; stopping after the audit." -ForegroundColor Yellow
    exit 0
}

# ---------------------------------------------------------------- venv + install
Section "Virtual environment"
$venvPy = Join-Path $venv 'Scripts\python.exe'
if (Test-Path $venvPy) {
    Ok "reusing existing venv at $venv"
} else {
    Write-Host "  creating $venv ..."
    Run-Py -Argv $pyArgv -ArgList @('-m', 'venv', $venv)
    if (-not (Test-Path $venvPy)) { Bad "venv creation failed"; exit 1 }
    Ok "created $venv"
}

Write-Host "`n  upgrading pip ..."
& $venvPy -m pip install --upgrade pip --quiet
if ($LASTEXITCODE -ne 0) { Warn2 "pip upgrade failed; continuing" }

Section "Installing packages"
foreach ($p in @('cuda-tile[tileiras]', 'numpy', 'cupy-cuda13x')) {
    Write-Host "  pip install $p ..."
    & $venvPy -m pip install --upgrade $p
    if ($LASTEXITCODE -ne 0) { Bad "pip install $p failed" }
}
if ($script:fail.Count -gt 0) {
    Write-Host "`nInstall failures above. Common causes:" -ForegroundColor Red
    Write-Host "  - no wheel for this Python version (retry with a 3.12 interpreter)"
    Write-Host "  - cupy-cuda13x unavailable: drop cupy and use torch instead (see PLAN.md section 3.2)"
    exit 1
}

# ---------------------------------------------------------------- smoke test
Section "Smoke test (the real gate)"
& $venvPy (Join-Path $PSScriptRoot 'smoke_cutile.py')
$rc = $LASTEXITCODE

Write-Host ""
if ($rc -eq 0) {
    Write-Host "=====================================================" -ForegroundColor Green
    Write-Host " GATE PASSED - cuTile runs on this GPU." -ForegroundColor Green
    Write-Host "=====================================================" -ForegroundColor Green
    Write-Host "Next: save the evidence, then start PLAN.md day 1:"
    Write-Host "  .\.venv\Scripts\python.exe -m pip freeze > docs\requirements.lock.txt"
} else {
    Write-Host "=====================================================" -ForegroundColor Red
    Write-Host " GATE FAILED - cuTile did not run." -ForegroundColor Red
    Write-Host "=====================================================" -ForegroundColor Red
    Write-Host "Check the driver version and compute capability first."
    Write-Host "If both satisfy the requirements and this still fails, switch to a cloud/lab GPU today (PLAN.md R1/R2)."
}
exit $rc
