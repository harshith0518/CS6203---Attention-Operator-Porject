#requires -Version 5.1
<#
.SYNOPSIS
    CS-6023 cuTile environment setup, with step checkpoints.

.DESCRIPTION
    Nine numbered steps. Each prints its number, percentage, a progress bar, what it is
    doing right now, and how long it should take. Every network call is bounded, so a bad
    connection fails in seconds instead of hanging.

    Prints a per-step timing table at the end.

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File .\scripts\setup_env.ps1
    powershell -ExecutionPolicy Bypass -File .\scripts\setup_env.ps1 -AuditOnly
#>
param(
    [switch]$AuditOnly,
    [switch]$Rebuild,
    [string]$PythonExe = ""
)

$ErrorActionPreference = 'Continue'
$root = Split-Path -Parent $PSScriptRoot
$venv = Join-Path $root '.venv'
$script:fail = @()
$script:warn = @()
$script:times = @()
$TOTAL = 9
$script:started = Get-Date

# ------------------------------------------------------------------ machinery
function Bar([int]$pct, [int]$width = 24) {
    $fill = [int][math]::Floor($pct / 100.0 * $width)
    ('#' * $fill) + ('.' * ($width - $fill))
}

function Step([int]$n, [string]$name, [string]$expected) {
    $pct = [int]((($n - 1) / $TOTAL) * 100)
    Write-Host ''
    Write-Host ("[{0}/{1}] {2,3}%  {3}  {4}" -f $n, $TOTAL, $pct, (Bar $pct), $name) -ForegroundColor Cyan
    Write-Host ("        expected: {0}" -f $expected) -ForegroundColor DarkGray
    $script:stepStart = Get-Date
    $script:stepName = $name
}

function Done() {
    $s = [math]::Round(((Get-Date) - $script:stepStart).TotalSeconds, 1)
    Write-Host ("        done in {0}s" -f $s) -ForegroundColor Green
    $script:times += [pscustomobject]@{ Step = $script:stepName; Seconds = $s }
}

function Failed([string]$why) {
    $s = [math]::Round(((Get-Date) - $script:stepStart).TotalSeconds, 1)
    Write-Host ("        FAILED after {0}s: {1}" -f $s, $why) -ForegroundColor Red
    $script:fail += $why
    $script:times += [pscustomobject]@{ Step = $script:stepName; Seconds = $s }
}

function Ok($t)    { Write-Host ("        ok   {0}" -f $t) -ForegroundColor Green }
function Warn2($t) { Write-Host ("        warn {0}" -f $t) -ForegroundColor Yellow; $script:warn += $t }
function Bad($t)   { Write-Host ("        FAIL {0}" -f $t) -ForegroundColor Red; $script:fail += $t }

Write-Host "CS-6023 cuTile environment check" -ForegroundColor White
Write-Host ("workspace: {0}" -f $root)
Write-Host ("started  : {0}" -f $script:started.ToString("HH:mm:ss"))
Write-Host ""
Write-Host "  If output ever freezes mid-step, the console has a text selection (QuickEdit)." -ForegroundColor DarkGray
Write-Host "  Click inside the window and press Esc. The process is blocked, not slow." -ForegroundColor DarkGray
if ($Rebuild) { Write-Host "  -Rebuild: the existing .venv will be deleted first." -ForegroundColor Yellow }

# ------------------------------------------------------------ temp directory
# pip creates its build tracker in TEMP. A DSH session rewrites TEMP/TMP to a
# per-session folder that other shells cannot use, and pip then reports
# "No usable temporary directory found" and installs nothing. Pin it to a
# directory we own, before any pip call.
$tmpDir = Join-Path $root 'tmp'
if (-not (Test-Path $tmpDir)) { New-Item -ItemType Directory -Path $tmpDir -Force | Out-Null }
$env:TEMP = $tmpDir
$env:TMP = $tmpDir
$env:TMPDIR = $tmpDir
Write-Host ("temp dir : {0}" -f $tmpDir) -ForegroundColor DarkGray

# CuPy warns "CUDA path could not be detected" and the TileIR compiler wants a
# CUDA root. The wheels put one in site-packages, so point both at it.
$cudaRoot = Join-Path $root '.venv\Lib\site-packages\nvidia\cu13'
if (Test-Path $cudaRoot) { $env:CUDA_PATH = $cudaRoot }
$env:CUPY_CACHE_DIR = Join-Path $tmpDir 'cupy'
$env:CUDA_TILE_CACHE_DIR = Join-Path $tmpDir 'cuda_tile'
New-Item -ItemType Directory -Force $env:CUPY_CACHE_DIR, $env:CUDA_TILE_CACHE_DIR | Out-Null

# ================================================================ 1 GPU/driver
Step 1 "GPU and driver" "2-5 s"
if (Get-Command nvidia-smi -ErrorAction SilentlyContinue) {
    $rows = & nvidia-smi --query-gpu=name,driver_version,memory.total,compute_cap --format=csv,noheader
    foreach ($row in ($rows -split "`r?`n" | Where-Object { $_.Trim() })) {
        $f = $row.Split(',') | ForEach-Object { $_.Trim() }
        if ($f.Count -lt 4) { continue }
        Ok ("{0}  |  {1}  |  CC {2}" -f $f[0], $f[2], $f[3])
        if ($f[3] -match '^(\d+)' -and [int]$Matches[1] -lt 8) {
            Bad ("compute capability {0} is below 8.x - cuTile will not run" -f $f[3])
        }
        $drvMajor = 0
        if ($f[1] -match '^(\d+)') { $drvMajor = [int]$Matches[1] }
        if ($drvMajor -lt 580) {
            Bad ("driver {0} < 580 - cuTile cannot run" -f $f[1])
        } elseif ($drvMajor -lt 582) {
            Ok ("driver {0} - runs fine; >= 582.16 for detailed ncu on cuTile" -f $f[1])
        } else {
            Ok ("driver {0} - satisfies runtime and profiling" -f $f[1])
        }
    }
} else { Bad "nvidia-smi not found" }
if ($script:fail.Count -eq 0) { Done }

# ============================================================ 2 toolkit/MSVC
Step 2 "CUDA Toolkit and MSVC" "1-3 s"
if (Get-Command nvcc -ErrorAction SilentlyContinue) {
    $nv = (& nvcc --version) -join "`n"
    if ($nv -match 'release (\d+)\.(\d+)') {
        Ok ("nvcc: CUDA {0}.{1}" -f $Matches[1], $Matches[2])
    }
} else {
    Warn2 "nvcc not on PATH - only the CUDA C++ half needs it, not this gate"
}
$vswhere = "${env:ProgramFiles(x86)}\Microsoft Visual Studio\Installer\vswhere.exe"
if (Test-Path $vswhere) {
    $vsp = & $vswhere -latest -products * -requires Microsoft.VisualStudio.Component.VC.Tools.x86.x64 -property installationPath
    if ($vsp) { Ok "MSVC present" } else { Warn2 "VS found but no C++ workload" }
} else { Warn2 "no MSVC - nvcc would have no host compiler" }
Done

# ============================================================== 3 disk + py
Step 3 "disk space and Python" "1-3 s"
try {
    $free = [math]::Round((Get-Volume -DriveLetter C -ErrorAction Stop).SizeRemaining / 1GB, 1)
    if ($free -lt 12) { Warn2 ("only {0} GB free" -f $free) } else { Ok ("{0} GB free" -f $free) }
} catch { Warn2 "could not read disk space" }

$cands = @()
if ($PythonExe) { $cands += , @($PythonExe) }
$cands += , @('py', '-3.12')
$cands += , @('py', '-3')
$cands += , @('python')
$script:pyArgv = $null
foreach ($c in $cands) {
    if (-not (Get-Command $c[0] -ErrorAction SilentlyContinue)) { continue }
    try {
        if ($c.Count -gt 1) { $v = (& $c[0] $c[1..($c.Count - 1)] -c "import sys;print('%d.%d'%sys.version_info[:2])" 2>$null) }
        else { $v = (& $c[0] -c "import sys;print('%d.%d'%sys.version_info[:2])" 2>$null) }
        if ($LASTEXITCODE -eq 0 -and $v) { $script:pyArgv = $c; Ok ("python {0} via '{1}'" -f $v.Trim(), ($c -join ' ')); break }
    } catch { }
}
if (-not $script:pyArgv) { Bad "no usable python found" }
Done

if ($script:fail.Count -gt 0) {
    Write-Host ''
    Write-Host "Stopping before anything is downloaded. Fix the failures above and re-run." -ForegroundColor Red
    exit 1
}

# ================================================================ 4 network
Step 4 "PyPI reachability (this is where it hung before)" "<= 12 s, hard timeout"
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
$probe = $null
try {
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    $probe = Invoke-WebRequest -Uri "https://pypi.org/pypi/cuda-tile/json" -TimeoutSec 12 -UseBasicParsing
    $sw.Stop()
    Ok ("pypi.org answered HTTP {0} in {1} ms" -f $probe.StatusCode, $sw.ElapsedMilliseconds)
    Done
} catch {
    Failed ("cannot reach pypi.org - {0}" -f $_.Exception.Message.Split([Environment]::NewLine)[0])
    Write-Host ''
    Write-Host "  Nothing was installed, because nothing can be. Check, in this order:" -ForegroundColor Yellow
    Write-Host "    1. Are you behind a proxy?  `$env:HTTPS_PROXY  and  pip config list" -ForegroundColor Yellow
    Write-Host "    2. Does a browser open https://pypi.org/simple/ ?" -ForegroundColor Yellow
    Write-Host "    3. Try once by hand:  .\.venv\Scripts\python.exe -m pip download numpy -d `$env:TEMP\t" -ForegroundColor Yellow
    Write-Host "    4. On campus wifi, the mirror may be faster:  -i https://pypi.org/simple" -ForegroundColor Yellow
    Write-Host ''
    Write-Host "  A stalled install is a network problem, not a code problem. Do not debug the script." -ForegroundColor Yellow
    exit 3
}

if ($AuditOnly) {
    Write-Host ''
    Write-Host "-AuditOnly given: environment checks passed, nothing installed." -ForegroundColor Yellow
    exit 0
}

# ==================================================================== 5 venv
Step 5 "virtual environment" "5-20 s"
$venvPy = Join-Path $venv 'Scripts\python.exe'
if ($Rebuild -and (Test-Path $venv)) {
    Write-Host "        -Rebuild given: removing the old .venv ..." -ForegroundColor Yellow
    Remove-Item $venv -Recurse -Force -ErrorAction SilentlyContinue
    if (Test-Path $venv) { Failed ".venv could not be removed - close any terminal using it, then retry"; exit 1 }
}
if (Test-Path $venvPy) {
    Ok "reusing .venv"
} else {
    Write-Host "        creating .venv ..."
    if ($script:pyArgv.Count -gt 1) { & $script:pyArgv[0] $script:pyArgv[1..($script:pyArgv.Count - 1)] -m venv $venv }
    else { & $script:pyArgv[0] -m venv $venv }
    if (-not (Test-Path $venvPy)) { Failed "venv creation failed"; exit 1 }
    Ok "created .venv"
}
# a venv that cannot be written to is the failure mode that looks like a hang
try {
    $probe = Join-Path $venv 'write-probe.tmp'
    Set-Content $probe "x" -ErrorAction Stop
    Remove-Item $probe -Force -ErrorAction SilentlyContinue
    Ok "venv is writable"
} catch {
    Failed "cannot write inside .venv - re-run with -Rebuild"
    exit 1
}
# pip dies early and confusingly if it cannot make a temp dir; prove it here
& $venvPy -c "import tempfile; tempfile.mkdtemp()" 2>$null
if ($LASTEXITCODE -eq 0) {
    Ok "temp directory usable"
} else {
    Failed "python cannot create a temp directory in $tmpDir - pip will fail"
    exit 1
}
Done

# ------------------------------------------------------------ pip settings
$env:PIP_DISABLE_PIP_VERSION_CHECK = "1"
$env:PIP_DEFAULT_TIMEOUT = "15"
$env:PIP_RETRIES = "2"
$pipCommon = @('--no-input', '--timeout', '15', '--retries', '2', "--progress-bar=on")

# Runs pip with a start stamp and a live heartbeat, so "no output" is never
# ambiguous. If the heartbeat stops too, the console is frozen (see the
# QuickEdit note at the top), not pip.
function Invoke-Pip {
    param([string[]]$PipArgs, [string]$What)
    $t0 = Get-Date
    Write-Host ("        {0}  starting at {1}" -f $What, $t0.ToString("HH:mm:ss")) -ForegroundColor DarkGray
    $job = Start-Job -ScriptBlock {
        param($py, $a)
        & $py -m pip @a 2>&1
        $LASTEXITCODE
    } -ArgumentList $venvPy, $PipArgs
    while ($job.State -notin @('Completed', 'Failed', 'Stopped')) {
        Start-Sleep -Seconds 20
        if ($job.State -notin @('Completed', 'Failed', 'Stopped')) {
            $waited = [int]((Get-Date) - $t0).TotalSeconds
            Write-Host ("        ... {0}s elapsed, still working. If this line stops appearing, the console is frozen: click the window and press Esc." -f $waited) -ForegroundColor DarkGray
        }
    }
    $out = Receive-Job $job -Wait -AutoRemoveJob
    $code = 0
    if ($out.Count -gt 0) {
        $code = $out[-1]
        $out[0..($out.Count - 1)] | Where-Object { $_ -ne $code } | ForEach-Object { Write-Host ("        $_") }
    }
    return $code
}

# =============================================================== 6 pip upgrade
Step 6 "upgrading pip (optional - a failure here is not fatal)" "5-20 s"
$code = Invoke-Pip -PipArgs (@('install', '--upgrade', 'pip') + $pipCommon) -What "pip upgrade"
if ($code -eq 0) { Ok "pip ready"; Done } else { Warn2 "pip upgrade skipped, continuing"; Done }

# ============================================================ 7 cuda-tile
Step 7 "installing cuda-tile[tileiras] (the big one, ~0.7-1 GB)" "2-15 min"
$code = Invoke-Pip -PipArgs (@('install', '--upgrade', 'cuda-tile[tileiras]') + $pipCommon) -What "cuda-tile + TileIR compiler + ptxas + libnvvm"
if ($code -eq 0) { Ok "cuda-tile installed"; Done } else { Failed "cuda-tile install failed - paste the pip lines above"; exit 1 }

# ========================================================= 8 numpy + cupy
Step 8 "installing numpy and cupy-cuda13x (~0.3-0.5 GB)" "1-8 min"
$code = Invoke-Pip -PipArgs (@('install', '--upgrade', 'numpy', 'cupy-cuda13x') + $pipCommon) -What "numpy + cupy-cuda13x"
if ($code -eq 0) {
    Ok "numpy + cupy installed"
} else {
    Warn2 "cupy-cuda13x failed - not fatal, the harness can use torch instead"
    $code = Invoke-Pip -PipArgs (@('install', '--upgrade', 'numpy') + $pipCommon) -What "numpy only"
}
Done

# ============================================================= 9 smoke test
Step 9 "smoke test - the actual gate" "10-40 s (JIT compiles on first run)"
& $venvPy (Join-Path $PSScriptRoot 'smoke_cutile.py')
$rc = $LASTEXITCODE
if ($rc -eq 0) { Ok "cuTile runs on this GPU"; Done } else { Failed "smoke test exit $rc" }

# ------------------------------------------------------------------ summary
$total = [math]::Round(((Get-Date) - $script:started).TotalSeconds, 1)

Write-Host ''
Write-Host "==================== summary ====================" -ForegroundColor White
$script:times | ForEach-Object { "  {0,-52} {1,7}s" -f $_.Step, $_.Seconds }
Write-Host ("  {0,-52} {1,7}s" -f "TOTAL", $total)
Write-Host ''

if ($rc -eq 0 -and $script:fail.Count -eq 0) {
    Write-Host "GATE PASSED - cuTile runs on this GPU." -ForegroundColor Green
    Write-Host "Next:"
    Write-Host "  .\.venv\Scripts\python.exe -m pip freeze > docs\requirements.lock.txt"
    Write-Host "  then start PLAN.md day 1 (bench.py)."
} else {
    Write-Host "GATE NOT PASSED." -ForegroundColor Red
    if ($script:fail.Count -gt 0) { $script:fail | ForEach-Object { Write-Host ("  - {0}" -f $_) -ForegroundColor Red } }
    Write-Host ""
    Write-Host "If the smoke test failed but everything installed, the problem is cuTile on"
    Write-Host "this GPU, not the environment - paste the smoke-test output."
}
exit $(if ($rc -eq 0 -and $script:fail.Count -eq 0) { 0 } else { 1 })

