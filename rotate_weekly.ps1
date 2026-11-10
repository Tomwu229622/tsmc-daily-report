# rotate_weekly.ps1 -- weekly (Monday 08:30) wrapper for rotate_archive.py
# Moves IR_CALENDAR / EPS triggers entries beyond the working-set limits from
# generate_tsmc_html_report.py into history/*.json, then commits and pushes.
# Runs BEFORE the 09:00 daily report so that session starts from a clean tree.
# ASCII-only on purpose (PowerShell 5.1 + cp950 console).
#
# Usage:  powershell -NoProfile -ExecutionPolicy Bypass -File rotate_weekly.ps1 [-DryRun]
param([switch]$DryRun)

$ErrorActionPreference = "Stop"
$Proj    = Split-Path -Parent $MyInvocation.MyCommand.Path
$LogDir  = Join-Path $env:LOCALAPPDATA "tsmc_rotate"
$BakDir  = Join-Path $LogDir "bak"
$LogFile = Join-Path $LogDir "rotate_weekly.log"
New-Item -ItemType Directory -Force $LogDir | Out-Null
New-Item -ItemType Directory -Force $BakDir | Out-Null

function Log($msg) {
    $line = "{0} {1}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $msg
    Add-Content -Path $LogFile -Value $line -Encoding UTF8
    Write-Output $line
}

Set-Location $Proj
Log "start dryrun=$DryRun proj=$Proj"

# python: prefer PATH, fall back to py launcher
$py = $null
$cmd = Get-Command python -ErrorAction SilentlyContinue
if ($cmd) { $py = $cmd.Source }
if (-not $py) { $cmd = Get-Command py -ErrorAction SilentlyContinue; if ($cmd) { $py = $cmd.Source } }
if (-not $py) { Log "FAIL: python not found"; exit 1 }

# guard: targets must be clean (another process may be mid-edit)
$dirty = git status --porcelain -- generate_tsmc_html_report.py history
if ($dirty) { Log "SKIP: working tree dirty on targets: $dirty"; exit 0 }

# run rotate
$args = @("-X", "utf8", "rotate_archive.py", "--dir", $Proj, "--backup-dir", $BakDir)
if ($DryRun) { $args += "--dry-run" }
$out = & $py @args 2>&1
$rc = $LASTEXITCODE
foreach ($l in $out) { Log "rotate> $l" }
if ($rc -ne 0) { Log "FAIL: rotate exit $rc"; exit 1 }
if (($out -join "`n") -match "nothing to rotate") { Log "done: nothing to rotate"; exit 0 }
if ($DryRun) { Log "done: dry-run only, no commit"; exit 0 }

# commit + push only if files actually changed
$changed = git status --porcelain -- generate_tsmc_html_report.py history
if (-not $changed) { Log "done: rotate reported work but no file changed"; exit 0 }
$today = Get-Date -Format "yyyy-MM-dd"
$msgFile = Join-Path $LogDir "commitmsg.txt"
$body = @("chore: rotate IR/EPS archive ($today)", "", "Automated weekly run of rotate_archive.py (rotate_weekly.ps1).") + ($out | ForEach-Object { "  $_" })
Set-Content -Path $msgFile -Value $body -Encoding UTF8
git add generate_tsmc_html_report.py history/ir_calendar_archive.json history/eps_triggers_archive.json
git commit -q -F $msgFile
if ($LASTEXITCODE -ne 0) { Log "FAIL: git commit exit $LASTEXITCODE"; exit 1 }
$hash = git log -1 --format=%h
Log "committed $hash"
git push origin master 2>&1 | ForEach-Object { Log "push> $_" }
if ($LASTEXITCODE -ne 0) { Log "WARN: push failed (exit $LASTEXITCODE), commit kept locally"; exit 1 }
Log "done: pushed $hash"
exit 0
