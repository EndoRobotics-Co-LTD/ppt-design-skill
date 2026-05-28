# PPTMaker bootstrap installer
#
# Usage (PowerShell one-liner):
#   irm https://raw.githubusercontent.com/EndoRobotics-Co-LTD/ppt-design-skill/lee-dev/install.ps1 | iex
#
# What this does:
#   1) Check Python and Git are installed
#   2) Clone the repo into the correct location (~/.claude/skills/ppt-design-skill/) — folder name auto-handled
#   3) Run setup.ps1 (install Python dependencies + verify)
#
# ASCII-only messages so this works under PowerShell 5.1 cp949/cp1252 codepages
# without BOM/encoding issues.

$ErrorActionPreference = "Stop"

Write-Host ""
Write-Host "==> PPTMaker bootstrap installer" -ForegroundColor Cyan
Write-Host ""

# ----- config -----
$repoUrl = "https://github.com/EndoRobotics-Co-LTD/ppt-design-skill.git"
$repoBranch = "lee-dev"  # TODO: change to "main" after merge
$skillRoot = Join-Path $env:USERPROFILE ".claude\skills"
$skillDir = Join-Path $skillRoot "ppt-design-skill"

# ----- 1) prereq check -----
Write-Host "[1/4] Checking prerequisites..." -ForegroundColor Cyan

try {
    $pyVersion = (python --version 2>&1)
    if ($LASTEXITCODE -ne 0) { throw "python --version failed" }
    Write-Host "  Python: $pyVersion" -ForegroundColor Green
} catch {
    Write-Host "  [X] Python not found." -ForegroundColor Red
    Write-Host "      Install 3.11+ from https://python.org, then restart PowerShell and re-run." -ForegroundColor Yellow
    return
}

try {
    $gitVersion = (git --version 2>&1)
    if ($LASTEXITCODE -ne 0) { throw "git --version failed" }
    Write-Host "  Git:    $gitVersion" -ForegroundColor Green
} catch {
    Write-Host "  [X] Git not found." -ForegroundColor Red
    Write-Host "      Install from https://git-scm.com/download/win, restart PowerShell, re-run." -ForegroundColor Yellow
    return
}

# ----- 2) prepare skills directory -----
Write-Host ""
Write-Host "[2/4] Preparing skills directory..." -ForegroundColor Cyan
if (-not (Test-Path $skillRoot)) {
    Write-Host "  Creating $skillRoot"
    New-Item -ItemType Directory -Path $skillRoot -Force | Out-Null
}

# ----- 3) check if already installed -----
if (Test-Path $skillDir) {
    Write-Host ""
    Write-Host "  [!] Already installed: $skillDir" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "  To update:" -ForegroundColor Yellow
    Write-Host "    cd $skillDir" -ForegroundColor White
    Write-Host "    git pull" -ForegroundColor White
    Write-Host "    python -m pip install -e . --upgrade" -ForegroundColor White
    Write-Host ""
    Write-Host "  To reinstall, remove the existing folder first:" -ForegroundColor Yellow
    Write-Host "    Remove-Item -Recurse -Force $skillDir" -ForegroundColor White
    return
}

# ----- 4) clone (explicit destination) -----
Write-Host ""
Write-Host "[3/4] Cloning repo..." -ForegroundColor Cyan
Write-Host "  $repoUrl"
Write-Host "  -> $skillDir"
Write-Host "  (branch: $repoBranch)"

# Run git clone without redirecting stderr — git writes "Cloning into..." to
# stderr, which PowerShell 5.1 wraps into ErrorRecord and triggers $ErrorActionPreference="Stop".
# Temporarily relax error handling and check exit code instead.
$prevErr = $ErrorActionPreference
$ErrorActionPreference = "Continue"
& git clone --branch $repoBranch $repoUrl $skillDir
$cloneExit = $LASTEXITCODE
$ErrorActionPreference = $prevErr

if ($cloneExit -ne 0) {
    Write-Host ""
    Write-Host "  [X] Clone failed (exit $cloneExit)" -ForegroundColor Red
    Write-Host ""
    Write-Host "  Possible causes:" -ForegroundColor Yellow
    Write-Host "    - No GitHub access: contact Strategic Planning Team (Eunsang Lee)"
    Write-Host "    - Network/firewall issue: contact internal IT"
    return
}
Write-Host "  [OK] Clone complete" -ForegroundColor Green

# ----- 5) run setup.ps1 -----
$setupPath = Join-Path $skillDir "setup.ps1"
if (-not (Test-Path $setupPath)) {
    Write-Host ""
    Write-Host "  [X] setup.ps1 not found: $setupPath" -ForegroundColor Red
    Write-Host "      Repo structure may have changed. Contact eunsang.lee@endorobo.com" -ForegroundColor Yellow
    return
}

Write-Host ""
Write-Host "[4/4] Installing Python packages + verification..." -ForegroundColor Cyan
& $setupPath
if ($LASTEXITCODE -ne 0) {
    Write-Host ""
    Write-Host "  [X] setup.ps1 returned error." -ForegroundColor Red
    Write-Host "      Install folder left in place: $skillDir" -ForegroundColor Yellow
    Write-Host "      Try manually: cd $skillDir; python -m pip install -e ." -ForegroundColor Yellow
    return
}

# ----- success -----
Write-Host ""
Write-Host "================================================" -ForegroundColor Green
Write-Host " Installation complete" -ForegroundColor Green
Write-Host "================================================" -ForegroundColor Green
Write-Host ""
Write-Host "Next steps:" -ForegroundColor Cyan
Write-Host "  1) Fully restart Claude Code"
Write-Host "  2) Run Claude Code from any folder"
Write-Host "  3) Type your request in chat — Korean or English natural language"
Write-Host "     Examples: /pptmaker  |  make me an IR deck in EndoRobotics standard"
Write-Host "     See README.md for the full conversation flow."
Write-Host ""
Write-Host "Skill location: $skillDir" -ForegroundColor White
Write-Host ""
