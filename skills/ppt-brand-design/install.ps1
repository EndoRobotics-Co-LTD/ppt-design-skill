# ppt-brand-design bootstrap installer
#
# Usage (PowerShell one-liner):
#   irm https://raw.githubusercontent.com/EndoRobotics-Co-LTD/endo-skills/main/skills/ppt-brand-design/install.ps1 | iex
#
# What this does:
#   1) Check Python and Git are installed
#   2) Clone endo-skills repo to a temp folder, extract skills/ppt-brand-design/
#      → ~/.claude/skills/ppt-brand-design/  (correct location for Claude Code)
#   3) Run setup.ps1 (install Python dependencies + verify)
#
# Alternative: `npx skills add EndoRobotics-Co-LTD/endo-skills -s ppt-brand-design`
# (npx 만 사용, Python deps 는 수동으로 `pip install -e .` 필요)
#
# ASCII-only messages so this works under PowerShell 5.1 cp949/cp1252 codepages.

$ErrorActionPreference = "Stop"

Write-Host ""
Write-Host "==> ppt-brand-design bootstrap installer" -ForegroundColor Cyan
Write-Host ""

# ----- config -----
$repoUrl = "https://github.com/EndoRobotics-Co-LTD/endo-skills.git"
$repoBranch = "main"
$skillName = "ppt-brand-design"
$repoSkillPath = "skills/$skillName"
$skillRoot = Join-Path $env:USERPROFILE ".claude\skills"
$skillDir = Join-Path $skillRoot $skillName

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
    Write-Host "    Remove-Item -Recurse -Force $skillDir" -ForegroundColor White
    Write-Host "    Then re-run this installer" -ForegroundColor White
    Write-Host ""
    Write-Host "  (Note: endo-skills repo is multi-skill — direct git pull from $skillDir won't work" -ForegroundColor Yellow
    Write-Host "   because only the ppt-brand-design subfolder was extracted.)" -ForegroundColor Yellow
    return
}

# ----- 4) clone + extract skills/ppt-brand-design/ -----
Write-Host ""
Write-Host "[3/4] Cloning endo-skills repo and extracting $skillName..." -ForegroundColor Cyan
Write-Host "  $repoUrl (branch: $repoBranch)"
Write-Host "  -> $skillDir"

$tempDir = Join-Path $env:TEMP ("endo-skills-" + [Guid]::NewGuid().ToString())

$prevErr = $ErrorActionPreference
$ErrorActionPreference = "Continue"
& git clone --depth 1 --branch $repoBranch $repoUrl $tempDir
$cloneExit = $LASTEXITCODE
$ErrorActionPreference = $prevErr

if ($cloneExit -ne 0) {
    Write-Host ""
    Write-Host "  [X] Clone failed (exit $cloneExit)" -ForegroundColor Red
    Write-Host ""
    Write-Host "  Possible causes:" -ForegroundColor Yellow
    Write-Host "    - Network/firewall issue: contact internal IT"
    Write-Host "    - Repo doesn't exist or visibility changed"
    if (Test-Path $tempDir) { Remove-Item -Recurse -Force $tempDir }
    return
}

$sourceSkillPath = Join-Path $tempDir $repoSkillPath
if (-not (Test-Path $sourceSkillPath)) {
    Write-Host "  [X] Skill folder not found in repo: $repoSkillPath" -ForegroundColor Red
    Remove-Item -Recurse -Force $tempDir
    return
}

# Move skills/ppt-brand-design contents to ~/.claude/skills/ppt-brand-design/
Move-Item -Path $sourceSkillPath -Destination $skillDir
# Cleanup temp clone
Remove-Item -Recurse -Force $tempDir
Write-Host "  [OK] Extracted to $skillDir" -ForegroundColor Green

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
Write-Host " Installation complete: $skillName" -ForegroundColor Green
Write-Host "================================================" -ForegroundColor Green
Write-Host ""
Write-Host "Next steps:" -ForegroundColor Cyan
Write-Host "  1) Fully restart Claude Code"
Write-Host "  2) Run Claude Code from any folder"
Write-Host "  3) Type your request in chat (Korean or English natural language)"
Write-Host "     Examples: 'PPT 만들어줘'  |  'make me an IR deck'"
Write-Host "     See README.md for the full conversation flow."
Write-Host ""
Write-Host "Skill location: $skillDir" -ForegroundColor White
Write-Host ""
