# PPTMaker setup script (Windows PowerShell)
#
# This script is called by install.ps1 after cloning. You can also run it directly:
#   & "$env:USERPROFILE\.claude\skills\pptmaker\setup.ps1"
#
# What it does:
#   1) Verify Python is available
#   2) Install Python dependencies via pip (python-pptx, pywin32, pptmaker)
#   3) Smoke test the import
#
# ASCII-only messages to avoid cp949/cp1252 codepage issues under PowerShell 5.1.

$ErrorActionPreference = "Stop"

Write-Host "==> PPTMaker setup" -ForegroundColor Cyan

# 1) Check Python
try {
    $pyVersion = (python --version 2>&1)
    if ($LASTEXITCODE -ne 0) { throw "python --version failed" }
    Write-Host "  Python found: $pyVersion"
} catch {
    Write-Host "  [X] Python not found. Install 3.11+ from https://python.org and re-run." -ForegroundColor Red
    exit 1
}

# 2) Locate skill dir (this script's location)
$skillDir = Split-Path -Parent $MyInvocation.MyCommand.Path
Write-Host "  Skill dir: $skillDir"

# 3) Verify pyproject.toml
$pyproject = Join-Path $skillDir "pyproject.toml"
if (-not (Test-Path $pyproject)) {
    Write-Host "  [X] pyproject.toml not found. Re-clone the repo." -ForegroundColor Red
    exit 1
}

# 4) Install Python dependencies
Write-Host "==> Installing Python packages (python-pptx, pywin32, pptmaker)..." -ForegroundColor Cyan
$prevErr = $ErrorActionPreference
$ErrorActionPreference = "Continue"
& python -m pip install -e $skillDir
$pipExit = $LASTEXITCODE
$ErrorActionPreference = $prevErr

if ($pipExit -ne 0) {
    Write-Host "  [X] pip install failed (exit $pipExit)" -ForegroundColor Red
    exit 1
}

# 5) Smoke test
Write-Host "==> Verifying install..." -ForegroundColor Cyan
$env:PYTHONIOENCODING = "utf-8"
$check = & python -c "from pptmaker import TOKENS; print('OK', TOKENS.colors.accent1, TOKENS.fonts.major)"
if ($LASTEXITCODE -ne 0) {
    Write-Host "  [X] Verification failed:" -ForegroundColor Red
    Write-Host $check
    exit 1
}
Write-Host "  $check"

# 6) Claude Code skills path check
$claudeSkillsRoot = Join-Path $env:USERPROFILE ".claude\skills"
if ($skillDir -like "$claudeSkillsRoot*") {
    Write-Host "  [OK] Located under user-global skills path — auto-detected by Claude Code" -ForegroundColor Green
} else {
    Write-Host "  [!] This directory is NOT under ~/.claude/skills/" -ForegroundColor Yellow
    Write-Host "      For Claude Code auto-detection, do one of:"
    Write-Host "      1) Re-clone into $claudeSkillsRoot\pptmaker (recommended: use install.ps1)"
    Write-Host "      2) Create a junction: New-Item -ItemType Junction -Path $claudeSkillsRoot\pptmaker -Target $skillDir"
}

Write-Host ""
Write-Host "==> Setup complete" -ForegroundColor Green
Write-Host ""
Write-Host "Next steps:" -ForegroundColor Cyan
Write-Host "  1) Fully restart Claude Code"
Write-Host "  2) Open Claude Code from any folder, type your request in chat"
Write-Host "     (Korean or English natural language both work)"
Write-Host "  3) If PowerPoint is open, slides will appear live as they are built"
Write-Host ""
