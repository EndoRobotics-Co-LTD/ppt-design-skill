# PPTMaker 설치 스크립트 (Windows PowerShell)
# Usage:
#   git clone <url> "$env:USERPROFILE\.claude\skills\pptmaker"
#   & "$env:USERPROFILE\.claude\skills\pptmaker\setup.ps1"

$ErrorActionPreference = "Stop"

Write-Host "==> PPTMaker 설치 시작" -ForegroundColor Cyan

# 1) Python 확인
try {
    $pyVersion = (python --version 2>&1)
    Write-Host "  Python 발견: $pyVersion"
} catch {
    Write-Host "  Python을 찾을 수 없습니다. https://python.org 에서 3.11+ 설치 후 재실행하세요." -ForegroundColor Red
    exit 1
}

# 2) 스킬 디렉토리 위치 확인 (이 스크립트가 있는 곳)
$skillDir = Split-Path -Parent $MyInvocation.MyCommand.Path
Write-Host "  스킬 디렉토리: $skillDir"

# 3) pyproject.toml 존재 확인
$pyproject = Join-Path $skillDir "pyproject.toml"
if (-not (Test-Path $pyproject)) {
    Write-Host "  pyproject.toml을 찾을 수 없습니다. 저장소를 다시 클론하세요." -ForegroundColor Red
    exit 1
}

# 4) Python 의존성 설치
Write-Host "==> Python 패키지 설치 (python-pptx, pywin32, pptmaker)..." -ForegroundColor Cyan
$installResult = python -m pip install -e $skillDir 2>&1
if ($LASTEXITCODE -ne 0) {
    Write-Host "  설치 실패. 출력:" -ForegroundColor Red
    Write-Host $installResult
    exit 1
}

# 5) Smoke test
Write-Host "==> 설치 검증..." -ForegroundColor Cyan
$env:PYTHONIOENCODING = "utf-8"
$check = python -c "from pptmaker import TOKENS; print('OK', TOKENS.colors.accent1, TOKENS.fonts.major)" 2>&1
if ($LASTEXITCODE -ne 0) {
    Write-Host "  검증 실패:" -ForegroundColor Red
    Write-Host $check
    exit 1
}
Write-Host "  $check"

# 6) Claude Code skills 경로 확인
$claudeSkillsRoot = Join-Path $env:USERPROFILE ".claude\skills"
if ($skillDir -like "$claudeSkillsRoot*") {
    Write-Host "  ✓ Claude Code 사용자 전역 skills 경로에 설치되어 있음 — 자동 검출됨" -ForegroundColor Green
} else {
    Write-Host "  ⚠ 이 디렉토리가 ~/.claude/skills/ 안에 없습니다." -ForegroundColor Yellow
    Write-Host "    스킬을 자동 인식하려면 다음 중 하나:" -ForegroundColor Yellow
    Write-Host "    1) 저장소를 $claudeSkillsRoot\pptmaker 로 다시 클론"
    Write-Host "    2) 또는 .claude/skills/pptmaker/SKILL.md 를 그 경로로 심볼릭 링크"
}

Write-Host ""
Write-Host "==> 설치 완료" -ForegroundColor Green
Write-Host ""
Write-Host "다음 단계:" -ForegroundColor Cyan
Write-Host "  1) Claude Code를 다시 시작하세요"
Write-Host "  2) 대화창에 /pptmaker 또는 'PPT 만들어줘' 입력"
Write-Host "  3) PowerPoint 데스크톱이 켜져 있으면 라이브 모드로 즉시 슬라이드 생성"
Write-Host ""
