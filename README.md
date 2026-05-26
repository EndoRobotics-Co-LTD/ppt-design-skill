# PPTMaker — EndoRobotics 전사 PPT 빌더

> AI가 EndoRobotics 사내 표준 디자인으로 PPT를 만들어주는 Claude Code Skill.
> 모든 슬라이드가 같은 폰트·색·요소 배치로 자동 통일됩니다.

---

## ⚡ 빠른 설치 (전사 직원용)

PowerShell 한 줄로 끝납니다:

```powershell
git clone https://github.com/<org>/PPTMaker.git "$env:USERPROFILE\.claude\skills\pptmaker"; & "$env:USERPROFILE\.claude\skills\pptmaker\setup.ps1"
```

> ⚠️ `<org>` 부분은 사내 GitHub Organization 이름으로 교체하세요. 사내 배포 담당자에게 정확한 URL을 받으세요.

설치 후 Claude Code를 다시 시작하면 `/pptmaker` 명령어로 호출 가능합니다.

---

## 사전 요구사항

| 항목 | 버전 |
|---|---|
| Windows | 10 또는 11 |
| PowerPoint (Microsoft 365) | 데스크톱 버전 |
| Python | 3.11 이상 |
| Claude Code | 최신 버전 |

설치 전 PowerShell에서 확인:
```powershell
python --version    # 3.11 이상이어야 함
git --version       # 설치되어 있어야 함
```

---

## 설치 단계 (수동)

자동 설치 스크립트가 실패하면 수동으로:

```powershell
# 1) 스킬 폴더에 클론
git clone https://github.com/<org>/PPTMaker.git $env:USERPROFILE\.claude\skills\pptmaker

# 2) Python 의존성 설치
cd $env:USERPROFILE\.claude\skills\pptmaker
python -m pip install -e .

# 3) 설치 확인 — 다음 명령이 에러 없이 동작해야 함
python -c "from pptmaker import TOKENS; print('OK', TOKENS.colors.accent1)"
```

기대 출력: `OK #156082`

---

## 사용법

### 1) Claude Code 시작
어느 폴더에서든 Claude Code를 실행합니다.

### 2) PPTMaker 호출
대화창에 다음 중 하나를 입력:
- `/pptmaker`
- "PPT 만들어줘"
- "주간 보고 PPT 만들어줘"
- "EndoRobotics 표준으로 IR 자료 만들어줘"

### 3) PowerPoint 데스크톱과 함께
PowerPoint가 켜져 있으면 AI가 실시간으로 슬라이드를 추가하는 것을 화면에서 직접 보실 수 있습니다. 중간에 수동으로 편집해도 됩니다.

---

## 디자인 시스템 요약

| 항목 | 사내 표준 |
|---|---|
| 슬라이드 크기 | 33.87 × 19.05 cm (16:9) |
| 폰트 | 맑은 고딕 (강제) |
| 메인 컬러 | accent1 `#156082` (틸 블루) |
| 다크 컬러 | `#0E2841` (네이비) |
| 14종 표준 레이아웃 | 표지·목차·섹션구분·1단·2단·KPI·표·차트·비교·인용·텍스트+이미지·풀이미지·프로세스·마무리 |

상세 규칙은 [`.claude/skills/pptmaker/SKILL.md`](.claude/skills/pptmaker/SKILL.md) 참조.

---

## 업데이트

새 디자인 가이드/레이아웃이 배포되면:

```powershell
cd $env:USERPROFILE\.claude\skills\pptmaker
git pull
python -m pip install -e . --upgrade
```

Claude Code를 재시작하면 변경사항이 자동 반영됩니다.

---

## 디자인 표준 변경 (관리자용)

회사 표준 디자인이 변경되었거나 새 레이아웃을 추가할 때:

1. `Reference/` 폴더에 새 .pptx 표준 템플릿 갱신
2. `python scripts/analyze_templates.py` → `design_tokens.json` 재생성
3. 변경 사항 git push
4. 전사 공지 (직원들은 `git pull`로 받음)

새 본문 레이아웃 추가:
1. `src/pptmaker/slides/` 에 새 모듈 작성 — 반드시 `_common.add_chrome_slide()` 진입점 사용
2. `Type/Role/Space/Layout` 스케일만 사용 (magic number 금지)
3. `LiveSession`에 편의 메서드 추가
4. `SKILL.md` 표 갱신

---

## 문제 발생 시

| 증상 | 해결 |
|---|---|
| `/pptmaker` 자동완성이 안 보임 | Claude Code 재시작. 그래도 안 되면 `~/.claude/skills/pptmaker/SKILL.md` 파일이 존재하는지 확인 |
| Python 임포트 에러 | `cd $env:USERPROFILE\.claude\skills\pptmaker; python -m pip install -e .` 재실행 |
| 차트 데이터 표가 이미 열려 있음 | PowerPoint에서 떠 있는 mini Excel 창을 수동으로 닫기 |
| 한글이 콘솔에 깨져 보임 | `$env:PYTHONIOENCODING="utf-8"` 설정 (PPT 안의 한글은 정상) |

추가 문의: 전략기획팀 이은상 (eunsang.lee@endorobo.com)

---

## 기여 / 개선 제안

사내 GitHub Issues에 등록하거나 전략기획팀에 직접 문의해주세요.
새 레이아웃·디자인 표준 갱신 PR도 환영합니다.

---

## 라이센스

EndoRobotics 사내용 (외부 배포 금지).
