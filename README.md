# PPTMaker — EndoRobotics 전사 PPT 빌더

> AI가 EndoRobotics 사내 표준 디자인으로 PPT를 만들어주는 Claude Code Skill.
> **그 누가 어디서 쓰든 동일한 디자인 가이드**가 자동 적용되고, 직원은 PowerPoint와 Claude Code 대화창을 **동시에 띄워두고 실시간으로 페어 작업**합니다.

---

## 목차

1. [한 줄 요약](#한-줄-요약)
2. [사전 요구사항](#-사전-요구사항)
3. [설치 (1회)](#-설치-1회)
4. [첫 PPT 만들기 — 단계별](#-첫-ppt-만들기--단계별)
5. [두 가지 테마](#-두-가지-테마)
6. [두 가지 시작 시나리오](#-두-가지-시작-시나리오)
7. [저장 동작 (안전장치)](#-저장-동작-안전장치)
8. [업데이트](#-업데이트)
9. [문제 해결](#-문제-해결)
10. [관리자용 (디자인 표준 변경)](#-관리자용-디자인-표준-변경)
11. [문의](#-문의)

---

## 한 줄 요약

PowerShell 한 줄로 끝:

```powershell
irm https://raw.githubusercontent.com/EndoRobotics-Co-LTD/endo-claude-skill-ppt-maker/lee-dev/install.ps1 | iex
```

→ Claude Code 재시작 → 대화창에 **"PPT 만들어줘"** 입력 → 끝.

> 💡 `irm | iex` 패턴은 부트스트랩 스크립트를 다운로드해서 바로 실행합니다. PowerShell 실행 정책 설정 변경 불필요. 내부적으로 올바른 위치(`~/.claude/skills/pptmaker/`)에 클론하고 의존성을 설치합니다.

---

## 📋 사전 요구사항

| 항목 | 버전/조건 |
|---|---|
| **OS** | Windows 10 또는 11 |
| **PowerPoint** | Microsoft 365 (데스크톱 버전 — 웹/모바일 불가) |
| **Python** | 3.11 이상 |
| **Git** | 설치되어 있어야 함 |
| **Claude Code** | 최신 버전 |
| **GitHub 접근권** | `EndoRobotics-Co-LTD/endo-claude-skill-ppt-maker` 리포 접근 권한 |

설치 전 PowerShell에서 확인:

```powershell
python --version    # 3.11 이상이어야 함 (Python 3.11.5 같은 출력)
git --version       # git version 2.x.x 같은 출력
```

⚠️ Python이 없거나 3.10 이하라면 https://python.org 에서 3.11+ 설치 후 진행.
⚠️ GitHub 접근권이 없으면 **전략기획팀 이은상**에게 요청.

---

## ⚡ 설치 (1회)

### 방법 A — 부트스트랩 인스톨러 (권장)

PowerShell을 열고 **한 줄**:

```powershell
irm https://raw.githubusercontent.com/EndoRobotics-Co-LTD/endo-claude-skill-ppt-maker/lee-dev/install.ps1 | iex
```

이 명령이 자동으로:
1. Python/Git 설치 여부 확인
2. 스킬을 **정확한 위치** (`~/.claude/skills/pptmaker/`) 에 클론 — 폴더 이름 자동 처리
3. Python 패키지 설치 (`python-pptx`, `pywin32`, `pptmaker`)
4. 설치 검증

마지막에 `전체 설치 완료` 메시지가 보이면 OK.

> ⚠️ **왜 `git clone <url>` 직접 안 됨?**
> 기본 `git clone`은 리포 이름(`endo-claude-skill-ppt-maker`) 그대로 폴더를 만드는데, Claude Code는 `~/.claude/skills/pptmaker/` 를 기대합니다. 이름이 다르면 스킬을 못 찾아요. `install.ps1`이 이 문제를 자동으로 해결합니다.

### 방법 B — 수동 (자동이 실패하면)

```powershell
# 1) 스킬 폴더에 클론 — 두 번째 인자에 목적지 명시 (중요!)
git clone https://github.com/EndoRobotics-Co-LTD/endo-claude-skill-ppt-maker.git $env:USERPROFILE\.claude\skills\pptmaker

# 2) Python 의존성 설치
cd $env:USERPROFILE\.claude\skills\pptmaker
python -m pip install -e .

# 3) 설치 검증
python -c "from pptmaker import TOKENS; print('OK', TOKENS.colors.accent1)"
```

기대 출력: `OK #156082`

> 💡 `git clone <url> <목적지>` 형식이 핵심. 목적지를 생략하면 `endo-claude-skill-ppt-maker/` 폴더가 만들어져서 Claude Code가 못 찾습니다.

### 설치 후 한 번

**Claude Code를 한 번 종료했다가 다시 실행하세요.** 그래야 새 스킬이 인식됩니다.

---

## 🎬 첫 PPT 만들기 — 단계별

### 1단계 — Claude Code 실행

아무 폴더에서 Claude Code를 실행하면 됩니다. PowerPoint를 미리 켜둘 필요는 없어요 (Claude가 알아서 켭니다).

### 2단계 — PPT 요청

대화창에 자연어로 요청:

```
PPT 만들어줘
```

또는 더 구체적으로:

```
EndoRobotics 표준으로 2026년 1분기 사업 보고 PPT 만들어줘
```

```
어제 만들던 ~/Desktop/2026Q1.pptx 에 KPI 카드 한 장 추가해줘
```

```
주간 보고 PPT 만들어줘. 이번주 매출 +18% 라는 게 핵심이야.
```

### 3단계 — Claude의 질문에 답하기

Claude가 순서대로 짧게 묻습니다:

**Q1. 새 PPT인가요, 기존 작업 이어가는 건가요?**
- "**새 PPT**" → 빈 사본을 만들어 시작
- "**~/Desktop/foo.pptx 이어서**" → 그 파일 열어서 추가 작업

**Q2. (새 PPT일 때) 어느 테마요?**
- "**theme1**" — Formal/Corporate (경영 보고, IR, 내부 보고서)
- "**theme2**" — Visual/Branded (외부 발표, 학회, 제품 마케팅)
- 잘 모르겠으면 **theme1** (기본값)

**Q3. 무슨 내용? 어떤 슬라이드들?**

답하다 보면 Claude가 슬라이드 개요를 먼저 제안할 거예요. OK 하면 본격 생성 시작.

### 4단계 — 실시간 페어 작업

이게 핵심 시나리오입니다.

1. **Claude가 PowerPoint를 자동으로 켭니다** (또는 이미 켜져있으면 거기 연결)
2. 슬라이드를 한 장씩 추가합니다
3. **사용자는 PowerPoint 화면에서 슬라이드가 실시간으로 생기는 걸 봅니다**
4. 중간에:
   - 사용자가 PowerPoint에서 **직접 편집해도 OK** (텍스트 수정, 도형 이동 등)
   - Claude에게 "이 슬라이드 다시 만들어줘", "한 장 더 추가" 등 자연어로 지시
   - "다음은 KPI 카드로 / 표로 / 비교 카드로" 같은 형식 지정 가능

### 5단계 — 저장

**Claude는 사용자가 "저장해줘"라고 말할 때까지 자동 저장하지 않습니다.**

저장 옵션:
- "**저장해줘**" → 기본 위치 (`~/Documents/PPTMaker/<이름>_<날짜시간>.pptx`)
- "**~/Desktop/2026Q1.pptx 로 저장해줘**" → 명시 경로
- **사용자가 PowerPoint에서 직접 Ctrl+S** → 안전 (사본만 영향, 템플릿 보호)

---

## 🎨 두 가지 테마

| 테마 | 비주얼 특성 | 적합한 용도 |
|---|---|---|
| **theme1** | 네이비 헤더 바 + 흰색 Bold 타이틀, 흰 본문 배경 | 경영 보고, IR 자료, 내부 보고서, 주간/월간 리포트 |
| **theme2** | 의료기기 사진 풀-블리드 표지/마무리, 사진 띠 헤더, 흰색 Bold 타이틀 | 외부 발표, 학회 자료, 제품 마케팅, 데모 자료 |

**두 테마 모두 14종 본문 레이아웃**을 동일하게 지원:

```
표지 · 목차 · 섹션 구분 · 1단 · 2단 · KPI 카드 · 표 · 차트
비교 · 인용 메시지 · 텍스트+이미지 · 풀이미지 · 프로세스 · 마무리
```

### 디자인 시스템 요약

| 항목 | 사내 표준 |
|---|---|
| 슬라이드 크기 | 33.87 × 19.05 cm (16:9) |
| 폰트 | 맑은 고딕 (강제) |
| 메인 컬러 | accent1 `#156082` (틸 블루) |
| 다크 컬러 | `#0E2841` (네이비) |
| 14종 레이아웃 | 모두 코드+템플릿에 강제됨 — 직원이 임의로 깰 수 없음 |

상세 규칙: [`SKILL.md`](SKILL.md)

---

## 🚀 두 가지 시작 시나리오

### 시나리오 A — 새 PPT 시작 (가장 흔함)

```
사용자: "EndoRobotics 표준으로 1분기 보고 PPT 만들어줘"
Claude: "새 PPT인가요? 어느 테마(theme1/theme2)요?"
사용자: "새 PPT, theme1"
Claude: "어디에 저장할까요? (기본: ~/Documents/PPTMaker/)"
사용자: "기본 위치 좋아"
→ Claude가 layouts/theme1.pptx 를 사본으로 복사 → PowerPoint로 사본을 열음
→ 슬라이드 생성 시작
```

**핵심 안전장치**: 사본을 만들어 작업하므로 **layouts/ 원본 템플릿은 절대 안 건드립니다.**

### 시나리오 B — 기존 파일 이어가기

```
사용자: "어제 만들던 ~/Desktop/2026Q1.pptx 에 KPI 카드 추가해줘"
Claude: (그 파일을 PowerPoint로 열고 끝에 슬라이드 append)
```

기존 슬라이드는 그대로 보존, 추가만 합니다.

---

## 💾 저장 동작 (안전장치)

| 시나리오 | 결과 |
|---|---|
| 새 PPT 시작 | `layouts/theme*.pptx` 사본 자동 생성 → 사본을 PowerPoint로 열음 |
| Claude가 슬라이드 추가 중 | 사본에만 변경 적용. **자동 저장 안 함** |
| 사용자가 PowerPoint에서 Ctrl+S | **사본만 저장**. 템플릿 영향 없음 ✓ |
| 사용자가 "저장해줘" | 사본을 명시 경로 또는 기본 위치에 저장 |
| 작업 사본 경로가 `layouts/` 안이면? | `ValueError`로 **차단** (이중 안전장치) |

**핵심**: `layouts/theme1.pptx`, `layouts/theme2.pptx` 는 SSOT(Single Source of Truth)이고 **절대 직접 수정되지 않습니다.**

---

## 🔄 업데이트

새 디자인 가이드 / 새 레이아웃 / 버그 수정이 배포되면:

```powershell
cd $env:USERPROFILE\.claude\skills\pptmaker
git pull
python -m pip install -e . --upgrade
```

Claude Code 재시작 → 변경사항 자동 반영.

새 디자인 가이드 배포 안내는 **전사 공지 또는 GitHub Releases** 에서 확인하세요.

---

## 🛠 문제 해결

| 증상 | 해결 |
|---|---|
| `/pptmaker` 자동완성이 안 보임 | Claude Code 완전 종료 후 재시작. 그래도 안 되면 `~/.claude/skills/pptmaker/SKILL.md` 파일 존재 확인. 폴더 이름이 `endo-claude-skill-ppt-maker` 같이 되어있으면 `pptmaker` 로 rename 하거나 install.ps1로 재설치 |
| Python 임포트 에러 (`ModuleNotFoundError: pptmaker`) | `cd $env:USERPROFILE\.claude\skills\pptmaker; python -m pip install -e .` 재실행 |
| PowerPoint COM 에러 | PowerPoint 데스크톱앱을 한 번 직접 실행해 정상 동작 확인 후 재시도. Microsoft 365 라이선스 활성화 필요 |
| 차트 데이터 표(mini Excel)가 안 닫힘 | 수동으로 닫기. 다음 슬라이드는 정상 작동 |
| 한글이 콘솔에 깨져 보임 | `$env:PYTHONIOENCODING="utf-8"` 설정. PPT 안의 한글은 항상 정상 |
| `ValueError: layouts/ 템플릿 영역의 파일은 직접 작업할 수 없습니다` | **정상 동작**. 템플릿 보호 가드가 작동한 것. 다른 경로(예: `~/Desktop/...`)로 working_path 지정 |
| `LiveSession은 Windows + PowerPoint 환경에서만 사용 가능` | macOS/Linux는 미지원. Windows + Microsoft 365 PowerPoint 필요 |
| git clone 시 `Authentication failed` | GitHub 접근권 없음. 전략기획팀 이은상에게 권한 요청 |

---

## 🔧 관리자용 (디자인 표준 변경)

### 디자인 가이드 갱신

회사 표준 디자인이 변경됐을 때:

1. `layouts/theme1.pptx` 또는 `layouts/theme2.pptx` 를 직접 편집
2. `python scripts/analyze_templates.py` 실행 → `design_tokens.json` 자동 갱신
3. 변경사항 검증 (시각적 QA)
4. Git에 push → 전사 공지 → 직원들은 `git pull` 로 받음

### 새 본문 레이아웃 추가

1. `src/pptmaker/slides/` 에 새 모듈 작성
2. 반드시 `_common.add_chrome_slide()` 진입점 사용 (chrome 자동 상속)
3. `Type / Role / Space / Layout` 스케일만 사용 (magic number 금지)
4. `LiveSession` 에 편의 메서드 (`add_xxx`) 추가
5. `SKILL.md` 표 갱신

상세: `SKILL.md` 의 §7 "새 레이아웃 추가 시 체크리스트"

---

## 📞 문의

| 종류 | 연락처 |
|---|---|
| 사용 문의 / 버그 리포트 | 전략기획팀 이은상 (eunsang.lee@endorobo.com) |
| 디자인 표준 변경 요청 | GitHub Issues 또는 전략기획팀 |
| 새 레이아웃 / 기능 제안 | GitHub PR 또는 전략기획팀 |

---

## 📜 라이센스

EndoRobotics 사내용 — 외부 배포 금지.
