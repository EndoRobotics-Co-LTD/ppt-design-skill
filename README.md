# endo-skills — EndoRobotics Claude Code Skills

> EndoRobotics 사내 표준 업무를 Claude Code 에 박아넣은 **공식 Skill 카탈로그**.
> 직원 누구나 자기 PC에서 동일한 사내 표준으로 작업할 수 있도록, 디자인 가이드·룰·워크플로우를 코드로 박아둔 스킬들을 모아두는 곳.

[![License: Internal](https://img.shields.io/badge/license-Internal-blue.svg)](#-라이센스)

---

## 🎯 이 리포가 뭐예요?

**Claude Code Skill** 은 Claude 에게 "특정 업무를 사내 표준대로 처리하는 법" 을 박아두는 패키지입니다. 한 번 설치하면 그 PC의 Claude 가 해당 업무를 일관된 방식으로 처리합니다.

`endo-skills` 는 EndoRobotics 가 사용하는 **모든 사내 표준 Skill 들의 단일 카탈로그**입니다. 지금은 PPT 빌더 한 개로 시작하지만, 향후 ERP 연동, 보고서 자동화, 데이터 정리 등 회사가 자주 하는 업무를 하나씩 추가해 나갑니다.

**핵심 가치**:
- ✅ 누가 쓰든 동일한 사내 표준 결과물
- ✅ 직원은 자기 자리에서 자연어로 요청만
- ✅ 디자인·룰 변경은 PR 한 번으로 전사 반영
- ✅ 코드에 박혀있어 직원이 임의로 깰 수 없음

---

## 📚 Skill 카탈로그

| Skill | 설명 | 상태 | 링크 |
|---|---|---|---|
| **ppt-brand-design** | EndoRobotics 표준 PPT 빌더. 14종 레이아웃, 듀얼 테마(theme1/theme2), PowerPoint 데스크톱과 실시간 페어 워크플로우 | ✅ v0.1.0 | [`skills/ppt-brand-design/`](skills/ppt-brand-design/) |
| *(추후 추가 예정)* | 회사가 자주 하는 다른 업무들 | 🔜 | — |

각 스킬의 상세 사용법은 해당 폴더의 `README.md` 참조.

---

## 🚀 0-to-1 가이드 — 처음부터 끝까지

### Step 0 — 환경 확인

| 항목 | 버전/조건 | 확인 명령 |
|---|---|---|
| **OS** | Windows 10 또는 11 | `winver` |
| **Python** | 3.11 이상 | `python --version` |
| **Git** | 설치되어 있어야 함 | `git --version` |
| **Node.js** | 18+ (npx skills 사용 시) | `node --version` |
| **PowerPoint** | Microsoft 365 (ppt-brand-design 사용 시) | — |

⚠️ 누락된 게 있으면:
- Python: https://python.org → 3.11+ 설치
- Git: https://git-scm.com/download/win
- Node.js: https://nodejs.org/

### Step 1 — Claude Code 설치

아직 안 깔려있다면:

```bash
npm install -g @anthropic-ai/claude-code
```

설치 후 PowerShell 한 번 닫았다 다시 열기.

### Step 2 — 원하는 Skill 설치

각 스킬은 독립적으로 설치합니다. 가장 표준적인 방법은 `npx skills add`:

```bash
npx skills add EndoRobotics-Co-LTD/endo-skills -s <skill-name>
```

예시 — `ppt-brand-design` 설치:

```bash
npx skills add EndoRobotics-Co-LTD/endo-skills -s ppt-brand-design
```

이 명령은:
- 현재 폴더의 `.claude/skills/<skill-name>/` 에 스킬을 다운로드 (project-local)
- 글로벌 설치는 `-g` 추가: `npx skills add ... -g` → `~/.claude/skills/<skill-name>/`

⚠️ **Python 의존성이 있는 스킬은 추가 설치 필요**. 각 스킬의 README 에 안내. `ppt-brand-design` 의 경우:

```powershell
cd $env:USERPROFILE\.claude\skills\ppt-brand-design
python -m pip install -e .
```

**또는** Python 의존성도 한 방에 설치하고 싶다면 스킬별 PowerShell 부트스트랩 사용 (예: `ppt-brand-design`):

```powershell
irm https://raw.githubusercontent.com/EndoRobotics-Co-LTD/endo-skills/main/skills/ppt-brand-design/install.ps1 | iex
```

### Step 3 — Claude Code 재시작

설치 후 Claude Code 를 **완전히 종료한 뒤 다시 실행**. 그래야 새 스킬이 인식됩니다.

### Step 4 — 스킬 사용 (자연어로 요청)

Claude Code 대화창에 평소처럼 요청하면, 관련된 스킬이 자동으로 발동됩니다.

`ppt-brand-design` 예시:

```
PPT 만들어줘
```

```
EndoRobotics 표준으로 1분기 사업 보고 PPT 만들어줘
```

```
어제 만든 ~/Desktop/2026Q1.pptx 에 KPI 카드 한 장 추가해줘
```

→ Claude 가 사내 디자인 표준으로 PPT 를 생성/수정합니다.

---

## 🔄 업데이트

새 스킬이 추가되거나 기존 스킬이 갱신되면 **재설치** 로 받습니다:

```powershell
# 1) 기존 스킬 제거
Remove-Item -Recurse -Force $env:USERPROFILE\.claude\skills\<skill-name>

# 2) 다시 설치
npx skills add EndoRobotics-Co-LTD/endo-skills -s <skill-name>
# 또는 (Python 의존성 자동)
irm https://raw.githubusercontent.com/EndoRobotics-Co-LTD/endo-skills/main/skills/<skill-name>/install.ps1 | iex
```

배포 안내는 **사내 공지** 또는 [GitHub Releases](https://github.com/EndoRobotics-Co-LTD/endo-skills/releases) 에서 확인.

---

## 🗂 리포 구조

```
endo-skills/
├── README.md                  ← 지금 이 파일 (전체 스킬 카탈로그)
└── skills/
    └── ppt-brand-design/      ← PPT 빌더 스킬
        ├── SKILL.md           ← Claude Code 가 자동 인식하는 진입점
        ├── README.md          ← 스킬별 상세 사용법
        ├── install.ps1        ← 스킬 부트스트랩 (Python deps 포함)
        ├── setup.ps1          ← Python 패키지 설치 + 검증
        ├── src/, layouts/, scripts/, tests/
        └── pyproject.toml
```

각 스킬은 **자기 자신의 의존성·설치 스크립트·README 를 자기 폴더 안에 모두 포함**합니다. 그래서 어느 스킬을 설치하든 다른 스킬과 충돌하지 않고 독립적으로 동작합니다.

---

## 🛠 새 스킬 추가 (관리자용)

회사에서 새로운 사내 표준 업무를 Claude 에 박아넣고 싶을 때:

1. `skills/<new-skill-name>/` 폴더 생성
2. 최소 구성:
   - `SKILL.md` — 진입점 (frontmatter `name`, `description` 필수)
   - `README.md` — 사용자용 상세 가이드
   - 필요 시 `install.ps1`, `setup.ps1`, 코드, 리소스
3. 본 리포 README 의 **Skill 카탈로그** 표에 항목 추가
4. PR → main 머지 → 전사 배포

상세 가이드는 향후 별도 `CONTRIBUTING.md` 로 정리 예정. 그 전까지는 `ppt-brand-design/` 구조를 참고.

> 💡 [Claude Code Skills 공식 가이드라인](https://docs.claude.com/en/docs/claude-code/skills) 과 [vercel-labs/skills](https://github.com/vercel-labs/skills) 의 패턴을 따릅니다.

---

## ❓ 문의

| 종류 | 연락처 |
|---|---|
| 사용 문의 / 버그 리포트 | 전략기획팀 이은상 (eunsang.lee@endorobo.com) |
| 새 스킬 제안 / 기여 | GitHub Issues 또는 PR |
| 기존 스킬 변경 요청 | 해당 스킬 폴더의 Issue 또는 전략기획팀 |

---

## 📜 라이센스

EndoRobotics 사내용 — 외부 배포 금지. 본 리포에 포함된 디자인 가이드, 템플릿, 코드는 EndoRobotics 의 자산입니다.
