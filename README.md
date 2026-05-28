# endo-skills — EndoRobotics Claude Code Skills

> EndoRobotics 사내 표준 업무를 Claude Code 에 박아넣은 **공식 Skill 카탈로그**.
> 직원은 자기 PC의 Claude 에게 자연어로 요청만 하면 됩니다. 사내 표준은 코드와 SKILL.md 에 박혀있어, 누가 쓰든 동일한 결과물이 나옵니다.

[![License: Internal](https://img.shields.io/badge/license-Internal-blue.svg)](#-라이센스)

---

## 🎯 이 리포가 뭐예요?

**Claude Code Skill** = Claude 에게 특정 업무를 사내 표준대로 처리하는 법을 알려주는 패키지.
**`endo-skills`** = EndoRobotics 의 모든 사내 표준 Skill 들이 모여있는 단일 카탈로그.

직원은 이 카탈로그에서 필요한 스킬을 자기 Claude 프로젝트에 **추가(import)** 만 하면 됩니다. 추가 후 자연어로 요청하면 Claude 가 알아서 처리합니다. 설치 도중·이후에 추가 환경 세팅(파이썬 패키지 설치 등)이 필요하면 **스킬과 Claude 가 같이 풀어줍니다** — 사용자가 따로 매뉴얼을 읽지 않아도 OK.

**핵심 가치**:
- ✅ 누가 쓰든 동일한 사내 표준 결과물
- ✅ 직원은 자기 자리에서 자연어로 요청만
- ✅ 사내 표준 변경은 PR 한 번으로 전사 반영

---

## 📚 Skill 카탈로그

| Skill | 한 줄 설명 | 상태 | 상세 |
|---|---|---|---|
| **`ppt-brand-design`** | EndoRobotics 표준 PPT 빌더. 14종 레이아웃, 듀얼 테마(theme1/theme2), PowerPoint 데스크톱과 실시간 페어 워크플로우 | ✅ v0.1.0 | [`skills/ppt-brand-design/`](skills/ppt-brand-design/) |
| *(추후 추가)* | 회사가 자주 하는 다른 업무들 — 보고서, ERP 연동, 데이터 정리 등 | 🔜 | — |

각 스킬의 사용법·요구사항·문제 해결은 해당 폴더의 `README.md` 에서 확인하세요.

---

## 🚀 스킬을 내 프로젝트에 import 하는 법

```bash
npx skills add EndoRobotics-Co-LTD/endo-skills -s <skill-name>
```

이 한 줄이 endo-skills 리포에서 해당 스킬 폴더 하나만 추출해, 호출한 위치의 Claude 프로젝트에 박아넣습니다.

### 예시 — `ppt-brand-design` 추가

```bash
# 1) Claude Code 를 실행할 프로젝트 폴더로 이동
cd ~/my-project

# 2) 카탈로그에서 ppt-brand-design 만 골라 import
npx skills add EndoRobotics-Co-LTD/endo-skills -s ppt-brand-design
```

→ `./.claude/skills/ppt-brand-design/` 에 스킬이 들어옵니다 (project-local).

### 설치 옵션

| 옵션 | 동작 | 언제 |
|---|---|---|
| 기본 | 현재 폴더의 `.claude/skills/<skill>/` 에 설치 (project-local) | 특정 프로젝트에서만 쓸 때 |
| `-g` 추가 | 글로벌 `~/.claude/skills/<skill>/` 에 설치 | 모든 프로젝트에서 쓸 때 |
| 여러 스킬 한 번에 | `-s a -s b` 처럼 반복 | 여러 개 동시 import |

설치 후 **Claude Code 를 한 번 종료했다가 다시 실행**해야 새 스킬이 인식됩니다.

### 그 다음은?

Claude Code 대화창에서 평소처럼 요청만 하면 됩니다. 스킬이 자동 발동되어 사내 표준대로 처리합니다.

```
PPT 만들어줘
```

```
1분기 사업 보고 슬라이드 한 장 추가해줘
```

스킬이 처음 호출될 때 **추가 세팅이 필요하면 (예: 파이썬 패키지 설치) Claude 가 직접 안내**합니다. 별도의 매뉴얼을 미리 읽을 필요 없음.

---

## 🔄 업데이트 — 최신 버전 받기

스킬은 시간이 지나며 갱신됩니다 (디자인 표준 변경, 새 레이아웃, 버그 수정 등). 같은 명령어로 재설치하면 최신본을 받습니다.

```bash
# 기존 설치 제거 후 다시 import
rm -rf ./.claude/skills/<skill-name>            # 또는 ~/.claude/skills/<skill-name>
npx skills add EndoRobotics-Co-LTD/endo-skills -s <skill-name>
```

배포 안내는 **사내 공지** 또는 [GitHub Releases](https://github.com/EndoRobotics-Co-LTD/endo-skills/releases) 에서 확인.

---

## 🗂 리포 구조 (한눈에)

```
endo-skills/
├── README.md                  ← 지금 이 파일 (카탈로그 + import 가이드)
└── skills/
    └── <skill-name>/          ← 각 스킬은 자기 폴더 안에 자급자족
        ├── SKILL.md           ← Claude Code 가 인식하는 진입점
        ├── README.md          ← 스킬별 상세 (요구사항·설치·트러블슈팅)
        └── ...                ← 스킬 코드, 리소스, 설치 스크립트 등
```

**원칙**: 각 스킬은 자기 의존성·설치 스크립트·문서·코드를 모두 자기 폴더 안에 포함합니다. 사용자가 어떤 스킬을 import 하든 다른 스킬과 충돌하지 않고 독립적으로 동작합니다.

---

## 🛠 새 스킬 추가하기 (기여자용)

회사에서 새로운 사내 표준 업무를 Claude 에 박아넣고 싶을 때:

1. `skills/<new-skill-name>/` 폴더 신규 생성
2. 최소 구성:
   - `SKILL.md` — 진입점 (frontmatter `name`, `description` 필수)
   - `README.md` — 사용자용 상세 가이드 (요구사항·설치·사용법·문제 해결)
   - 필요 시 설치 스크립트(`install.ps1` 등), 코드, 리소스
3. 본 리포 README 의 **Skill 카탈로그** 표에 한 줄 추가
4. PR → main 머지 → 전사 배포

상세 가이드는 향후 `CONTRIBUTING.md` 로 정리 예정. 그 전까지는 `ppt-brand-design/` 구조를 참고.

> 💡 [Claude Code Skills 공식 가이드라인](https://docs.claude.com/en/docs/claude-code/skills) 과 [vercel-labs/skills](https://github.com/vercel-labs/skills) 의 패턴을 따릅니다.

---

## ❓ 문의

| 종류 | 연락처 |
|---|---|
| 카탈로그·import 관련 일반 문의 | 전략기획팀 이은상 (eunsang.lee@endorobo.com) |
| 특정 스킬의 사용 문의·버그 | 해당 스킬 폴더의 README 또는 GitHub Issues |
| 새 스킬 제안·기여 | GitHub PR |

---

## 📜 라이센스

EndoRobotics 사내용 — 외부 배포 금지. 본 리포에 포함된 모든 스킬, 디자인 가이드, 템플릿, 코드는 EndoRobotics 의 자산입니다.
