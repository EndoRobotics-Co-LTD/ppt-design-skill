---
name: ppt-brand-design
description: EndoRobotics 전사 표준 PPT 빌더 — 사내 디자인 표준(맑은 고딕, accent1 #156082, 16:9, theme1/theme2 chrome)을 강제하며 14종 레이아웃으로 PPT를 대화형 생성·수정한다. 사용자가 PPT·슬라이드·덱(deck)·발표자료·IR·주간보고·경영보고·사외발표를 만들거나, "한 장 더 추가해줘"·"이 부분 바꿔줘"·"슬라이드 순서 바꿔줘"·"새 레이아웃 등록" 같은 PPT 수정·확장 요청을 할 때 반드시 호출. 사용자가 EndoRobotics 표준을 명시하지 않아도 일단 진입해 적용 여부를 묻는다. Windows + PowerPoint 데스크톱 + Claude Code 대화창 실시간 페어 워크플로우.
---

# PPTMaker — EndoRobotics 전사 PPT 빌더

사내 디자인 표준을 자동 강제하면서 PPT를 대화형으로 빌드·수정한다. PowerPoint 데스크톱과 Claude Code 대화창을 동시에 띄워두고 슬라이드가 실시간으로 추가되는 것을 사용자가 화면에서 직접 보는 게 핵심 시나리오.

상세 사용법·설치·트러블슈팅은 [README.md](README.md). 디자인 토큰(컬러·타이포·간격) 상세 표는 [`src/pptmaker/slides/_common.py`](src/pptmaker/slides/_common.py)의 `Type`/`Role`/`Space`/`Layout` 클래스 docstring에 있다.

**기본 모드**: `pptmaker.live_session.LiveSession` (라이브, PowerPoint COM).
**오프라인 모드**: 사용자가 "파일만 받을게" 식으로 명시하면 `pptmaker.offline_builder.OfflineBuilder` (python-pptx).

---

## 0. 첫 호출 — 의존성 자가 점검 (반드시 가장 먼저, 세션당 1회)

`endo-skills` 루트 README 는 import 가이드만 다룬다 (사용자가 매뉴얼 미리 안 읽어도 되게). 환경 셋업은 **이 스킬이 첫 호출 시점에 자체 해결**하는 게 설계다. 건너뛰면 사용자가 의미 없는 ImportError 만 보고 멈춘다.

세션당 1회만 점검 — 같은 세션에서 통과 후엔 반복 안 함.

### 절차

1. **Import 테스트**:
   ```bash
   python -c "from pptmaker import TOKENS; print('OK', TOKENS.colors.accent1)"
   ```
   `OK #156082` 나오면 §1로 진행. 실패면 자가 복구.

2. **자가 복구** — 사용자에게 한 줄 보고 후 동의받고 진행:
   > "이 스킬 처음 쓰시는 것 같아요. Python 패키지 자동 설치할게요 (1분 내외). 진행할까요?"

   순차 시도:
   - **a.** `& "<skill-folder>/setup.ps1"` (표준 경로 — pip install + 검증 한 번에). `<skill-folder>` = 이 SKILL.md 들어있는 폴더 (`~/.claude/skills/ppt-brand-design/` 또는 `./.claude/skills/ppt-brand-design/`).
   - **b.** setup.ps1 없거나 실패 시: `cd "<skill-folder>"; python -m pip install -e .`
   - **c.** Python 자체가 없거나 3.10 이하면 사용자에게 "Python 3.11+ 필요. https://python.org 설치 후 PowerShell 다시 열고 같은 요청" 안내.

   복구 끝나면 §1로 진행 ("설치 끝, 이어서 PPT 작업 시작합니다" 한 줄).

3. **PowerPoint 데스크톱**: `LiveSession.__init__` 안에서 COM 연결 시도하므로 사전 점검 불필요. 실패 시 에러 메시지에 "PowerPoint 데스크톱(Microsoft 365)" 키워드 그대로 사용자에게 전달 + "PowerPoint 한 번 직접 실행해 정상 동작 확인 후 재시도" 안내.

---

## 1. 대화 흐름

사용자에게 한 번에 다 묻지 말고 짧게 끊어 진행. 각 단계는 사용자 확인 후 다음으로.

### 1.1 스킬 적용 여부 게이트 (반드시 가장 먼저)

다른 어떤 질문보다 먼저 묻는다:

> "EndoRobotics PPT 스킬(사내 디자인 표준)을 적용할까요?"

- **Yes (적용)** → 이 스킬 활성화, 아래 진행. **사내 디자인 표준(theme1/theme2 chrome, 14종 레이아웃, 맑은 고딕, 강제 색상)을 100% 강제**. 사용자가 폰트·색상·레이아웃을 깨려 해도 거부하고 표준으로 안내. (사외 발표·IR·내부 공식 보고 등 공식 자료에 적합)
- **No (적용 안 함)** → 이 스킬은 **사용하지 않는다**. 사용자에게 "EndoRobotics 디자인 가이드라인 우회. 일반 PPT 작성으로 진행"이라고 명확히 안내한 뒤 일반 모드로 전환 (Claude 자체 판단 또는 `anthropic-skills:pptx`). (개인 작업·임시 자료·연습용 등에 적합)

게이트의 이유: 비공식·임시 자료에까지 표준이 강제돼 사용자가 답답해지는 것을 방지하면서, 공식 자료의 디자인 일관성은 무조건 보장하기 위함.

### 1.2 시나리오 분기 — 새 PPT vs 후속 수정

- **A. 새 PPT** → `LiveSession.new(theme=, working_path=, name_hint=)`
  - 템플릿을 새 경로로 자동 복사하고 그 사본을 PowerPoint로 연다. `working_path` 생략 시 `~/Documents/PPTMaker/<이름>_<timestamp>.pptx` 자동 생성.
  - 이 호출은 **layouts/ 의 템플릿 원본을 절대 안 건드림**. 안전 가드: 같은 working_path가 이미 존재하면 `FileExistsError`, PowerPoint에 이미 열려있으면 `RuntimeError`로 차단한다. silent overwrite 사고 방지.
- **B. 후속 수정·추가 요청 → 무조건 `LiveSession.open_existing(path)`**
  - 사용자가 이미 만든 PPT에 대해 "이 부분 바꿔줘"·"한 장 더 추가"·"순서 바꿔줘" 같은 요청을 하면 **절대 `new()`를 부르지 않는다**. `new()`는 처음부터 재생성이라 사용자가 PowerPoint에서 수동 편집한 내용을 모두 날린다. `open_existing`은 기존 슬라이드를 보존한 채 부분 수정·추가만 적용한다.
  - 사용자가 명시적으로 "처음부터 다시 만들어줘"라고 할 때만 `new(force_overwrite=True)`.

### 1.3 테마 선택 (시나리오 A에서만)

- **theme1** — 네이비 헤더 바 + 흰색 Bold 타이틀. Formal/Corporate. (경영 보고, IR, 내부 보고서)
- **theme2** — 의료기기 사진 풀-블리드 표지 + 사진 띠 헤더 + 흰색 Bold 타이틀. Visual/Branded. (외부 발표, 학회, 제품 마케팅)
- 명시 없으면 한 번 물어보고 진행. 기본값 `theme1`.

### 1.4 목적·내용·개요

PPT 유형(경영 보고·주간·교육 등), 핵심 메시지, 슬라이드 개요를 수집한 뒤 AI가 슬라이드 타이틀 리스트를 먼저 제시해 사용자 승인을 받는다.

**사용자 커스텀 레이아웃 우선 활용**: 개요 작성 시 `session.list_custom_templates(theme=...)`로 사용자 등록 레이아웃을 먼저 확인. 사용자가 "타임라인 슬라이드"라고 하면 표준 process_flow 대신 등록된 `timeline_4step` 같은 커스텀을 우선 쓴다.

### 1.5 순차 생성 + 시각 fit 확인

한 번에 1~3장씩만 만들어 사용자가 PowerPoint 화면에서 확인할 시간을 둔다. 중간에 사용자가 PowerPoint에서 직접 편집해도 OK — Claude는 그 위에서 다음 작업을 이어간다.

매 `add_xxx()` 호출 끝에 자동으로 시각 health check + auto fit이 적용된다 (텍스트 오버플로우 시 박스 확장 → 폰트 한 단계 축소). 적용된 fix는 콘솔에 로그된다.

**자동 fit으로 해결 안 되는 issue가 남으면 반드시 사용자에게 옵션 제시**:
- (A) 텍스트 축약 또는 요약
- (B) 다음 슬라이드로 분할
- (C) 디자인 룰 한 단계 더 완화 (CAPTION 사이즈로 더 축소)
- (D) 그대로 두기

확인 방법: `session._last_remaining_issues` 가 비어있지 않으면 위 안내. silent로 넘어가지 말 것.

### 1.6 저장

사용자가 "저장해줘"라고 **명시할 때까지** 절대 자동 저장하지 않는다. 시나리오 A는 시작 시점에 이미 작업 경로 사본이 만들어진 상태라 사용자가 PowerPoint에서 Ctrl+S를 눌러도 안전 (템플릿은 안 건드림).

저장 옵션:
- "저장해줘" → `session.save()` (현재 working_path)
- "~/Desktop/X.pptx 로 저장" → `session.save_as(path)`

---

## 2. 핵심 룰 (위반 금지)

이 룰들은 단순 행동 지침이 아니라 사용자 작업 안전과 사내 표준 일관성을 보장하는 안전장치다. 깨면 직원 작업물이 손상되거나 표준이 무너진다.

1. **`layouts/` 원본 절대 안 건드림** — 사내 디자인 SSOT. 모든 작업은 `LiveSession.new()`가 만든 사본에서. 안전 가드(`FileExistsError`/`RuntimeError`)를 우회하지 말 것. 안전 가드가 발동하면 사용자에게 `open_existing()` 사용을 안내.

2. **콘텐츠 SSOT는 `.pptx` 파일 자체.** `.py` 스크립트에 "전체 콘텐츠 정의"를 보관하지 않는다. 매 수정 요청 시 임시 작업 스크립트를 만들어 실행하고 끝나면 정리 — 작업 폴더에 `.py` 누적 금지. `.py` 콘텐츠 보관 시 `.pptx`와 sync 깨져 stale 위험.

3. **디자인 토큰 강제** — `_common.py`의 `Type`/`Role`/`Space`/`Layout` 토큰만 사용. 매직 넘버 금지. 폰트는 맑은 고딕(`FONT_MAJOR`/`FONT_MINOR`)만. 컬러는 `Role.*` 경유, HEX 직접 참조 금지. 토큰을 깨면 다른 슬라이드와 시각 일관성이 무너진다.

4. **Chrome 직접 그리기 금지** — 헤더 바·로고·풋바 등 chrome은 `layouts/theme*.pptx`의 슬라이드 마스터에서 자동 상속된다. 모든 본문 슬라이드는 `add_chrome_slide()` 진입점 또는 슬라이드 복제 방식으로 생성. chrome을 손으로 그리면 테마 갱신 시 자동 반영이 깨진다.

5. **시각 fit 우선 (디자인 vs 컨텐츠 충돌 시)** — Type 스케일을 일시적으로 한 단계 깨더라도 텍스트가 보이게 한다. 자동 fit이 처리하고, 남는 issue는 §1.5의 4개 옵션으로 사용자에게 묻는다. silent 변경 금지 — 로그/콘솔에 명시.

6. **한국어 우선** — 모든 안내·확인 메시지는 한국어. PPT 내부 텍스트도 사용자가 다른 언어를 명시하지 않으면 한국어 기본.

---

## 3. 사용 가능한 14종 레이아웃

모두 `LiveSession`의 편의 메서드로 호출. 메서드 시그니처 상세는 `live_session.py` 참조.

| # | 레이아웃 | 호출 메서드 |
|---|---|---|
| - | **Cover** | `add_cover(title, presenter=, organization=, date=, subtitle=)` |
| A1 | Agenda (목차) | `add_agenda(items, title="목차")` |
| A2 | Section Divider | `add_section_divider(title, number=, subtitle=)` |
| A3 | Single Column | `add_single_column(title, bullets=, tiles=, section_no=)` |
| A4 | Two Column | `add_two_column(title, left_sub, left_items, right_sub, right_items, use_tiles=True, section_no=)` |
| A5 | Highlight Quote | `add_highlight_quote(message, attribution=, section_no=)` |
| B1 | KPI Cards | `add_kpi_cards(title, cards, section_no=)` — 2~4개 카드 |
| B2 | Data Table | `add_data_table(title, rows, section_no=)` |
| B3 | Chart | `add_chart(title, categories, series, chart_type="column", section_no=)` |
| B4 | Comparison | `add_comparison(title, left, right, arrow_label=, section_no=)` — Before/After |
| C1 | Text + Image | `add_text_image(title, image_path, subtitle, bullets, image_side="left", section_no=)` |
| C2 | Full Image | `add_full_image(image_path, caption=, title=, section_no=)` |
| C3 | Process Flow | `add_process_flow(title, steps, section_no=)` — 2~6단계 |
| - | **Closing** | `add_closing(message="Thank you", ...)` |

`section_no=N`을 주면 타이틀에 "NN. " prefix가 자동 부여된다 (예: "04. 핵심 성과").

`single_column` / `two_column`의 `tiles=` 옵션: Reference 표준 타일 카드 (번호 + Bold 제목 + 설명 + 좌측 컬러 보더). `bullets=`는 옛 불릿 리스트.

---

## 4. 부분 수정 메서드 (open_existing 후 사용)

이미 만든 PPT의 일부만 손볼 때. 사용자가 PowerPoint에서 수동 편집한 내용도 그대로 보존된다.

| 메서드 | 용도 |
|---|---|
| `replace_text_in_slide(idx, find, replace)` | 슬라이드 N의 텍스트 부분 치환 |
| `replace_text_global(find, replace)` | 전 슬라이드 텍스트 치환 |
| `delete_slide(idx)` | 슬라이드 삭제 |
| `move_slide(from_idx, to_idx)` | 슬라이드 순서 변경 |
| `get_slide_text(idx)` | 슬라이드 본문 미리보기 (어떤 텍스트가 있는지 확인용) |
| `add_xxx(...)` | 새 슬라이드 추가 — closing 직전에 삽입 (기본값) |

---

## 5. 사용자 커스텀 레이아웃 (PC 단위 fork)

각 사용자가 사내 표준 외 자기 자주 쓰는 레이아웃을 추가할 수 있다. 저장 위치는 `~/.claude/skills/ppt-design-skill/user_layouts/<theme>_user.pptx` (`.gitignore` 포함, git pull 안전).

**활용 (이미 등록된 레이아웃)**:
- `session.list_custom_templates(theme=)` — 등록된 템플릿 목록
- `session.add_custom(name, mapping)` — 등록된 시안을 복제 + `{{key}}` 자동 치환

**대화형 등록 (Phase 2 — 새 레이아웃 추가 요청 시)**:
1. 이름·테마·텍스트 필드를 사용자와 합의한다.
2. 현재 작업 중인 PPT에 시안 슬라이드를 한 장 추가 (텍스트박스에 `{{title}}`, `{{m1.label}}` 같은 placeholder 박은 채로).
3. 사용자가 PowerPoint에서 직접 디자인 다듬고 "OK"하면:
4. `session.register_layout(name, theme=, description=)` 한 줄 호출 — `user_layouts/<theme>_user.pptx`에 슬라이드 추가, 노트에 `pptmaker:custom name=... placeholders=...` 메타 자동 박음, `{{key}}` 자동 감지, `manifest.json` 갱신.
5. 이후부터 사용자가 그 패턴 요청 시 `add_custom()`으로 재사용.

상세 워크플로우 예시는 [README.md](README.md) "🧩 사용자 커스텀 레이아웃" 섹션.

---

## 6. PPT 유형별 권장 슬라이드 순서

| 유형 | 권장 순서 |
|---|---|
| **경영 보고 / IR** | Cover → Agenda → Section Divider → KPI → Chart → Table → Highlight Quote → Closing |
| **제품 / 세일즈** | Cover → Agenda → Full Image → Text+Image (×N) → Comparison → Highlight Quote → Closing |
| **주간 보고** | Cover → Agenda → Single Column (완료/진행/예정) → KPI → Table → Highlight Quote → Closing |
| **교육 / 온보딩** | Cover → Agenda → Section Divider → Process Flow → Two Column → Text+Image → Closing |

---

## 7. 디자인 변경 / 새 레이아웃 추가

개별 슬라이드 수정 ❌ → `_common.py`의 `Type`/`Role`/`Space`/`Layout` 클래스 값 수정 ⭕. 한 곳을 바꾸면 모든 레이아웃에 자동 반영된다.

새 본문 레이아웃 추가 절차:
1. `src/pptmaker/slides/` 에 새 모듈 작성 — 반드시 `_common.add_chrome_slide()` 진입점 사용
2. `Type`/`Role`/`Space`/`Layout` 만 사용 (매직 넘버 금지)
3. `LiveSession`에 `add_xxx()` 편의 메서드 추가
4. 이 SKILL.md의 §3 표 갱신

---

## 참조

- [README.md](README.md) — 설치, 사용 예시, 트러블슈팅, 디렉토리 구조
- [`src/pptmaker/slides/_common.py`](src/pptmaker/slides/_common.py) — 디자인 토큰 (`Type`/`Role`/`Space`/`Layout`) 상세 표
- [`src/pptmaker/live_session.py`](src/pptmaker/live_session.py) — 모든 메서드 시그니처
- [`design_tokens.json`](design_tokens.json) — layouts/ 분석 결과 (자동 생성)
