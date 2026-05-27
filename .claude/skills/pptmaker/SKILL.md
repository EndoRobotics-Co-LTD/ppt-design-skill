---
name: pptmaker
description: EndoRobotics 전사 PPT 빌더. 사내 표준 디자인(맑은 고딕, accent1 #156082, 16:9, 표준 chrome)을 강제하면서 14종 레이아웃으로 PPT를 대화형 생성한다. 사용자 요청에 "PPT 만들어줘", "표지 슬라이드", "KPI 카드", "주간 보고", "발표 자료", "회사 표준 PPT" 등이 포함되면 호출. PowerPoint 데스크톱과 실시간 연동되는 라이브 모드를 기본으로 한다.
---

# PPTMaker — EndoRobotics 전사 PPT 빌더

이 스킬은 EndoRobotics 사내 디자인 표준을 자동으로 강제하면서 AI가 PPT를 만들어주는 도구다.
사용자가 `/pptmaker` 또는 자연어로 PPT 작성을 요청하면 이 문서의 규칙을 따라 작업한다.

---

## 0. 사용 가능한 두 백엔드

| 모드 | 진입점 | 사용 시점 |
|---|---|---|
| **라이브 (기본)** | `pptmaker.live_session.LiveSession` (pywin32 COM) | PowerPoint 데스크톱이 켜진 상태에서 실시간 슬라이드 생성·확인 |
| **오프라인** | `pptmaker.offline_builder.OfflineBuilder` (python-pptx) | PowerPoint 없이 .pptx 파일만 생성 (배치, CI 환경) |

기본은 **라이브 모드**. 사용자가 명시적으로 "파일만 받을게" / "백그라운드로 만들어줘"라고 하면 오프라인 모드.

---

## 1. 대화 흐름 (Conversational Flow)

사용자에게 모든 정보를 한 번에 묻지 말고 짧게 끊어서 진행한다.

1. **테마 선택** — **반드시 첫 단계에서 물어본다.** 사용자가 어느 테마를 원하는지 확인:
   - **theme1** — 네이비 헤더 바 + 흰색 Bold 타이틀. Formal/Corporate. (경영 보고, IR, 내부 보고서)
   - **theme2** — 의료기기 사진 풀-블리드 표지 + 사진 띠 헤더 + 흰색 Bold 타이틀. Visual/Branded. (외부 발표, 제품 마케팅, 학회 자료)
   - 사용자가 명시 안 했으면 한 번 물어보고 진행. 기본값은 `theme1`.
2. **목적/유형 확인** — 경영 보고, 제품 세일즈, 주간 보고, 교육 등 어떤 PPT인지
3. **자료/주제 수집** — 핵심 메시지, 슬라이드 개요, 첨부 자료 경로
4. **개요 슬라이드 제안** — AI가 슬라이드 타이틀 리스트를 먼저 제시 → 사용자 승인
5. **순차 생성** — 한 번에 1~3장씩 만들고 사용자가 확인할 시간을 둔다
6. **저장 시점** — 사용자가 "저장해줘"라고 명시할 때까지 절대 자동 저장하지 않음

### 1.1 테마 선택 코드 진입점

```python
# Production 빌드 (데모 컨텐츠 제거 후 시작 — 결과물에 데모 안 따라옴)
from pptmaker.live_session import LiveSession
from pptmaker.slides._common import TileCard

with LiveSession.open(theme="theme1", from_blank=True) as session:
    session.add_cover("발표 제목", presenter="이은상", organization="전략기획팀", date="2026.05.27")
    session.add_agenda(["회사 소개", "1분기 실적", "주요 전략", "Q&A"])
    session.add_section_divider("1분기 실적", number=1, subtitle="Q1 Review")
    session.add_single_column(
        "핵심 성과 요약",
        tiles=[
            TileCard(number=1, title="매출 +18% YoY", desc="분기 최대 실적"),
            TileCard(number=2, title="신규 고객 12사", desc="글로벌 5개사 포함"),
        ],
        section_no=4,  # 타이틀에 "04. " prefix 자동 부여
    )
    session.add_two_column(
        "시장 환경 변화",
        "기회 요인", ["글로벌 헬스케어 수요 확대", "정부 R&D 지원 강화"],
        "리스크 요인", ["부품 가격 상승", "환율 변동성"],
        section_no=5,
    )
    session.add_closing("Thank you")
    session.save_as("output.pptx")

# 오프라인 모드 (PowerPoint 없이 빌드)
from pptmaker.offline_builder import OfflineBuilder
builder = OfflineBuilder(theme="theme2")
```

**핵심 옵션**:
- `from_blank=True` — production 빌드 권장. 데모 슬라이드 제거하고 cover/closing 템플릿만 유지.
- `tiles=...` (single_column / two_column) — Reference 4페이지 표준 타일 카드 (번호 + Bold 제목 + 설명 + 좌측 컬러 보더).
- `section_no=N` — 본문 슬라이드 타이틀에 "NN. " prefix 자동 부여.

`pptmaker.themes.list_themes()` 로 사용 가능 테마 목록을 런타임에 확인할 수 있다.

---

## 2. 디자인 시스템 (강제 규칙)

**이 규칙은 절대 위반 금지. 모든 슬라이드 빌더는 아래 스케일만 사용한다.**
구현: `src/pptmaker/slides/_common.py`의 `Type`, `Role`, `Space`, `Layout` 클래스.

### 2.1 Chrome (모든 슬라이드 공통 프레임)

- **본문 슬라이드**: Reference의 **"제목만" 레이아웃**을 강제 상속 → 헤더 바·로고·하단 구분선·보안 표시·우측 엠블럼이 자동 포함
- **표지/마무리**: Reference의 **1번 슬라이드(또는 Thank you 슬라이드)를 복제** → 풀-블리드 배경 보존
- **Chrome 제목**: 항상 **Bold**, 폰트는 맑은 고딕 강제
- **새 레이아웃 추가 시**: 반드시 `add_chrome_slide()` 또는 슬라이드 복제 방식 사용. chrome을 직접 그리기 **금지**

### 2.2 Typography — `Type` 클래스 (pt 단위)

| 키 | 값 | 용도 |
|---|---|---|
| `QUOTE_MARK` | 100 | 거대한 따옴표 글리프 |
| `HERO` | 72 | section_divider 큰 번호 |
| `COVER_TITLE` | 56 | 표지·마무리 메인 제목 |
| `DISPLAY` | 40 | section_divider 제목, 매우 큰 강조 |
| `H1` | 28 | KPI value, agenda 번호 |
| `H2` | 20 | 소제목, 표지 부제목 |
| `H3` | 16 | 카드 제목, agenda 항목 |
| `BODY` | 14 | 일반 본문 |
| `CAPTION` | 11 | 캡션, 표 셀, note |
| `MICRO` | 9 | eyebrow, footer |

**폰트**: 맑은 고딕(`FONT_MAJOR`/`FONT_MINOR`)만 사용. 다른 폰트 사용 금지.

### 2.3 Color — `Role` 클래스 (의미 기반)

| 역할 | HEX | 용도 |
|---|---|---|
| `TEXT_PRIMARY` | `#000000` | 본문 텍스트 |
| `TEXT_SECONDARY` | `#595959` | 보조 텍스트, 캡션 |
| `TEXT_INVERSE` | `#FFFFFF` | 어두운 배경 위 (표지·마무리) |
| `BRAND_PRIMARY` | `#156082` | 메인 강조 (accent1, 틸 블루) |
| `BRAND_DARK` | `#0E2841` | 다크 네이비 |
| `SEMANTIC_NEGATIVE` | `#A02B93` | Before / 부정적 비교 (퍼플) |
| `SEMANTIC_POSITIVE` | `#196B24` | After / 긍정적 비교 (그린) |
| `ACCENT_SERIES` | 6색 튜플 | 차트·카드 순환 |
| `BG_SUBTLE` | `#F4F4F4` | 카드, 표 줄무늬 |
| `BG_DIVIDER` | `#DDDDDD` | 칼럼 구분선 |
| `BG_CAPTION_OVERLAY` | `#0E2841` | 풀스크린 이미지 캡션 띠 |

**HEX 코드 직접 참조 금지. 항상 `Role.*`만.**

### 2.4 Spacing — `Space` 클래스 (4pt 기반)

`XS=4, SM=8, MD=12, LG=16, XL=24, XXL=32, XXXL=48` — 임의 값 금지.

### 2.5 Layout 배치 표준 — `Layout` 클래스

| 키 | 값 | 의미 |
|---|---|---|
| `MARGIN_X` | 48 | 슬라이드 좌우 마진 |
| `CARD_H` | 220 | KPI 카드 표준 높이 |
| `CARD_GUTTER` | 24 | 카드 사이 거터 |
| `COLUMN_GUTTER` | 32 | 좌우 칼럼 거터 |
| `STEP_BOX_H` | 120 | 프로세스 단계 박스 높이 |
| `ACCENT_LINE_LEN` | 40 | 소제목 아래 짧은 라인 |

**불릿**: `bullet_prefix(level)` 강제 사용 — level 0 = "• ", level 1+ = 들여쓰기 + "– ".

---

## 3. 사용 가능한 14종 레이아웃

모두 `LiveSession`의 편의 메서드로 호출.

| # | 레이아웃 | 호출 메서드 | 용도 |
|---|---|---|---|
| - | **Cover** | `session.add_cover(title, presenter, organization, date)` | 표지 (Reference 1번 복제) |
| A1 | Agenda | `session.add_agenda(items, title)` | 목차 |
| A2 | Section Divider | `session.add_section_divider(title, number, subtitle)` | 챕터 구분 |
| A3 | Single Column | `session.add_single_column(title, bullets, eyebrow=)` | 본문 1단 |
| A4 | Two Column | `session.add_two_column(title, l_sub, l_bullets, r_sub, r_bullets)` | 본문 2단 |
| A5 | Highlight Quote | `session.add_highlight_quote(message, attribution=)` | 핵심 메시지 |
| B1 | KPI Cards | `session.add_kpi_cards(title, cards)` | 2~4개 지표 카드 |
| B2 | Data Table | `session.add_data_table(title, rows)` | 표 |
| B3 | Chart | `session.add_chart(title, categories, series, chart_type=)` | 차트 |
| B4 | Comparison | `session.add_comparison(title, left, right, arrow_label=)` | Before/After |
| C1 | Text + Image | `session.add_text_image(title, image_path, subtitle, bullets, image_side=)` | 텍스트+이미지 |
| C2 | Full Image | `session.add_full_image(image_path, caption=, title=)` | 풀스크린 이미지 |
| C3 | Process Flow | `session.add_process_flow(title, steps)` | 프로세스 다이어그램 |
| - | **Closing** | `session.add_closing(message, presenter, organization, date)` | 마무리 (Thank You) |

---

## 4. PPT 유형별 권장 구성

| PPT 유형 | 권장 슬라이드 순서 |
|---|---|
| **경영 보고 / IR** | Cover → Agenda → Section Divider → KPI → Chart → Table → Highlight Quote → Closing |
| **제품 / 세일즈** | Cover → Agenda → Full Image → Text+Image (×N) → Comparison → Highlight Quote → Closing |
| **주간 보고** | Cover → Agenda → Single Column (완료/진행/예정) → KPI → Table → Highlight Quote → Closing |
| **교육 / 온보딩** | Cover → Agenda → Section Divider → Process Flow → Two Column (개념) → Text+Image → Closing |

---

## 5. 코드 사용 예시

### 5.1 라이브 모드 (기본)
```python
from pathlib import Path
from pptmaker.live_session import LiveSession

with LiveSession.open(theme="theme1") as session:  # 또는 theme="theme2"
    # 표지
    session.add_cover(
        title="2026년 1분기 사업 보고",
        presenter="이은상",
        organization="전략기획팀",
        date="2026.05.26",
    )
    # 목차
    session.add_agenda(["1분기 실적", "주요 이슈", "다음 분기 계획", "Q&A"])
    # KPI 3장
    session.add_kpi_cards("핵심 지표", cards=[
        {"label": "매출 (YoY)", "value": "+18%", "note": "전년 동기 대비"},
        {"label": "신규 고객", "value": "12사", "note": "글로벌 5개사 포함"},
        {"label": "EBITDA", "value": "₩2.3B", "note": "마진 14%"},
    ])
    # 차트
    session.add_chart(
        "분기별 매출 추이",
        categories=["Q1", "Q2", "Q3", "Q4"],
        series=[{"name": "매출(억원)", "values": [82, 95, 108, 121]}],
        chart_type="column",
    )
    # 마무리
    session.add_closing("Thank You", presenter="이은상", date="2026.05.26")
    # 저장은 사용자가 "저장"이라고 했을 때만!
```

### 5.2 오프라인 모드
```python
from pptmaker.offline_builder import OfflineBuilder

builder = OfflineBuilder()
builder.add_title_only_slide("회사 소개")
builder.add_body_slide("미션", ["환자와 의료진을 위한 로봇", "..."])
out = builder.save("output/회사소개.pptx")
```

---

## 6. 작업 원칙 (Workflow Rules)

| 규칙 | 설명 |
|---|---|
| **사용자 수동 편집 보호** | 라이브 모드에서 매 작업 전 슬라이드 상태를 읽어 사용자가 손댄 내용을 덮어쓰지 않는다 |
| **저장 명시성** | `session.save()`는 사용자 명시 지시 후에만. 절대 자동 저장 금지 |
| **점진적 생성** | 한 번에 전체 PPT 생성 금지. 1~3장씩 만들고 사용자 확인 받기 |
| **한국어 우선** | 모든 안내·확인 메시지는 한국어로 |
| **컬러 검증** | 의심스러운 색이 들어가면 `pptmaker.theme_guard.audit_color`로 감사 후 사용자에게 알린다 |
| **빈 슬라이드 자동 청소** | 생성 작업 후 "제목만" 레이아웃 + 텍스트 없음 슬라이드는 식별해 삭제 (`scripts/remove_empty_slides.py` 참조) |

---

## 7. 디자인 변경이 필요할 때

**개별 슬라이드 수정 ❌ → `_common.py`의 `Type/Role/Space/Layout` 클래스 값 수정 ⭕**
→ 한 번 바꾸면 모든 레이아웃에 자동 반영.

예시:
- 본문 글자 14 → 15pt: `Type.BODY = 15`
- 메인 컬러 변경: `Role.BRAND_PRIMARY` 또는 `design_tokens.json` 갱신 → `python scripts/analyze_templates.py` 재실행
- 좌우 마진 축소: `Layout.MARGIN_X = 40`

---

## 8. 트러블슈팅

### 8.1 "차트 데이터 표가 이미 열려 있습니다" 에러
이전 차트 생성 시 워크북이 안 닫혔음. 해결:
```python
from scripts.cleanup_slides import cleanup_chart_workbooks  # 또는
```
또는 PowerPoint에서 차트 데이터 편집창 수동으로 닫고 재시도.

### 8.2 슬라이드 1이 없어서 cover/closing 실패
`cover`·`closing` 빌더는 슬라이드 1을 배경 템플릿으로 복제한다. 슬라이드 1을 삭제했다면 Reference .pptx를 새로 열어 시작.

### 8.3 한글이 깨져 보임
PowerShell stdout 인코딩 문제. 스크립트 실행 시 `$env:PYTHONIOENCODING="utf-8"` 먼저 설정. PowerPoint 안의 한글은 정상.

### 8.4 새 레이아웃을 추가하고 싶음
1. `src/pptmaker/slides/`에 새 모듈 추가 (예: `timeline.py`)
2. `add_to_live(session, ...)` 함수 구현 — 반드시 `add_chrome_slide()` 진입점 사용
3. `Type/Role/Space/Layout` 만 사용 (magic number 금지)
4. `slides/__init__.py`에 등록
5. `LiveSession`에 편의 메서드 추가
6. 이 SKILL.md 표 갱신

---

## 9. 핵심 파일 위치

```
PPTMaker/
├── layouts/                                ← 테마 템플릿 SSOT (.pptx)
│   ├── theme1.pptx                         ← theme1 — 14종 레이아웃 전부
│   └── theme2.pptx                         ← theme2 — 14종 레이아웃 전부
├── design_tokens.json                      ← layouts/ 분석 결과
├── src/pptmaker/
│   ├── design_tokens.py                    ← TOKENS 객체
│   ├── live_session.py                     ← 라이브 모드 + 14종 편의 메서드
│   ├── offline_builder.py                  ← 오프라인 모드
│   ├── theme_guard.py                      ← 컬러·폰트 감사
│   └── slides/
│       ├── _common.py                      ← Type/Role/Space/Layout + 헬퍼
│       ├── cover.py · closing.py           ← 표지·마무리 (슬라이드 복제 방식)
│       ├── agenda.py · section_divider.py · single_column.py · ...
│       └── ... (12 본문 레이아웃)
├── scripts/
│   ├── analyze_templates.py                ← layouts/ 분석 → design_tokens.json
│   ├── inspect_master.py                   ← 슬라이드 마스터 검사
│   ├── verify_all.py                       ← 14종 검증 스크립트
│   └── remove_empty_slides.py              ← 빈 슬라이드 자동 청소
└── .claude/skills/pptmaker/SKILL.md        ← (이 파일)
```

---

## 10. 버전 / 변경 관리

- **2026.02 버전 기준** 디자인 토큰 추출 완료
- `layouts/theme*.pptx`를 커스텀했을 때: `python scripts/analyze_templates.py` 실행 → `design_tokens.json` 자동 갱신
- 이 스킬의 갱신 = git pull (전사 배포 시)
