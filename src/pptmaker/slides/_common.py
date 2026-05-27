"""슬라이드 레이아웃 공통 헬퍼.

⚠️ 중요 — 디자인 일관성 강제 규칙:
모든 본문 슬라이드는 Reference 템플릿의 "제목만" 레이아웃 위에 생성된다.
이 레이아웃에 회사 표준 chrome(상단 헤더바 + 로고 + 하단 구분선 +
보안 표시 + 우측 하단 엠블럼)이 모두 들어있어, 슬라이드별로 chrome을
다시 그리지 않아도 자동 상속된다. 이는 사용자 명시 요구사항으로,
어떤 새 레이아웃을 만들어도 반드시 add_chrome_slide()를 진입점으로 한다.

좌표계: points(pt). 16:9 슬라이드 = 960 × 540 pt.
"""
from __future__ import annotations

from dataclasses import dataclass
from typing import Iterable

from pptmaker.design_tokens import TOKENS

# ── PowerPoint COM 상수 ─────────────────────────────────────────────
MSO_TRUE = -1
MSO_FALSE = 0

MSO_TEXT_HORIZONTAL = 1  # msoTextOrientation

# msoAutoShapeType
MSO_SHAPE_RECTANGLE = 1
MSO_SHAPE_ROUNDED_RECTANGLE = 5
MSO_SHAPE_RIGHT_ARROW = 13
MSO_SHAPE_OVAL = 9

# ppParagraphAlignment
PP_ALIGN_LEFT = 1
PP_ALIGN_CENTER = 2
PP_ALIGN_RIGHT = 3
PP_ALIGN_JUSTIFY = 4

MSO_PLACEHOLDER = 14  # msoShapeType

# PpSlideLayout fallback enums
PP_LAYOUT_TITLE_ONLY = 11
PP_LAYOUT_BLANK = 12

# ── 슬라이드 좌표계 (pt) ────────────────────────────────────────────
SLIDE_W = 960.0
SLIDE_H = 540.0

# Chrome (Reference 레이아웃 '제목만'의 영역) — 콘텐츠 절대 침범 금지
CHROME_TOP_H = 62.4         # 상단 헤더 바 높이 (2.2cm)
CHROME_BOTTOM_Y = 516.4     # 하단 구분선 y (18.22cm)
CHROME_TITLE_LEFT = 74.0    # 타이틀 placeholder x (2.61cm)
CHROME_TITLE_TOP = 3.1      # 타이틀 placeholder y (0.11cm)
CHROME_TITLE_W = 828.0      # 타이틀 placeholder width (29.21cm)
CHROME_TITLE_H = 56.0       # 타이틀 placeholder height (1.97cm)

# 본문 영역 (chrome 안쪽)
MARGIN_X = 48.0
BODY_TOP = CHROME_TOP_H + 18      # 헤더 바 아래 약간 간격
BODY_BOTTOM = CHROME_BOTTOM_Y - 8  # 하단 구분선 위 약간 간격
BODY_H = BODY_BOTTOM - BODY_TOP    # 약 428 pt


# ── 컬러 변환 ────────────────────────────────────────────────────────
def hex_to_msorgb(hex_str: str) -> int:
    """#RRGGBB → MSO Long (R + G<<8 + B<<16)."""
    h = hex_str.lstrip("#")
    r, g, b = int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)
    return r + (g << 8) + (b << 16)


# 미리 계산된 브랜드 컬러 (저수준 — 직접 참조 금지, Role을 통해 사용)
class C:
    ACCENT1 = hex_to_msorgb(TOKENS.colors.accent1)
    ACCENT2 = hex_to_msorgb(TOKENS.colors.accent2)
    ACCENT3 = hex_to_msorgb(TOKENS.colors.accent3)
    ACCENT4 = hex_to_msorgb(TOKENS.colors.accent4)
    ACCENT5 = hex_to_msorgb(TOKENS.colors.accent5)
    ACCENT6 = hex_to_msorgb(TOKENS.colors.accent6)
    DARK = hex_to_msorgb(TOKENS.colors.dark)
    LIGHT = hex_to_msorgb(TOKENS.colors.light)
    TEXT_DARK = hex_to_msorgb(TOKENS.colors.text_dark)
    TEXT_LIGHT = hex_to_msorgb(TOKENS.colors.text_light)
    GRAY_TEXT = hex_to_msorgb("#595959")


FONT_MAJOR = TOKENS.fonts.major
FONT_MINOR = TOKENS.fonts.minor


# ──────────────────────────────────────────────────────────────────────
# 브랜드 디자인 시스템 (Type / Role / Space / Layout)
# ──────────────────────────────────────────────────────────────────────
# 모든 레이아웃은 magic number 대신 아래 스케일만 사용해야 한다.
# 매직 넘버 사용 = 디자인 일관성 위반 = 사용자 명시 금지 사항.
# ──────────────────────────────────────────────────────────────────────

class Type:
    """타이포그래피 스케일 — 폰트 크기(pt). 이 값들 외에는 사용 금지."""
    QUOTE_MARK = 100    # highlight_quote의 거대한 따옴표 글리프
    HERO = 72           # section_divider의 큰 번호
    COVER_TITLE = 56    # 표지 메인 제목
    DISPLAY = 40        # section_divider 제목, 매우 큰 강조
    H1 = 28             # 큰 강조 (KPI value, agenda 번호)
    H2 = 20             # 소제목 (subtitle, eyebrow 강조)
    H3 = 16             # 본문 강조 (카드 제목, agenda 항목, comparison 카드 제목)
    BODY = 14           # 일반 본문 텍스트
    CAPTION = 11        # 캡션, 표 셀, kpi label/note, chart note
    MICRO = 9           # eyebrow, footer, 보안 표시


class Role:
    """컬러 역할 — 의미 기반. C.* 직접 참조 금지, 항상 Role을 거친다."""
    # 텍스트
    TEXT_PRIMARY = C.TEXT_DARK           # 본문 텍스트
    TEXT_SECONDARY = C.GRAY_TEXT         # 보조 텍스트, 캡션, eyebrow
    TEXT_INVERSE = C.TEXT_LIGHT          # 어두운 배경 위 텍스트 (흰색)
    # 브랜드
    BRAND_PRIMARY = C.ACCENT1            # 메인 강조 색
    BRAND_DARK = C.DARK                  # 다크 네이비
    # 의미적
    SEMANTIC_NEGATIVE = C.ACCENT5        # before, 부정적 비교 (퍼플)
    SEMANTIC_POSITIVE = C.ACCENT3        # after, 긍정적 비교 (그린)
    # 차트/카드용 액센트 (인덱스 0~5)
    ACCENT_SERIES = (
        C.ACCENT1, C.ACCENT2, C.ACCENT3, C.ACCENT4, C.ACCENT5, C.ACCENT6,
    )
    # 배경
    BG_SUBTLE = hex_to_msorgb("#F4F4F4")     # 카드, 표 줄무늬
    BG_DIVIDER = hex_to_msorgb("#DDDDDD")    # 구분선
    BG_CAPTION_OVERLAY = C.DARK              # 풀스크린 이미지 위 캡션 띠


class Space:
    """간격 스케일 — pt 단위. 4pt 기반."""
    XS = 4
    SM = 8
    MD = 12
    LG = 16
    XL = 24
    XXL = 32
    XXXL = 48


class Layout:
    """요소 배치 표준 — 모든 레이아웃이 따라야 할 고정 수치."""
    # 슬라이드 마진
    MARGIN_X = 48.0                    # 좌우 마진
    # 카드 (KPI, comparison)
    CARD_H = 220.0                     # KPI 카드 표준 높이
    CARD_GUTTER = Space.XL             # 카드 사이 거터
    CARD_ACCENT_BAR = 4.0              # 카드 상단 액센트 바 두께
    CARD_LABEL_BAR_H = 32.0            # comparison 카드 상단 라벨 바 높이
    CARD_INNER_PAD = Space.MD          # 카드 내부 패딩
    # 칼럼
    COLUMN_GUTTER = Space.XXL          # 좌우 칼럼 사이 거터
    # 프로세스 단계
    STEP_BOX_H = 120.0                 # 단계 박스 높이
    STEP_ARROW_W = 28.0                # 화살표 폭
    STEP_GAP = Space.SM                # 박스-화살표 간격
    # 액센트 라인 (subtitle 아래 짧은 라인)
    ACCENT_LINE_LEN = 40.0
    ACCENT_LINE_WEIGHT = 2.0
    # 구분선 (칼럼 사이 등)
    DIVIDER_WEIGHT = 1.0
    # 라인 스페이싱
    LINE_SPACING_BODY = 1.3
    LINE_SPACING_QUOTE = 1.4


# ── 불릿 ──────────────────────────────────────────────────────────────
def bullet_prefix(level: int = 0) -> str:
    """불릿 텍스트 prefix — 모든 레이아웃이 동일 형식 사용."""
    if level <= 0:
        return "• "
    return "  " * level + "– "


def standard_accent_line(slide, left: float, top: float,
                         length: float = Layout.ACCENT_LINE_LEN):
    """소제목 아래 표준 액센트 라인 (BRAND_PRIMARY, weight 2.0)."""
    return add_line(slide, left, top, left + length, top,
                    color=Role.BRAND_PRIMARY, weight=Layout.ACCENT_LINE_WEIGHT)


# ── 슬라이드 생성: 단일 진입점 ───────────────────────────────────────
# 모든 본문 레이아웃은 반드시 이 함수를 통해 슬라이드를 추가해야 한다.
def add_chrome_slide(session, title: str | None = None):
    """Reference '제목만' 레이아웃으로 슬라이드 추가.

    이 레이아웃에 회사 표준 chrome(헤더 바, 로고, 하단 구분선, 보안 표시,
    우측 하단 엠블럼)이 통째로 들어있어 모든 슬라이드에 자동 적용된다.

    session.blank_mode가 True면 본문 슬라이드를 closing 직전에 삽입
    (closing 템플릿을 마지막 자리에 보존). 아니면 끝에 append.

    Returns: (slide, idx) — slide는 COM 객체, idx는 1-based 슬라이드 번호.
    """
    # blank 모드면 closing(마지막) 직전에 삽입 — closing이 항상 마지막에 남도록
    if getattr(session, "blank_mode", False) and session.slide_count >= 1:
        idx = session.slide_count  # 마지막 자리 (closing 위치) — 삽입 후 closing은 idx+1로 밀림
    else:
        idx = session.slide_count + 1

    layout = session._find_custom_layout("제목만", "Title Only")
    if layout is not None:
        slide = session._prs.Slides.AddSlide(idx, layout)
    else:
        # Reference에 '제목만'이 없는 경우 fallback (사실상 일어나면 안 됨)
        slide = session._prs.Slides.Add(idx, PP_LAYOUT_TITLE_ONLY)

    # 타이틀 placeholder 채우기 — 두 테마 모두 흰색 Bold 강제.
    # theme1: 네이비 헤더 바 위 흰색. theme2: 사진 띠 위 흰색.
    if title is not None and slide.Shapes.HasTitle == MSO_TRUE:
        try:
            tr = slide.Shapes.Title.TextFrame.TextRange
            tr.Text = title
            tr.Font.Name = FONT_MAJOR
            tr.Font.NameFarEast = FONT_MAJOR
            tr.Font.Bold = MSO_TRUE
            tr.Font.Color.RGB = Role.TEXT_INVERSE  # 흰색 강제
        except Exception:
            pass

    return slide, idx


def title_with_section_no(idx: int | str | None, text: str) -> str:
    """본문 슬라이드 타이틀에 섹션 넘버 prefix 부여 ("04. 핵심 성과 요약").

    Reference Theme 표준: 본문 슬라이드 타이틀은 "NN. " prefix 가짐.
    idx=None이면 prefix 없이 그대로 반환.
    """
    if idx is None:
        return text
    if isinstance(idx, int):
        return f"{idx:02d}. {text}"
    return f"{idx}. {text}"


# ── 기본 도형 추가 ──────────────────────────────────────────────────
def add_textbox(slide, left, top, width, height, text="", *,
                font=FONT_MINOR, size=14, bold=False, color=C.TEXT_DARK,
                align=PP_ALIGN_LEFT, line_spacing=1.2):
    """텍스트 박스 추가 + 텍스트/폰트/색상/정렬 설정."""
    shape = slide.Shapes.AddTextbox(MSO_TEXT_HORIZONTAL, left, top, width, height)
    tf = shape.TextFrame
    tf.MarginLeft = 0
    tf.MarginRight = 0
    tf.MarginTop = 0
    tf.MarginBottom = 0
    tr = tf.TextRange
    tr.Text = text
    tr.Font.Name = font
    tr.Font.NameFarEast = font
    tr.Font.Size = size
    tr.Font.Bold = MSO_TRUE if bold else MSO_FALSE
    tr.Font.Color.RGB = color
    tr.ParagraphFormat.Alignment = align
    try:
        tr.ParagraphFormat.LineRuleWithin = MSO_TRUE
        tr.ParagraphFormat.SpaceWithin = line_spacing
    except Exception:
        pass
    return shape


def add_rect(slide, left, top, width, height, *,
             fill=C.ACCENT1, line=None, shape_type=MSO_SHAPE_RECTANGLE):
    """채워진 사각형(또는 다른 도형). line=None이면 테두리 숨김."""
    shape = slide.Shapes.AddShape(shape_type, left, top, width, height)
    shape.Fill.Visible = MSO_TRUE
    shape.Fill.Solid()
    shape.Fill.ForeColor.RGB = fill
    if line is None:
        shape.Line.Visible = MSO_FALSE
    else:
        shape.Line.Visible = MSO_TRUE
        shape.Line.ForeColor.RGB = line
    try:
        shape.TextFrame.MarginLeft = 6
        shape.TextFrame.MarginRight = 6
        shape.TextFrame.MarginTop = 6
        shape.TextFrame.MarginBottom = 6
    except Exception:
        pass
    return shape


def add_line(slide, x1, y1, x2, y2, *, color=C.ACCENT1, weight=2.0):
    """단순 라인 — 구분선 등."""
    line = slide.Shapes.AddLine(x1, y1, x2, y2)
    line.Line.ForeColor.RGB = color
    line.Line.Weight = weight
    return line


# ── 본문 영역 내 보조 컴포지션 ──────────────────────────────────────
def draw_section_intro(slide, *, eyebrow: str | None = None,
                       subtitle: str | None = None):
    """본문 상단(타이틀 바로 아래)에 부가 정보 — eyebrow(작은 라벨)와 subtitle.

    Chrome의 타이틀은 이미 상단 헤더에 표시되어 있고, 본문 영역에서
    이를 보조하는 작은 텍스트를 더할 때 사용. 모든 레이아웃이 일관된
    위치에 부가 정보를 넣도록 강제.
    """
    y = BODY_TOP
    if eyebrow:
        add_textbox(
            slide, MARGIN_X, y, SLIDE_W - 2 * MARGIN_X, 14,
            text=eyebrow, font=FONT_MINOR, size=10,
            color=C.GRAY_TEXT, align=PP_ALIGN_LEFT,
        )
        y += 16
    if subtitle:
        add_textbox(
            slide, MARGIN_X, y, SLIDE_W - 2 * MARGIN_X, 22,
            text=subtitle, font=FONT_MAJOR, size=14, bold=False,
            color=C.ACCENT1, align=PP_ALIGN_LEFT,
        )
        y += 24
    # accent 짧은 라인 (선택)
    add_line(
        slide, MARGIN_X, y + 2,
        MARGIN_X + 48, y + 2,
        color=C.ACCENT1, weight=2.0,
    )
    return y + 14  # 다음 컨텐츠 시작 y


@dataclass
class Card:
    title: str
    value: str | None = None
    body: str | None = None
    accent: int = C.ACCENT1


def draw_card(slide, left, top, width, height, card: Card):
    """KPI/포인트 카드: 상단 accent bar + 라벨 + 큰 값 + 보조 본문.

    표준 디자인: BG_SUBTLE 배경, 4pt 상단 accent bar, 라벨/값/본문 3단 구성.
    """
    pad = Layout.CARD_INNER_PAD
    # 배경
    add_rect(slide, left, top, width, height,
             fill=Role.BG_SUBTLE, shape_type=MSO_SHAPE_ROUNDED_RECTANGLE)
    # 상단 액센트 바
    add_rect(slide, left, top, width, Layout.CARD_ACCENT_BAR, fill=card.accent)
    # 라벨 (CAPTION, 보조 컬러)
    add_textbox(slide, left + pad, top + 14, width - 2 * pad, 18,
                text=card.title, font=FONT_MINOR, size=Type.CAPTION,
                color=Role.TEXT_SECONDARY)
    # 큰 값 (H1, 강조)
    if card.value:
        add_textbox(slide, left + pad, top + 36, width - 2 * pad, 48,
                    text=card.value, font=FONT_MAJOR, size=Type.H1, bold=True,
                    color=Role.BRAND_DARK)
    # 보조 본문 (CAPTION)
    if card.body:
        add_textbox(slide, left + pad, top + 90, width - 2 * pad, height - 100,
                    text=card.body, font=FONT_MINOR, size=Type.CAPTION,
                    color=Role.TEXT_PRIMARY)


def iterate_bullets(text_range, lines: Iterable[str], *, size: int = 16, color: int = C.TEXT_DARK):
    """TextRange에 한 줄씩 추가."""
    lines = list(lines)
    if not lines:
        return
    text_range.Text = lines[0]
    for line in lines[1:]:
        text_range.InsertAfter("\r" + line)
    text_range.Font.Name = FONT_MINOR
    text_range.Font.NameFarEast = FONT_MINOR
    text_range.Font.Size = size
    text_range.Font.Color.RGB = color


# ──────────────────────────────────────────────────────────────────────
# 새 시각 패턴 helpers (Reference 4페이지 디자인 표준)
# ──────────────────────────────────────────────────────────────────────

@dataclass
class TileCard:
    """타일 카드: 큰 faded 번호 + Bold 제목 + 보조 설명 + 좌측 컬러 보더."""
    number: int | str
    title: str
    desc: str | None = None
    accent: int = C.ACCENT1  # 좌측 보더 컬러
    faded_num_color: int | None = None  # None이면 accent 기반 자동


def _lighten(rgb_long: int, alpha: float = 0.65) -> int:
    """MSO RGB long을 흰색 방향으로 alpha 만큼 lighten."""
    r = rgb_long & 0xFF
    g = (rgb_long >> 8) & 0xFF
    b = (rgb_long >> 16) & 0xFF
    nr = int(r + (255 - r) * alpha)
    ng = int(g + (255 - g) * alpha)
    nb = int(b + (255 - b) * alpha)
    return nr + (ng << 8) + (nb << 16)


def draw_tile_card(slide, left: float, top: float, width: float, height: float,
                   card: TileCard):
    """Reference 4페이지 표준 타일 카드.

    구조: 좌측 컬러 보더 (4pt) + 회색 카드 배경 + 큰 faded 번호 + Bold 제목 + 본문 설명.
    """
    pad = Layout.CARD_INNER_PAD
    border_w = 4.0  # 좌측 보더 두께 (pt)

    # 카드 배경
    add_rect(slide, left, top, width, height,
             fill=Role.BG_SUBTLE, shape_type=MSO_SHAPE_ROUNDED_RECTANGLE)
    # 좌측 컬러 보더 (둥근 사각형 위에 얇은 사각형으로 덮음)
    add_rect(slide, left, top, border_w, height, fill=card.accent)

    # 큰 faded 번호 (배경처럼 — 텍스트 hierarchy의 보조 요소)
    num_color = card.faded_num_color if card.faded_num_color is not None else _lighten(card.accent, 0.7)
    num_str = f"{card.number:02d}" if isinstance(card.number, int) else str(card.number)
    add_textbox(slide,
                left + border_w + Space.SM, top + pad,
                52, height - 2 * pad,
                text=num_str, font=FONT_MAJOR, size=Type.H1, bold=True,
                color=num_color, align=PP_ALIGN_LEFT)

    # 제목 (H3 Bold)
    text_x = left + border_w + 60
    text_w = width - (text_x - left) - pad
    add_textbox(slide,
                text_x, top + pad + 2,
                text_w, 24,
                text=card.title, font=FONT_MAJOR, size=Type.H3, bold=True,
                color=Role.TEXT_PRIMARY, align=PP_ALIGN_LEFT)
    # 설명 (CAPTION)
    if card.desc:
        add_textbox(slide,
                    text_x, top + pad + 28,
                    text_w, height - pad - 30,
                    text=card.desc, font=FONT_MINOR, size=Type.CAPTION,
                    color=Role.TEXT_SECONDARY, align=PP_ALIGN_LEFT)


def draw_column_eyebrow(slide, left: float, top: float, width: float,
                        header_text: str, *, accent: int = Role.BRAND_PRIMARY):
    """원형 아이콘 + Bold 헤더 + 굵은 underline (Reference 4페이지 컬럼 eyebrow).

    accent 컬러로 통일 — 좌/우 컬럼 구분 시 다른 accent 전달.
    Returns: underline 아래 y 좌표 (다음 컨텐츠 시작점).
    """
    d_circle = 26.0  # 원 직경
    # 원형 아이콘
    add_rect(slide, left, top, d_circle, d_circle,
             fill=accent, shape_type=MSO_SHAPE_OVAL)
    # Bold 헤더 텍스트
    add_textbox(slide,
                left + d_circle + Space.SM, top + 2,
                width - d_circle - Space.SM, d_circle - 2,
                text=header_text, font=FONT_MAJOR, size=Type.H2, bold=True,
                color=accent, align=PP_ALIGN_LEFT)
    # 굵은 가로 underline
    underline_y = top + d_circle + Space.XS
    add_rect(slide, left, underline_y, width, 2.5,
             fill=accent)
    return underline_y + Space.LG  # 다음 컨텐츠 시작 y


def draw_toc_item(slide, left: float, top: float, width: float,
                  number: int, text: str, *,
                  number_color: int = Role.BRAND_PRIMARY,
                  text_color: int = Role.TEXT_PRIMARY,
                  row_height: float = 56.0):
    """Reference 2단 목차 항목: 번호 + 텍스트 + 아래 가로 구분선.

    구조: "01" Bold 컬러 + 텍스트 + 아래 얇은 가로선.
    """
    num_w = 36.0
    # 번호
    add_textbox(slide, left, top, num_w, row_height - 12,
                text=f"{number:02d}", font=FONT_MAJOR, size=Type.H2, bold=True,
                color=number_color, align=PP_ALIGN_LEFT)
    # 텍스트
    add_textbox(slide, left + num_w + Space.MD, top,
                width - num_w - Space.MD, row_height - 12,
                text=text, font=FONT_MAJOR, size=Type.H3, bold=False,
                color=text_color, align=PP_ALIGN_LEFT)
    # 아래 구분선 (얇은 회색)
    line_y = top + row_height - 6
    add_rect(slide, left, line_y, width, 0.8,
             fill=Role.BG_DIVIDER)
