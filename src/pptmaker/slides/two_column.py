"""A4. 본문 2단 — Reference 4페이지 표준: 원형 아이콘 헤더 + 타일 카드.

Type/Role/Space/Layout 스케일만 사용.

두 가지 형태 지원:
  - bullets 모드: 각 컬럼이 불릿 리스트 (예전 형식, backward compat)
  - tiles 모드: 각 컬럼이 타일 카드 (Reference 표준)
"""
from __future__ import annotations

from pptmaker.slides import _common as C


# Reference 4페이지 표준 — 좌/우 컬럼 액센트
LEFT_ACCENT = C.Role.BRAND_PRIMARY     # 블루
RIGHT_ACCENT = C.hex_to_msorgb("#C00504")  # 레드 (Reference 우측 컬럼 매칭)


def add_to_live(
    session,
    title: str,
    left_sub: str,
    left_items: list,
    right_sub: str,
    right_items: list,
    *,
    use_tiles: bool = True,
    section_no: int | None = None,
    page: int | None = None,
    # 하위 호환: 옛 API의 left_bullets/right_bullets도 받는다
    left_bullets: list[str] | None = None,
    right_bullets: list[str] | None = None,
):
    """2단 본문 슬라이드.

    Args:
        title: 슬라이드 타이틀.
        left_sub / right_sub: 각 컬럼의 헤더 텍스트 (예: "기회 요인" / "리스크 요인").
        left_items / right_items: 컬럼 항목 — use_tiles에 따라 의미가 달라짐.
            - use_tiles=True (기본): TileCard/dict/str 리스트 → 타일 카드로 렌더
            - use_tiles=False: 문자열 리스트 → 불릿 리스트로 렌더
        use_tiles: True (기본) = Reference 표준 타일 카드. False = 옛 불릿 리스트.
        section_no: 타이틀에 "NN. " prefix 추가.
        left_bullets / right_bullets: 옛 API 호환. 지정 시 left_items/right_items 무시.
    """
    # 옛 API 호환
    if left_bullets is not None:
        left_items = left_bullets
        use_tiles = False
    if right_bullets is not None:
        right_items = right_bullets
        use_tiles = False

    full_title = C.title_with_section_no(section_no, title)
    slide, idx = C.add_chrome_slide(session, title=full_title)

    gutter = C.Layout.COLUMN_GUTTER
    inner_w = C.SLIDE_W - 2 * C.Layout.MARGIN_X - gutter
    col_w = inner_w / 2
    left_x = C.Layout.MARGIN_X
    right_x = left_x + col_w + gutter
    top = C.BODY_TOP

    # 컬럼 eyebrow (원형 아이콘 + Bold 헤더 + 굵은 underline)
    after_left = C.draw_column_eyebrow(slide, left_x, top, col_w, left_sub, accent=LEFT_ACCENT)
    after_right = C.draw_column_eyebrow(slide, right_x, top, col_w, right_sub, accent=RIGHT_ACCENT)
    body_y = max(after_left, after_right)

    body_h = C.BODY_BOTTOM - body_y - C.Space.MD

    # 본문 — 타일 또는 불릿
    if use_tiles:
        _render_tile_column(slide, left_x, body_y, col_w, body_h, left_items, LEFT_ACCENT)
        _render_tile_column(slide, right_x, body_y, col_w, body_h, right_items, RIGHT_ACCENT)
    else:
        _render_bullet_column(slide, left_x, body_y, col_w, body_h, left_items)
        _render_bullet_column(slide, right_x, body_y, col_w, body_h, right_items)

    session.focus(idx)
    return idx


def _to_tile(item, idx: int, default_accent: int) -> C.TileCard:
    if isinstance(item, C.TileCard):
        return item
    if isinstance(item, dict):
        return C.TileCard(
            number=item.get("number", idx),
            title=item.get("title", ""),
            desc=item.get("desc"),
            accent=item.get("accent", default_accent),
        )
    return C.TileCard(number=idx, title=str(item), desc=None, accent=default_accent)


def _render_tile_column(slide, x: float, y: float, w: float, h: float,
                        items: list, accent: int):
    n = max(1, len(items))
    gap = C.Space.MD
    tile_h = min(72.0, (h - gap * (n - 1)) / n)
    for i, raw in enumerate(items):
        card = _to_tile(raw, i + 1, accent)
        C.draw_tile_card(slide, x, y + i * (tile_h + gap), w, tile_h, card)


def _render_bullet_column(slide, x: float, y: float, w: float, h: float,
                          bullets: list):
    body = slide.Shapes.AddTextbox(C.MSO_TEXT_HORIZONTAL, x, y, w, h)
    body.TextFrame.MarginLeft = 0
    body.TextFrame.MarginRight = 0
    tr = body.TextFrame.TextRange
    for i, b in enumerate(bullets):
        line = C.bullet_prefix(0) + str(b)
        if i == 0:
            tr.Text = line
        else:
            tr.InsertAfter("\r" + line)
    tr.Font.Name = C.FONT_MINOR
    tr.Font.NameFarEast = C.FONT_MINOR
    tr.Font.Size = C.Type.BODY
    tr.Font.Color.RGB = C.Role.TEXT_PRIMARY
    try:
        tr.ParagraphFormat.SpaceWithin = C.Layout.LINE_SPACING_BODY
    except Exception:
        pass
