"""A3. 본문 1단 — 두 가지 형태 지원:
  - bullets: 옛 불릿 리스트
  - tiles: Reference 4페이지 표준 타일 카드 (번호 + Bold 제목 + 설명 + 좌측 보더)

Type/Role/Space/Layout 스케일만 사용.
"""
from __future__ import annotations

from typing import Union

from pptmaker.slides import _common as C

Bullet = Union[str, dict]
Tile = Union[C.TileCard, dict, str]


def _to_tile(item: Tile, idx: int, default_accent: int) -> C.TileCard:
    """dict/str을 TileCard로 정규화."""
    if isinstance(item, C.TileCard):
        return item
    if isinstance(item, dict):
        return C.TileCard(
            number=item.get("number", idx),
            title=item.get("title", ""),
            desc=item.get("desc"),
            accent=item.get("accent", default_accent),
        )
    # 문자열 — title만 있는 카드
    return C.TileCard(number=idx, title=str(item), desc=None, accent=default_accent)


def add_to_live(
    session,
    title: str,
    bullets: list[Bullet] | None = None,
    *,
    tiles: list[Tile] | None = None,
    eyebrow: str | None = None,
    subtitle: str | None = None,
    section_no: int | None = None,
    page: int | None = None,
):
    """1단 본문 슬라이드.

    Args:
        title: 슬라이드 타이틀.
        bullets: 불릿 항목 리스트 (예전 형식). tiles와 동시 지정 금지.
        tiles: 타일 카드 항목 리스트 (Reference 표준). TileCard/dict/str 가능.
        eyebrow: 본문 상단 작은 라벨.
        subtitle: 본문 상단 부제.
        section_no: 타이틀에 "NN. " prefix 추가할 섹션 번호.
    """
    full_title = C.title_with_section_no(section_no, title)
    slide, idx = C.add_chrome_slide(session, title=full_title)

    body_start_y = C.BODY_TOP
    if eyebrow or subtitle:
        body_start_y = C.draw_section_intro(slide, eyebrow=eyebrow, subtitle=subtitle)

    if tiles is not None:
        _render_tiles(slide, tiles, body_start_y)
    else:
        _render_bullets(slide, bullets or [], body_start_y)

    session.focus(idx)
    return idx


def _render_tiles(slide, tiles_in, body_start_y: float):
    """타일 카드 N개를 세로로 쌓아 풀폭으로 렌더."""
    body_w = C.SLIDE_W - 2 * C.Layout.MARGIN_X
    body_h = C.BODY_BOTTOM - body_start_y - C.Space.MD
    n = max(1, len(tiles_in))
    gap = C.Space.MD
    tile_h = min(76.0, (body_h - gap * (n - 1)) / n)

    for i, raw in enumerate(tiles_in):
        card = _to_tile(raw, i + 1, C.Role.BRAND_PRIMARY)
        y = body_start_y + i * (tile_h + gap)
        C.draw_tile_card(slide, C.Layout.MARGIN_X, y, body_w, tile_h, card)


def _render_bullets(slide, bullets_in, body_start_y: float):
    """기존 불릿 리스트 렌더 (backward compat)."""
    body = slide.Shapes.AddTextbox(
        C.MSO_TEXT_HORIZONTAL,
        C.Layout.MARGIN_X, body_start_y,
        C.SLIDE_W - 2 * C.Layout.MARGIN_X, C.BODY_BOTTOM - body_start_y,
    )
    body.TextFrame.MarginLeft = 0
    body.TextFrame.MarginRight = 0
    tr = body.TextFrame.TextRange
    tr.Text = ""

    for i, item in enumerate(bullets_in):
        if isinstance(item, str):
            text, level = item, 0
        else:
            text = item.get("text", "")
            level = int(item.get("level", 0))
        chunk = C.bullet_prefix(level) + text
        if i == 0:
            tr.Text = chunk
        else:
            tr.InsertAfter("\r" + chunk)

    tr.Font.Name = C.FONT_MINOR
    tr.Font.NameFarEast = C.FONT_MINOR
    tr.Font.Size = C.Type.H3
    tr.Font.Color.RGB = C.Role.TEXT_PRIMARY
    tr.ParagraphFormat.Alignment = C.PP_ALIGN_LEFT
    try:
        tr.ParagraphFormat.SpaceWithin = C.Layout.LINE_SPACING_BODY
    except Exception:
        pass
