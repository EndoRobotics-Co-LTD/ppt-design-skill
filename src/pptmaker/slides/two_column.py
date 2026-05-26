"""A4. 본문 2단 — Type/Role/Space/Layout 스케일만 사용."""
from __future__ import annotations

from pptmaker.slides import _common as C


def _draw_column(slide, left: float, width: float, top: float, body_h: float,
                 sub: str, bullets: list[str]):
    # 소제목 (H2)
    C.add_textbox(
        slide, left, top, width, 26,
        text=sub, font=C.FONT_MAJOR, size=C.Type.H2, bold=True,
        color=C.Role.BRAND_PRIMARY, align=C.PP_ALIGN_LEFT,
    )
    # 액센트 라인
    C.standard_accent_line(slide, left, top + 28)
    # 불릿 본문 (BODY)
    body = slide.Shapes.AddTextbox(
        C.MSO_TEXT_HORIZONTAL,
        left, top + 40, width, body_h - 40,
    )
    body.TextFrame.MarginLeft = 0
    body.TextFrame.MarginRight = 0
    tr = body.TextFrame.TextRange
    for i, b in enumerate(bullets):
        line = C.bullet_prefix(0) + b
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


def add_to_live(
    session,
    title: str,
    left_sub: str,
    left_bullets: list[str],
    right_sub: str,
    right_bullets: list[str],
    *,
    page: int | None = None,
):
    slide, idx = C.add_chrome_slide(session, title=title)

    gutter = C.Layout.COLUMN_GUTTER
    inner_w = C.SLIDE_W - 2 * C.Layout.MARGIN_X - gutter
    col_w = inner_w / 2

    _draw_column(slide, C.Layout.MARGIN_X, col_w, C.BODY_TOP, C.BODY_H,
                 left_sub, left_bullets)
    _draw_column(slide, C.Layout.MARGIN_X + col_w + gutter, col_w, C.BODY_TOP, C.BODY_H,
                 right_sub, right_bullets)

    # 중앙 세로 구분선
    C.add_line(
        slide,
        C.Layout.MARGIN_X + col_w + gutter / 2, C.BODY_TOP + C.Space.SM,
        C.Layout.MARGIN_X + col_w + gutter / 2, C.BODY_BOTTOM - C.Space.SM,
        color=C.Role.BG_DIVIDER, weight=C.Layout.DIVIDER_WEIGHT,
    )

    session.focus(idx)
    return idx
