"""B4. 비교 (Before / After) — Type/Role/Layout 스케일만 사용."""
from __future__ import annotations

from pptmaker.slides import _common as C


def add_to_live(
    session,
    title: str,
    left: dict,
    right: dict,
    *,
    arrow_label: str | None = None,
    page: int | None = None,
):
    slide, idx = C.add_chrome_slide(session, title=title)

    gutter = 80.0  # 가운데 화살표 영역 (특수)
    side_w = (C.SLIDE_W - 2 * C.Layout.MARGIN_X - gutter) / 2
    card_h = min(C.BODY_H - C.Space.XL, 320.0)
    card_top = C.BODY_TOP + (C.BODY_H - card_h) / 2

    def draw_side(left_x: float, data: dict, accent: int):
        # 카드 배경
        C.add_rect(slide, left_x, card_top, side_w, card_h,
                   fill=C.Role.BG_SUBTLE,
                   shape_type=C.MSO_SHAPE_ROUNDED_RECTANGLE)
        # 상단 라벨 바
        C.add_rect(slide, left_x, card_top, side_w, C.Layout.CARD_LABEL_BAR_H,
                   fill=accent)
        # 라벨 (BODY)
        C.add_textbox(
            slide, left_x + C.Space.MD, card_top + 6,
            side_w - 2 * C.Space.MD, 22,
            text=data.get("label", ""),
            font=C.FONT_MAJOR, size=C.Type.BODY, bold=True,
            color=C.Role.TEXT_INVERSE, align=C.PP_ALIGN_LEFT,
        )
        # 카드 제목 (H3)
        C.add_textbox(
            slide, left_x + C.Layout.CARD_INNER_PAD,
            card_top + C.Layout.CARD_LABEL_BAR_H + C.Space.LG,
            side_w - 2 * C.Layout.CARD_INNER_PAD, 36,
            text=data.get("title", ""),
            font=C.FONT_MAJOR, size=C.Type.H3, bold=True,
            color=C.Role.BRAND_DARK, align=C.PP_ALIGN_LEFT,
        )
        # 불릿 (CAPTION)
        body = slide.Shapes.AddTextbox(
            C.MSO_TEXT_HORIZONTAL,
            left_x + C.Layout.CARD_INNER_PAD,
            card_top + C.Layout.CARD_LABEL_BAR_H + C.Space.XXL + 28,
            side_w - 2 * C.Layout.CARD_INNER_PAD,
            card_h - C.Layout.CARD_LABEL_BAR_H - C.Space.XXL - 28 - C.Space.MD,
        )
        body.TextFrame.MarginLeft = 0
        body.TextFrame.MarginRight = 0
        tr = body.TextFrame.TextRange
        for i, b in enumerate(data.get("bullets", [])):
            line = C.bullet_prefix(0) + b
            if i == 0:
                tr.Text = line
            else:
                tr.InsertAfter("\r" + line)
        tr.Font.Name = C.FONT_MINOR
        tr.Font.NameFarEast = C.FONT_MINOR
        tr.Font.Size = C.Type.CAPTION
        tr.Font.Color.RGB = C.Role.TEXT_PRIMARY
        try:
            tr.ParagraphFormat.SpaceWithin = C.Layout.LINE_SPACING_BODY
        except Exception:
            pass

    # Before — SEMANTIC_NEGATIVE, After — SEMANTIC_POSITIVE
    draw_side(C.Layout.MARGIN_X, left, C.Role.SEMANTIC_NEGATIVE)
    draw_side(C.Layout.MARGIN_X + side_w + gutter, right, C.Role.SEMANTIC_POSITIVE)

    # 가운데 화살표
    arrow_w = 60.0
    arrow_h = 50.0
    arrow_x = C.Layout.MARGIN_X + side_w + (gutter - arrow_w) / 2
    arrow_y = card_top + (card_h - arrow_h) / 2
    C.add_rect(slide, arrow_x, arrow_y, arrow_w, arrow_h,
               fill=C.Role.BRAND_PRIMARY, shape_type=C.MSO_SHAPE_RIGHT_ARROW)

    if arrow_label:
        C.add_textbox(
            slide, C.Layout.MARGIN_X + side_w,
            card_top + card_h + C.Space.SM, gutter, 20,
            text=arrow_label, font=C.FONT_MINOR, size=C.Type.CAPTION,
            color=C.Role.TEXT_SECONDARY, align=C.PP_ALIGN_CENTER,
        )

    session.focus(idx)
    return idx
