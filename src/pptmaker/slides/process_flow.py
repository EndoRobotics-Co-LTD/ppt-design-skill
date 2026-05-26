"""C3. 프로세스 다이어그램 — Type/Role/Layout 스케일만 사용."""
from __future__ import annotations

from pptmaker.slides import _common as C


def add_to_live(
    session,
    title: str,
    steps: list[dict],
    *,
    page: int | None = None,
):
    if not (2 <= len(steps) <= 6):
        raise ValueError("프로세스 단계는 2~6개여야 합니다.")

    slide, idx = C.add_chrome_slide(session, title=title)

    n = len(steps)
    pad = C.Layout.CARD_INNER_PAD
    arrow_w = C.Layout.STEP_ARROW_W
    gap = C.Layout.STEP_GAP
    total_gap = arrow_w * (n - 1) + gap * 2 * (n - 1)
    box_w = (C.SLIDE_W - 2 * C.Layout.MARGIN_X - total_gap) / n
    box_h = C.Layout.STEP_BOX_H
    row_top = C.BODY_TOP + (C.BODY_H - box_h) / 2 - 10

    cursor = C.Layout.MARGIN_X
    for i, step in enumerate(steps):
        accent = C.Role.ACCENT_SERIES[i % len(C.Role.ACCENT_SERIES)]
        # 박스
        C.add_rect(slide, cursor, row_top, box_w, box_h,
                   fill=C.Role.BG_SUBTLE,
                   shape_type=C.MSO_SHAPE_ROUNDED_RECTANGLE)
        # 단계 라벨 (H3, accent 컬러)
        C.add_textbox(
            slide, cursor + pad, row_top + C.Space.MD,
            box_w - 2 * pad, 26,
            text=str(step.get("label", f"{i+1:02d}")),
            font=C.FONT_MAJOR, size=C.Type.H3, bold=True,
            color=accent, align=C.PP_ALIGN_LEFT,
        )
        # 단계 제목 (H3)
        C.add_textbox(
            slide, cursor + pad, row_top + 38,
            box_w - 2 * pad, 26,
            text=step.get("title", ""),
            font=C.FONT_MAJOR, size=C.Type.H3, bold=True,
            color=C.Role.BRAND_DARK, align=C.PP_ALIGN_LEFT,
        )
        # 설명 (CAPTION)
        if step.get("note"):
            C.add_textbox(
                slide, cursor + pad, row_top + 66,
                box_w - 2 * pad, 48,
                text=step["note"], font=C.FONT_MINOR, size=C.Type.CAPTION,
                color=C.Role.TEXT_SECONDARY, align=C.PP_ALIGN_LEFT,
                line_spacing=1.2,
            )

        # 다음 박스로 + 화살표
        next_x = cursor + box_w + gap
        if i < n - 1:
            arrow_y = row_top + (box_h - 28) / 2
            C.add_rect(slide, next_x, arrow_y, arrow_w, 28,
                       fill=C.Role.BRAND_PRIMARY,
                       shape_type=C.MSO_SHAPE_RIGHT_ARROW)
            cursor = next_x + arrow_w + gap

    session.focus(idx)
    return idx
