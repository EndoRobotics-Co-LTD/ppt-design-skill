"""A5. 핵심 메시지 — Type/Role/Space 스케일만 사용."""
from __future__ import annotations

from pptmaker.slides import _common as C


def add_to_live(
    session,
    message: str,
    *,
    title: str = "핵심 메시지",
    attribution: str | None = None,
    page: int | None = None,
):
    slide, idx = C.add_chrome_slide(session, title=title)

    center_y = (C.BODY_TOP + C.BODY_BOTTOM) / 2

    # 큰 따옴표 (QUOTE_MARK)
    C.add_textbox(
        slide, C.Layout.MARGIN_X, center_y - 110, 90, 90,
        text="“", font=C.FONT_MAJOR, size=C.Type.QUOTE_MARK, bold=True,
        color=C.Role.BRAND_PRIMARY, align=C.PP_ALIGN_LEFT,
    )

    # 메시지 (DISPLAY로 강조)
    C.add_textbox(
        slide, C.Layout.MARGIN_X + 60, center_y - 60,
        C.SLIDE_W - 2 * C.Layout.MARGIN_X - 60, 140,
        text=message, font=C.FONT_MAJOR, size=C.Type.DISPLAY, bold=True,
        color=C.Role.BRAND_DARK, align=C.PP_ALIGN_LEFT,
        line_spacing=C.Layout.LINE_SPACING_QUOTE,
    )

    # 출처 (BODY)
    if attribution:
        C.add_textbox(
            slide, C.Layout.MARGIN_X + 60, center_y + 90,
            C.SLIDE_W - 2 * C.Layout.MARGIN_X - 60, 24,
            text="— " + attribution, font=C.FONT_MINOR, size=C.Type.BODY,
            color=C.Role.TEXT_SECONDARY, align=C.PP_ALIGN_LEFT,
        )

    session.focus(idx)
    return idx
