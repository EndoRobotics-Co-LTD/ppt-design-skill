"""A2. 섹션 구분 — Type/Role/Space 스케일만 사용."""
from __future__ import annotations

from pptmaker.slides import _common as C


def add_to_live(
    session,
    title: str,
    *,
    number: int | str | None = None,
    subtitle: str | None = None,
    page: int | None = None,
):
    slide, idx = C.add_chrome_slide(session, title="구역 구분")

    center_y = (C.BODY_TOP + C.BODY_BOTTOM) / 2

    # 큰 번호 (HERO 크기, 브랜드)
    if number is not None:
        num_text = f"{int(number):02d}" if isinstance(number, int) or (isinstance(number, str) and str(number).isdigit()) else str(number)
        C.add_textbox(
            slide, C.Layout.MARGIN_X, center_y - 100, 400, 100,
            text=num_text, font=C.FONT_MAJOR, size=C.Type.HERO, bold=True,
            color=C.Role.BRAND_PRIMARY, align=C.PP_ALIGN_LEFT,
        )

    # 표준 액센트 라인
    C.standard_accent_line(slide, C.Layout.MARGIN_X, center_y + C.Space.XS)

    # 섹션 제목 (DISPLAY)
    C.add_textbox(
        slide, C.Layout.MARGIN_X, center_y + C.Space.LG + 2,
        C.SLIDE_W - 2 * C.Layout.MARGIN_X, 60,
        text=title, font=C.FONT_MAJOR, size=C.Type.DISPLAY, bold=True,
        color=C.Role.BRAND_DARK, align=C.PP_ALIGN_LEFT,
    )

    # 부제목 (H3)
    if subtitle:
        C.add_textbox(
            slide, C.Layout.MARGIN_X, center_y + 80, C.SLIDE_W - 2 * C.Layout.MARGIN_X, 24,
            text=subtitle, font=C.FONT_MINOR, size=C.Type.H3,
            color=C.Role.TEXT_SECONDARY, align=C.PP_ALIGN_LEFT,
        )

    session.focus(idx)
    return idx
