"""A1. 목차 (Agenda) — Type/Role/Space/Layout 스케일만 사용."""
from __future__ import annotations

from pptmaker.slides import _common as C


def add_to_live(session, items: list[str], *, title: str = "목차", page: int | None = None):
    slide, idx = C.add_chrome_slide(session, title=title)

    item_count = max(1, len(items))
    available_h = C.BODY_H - C.Space.XL
    line_h = min(48.0, available_h / item_count)
    start_y = C.BODY_TOP + C.Space.MD

    for i, item in enumerate(items, start=1):
        y = start_y + (i - 1) * line_h
        # 번호 (브랜드 컬러, H1)
        C.add_textbox(
            slide, C.Layout.MARGIN_X, y, 56, line_h - C.Space.SM,
            text=f"{i:02d}", font=C.FONT_MAJOR, size=C.Type.H1, bold=True,
            color=C.Role.BRAND_PRIMARY, align=C.PP_ALIGN_LEFT,
        )
        # 세로 구분선
        C.add_line(
            slide,
            C.Layout.MARGIN_X + 64, y + C.Space.SM,
            C.Layout.MARGIN_X + 64, y + line_h - C.Space.SM,
            color=C.Role.BRAND_PRIMARY, weight=1.5,
        )
        # 항목 텍스트 (H3, 본문)
        C.add_textbox(
            slide, C.Layout.MARGIN_X + 80, y,
            C.SLIDE_W - C.Layout.MARGIN_X * 2 - 80, line_h - C.Space.SM,
            text=item, font=C.FONT_MAJOR, size=C.Type.H3, bold=False,
            color=C.Role.TEXT_PRIMARY, align=C.PP_ALIGN_LEFT,
        )

    session.focus(idx)
    return idx
