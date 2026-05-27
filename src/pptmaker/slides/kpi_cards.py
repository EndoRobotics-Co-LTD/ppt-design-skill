"""B1. KPI 3단 카드 — Layout/Role 스케일만 사용."""
from __future__ import annotations

from pptmaker.slides import _common as C


def add_to_live(
    session,
    title: str,
    cards: list[dict],
    *,
    section_no: int | None = None,
    page: int | None = None,
):
    if not (2 <= len(cards) <= 4):
        raise ValueError("KPI 카드는 2~4개여야 합니다.")

    slide, idx = C.add_chrome_slide(session, title=C.title_with_section_no(section_no, title))

    gutter = C.Layout.CARD_GUTTER
    n = len(cards)
    total_w = C.SLIDE_W - 2 * C.Layout.MARGIN_X - gutter * (n - 1)
    card_w = total_w / n
    card_h = C.Layout.CARD_H
    card_top = C.BODY_TOP + (C.BODY_H - card_h) / 2

    for i, c in enumerate(cards):
        left = C.Layout.MARGIN_X + i * (card_w + gutter)
        accent = C.Role.ACCENT_SERIES[i % len(C.Role.ACCENT_SERIES)]
        C.draw_card(
            slide, left, card_top, card_w, card_h,
            C.Card(
                title=c.get("label", ""),
                value=c.get("value", ""),
                body=c.get("note", ""),
                accent=accent,
            ),
        )

    session.focus(idx)
    return idx
