"""A3. 본문 1단 — Type/Role/Space 스케일만 사용."""
from __future__ import annotations

from typing import Union

from pptmaker.slides import _common as C

Bullet = Union[str, dict]


def add_to_live(
    session,
    title: str,
    bullets: list[Bullet],
    *,
    eyebrow: str | None = None,
    subtitle: str | None = None,
    page: int | None = None,
):
    slide, idx = C.add_chrome_slide(session, title=title)

    body_start_y = C.BODY_TOP
    if eyebrow or subtitle:
        body_start_y = C.draw_section_intro(slide, eyebrow=eyebrow, subtitle=subtitle)

    body = slide.Shapes.AddTextbox(
        C.MSO_TEXT_HORIZONTAL,
        C.Layout.MARGIN_X, body_start_y,
        C.SLIDE_W - 2 * C.Layout.MARGIN_X, C.BODY_BOTTOM - body_start_y,
    )
    body.TextFrame.MarginLeft = 0
    body.TextFrame.MarginRight = 0
    tr = body.TextFrame.TextRange
    tr.Text = ""

    for i, item in enumerate(bullets):
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

    session.focus(idx)
    return idx
