"""C1. 텍스트 + 이미지 — Type/Role/Layout 스케일만 사용."""
from __future__ import annotations

from pathlib import Path

from pptmaker.slides import _common as C


def add_to_live(
    session,
    title: str,
    image_path: str | Path | None,
    subtitle: str,
    bullets: list[str],
    *,
    image_side: str = "left",
    section_no: int | None = None,
    page: int | None = None,
):
    slide, idx = C.add_chrome_slide(session, title=C.title_with_section_no(section_no, title))

    gutter = C.Space.XXXL - 8.0  # 40
    side_w = (C.SLIDE_W - 2 * C.Layout.MARGIN_X - gutter) / 2
    side_top = C.BODY_TOP
    side_h = C.BODY_H

    if image_side == "right":
        text_left = C.Layout.MARGIN_X
        img_left = C.Layout.MARGIN_X + side_w + gutter
    else:
        img_left = C.Layout.MARGIN_X
        text_left = C.Layout.MARGIN_X + side_w + gutter

    if image_path and Path(str(image_path)).exists():
        slide.Shapes.AddPicture(
            FileName=str(Path(image_path).resolve()),
            LinkToFile=C.MSO_FALSE,
            SaveWithDocument=C.MSO_TRUE,
            Left=img_left, Top=side_top, Width=side_w, Height=side_h,
        )
    else:
        C.add_rect(slide, img_left, side_top, side_w, side_h, fill=C.Role.BG_SUBTLE)
        C.add_textbox(
            slide, img_left, side_top + side_h / 2 - 10, side_w, 20,
            text="(이미지 자리)", font=C.FONT_MINOR, size=C.Type.CAPTION,
            color=C.Role.TEXT_SECONDARY, align=C.PP_ALIGN_CENTER,
        )

    # 소제목 (H2)
    C.add_textbox(
        slide, text_left, side_top, side_w, 28,
        text=subtitle, font=C.FONT_MAJOR, size=C.Type.H2, bold=True,
        color=C.Role.BRAND_PRIMARY, align=C.PP_ALIGN_LEFT,
    )
    C.standard_accent_line(slide, text_left, side_top + 32)

    # 불릿 (BODY)
    body = slide.Shapes.AddTextbox(
        C.MSO_TEXT_HORIZONTAL,
        text_left, side_top + 48, side_w, side_h - 50,
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
        tr.ParagraphFormat.SpaceWithin = C.Layout.LINE_SPACING_QUOTE
    except Exception:
        pass

    session.focus(idx)
    return idx
