"""C2. 풀스크린 이미지 + 캡션 — Type/Role 스케일만 사용."""
from __future__ import annotations

from pathlib import Path

from pptmaker.slides import _common as C


def add_to_live(
    session,
    image_path: str | Path | None,
    *,
    caption: str | None = None,
    title: str = "이미지",
    section_no: int | None = None,
    page: int | None = None,
):
    slide, idx = C.add_chrome_slide(session, title=C.title_with_section_no(section_no, title))

    img_left = 0
    img_top = C.BODY_TOP
    img_w = C.SLIDE_W
    img_h = C.BODY_H

    if image_path and Path(str(image_path)).exists():
        slide.Shapes.AddPicture(
            FileName=str(Path(image_path).resolve()),
            LinkToFile=C.MSO_FALSE,
            SaveWithDocument=C.MSO_TRUE,
            Left=img_left, Top=img_top, Width=img_w, Height=img_h,
        )
    else:
        C.add_rect(slide, img_left, img_top, img_w, img_h, fill=C.Role.BG_SUBTLE)
        C.add_textbox(
            slide, 0, img_top + img_h / 2 - 10, C.SLIDE_W, 20,
            text="(이미지 자리)", font=C.FONT_MINOR, size=C.Type.BODY,
            color=C.Role.TEXT_SECONDARY, align=C.PP_ALIGN_CENTER,
        )

    # 캡션 — 어두운 띠 오버레이
    if caption:
        cap_h = 36.0
        cap_top = img_top + img_h - cap_h
        C.add_rect(slide, img_left, cap_top, img_w, cap_h,
                   fill=C.Role.BG_CAPTION_OVERLAY)
        try:
            last = slide.Shapes(slide.Shapes.Count)
            last.Fill.Transparency = 0.25
        except Exception:
            pass
        C.add_textbox(
            slide, C.Layout.MARGIN_X, cap_top + C.Space.SM,
            img_w - 2 * C.Layout.MARGIN_X, 22,
            text=caption, font=C.FONT_MINOR, size=C.Type.CAPTION,
            color=C.Role.TEXT_INVERSE, align=C.PP_ALIGN_LEFT,
        )

    session.focus(idx)
    return idx
