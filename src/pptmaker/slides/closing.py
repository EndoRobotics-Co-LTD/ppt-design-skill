"""마무리 (Closing / Thank You) 레이아웃 — Reference 마무리 슬라이드 복제 방식.

cover와 동일하게 Reference의 표지/마무리 디자인을 100% 보존하면서 텍스트만 흰색으로 교체한다.

탐색 우선순위:
1) 현재 프레젠테이션에서 'Thank' 또는 '감사' 텍스트를 포함한 슬라이드
2) 1번 슬라이드 (cover) — fallback. cover와 동일한 풀-블리드 배경을 마무리에 재사용.
"""
from __future__ import annotations

from pptmaker.slides import _common as C


WHITE_RGB = C.Role.TEXT_INVERSE

# 표지와 동일한 스케일 사용 — 표지·마무리는 비주얼 페어
CLOSING_TITLE_SIZE = C.Type.COVER_TITLE   # 56pt
CLOSING_SUBTITLE_SIZE = C.Type.H2         # 20pt


def _force_text(shape, text: str, *, size: int, bold: bool = True,
                color: int = WHITE_RGB, align: int = C.PP_ALIGN_CENTER):
    try:
        if shape.HasTextFrame != C.MSO_TRUE:
            return
        tf = shape.TextFrame
        tr = tf.TextRange
        tr.Text = text
        tr.Font.Name = C.FONT_MAJOR
        tr.Font.NameFarEast = C.FONT_MAJOR
        tr.Font.Size = size
        tr.Font.Bold = C.MSO_TRUE if bold else C.MSO_FALSE
        tr.Font.Color.RGB = color
        tr.ParagraphFormat.Alignment = align
        try:
            tf.VerticalAnchor = 3  # msoAnchorMiddle
        except Exception:
            pass
    except Exception:
        pass


def _find_closing_template(session):
    """현재 프레젠테이션에서 마무리 슬라이드 후보를 찾는다.

    'Thank you' / '감사' / '끝' 텍스트를 포함한 슬라이드 우선,
    못 찾으면 1번 슬라이드(cover) 사용.
    """
    needles = ("Thank", "thank", "감사", "끝")
    for i in range(1, session.slide_count + 1):
        slide = session._prs.Slides(i)
        for shape in slide.Shapes:
            try:
                if shape.HasTextFrame != C.MSO_TRUE:
                    continue
                text = (shape.TextFrame.TextRange.Text or "").strip()
                if any(n in text for n in needles):
                    return slide
            except Exception:
                continue
    if session.slide_count >= 1:
        return session._prs.Slides(1)
    return None


def _find_main_textbox(slide):
    """슬라이드의 메인 텍스트 박스 — 위치 기준 가장 큰 TEXT_BOX 또는 placeholder 후보."""
    candidates = []
    for shape in slide.Shapes:
        try:
            if shape.HasTextFrame != C.MSO_TRUE:
                continue
            if shape.Type in (17, 14):  # TEXT_BOX or PLACEHOLDER
                candidates.append(shape)
        except Exception:
            continue
    if not candidates:
        return None
    # 면적 가장 큰 것
    candidates.sort(key=lambda s: -(s.Width * s.Height))
    return candidates[0]


def add_to_live(
    session,
    message: str = "Thank You",
    *,
    presenter: str | None = None,
    organization: str | None = None,
    date: str | None = None,
    subtitle: str | None = None,
):
    """마무리 슬라이드 추가.

    Args:
        message: 마무리 텍스트 (기본 "Thank You", "감사합니다" 등 자유)
        presenter, organization, date: 발표자 정보 (한 줄 부제목으로 합쳐 표시)
        subtitle: presenter 묶음 대신 직접 지정할 부제목
    """
    template = _find_closing_template(session)
    if template is None:
        raise ValueError("마무리 슬라이드 템플릿을 찾을 수 없습니다 (슬라이드 없음).")

    # 복제 → 끝으로 이동
    dup_range = template.Duplicate()
    target_idx = session.slide_count
    try:
        dup_range.MoveTo(target_idx)
    except Exception:
        pass

    new_slide = session._prs.Slides(target_idx)

    # 부제목 텍스트 구성
    if subtitle is None:
        parts = []
        if presenter:
            who = presenter if not organization else f"{presenter} · {organization}"
            parts.append(who)
        if date:
            parts.append(date)
        subtitle = "  |  ".join(parts) if parts else ""

    # 표준 위치 — 슬라이드 완전 중앙
    title_box_w = 800.0
    title_box_h = 140.0
    title_box_left = (C.SLIDE_W - title_box_w) / 2
    title_box_top = (C.SLIDE_H - title_box_h) / 2 - 24

    # 메인 텍스트 박스 재활용 또는 새로 생성
    main_box = _find_main_textbox(new_slide)
    if main_box is not None:
        try:
            main_box.Left = title_box_left
            main_box.Top = title_box_top
            main_box.Width = title_box_w
            main_box.Height = title_box_h
        except Exception:
            pass
        _force_text(
            main_box, message,
            size=CLOSING_TITLE_SIZE, bold=True,
            color=WHITE_RGB, align=C.PP_ALIGN_CENTER,
        )
    else:
        main_box = C.add_textbox(
            new_slide,
            title_box_left, title_box_top, title_box_w, title_box_h,
            text=message, font=C.FONT_MAJOR, size=CLOSING_TITLE_SIZE, bold=True,
            color=WHITE_RGB, align=C.PP_ALIGN_CENTER,
        )
        try:
            main_box.TextFrame.VerticalAnchor = 3
        except Exception:
            pass

    # 부제목
    if subtitle:
        sub_top = title_box_top + title_box_h + C.Space.MD
        sub_h = 36.0
        sub_box = C.add_textbox(
            new_slide,
            title_box_left, sub_top, title_box_w, sub_h,
            text=subtitle,
            font=C.FONT_MAJOR, size=CLOSING_SUBTITLE_SIZE, bold=False,
            color=WHITE_RGB, align=C.PP_ALIGN_CENTER,
        )
        try:
            sub_box.TextFrame.VerticalAnchor = 3
        except Exception:
            pass

    session.focus(target_idx)
    return target_idx
