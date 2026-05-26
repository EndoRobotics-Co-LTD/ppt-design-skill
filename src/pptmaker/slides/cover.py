"""표지 (Cover) 레이아웃 — 슬라이드 1 복제 방식.

EndoRobotics Reference의 1번 슬라이드(풀-블리드 배경 이미지 + Title 텍스트 박스)를
그대로 복제하고, 텍스트만 흰색으로 교체해 새 표지를 만든다.

이 방식의 장점:
- 배경 이미지를 픽셀 단위로 보존 (Reference 디자인 100% 유지)
- 새 풀스크린 이미지를 매번 끼워넣지 않아 파일 크기 안정
- 디자이너가 슬라이드 1만 수정하면 모든 표지에 자동 반영

흰색 텍스트 강제 — 배경이 어두운 이미지이므로 가독성을 위해 화이트로 고정.
"""
from __future__ import annotations

from pptmaker.slides import _common as C


WHITE_RGB = C.Role.TEXT_INVERSE

# 표지 표준 크기 — Type 스케일만 사용
COVER_TITLE_SIZE = C.Type.COVER_TITLE   # 56pt
COVER_SUBTITLE_SIZE = C.Type.H2         # 20pt


def _force_text(shape, text: str, *, size: int, bold: bool = True,
                color: int = WHITE_RGB, align: int = C.PP_ALIGN_CENTER):
    """텍스트 박스/placeholder에 텍스트 강제 적용 + 폰트·색 표준화."""
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
        # 텍스트 박스 내부 수직 중앙 정렬
        try:
            tf.VerticalAnchor = 3  # msoAnchorMiddle
        except Exception:
            pass
    except Exception:
        pass


def _find_title_textbox(slide):
    """슬라이드 1의 'Title:' 텍스트 박스 찾기.

    Reference에서 이 박스는 TEXT_BOX(type 17)이고 'Title' 텍스트를 가짐.
    못 찾으면 None — caller가 새로 만들어야 함.
    """
    candidates = []
    for shape in slide.Shapes:
        try:
            if shape.Type == 17 and shape.HasTextFrame == C.MSO_TRUE:  # TEXT_BOX
                text = (shape.TextFrame.TextRange.Text or "").strip()
                # Title placeholder 텍스트는 보통 "Title:" 또는 "제목"
                if "Title" in text or "제목" in text or text == "":
                    candidates.append(shape)
        except Exception:
            continue
    # 가장 위쪽에 있는 것을 우선
    if candidates:
        candidates.sort(key=lambda s: (s.Top, s.Left))
        return candidates[0]
    return None


def add_to_live(
    session,
    title: str,
    *,
    presenter: str | None = None,
    organization: str | None = None,
    date: str | None = None,
    subtitle: str | None = None,
):
    """표지 슬라이드 추가 — Reference 1번 슬라이드 복제 + 텍스트 흰색 교체.

    Args:
        title: 발표 제목 (필수)
        presenter: 발표자 이름
        organization: 소속 (presenter와 함께 한 줄로 결합)
        date: 발표 일자 문자열
        subtitle: presenter/organization/date 대신 직접 지정할 부제목
    """
    if session.slide_count < 1:
        raise ValueError("Reference 표지(1번 슬라이드)가 없는 프레젠테이션에서는 표지를 만들 수 없습니다.")

    # 1번 슬라이드 복제 → 바로 뒤에 새 슬라이드 생성
    source = session._prs.Slides(1)
    dup_range = source.Duplicate()

    # 복제본을 맨 끝으로 이동
    target_idx = session.slide_count  # Duplicate 후 카운트가 1 증가한 상태
    try:
        dup_range.MoveTo(target_idx)
    except Exception:
        # MoveTo 실패 시 2번 자리에 그대로 두기 (덜 이상적이지만 동작은 함)
        pass

    # 새 슬라이드 핸들 다시 확보
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

    # 표지 제목/부제목 표준 박스 — 슬라이드 완전 중앙
    title_box_w = 800.0
    title_box_h = 140.0
    title_box_left = (C.SLIDE_W - title_box_w) / 2          # 슬라이드 중앙
    title_box_top = (C.SLIDE_H - title_box_h) / 2 - 24      # 약간 위로

    # 기존 'Title:' 박스가 있으면 표준 위치로 재배치 + 텍스트 교체
    title_box = _find_title_textbox(new_slide)
    if title_box is not None:
        try:
            title_box.Left = title_box_left
            title_box.Top = title_box_top
            title_box.Width = title_box_w
            title_box.Height = title_box_h
        except Exception:
            pass
        _force_text(
            title_box, title,
            size=COVER_TITLE_SIZE, bold=True,
            color=WHITE_RGB, align=C.PP_ALIGN_CENTER,
        )
    else:
        # Fallback — 같은 표준 위치에 새로 생성
        title_box = C.add_textbox(
            new_slide,
            title_box_left, title_box_top, title_box_w, title_box_h,
            text=title, font=C.FONT_MAJOR, size=COVER_TITLE_SIZE, bold=True,
            color=WHITE_RGB, align=C.PP_ALIGN_CENTER,
        )
        try:
            title_box.TextFrame.VerticalAnchor = 3  # 수직 중앙
        except Exception:
            pass

    # 부제목 박스 — 제목 박스 바로 아래, 동일 폭, 중앙 정렬
    if subtitle:
        sub_top = title_box_top + title_box_h + C.Space.MD
        sub_h = 36.0
        sub_box = C.add_textbox(
            new_slide,
            title_box_left, sub_top, title_box_w, sub_h,
            text=subtitle,
            font=C.FONT_MAJOR, size=COVER_SUBTITLE_SIZE, bold=False,
            color=WHITE_RGB, align=C.PP_ALIGN_CENTER,
        )
        try:
            sub_box.TextFrame.VerticalAnchor = 3
        except Exception:
            pass

    session.focus(target_idx)
    return target_idx
