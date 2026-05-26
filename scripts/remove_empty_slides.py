"""실질적으로 비어있는 슬라이드 자동 식별 + 삭제.

판단 기준: '제목만' 레이아웃이면서 (텍스트 있는 도형이 0개 OR 모두 공백)
→ chrome만 있고 사용자 콘텐츠가 없는 슬라이드.

cover/closing(제목 슬라이드 레이아웃)은 절대 건드리지 않는다.
"""
from __future__ import annotations

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "src"))

from pptmaker.live_session import LiveSession

TEMPLATE = ROOT / "Reference" / "EndoRobotics_PPT Design 01_2026.02.pptx"


def is_empty_body_slide(slide) -> bool:
    """슬라이드가 '제목만' 레이아웃 + 콘텐츠 없음인지 판정."""
    try:
        layout_name = slide.CustomLayout.Name
    except Exception:
        return False
    if "제목만" not in layout_name and "Title Only" not in layout_name:
        return False

    # 텍스트가 의미 있게 채워진 도형이 하나라도 있으면 비어있지 않음
    for shape in slide.Shapes:
        try:
            if shape.HasTextFrame == -1:
                txt = (shape.TextFrame.TextRange.Text or "").strip()
                if txt:
                    return False
        except Exception:
            pass
    return True


def main():
    session = LiveSession.open(TEMPLATE)
    print(f"[INFO] 현재 슬라이드 수: {session.slide_count}")

    # 끝→앞 순서로 삭제 (인덱스 시프트 방지)
    targets = []
    for idx in range(1, session.slide_count + 1):
        slide = session._prs.Slides(idx)
        if is_empty_body_slide(slide):
            targets.append(idx)

    if not targets:
        print("[INFO] 비어있는 슬라이드 없음.")
        return

    print(f"[INFO] 빈 슬라이드 발견: {targets}")
    for idx in reversed(targets):
        try:
            session._prs.Slides(idx).Delete()
            print(f"  - 슬라이드 {idx} 삭제")
        except Exception as e:
            print(f"  - 슬라이드 {idx} 삭제 실패: {e}")

    print(f"[DONE] 현재 슬라이드 수: {session.slide_count}")


if __name__ == "__main__":
    main()
