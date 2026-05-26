"""표지 슬라이드 예시 — Reference 1번 슬라이드 복제 + 흰색 텍스트.

기존에 잘못 만든 표지 슬라이드(끝부분)가 있으면 삭제하고 새로 만든다.
"""
from __future__ import annotations

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "src"))

from pptmaker.live_session import LiveSession

TEMPLATE = ROOT / "Reference" / "EndoRobotics_PPT Design 01_2026.02.pptx"


def _delete_old_test_covers(session, since_idx: int = 12):
    """since_idx 이후 슬라이드를 끝→앞으로 삭제."""
    deleted = 0
    for idx in range(session.slide_count, since_idx - 1, -1):
        slide = session._prs.Slides(idx)
        # 끝 슬라이드의 첫 텍스트가 '2026년 1분기 사업 보고'면 삭제 (이전 테스트 결과)
        is_old_cover = False
        try:
            for shape in slide.Shapes:
                if shape.HasTextFrame == -1:
                    txt = (shape.TextFrame.TextRange.Text or "").strip()
                    if "사업 보고" in txt or txt == "Title:":
                        is_old_cover = True
                        break
        except Exception:
            pass
        if is_old_cover:
            try:
                slide.Delete()
                deleted += 1
            except Exception:
                pass
    return deleted


def main():
    session = LiveSession.open(TEMPLATE)
    print(f"[INFO] 현재 슬라이드 수: {session.slide_count}")

    removed = _delete_old_test_covers(session)
    if removed:
        print(f"[CLEANUP] 이전 표지 테스트 {removed}장 삭제. 현재 {session.slide_count}장")

    print("[STEP] 새 표지 슬라이드 추가 (Reference 1번 복제 + 흰색 텍스트)...")
    new_idx = session.add_cover(
        title="2026년 1분기 사업 보고",
        presenter="이은상",
        organization="전략기획팀",
        date="2026.05.26",
    )

    print(f"[DONE] 표지 슬라이드 {new_idx}번 추가.")
    print(f"       파일은 저장되지 않았습니다.")


if __name__ == "__main__":
    main()
