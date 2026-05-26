"""슬라이드 1~6 (Reference 원본) 삭제. 생성된 슬라이드만 남긴다.

⚠️ 주의: 슬라이드 1은 cover/closing add_to_live가 배경 복제용 템플릿으로 사용한다.
이걸 삭제하면 같은 세션 내 추가 호출 시 표지/마무리 디자인을 못 만든다.
다시 만들고 싶으면 Reference .pptx를 새로 열어 시작해야 함.
"""
from __future__ import annotations

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "src"))

from pptmaker.live_session import LiveSession

TEMPLATE = ROOT / "Reference" / "EndoRobotics_PPT Design 01_2026.02.pptx"


def main():
    session = LiveSession.open(TEMPLATE)
    print(f"[INFO] 현재 슬라이드 수: {session.slide_count}")

    if session.slide_count <= 6:
        print("[WARN] 삭제할 슬라이드가 충분하지 않습니다.")
        return

    deleted = 0
    for idx in range(6, 0, -1):  # 6, 5, 4, 3, 2, 1
        try:
            session._prs.Slides(idx).Delete()
            deleted += 1
        except Exception as e:
            print(f"[WARN] 슬라이드 {idx} 삭제 실패: {e}")

    print(f"[DONE] {deleted}장 삭제. 현재 슬라이드 수: {session.slide_count}")
    print("       (남은 슬라이드들이 1번부터 재번호됨)")


if __name__ == "__main__":
    main()
