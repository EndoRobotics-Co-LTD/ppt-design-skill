"""마무리 슬라이드 검증 — 현재 PowerPoint 끝에 1장 추가."""
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

    new_idx = session.add_closing(
        message="Thank You",
        presenter="이은상",
        organization="전략기획팀",
        date="2026.05.26",
    )
    print(f"[DONE] 마무리 슬라이드 {new_idx}번 추가.")


if __name__ == "__main__":
    main()
