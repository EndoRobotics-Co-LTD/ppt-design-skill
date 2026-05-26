"""현재 PowerPoint에 열려 있는 프레젠테이션의 1번 슬라이드 구조를 상세히 출력."""
from __future__ import annotations

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "src"))

from pptmaker.live_session import LiveSession

TEMPLATE = ROOT / "Reference" / "EndoRobotics_PPT Design 01_2026.02.pptx"


def main():
    session = LiveSession.open(TEMPLATE)
    slide = session._prs.Slides(1)
    print(f"Slide 1 — total shapes: {slide.Shapes.Count}")
    for i in range(1, slide.Shapes.Count + 1):
        shape = slide.Shapes(i)
        info = []
        info.append(f"#{i} type={shape.Type}")
        info.append(f"name='{shape.Name}'")
        info.append(f"pos=({shape.Left:.1f},{shape.Top:.1f})")
        info.append(f"size=({shape.Width:.1f},{shape.Height:.1f})")
        try:
            if shape.HasTextFrame == -1:
                txt = shape.TextFrame.TextRange.Text.strip().replace("\r", " | ")[:80]
                info.append(f"text='{txt}'")
        except Exception:
            pass
        try:
            ph_type = shape.PlaceholderFormat.Type
            info.append(f"ph_type={ph_type}")
        except Exception:
            pass
        print(" ".join(info))


if __name__ == "__main__":
    main()
