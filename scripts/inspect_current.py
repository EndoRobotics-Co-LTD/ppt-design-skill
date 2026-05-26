"""현재 PowerPoint에 열린 프레젠테이션의 각 슬라이드 요약."""
from __future__ import annotations

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "src"))

from pptmaker.live_session import LiveSession

TEMPLATE = ROOT / "Reference" / "EndoRobotics_PPT Design 01_2026.02.pptx"


def main():
    session = LiveSession.open(TEMPLATE)
    print(f"Total slides: {session.slide_count}\n")
    for i in range(1, session.slide_count + 1):
        slide = session._prs.Slides(i)
        n_shapes = slide.Shapes.Count
        try:
            layout = slide.CustomLayout.Name
        except Exception:
            layout = "?"
        title = ""
        try:
            if slide.Shapes.HasTitle == -1:
                title = slide.Shapes.Title.TextFrame.TextRange.Text.strip()[:40]
        except Exception:
            pass
        # 텍스트 일부 수집
        texts = []
        for shape in slide.Shapes:
            try:
                if shape.HasTextFrame == -1:
                    t = (shape.TextFrame.TextRange.Text or "").strip()
                    if t:
                        texts.append(t.replace("\r", " | ")[:50])
            except Exception:
                pass
        snippet = " // ".join(texts[:3])[:120]
        print(f"#{i}: layout='{layout}' shapes={n_shapes} title='{title}'")
        print(f"     text: {snippet}")


if __name__ == "__main__":
    main()
