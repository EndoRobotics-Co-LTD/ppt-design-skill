"""PowerPoint COM 연동 PoC.

목적:
1) 실행 중인 PowerPoint에 연결 가능한지 확인
2) 없으면 새로 띄우기
3) Reference 템플릿 한 개를 열고, 빈 본문 슬라이드 1장 추가
4) 추가된 슬라이드에 타이틀+본문 텍스트 채우기
5) 사용자가 화면으로 결과 즉시 확인 가능

이 스크립트가 동작하면 라이브 모드의 기술 전제가 모두 검증된다.
"""
from __future__ import annotations

import sys
import time
from pathlib import Path

try:
    import win32com.client  # type: ignore
except ImportError:
    sys.exit("pywin32가 설치되어 있지 않습니다. python -m pip install pywin32 후 다시 실행하세요.")

ROOT = Path(__file__).resolve().parent.parent
TEMPLATE = ROOT / "Reference" / "EndoRobotics_PPT Design 01_2026.02.pptx"

# PowerPoint MsoTriState
MSO_TRUE = -1
MSO_FALSE = 0

# PpSlideLayout enum (Microsoft.Office.Interop.PowerPoint)
PP_LAYOUT_TITLE = 1          # 제목 슬라이드
PP_LAYOUT_TEXT = 2           # 제목 및 내용
PP_LAYOUT_TITLE_ONLY = 11    # 제목만
PP_LAYOUT_BLANK = 12         # 빈 화면


def get_or_start_powerpoint():
    """이미 켜진 PowerPoint가 있으면 그것을, 없으면 새로 띄운 인스턴스를 반환."""
    try:
        app = win32com.client.GetActiveObject("PowerPoint.Application")
        print("[OK] 기존 PowerPoint 인스턴스에 연결")
    except Exception:
        app = win32com.client.Dispatch("PowerPoint.Application")
        print("[OK] 새 PowerPoint 인스턴스 시작")
    app.Visible = MSO_TRUE
    return app


def open_or_get_presentation(app, path: Path):
    """이미 열려있는 같은 파일이 있으면 그 Presentation을, 아니면 새로 연다."""
    abspath = str(path.resolve())
    for prs in app.Presentations:
        try:
            if str(Path(prs.FullName).resolve()).lower() == abspath.lower():
                print(f"[OK] 이미 열려있는 프레젠테이션 사용: {prs.Name}")
                return prs
        except Exception:
            continue
    prs = app.Presentations.Open(abspath, ReadOnly=MSO_FALSE, WithWindow=MSO_TRUE)
    print(f"[OK] 프레젠테이션 열기: {prs.Name}")
    return prs


def add_body_slide(prs, title_text: str, body_lines: list[str]):
    """'제목 및 내용' 레이아웃으로 슬라이드를 끝에 추가하고 텍스트를 채운다."""
    idx = prs.Slides.Count + 1
    # 슬라이드 마스터의 layout 중 '제목 및 내용' 찾기 (커스텀 마스터일 수 있으므로 이름 기반)
    layout = None
    try:
        designs = prs.Designs
        for d in designs:
            for cl in d.SlideMaster.CustomLayouts:
                if "제목 및 내용" in cl.Name or "Title and Content" in cl.Name:
                    layout = cl
                    break
            if layout:
                break
    except Exception:
        pass

    if layout is not None:
        slide = prs.Slides.AddSlide(idx, layout)
    else:
        slide = prs.Slides.Add(idx, PP_LAYOUT_TEXT)
    print(f"[OK] 슬라이드 {idx} 추가 (레이아웃: {slide.Layout if hasattr(slide, 'Layout') else 'custom'})")

    # 제목 채우기
    if slide.Shapes.HasTitle == MSO_TRUE:
        slide.Shapes.Title.TextFrame.TextRange.Text = title_text

    # 본문 placeholder 찾기 (인덱스 2 = body 가 일반적이지만 안전하게 순회)
    for shape in slide.Shapes:
        if shape.HasTextFrame != MSO_TRUE:
            continue
        if shape.Type == 14 and shape != slide.Shapes.Title:  # msoPlaceholder
            tr = shape.TextFrame.TextRange
            tr.Text = "\r".join(body_lines)
            break

    return slide


def main():
    if not TEMPLATE.exists():
        sys.exit(f"Reference 템플릿을 찾을 수 없습니다: {TEMPLATE}")

    app = get_or_start_powerpoint()
    prs = open_or_get_presentation(app, TEMPLATE)

    print(f"[INFO] 현재 슬라이드 수: {prs.Slides.Count}")

    print("[STEP] 본문 슬라이드 1장 추가 중...")
    add_body_slide(
        prs,
        title_text="PPTMaker COM 연동 PoC",
        body_lines=[
            "이 슬라이드는 Claude Code가 PowerPoint COM으로 추가했습니다",
            "디자인 토큰: 맑은 고딕 / accent1 #156082",
            "라이브 모드 검증 시점: " + time.strftime("%Y-%m-%d %H:%M:%S"),
        ],
    )

    # 마지막 슬라이드로 포커스 이동 (사용자가 결과를 바로 보도록)
    try:
        prs.Windows(1).View.GotoSlide(prs.Slides.Count)
    except Exception:
        pass

    print("[DONE] PowerPoint 화면을 확인하세요. 파일은 자동 저장하지 않았습니다.")


if __name__ == "__main__":
    main()
