"""라이브 모드: 실행 중인 PowerPoint를 pywin32 COM으로 조작.

LiveSession 객체 하나로 한 세션을 표현. with 블록으로 사용 권장:

    with LiveSession.open(template_path) as session:
        session.add_body_slide("제목", ["본문1", "본문2"])

핵심 원칙:
- 사용자의 수동 편집을 덮어쓰지 않기 위해 매 작업 전 현재 상태를 읽는다.
- 저장은 명시 호출시에만 (자동 저장 X). 안전을 위한 기본값.
- Windows 외 환경에서는 ImportError가 발생 — caller가 mode를 분기해야 한다.
"""
from __future__ import annotations

import sys
from dataclasses import dataclass
from pathlib import Path
from typing import Iterator

if sys.platform != "win32":
    raise ImportError("LiveSession은 Windows + PowerPoint 환경에서만 사용 가능합니다.")

import win32com.client  # type: ignore

MSO_TRUE = -1
MSO_FALSE = 0

# PpSlideLayout
PP_LAYOUT_TEXT = 2          # 제목 및 내용
PP_LAYOUT_TITLE_ONLY = 11   # 제목만
PP_LAYOUT_BLANK = 12        # 빈 화면

# msoShapeType
MSO_PLACEHOLDER = 14


@dataclass
class SlideRef:
    """슬라이드를 참조하는 가벼운 핸들."""
    index: int  # 1-based
    layout_name: str


class LiveSession:
    """실행 중 PowerPoint Presentation에 대한 컨텍스트."""

    def __init__(self, app, presentation):
        self._app = app
        self._prs = presentation

    # ---------- 진입/연결 ----------

    @classmethod
    def connect_or_start(cls):
        try:
            app = win32com.client.GetActiveObject("PowerPoint.Application")
        except Exception:
            app = win32com.client.Dispatch("PowerPoint.Application")
        app.Visible = MSO_TRUE
        return app

    @classmethod
    def open(cls, template_path: Path | str) -> "LiveSession":
        """템플릿을 열거나, 이미 열려있다면 그것에 연결한다."""
        app = cls.connect_or_start()
        target = str(Path(template_path).resolve())
        for prs in app.Presentations:
            try:
                if str(Path(prs.FullName).resolve()).lower() == target.lower():
                    return cls(app, prs)
            except Exception:
                continue
        prs = app.Presentations.Open(target, ReadOnly=MSO_FALSE, WithWindow=MSO_TRUE)
        return cls(app, prs)

    def __enter__(self) -> "LiveSession":
        return self

    def __exit__(self, exc_type, exc, tb) -> None:
        # 저장하지 않음. 사용자가 명시적으로 save()/save_as() 호출해야 함.
        return None

    # ---------- 읽기 ----------

    @property
    def slide_count(self) -> int:
        return int(self._prs.Slides.Count)

    def slides(self) -> Iterator[SlideRef]:
        for i in range(1, self.slide_count + 1):
            s = self._prs.Slides(i)
            yield SlideRef(index=i, layout_name=str(s.CustomLayout.Name) if hasattr(s, "CustomLayout") else "")

    def title_of(self, index: int) -> str:
        s = self._prs.Slides(index)
        if s.Shapes.HasTitle == MSO_TRUE:
            return str(s.Shapes.Title.TextFrame.TextRange.Text)
        return ""

    # ---------- 쓰기 ----------

    def _find_custom_layout(self, *needles: str):
        """슬라이드 마스터의 커스텀 레이아웃 중 이름에 needle 중 하나라도 포함된 것."""
        for d in self._prs.Designs:
            for cl in d.SlideMaster.CustomLayouts:
                name = str(cl.Name)
                if any(n in name for n in needles):
                    return cl
        return None

    def add_body_slide(self, title: str, body_lines: list[str]) -> SlideRef:
        """'제목 및 내용' 레이아웃 슬라이드를 끝에 추가."""
        idx = self.slide_count + 1
        layout = self._find_custom_layout("제목 및 내용", "Title and Content")
        if layout is not None:
            slide = self._prs.Slides.AddSlide(idx, layout)
        else:
            slide = self._prs.Slides.Add(idx, PP_LAYOUT_TEXT)

        if slide.Shapes.HasTitle == MSO_TRUE:
            slide.Shapes.Title.TextFrame.TextRange.Text = title

        for shape in slide.Shapes:
            if shape.HasTextFrame != MSO_TRUE:
                continue
            if shape.Type == MSO_PLACEHOLDER and shape != slide.Shapes.Title:
                shape.TextFrame.TextRange.Text = "\r".join(body_lines)
                break

        self.focus(idx)
        return SlideRef(index=idx, layout_name=str(slide.CustomLayout.Name))

    def add_title_only_slide(self, title: str) -> SlideRef:
        """간지/구역 머리글용 — '제목만' 레이아웃."""
        idx = self.slide_count + 1
        layout = self._find_custom_layout("제목만", "Title Only")
        if layout is not None:
            slide = self._prs.Slides.AddSlide(idx, layout)
        else:
            slide = self._prs.Slides.Add(idx, PP_LAYOUT_TITLE_ONLY)
        if slide.Shapes.HasTitle == MSO_TRUE:
            slide.Shapes.Title.TextFrame.TextRange.Text = title
        self.focus(idx)
        return SlideRef(index=idx, layout_name=str(slide.CustomLayout.Name))

    def focus(self, index: int) -> None:
        """현재 창에서 해당 슬라이드로 포커스 이동 — 사용자가 즉시 결과를 보도록."""
        try:
            self._prs.Windows(1).View.GotoSlide(index)
        except Exception:
            pass

    def save(self) -> None:
        self._prs.Save()

    def save_as(self, path: Path | str) -> None:
        self._prs.SaveAs(str(Path(path).resolve()))

    # ---------- 표준 레이아웃 편의 메서드 ----------
    # 각 메서드는 pptmaker.slides 모듈의 함수를 래핑한다.

    def add_cover(self, title: str, *, presenter: str | None = None,
                  organization: str | None = None, date: str | None = None,
                  subtitle: str | None = None) -> int:
        from pptmaker.slides import cover
        return cover.add_to_live(self, title, presenter=presenter,
                                 organization=organization, date=date, subtitle=subtitle)

    def add_closing(self, message: str = "Thank You", *,
                    presenter: str | None = None, organization: str | None = None,
                    date: str | None = None, subtitle: str | None = None) -> int:
        from pptmaker.slides import closing
        return closing.add_to_live(self, message, presenter=presenter,
                                   organization=organization, date=date, subtitle=subtitle)

    def add_agenda(self, items: list[str], *, title: str = "목차", page: int | None = None) -> int:
        from pptmaker.slides import agenda
        return agenda.add_to_live(self, items, title=title, page=page)

    def add_section_divider(self, title: str, *, number: int | str | None = None,
                            subtitle: str | None = None, page: int | None = None) -> int:
        from pptmaker.slides import section_divider
        return section_divider.add_to_live(self, title, number=number, subtitle=subtitle, page=page)

    def add_single_column(self, title: str, bullets: list, *,
                          eyebrow: str | None = None, page: int | None = None) -> int:
        from pptmaker.slides import single_column
        return single_column.add_to_live(self, title, bullets, eyebrow=eyebrow, page=page)

    def add_two_column(self, title: str, left_sub: str, left_bullets: list[str],
                       right_sub: str, right_bullets: list[str], *, page: int | None = None) -> int:
        from pptmaker.slides import two_column
        return two_column.add_to_live(self, title, left_sub, left_bullets,
                                       right_sub, right_bullets, page=page)

    def add_highlight_quote(self, message: str, *, attribution: str | None = None,
                            page: int | None = None) -> int:
        from pptmaker.slides import highlight_quote
        return highlight_quote.add_to_live(self, message, attribution=attribution, page=page)

    def add_kpi_cards(self, title: str, cards: list[dict], *, page: int | None = None) -> int:
        from pptmaker.slides import kpi_cards
        return kpi_cards.add_to_live(self, title, cards, page=page)

    def add_data_table(self, title: str, rows: list[list[str]], *, page: int | None = None) -> int:
        from pptmaker.slides import data_table
        return data_table.add_to_live(self, title, rows, page=page)

    def add_chart(self, title: str, categories: list[str], series: list[dict], *,
                  chart_type: str = "column", note: str | None = None, page: int | None = None) -> int:
        from pptmaker.slides import chart
        return chart.add_to_live(self, title, categories, series,
                                  chart_type=chart_type, note=note, page=page)

    def add_comparison(self, title: str, left: dict, right: dict, *,
                       arrow_label: str | None = None, page: int | None = None) -> int:
        from pptmaker.slides import comparison
        return comparison.add_to_live(self, title, left, right,
                                       arrow_label=arrow_label, page=page)

    def add_text_image(self, title: str, image_path, subtitle: str, bullets: list[str], *,
                       image_side: str = "left", page: int | None = None) -> int:
        from pptmaker.slides import text_image
        return text_image.add_to_live(self, title, image_path, subtitle, bullets,
                                       image_side=image_side, page=page)

    def add_full_image(self, image_path, *, caption: str | None = None,
                       title: str | None = None, page: int | None = None) -> int:
        from pptmaker.slides import full_image
        return full_image.add_to_live(self, image_path, caption=caption, title=title, page=page)

    def add_process_flow(self, title: str, steps: list[dict], *, page: int | None = None) -> int:
        from pptmaker.slides import process_flow
        return process_flow.add_to_live(self, title, steps, page=page)
