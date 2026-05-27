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

import shutil
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
    """실행 중 PowerPoint Presentation에 대한 컨텍스트.

    blank_mode: True면 데모 슬라이드(2 ~ N-1)가 제거되고 cover/closing 템플릿만 남는다.
                이 모드에서는 add_chrome_slide가 closing 직전에 본문을 삽입.
    """

    def __init__(self, app, presentation, *, blank_mode: bool = False):
        self._app = app
        self._prs = presentation
        self.blank_mode = blank_mode

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
    def new(
        cls,
        *,
        theme: str | None = None,
        working_path: Path | str | None = None,
        name_hint: str | None = None,
    ) -> "LiveSession":
        """**새 PPT 작업 시작 (권장 진입점).**

        layouts/ 의 템플릿을 working_path 로 **복사**한 다음, 그 사본을 연다.
        템플릿 원본은 절대 안 건드린다. blank_mode=True 자동 적용 (데모 본문 제거).

        Args:
            theme: 'theme1' | 'theme2'. None이면 DEFAULT_THEME.
            working_path: 작업 사본을 만들 경로. None이면 자동 생성
                ~/Documents/PPTMaker/<name>_<timestamp>.pptx
            name_hint: working_path 자동 생성 시 파일명 prefix (예: "2026Q1_보고").

        Raises:
            ValueError: working_path가 layouts/ 영역 안이면 차단.
        """
        from pptmaker import themes

        src = themes.template_path(theme)
        if working_path is None:
            working_path = themes.default_working_path(theme, name_hint=name_hint)
        working = Path(working_path).resolve()
        themes.assert_not_template(working)
        working.parent.mkdir(parents=True, exist_ok=True)
        # 템플릿 → 사본 복사 (원본 보호)
        shutil.copy2(str(src), str(working))

        app = cls.connect_or_start()
        prs = app.Presentations.Open(str(working), ReadOnly=MSO_FALSE, WithWindow=MSO_TRUE)
        session = cls(app, prs, blank_mode=True)
        session.clean_to_chrome()  # 데모 본문 자동 제거
        return session

    @classmethod
    def open_existing(cls, path: Path | str) -> "LiveSession":
        """**사용자가 기존에 작업하던 PPT를 이어서 작업.**

        이미 PowerPoint에서 열려있으면 그 인스턴스에 연결. 아니면 그 파일을 연다.
        blank_mode=False — 기존 슬라이드 보존, 끝에 append하는 기본 동작.

        Args:
            path: 기존 .pptx 파일 절대 경로.

        Raises:
            ValueError: layouts/ 템플릿을 가리키면 차단.
            FileNotFoundError: 파일 없음.
        """
        from pptmaker import themes

        target = Path(path).resolve()
        themes.assert_not_template(target)
        if not target.exists():
            raise FileNotFoundError(f"파일 없음: {target}")

        app = cls.connect_or_start()
        # 이미 열려있는지 확인 → 그 인스턴스 재사용
        for prs in app.Presentations:
            try:
                if str(Path(prs.FullName).resolve()).lower() == str(target).lower():
                    return cls(app, prs, blank_mode=False)
            except Exception:
                continue
        prs = app.Presentations.Open(str(target), ReadOnly=MSO_FALSE, WithWindow=MSO_TRUE)
        return cls(app, prs, blank_mode=False)

    @classmethod
    def open(
        cls,
        template_path: Path | str | None = None,
        theme: str | None = None,
        *,
        from_blank: bool = False,
    ) -> "LiveSession":
        """⚠️ Legacy API — 가능하면 `new()` 또는 `open_existing()` 사용.

        template_path/theme 으로 layouts/ 의 템플릿을 직접 열어버린다.
        템플릿 보호를 위해 새 API로 마이그레이션 권장.
        """
        from pptmaker import themes

        if template_path is None:
            template_path = themes.template_path(theme)
        app = cls.connect_or_start()
        target = str(Path(template_path).resolve())
        session = None
        for prs in app.Presentations:
            try:
                if str(Path(prs.FullName).resolve()).lower() == target.lower():
                    session = cls(app, prs, blank_mode=from_blank)
                    break
            except Exception:
                continue
        if session is None:
            prs = app.Presentations.Open(target, ReadOnly=MSO_FALSE, WithWindow=MSO_TRUE)
            session = cls(app, prs, blank_mode=from_blank)
        if from_blank:
            session.clean_to_chrome()
        return session

    def clean_to_chrome(self) -> None:
        """데모 본문 슬라이드(2 ~ N-1) 제거. cover(1) + closing(마지막)만 보존.

        이후 add_chrome_slide는 closing 직전에 본문을 삽입하므로 순서가 자동 유지됨.
        """
        n = self.slide_count
        if n < 3:
            return  # 본문이 없음 — 정리 불필요
        # 역순 삭제로 인덱스 안정성 유지
        for idx in range(n - 1, 1, -1):
            try:
                self._prs.Slides(idx).Delete()
            except Exception:
                pass
        self.blank_mode = True

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

    def add_single_column(self, title: str, bullets: list | None = None, *,
                          tiles: list | None = None,
                          eyebrow: str | None = None,
                          subtitle: str | None = None,
                          section_no: int | None = None,
                          page: int | None = None) -> int:
        """1단 본문 슬라이드. tiles 지정 시 Reference 표준 타일 카드, 아니면 불릿 리스트."""
        from pptmaker.slides import single_column
        return single_column.add_to_live(self, title, bullets, tiles=tiles,
                                         eyebrow=eyebrow, subtitle=subtitle,
                                         section_no=section_no, page=page)

    def add_two_column(self, title: str, left_sub: str, left_items: list,
                       right_sub: str, right_items: list, *,
                       use_tiles: bool = True,
                       section_no: int | None = None,
                       page: int | None = None) -> int:
        """2단 본문 슬라이드. use_tiles=True (기본) = Reference 원형 아이콘 + 타일 카드."""
        from pptmaker.slides import two_column
        return two_column.add_to_live(self, title, left_sub, left_items,
                                       right_sub, right_items,
                                       use_tiles=use_tiles, section_no=section_no, page=page)

    def add_highlight_quote(self, message: str, *, attribution: str | None = None,
                            section_no: int | None = None,
                            page: int | None = None) -> int:
        from pptmaker.slides import highlight_quote
        return highlight_quote.add_to_live(self, message, attribution=attribution,
                                            section_no=section_no, page=page)

    def add_kpi_cards(self, title: str, cards: list[dict], *,
                      section_no: int | None = None,
                      page: int | None = None) -> int:
        from pptmaker.slides import kpi_cards
        return kpi_cards.add_to_live(self, title, cards, section_no=section_no, page=page)

    def add_data_table(self, title: str, rows: list[list[str]], *,
                       section_no: int | None = None,
                       page: int | None = None) -> int:
        from pptmaker.slides import data_table
        return data_table.add_to_live(self, title, rows, section_no=section_no, page=page)

    def add_chart(self, title: str, categories: list[str], series: list[dict], *,
                  chart_type: str = "column", note: str | None = None,
                  section_no: int | None = None,
                  page: int | None = None) -> int:
        from pptmaker.slides import chart
        return chart.add_to_live(self, title, categories, series,
                                  chart_type=chart_type, note=note,
                                  section_no=section_no, page=page)

    def add_comparison(self, title: str, left: dict, right: dict, *,
                       arrow_label: str | None = None,
                       section_no: int | None = None,
                       page: int | None = None) -> int:
        from pptmaker.slides import comparison
        return comparison.add_to_live(self, title, left, right,
                                       arrow_label=arrow_label,
                                       section_no=section_no, page=page)

    def add_text_image(self, title: str, image_path, subtitle: str, bullets: list[str], *,
                       image_side: str = "left",
                       section_no: int | None = None,
                       page: int | None = None) -> int:
        from pptmaker.slides import text_image
        return text_image.add_to_live(self, title, image_path, subtitle, bullets,
                                       image_side=image_side,
                                       section_no=section_no, page=page)

    def add_full_image(self, image_path, *, caption: str | None = None,
                       title: str = "이미지",
                       section_no: int | None = None,
                       page: int | None = None) -> int:
        from pptmaker.slides import full_image
        return full_image.add_to_live(self, image_path, caption=caption, title=title,
                                       section_no=section_no, page=page)

    def add_process_flow(self, title: str, steps: list[dict], *,
                         section_no: int | None = None,
                         page: int | None = None) -> int:
        from pptmaker.slides import process_flow
        return process_flow.add_to_live(self, title, steps,
                                         section_no=section_no, page=page)

    # ---------- 사용자 커스텀 레이아웃 ----------

    def list_custom_templates(self, theme: str | None = None) -> list:
        """사용자가 user_layouts/ 에 등록한 커스텀 템플릿 목록 반환.

        Args:
            theme: None이면 모든 테마, 지정 시 그 테마만.
        Returns:
            CustomTemplate 객체 리스트 (name, source_file, placeholders 등).
        """
        from pptmaker import custom_templates
        return custom_templates.list_templates(theme=theme)

    def add_custom(self, name: str, mapping: dict[str, str] | None = None,
                   *, theme: str | None = None) -> SlideRef:
        """사용자 커스텀 레이아웃 슬라이드를 closing 직전에 추가.

        user_layouts/<theme>_user.pptx 안의 해당 이름 슬라이드를 PowerPoint
        프로세스 내에서 복제(임시 열기 → 복사 → 닫기)한 뒤, 현 프레젠테이션의
        closing 직전 위치에 삽입하고 {{key}} placeholder를 mapping[key]로 치환.

        Args:
            name: manifest 또는 슬라이드 노트에 등록된 템플릿 이름.
            mapping: {placeholder_key: text} dict. None이면 치환 없음 (raw 복제).
            theme: 명시 시 그 테마의 user .pptx 만 검색. None이면 자동.

        Raises:
            ValueError: 등록되지 않은 name 또는 user_layouts/ 파일 없음.
        """
        from pptmaker import custom_templates, themes

        tmpl = custom_templates.find_template(name, theme=theme)
        if tmpl is None:
            available = [t.name for t in custom_templates.list_templates(theme=theme)]
            raise ValueError(
                f"커스텀 템플릿 '{name}'을 찾을 수 없습니다.\n"
                f"  사용 가능: {available or '(없음)'}\n"
                f"  user_layouts/<theme>_user.pptx 안의 슬라이드 노트에 "
                f"'pptmaker:custom name={name} placeholders=...' 메타가 있어야 합니다."
            )

        src_path = themes.USER_LAYOUTS_DIR / tmpl.source_file
        if not src_path.exists():
            raise ValueError(f"user_layouts 파일 없음: {src_path}")

        # 소스 .pptx를 PowerPoint COM으로 열어 슬라이드 복사 → 현 prs로 paste
        src_prs = self._app.Presentations.Open(
            str(src_path.resolve()),
            ReadOnly=MSO_TRUE,
            WithWindow=MSO_FALSE,
        )
        try:
            src_slide = src_prs.Slides(tmpl.source_slide_index)
            src_slide.Copy()
        finally:
            src_prs.Close()

        # 현 프레젠테이션의 closing 직전 위치에 paste
        # blank_mode 가 True 면 closing이 마지막 슬라이드 → 그 직전(=slide_count)
        # 아니면 끝에 추가
        if getattr(self, "blank_mode", False) and self.slide_count >= 1:
            insert_at = self.slide_count  # closing 자리, 삽입 후 closing이 뒤로 밀림
        else:
            insert_at = self.slide_count + 1
        pasted = self._prs.Slides.Paste(insert_at)
        new_slide = self._prs.Slides(insert_at)

        # placeholder 치환
        if mapping:
            n = custom_templates.replace_placeholders_com(new_slide, mapping)
        else:
            n = 0

        self.focus(insert_at)
        return SlideRef(
            index=insert_at,
            layout_name=str(new_slide.CustomLayout.Name) if hasattr(new_slide, "CustomLayout") else f"custom:{name}",
        )
