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
        force_overwrite: bool = False,
    ) -> "LiveSession":
        """**새 PPT 작업 시작 (권장 진입점).**

        layouts/ 의 템플릿을 working_path 로 **복사**한 다음, 그 사본을 연다.
        템플릿 원본은 절대 안 건드린다. blank_mode=True 자동 적용 (데모 본문 제거).

        **안전 가드**:
        - working_path 파일이 이미 존재하면 FileExistsError (force_overwrite=True 로 덮어쓸 수 있지만,
          PowerPoint에서 사용자가 수동 편집한 내용이 사라질 수 있으니 주의).
        - working_path가 PowerPoint에서 이미 열려있으면 RuntimeError. 후속 수정·추가 작업은
          open_existing() 사용 권장.

        Args:
            theme: 'theme1' | 'theme2'. None이면 DEFAULT_THEME.
            working_path: 작업 사본을 만들 경로. None이면 자동 생성
                ~/Documents/PPTMaker/<name>_<timestamp>.pptx
            name_hint: working_path 자동 생성 시 파일명 prefix (예: "2026Q1_보고").
            force_overwrite: True면 기존 파일이 있어도 덮어씀 (사용자가 명시적으로
                "처음부터 다시 만들어줘"라고 한 경우에만).

        Raises:
            ValueError: working_path가 layouts/ 영역 안이면 차단.
            FileExistsError: working_path가 이미 존재 + force_overwrite=False.
            RuntimeError: working_path가 PowerPoint에 열려있음.
        """
        from pptmaker import themes

        src = themes.template_path(theme)
        if working_path is None:
            working_path = themes.default_working_path(theme, name_hint=name_hint)
        working = Path(working_path).resolve()
        themes.assert_not_template(working)

        # 안전 가드 1: 동일 경로의 .pptx가 이미 있으면 차단
        if working.exists() and not force_overwrite:
            raise FileExistsError(
                f"작업 파일이 이미 존재합니다: {working}\n"
                f"  - 이어서 작업하려면: LiveSession.open_existing(r'{working}')\n"
                f"  - 정말 덮어쓰려면: LiveSession.new(..., force_overwrite=True)\n"
                f"    (주의: PowerPoint에서 수동 편집한 내용이 모두 사라집니다)"
            )

        # 안전 가드 2: PowerPoint에서 이미 열려있으면 차단
        app = cls.connect_or_start()
        target_lower = str(working).lower()
        for prs in app.Presentations:
            try:
                opened_path = str(Path(prs.FullName).resolve()).lower()
            except Exception:
                continue
            if opened_path == target_lower:
                raise RuntimeError(
                    f"이 파일이 PowerPoint에 이미 열려있습니다: {working}\n"
                    f"  - 이어서 작업하려면: LiveSession.open_existing(r'{working}')\n"
                    f"  - 새로 만들려면 먼저 PowerPoint에서 이 파일을 닫아주세요."
                )

        working.parent.mkdir(parents=True, exist_ok=True)
        # 템플릿 → 사본 복사 (원본 보호)
        shutil.copy2(str(src), str(working))

        prs = app.Presentations.Open(str(working), ReadOnly=MSO_FALSE, WithWindow=MSO_TRUE)
        session = cls(app, prs, blank_mode=True)
        session.clean_to_chrome()  # 데모 본문 자동 제거
        return session

    @classmethod
    def open_existing(
        cls,
        path: Path | str,
        *,
        keep_last_slide: bool = True,
    ) -> "LiveSession":
        """**사용자가 기존에 작업하던 PPT를 이어서 작업.**

        이미 PowerPoint에서 열려있으면 그 인스턴스에 연결. 아니면 그 파일을 연다.
        기존 슬라이드는 모두 보존됨 (사용자 수동 편집 포함).

        Args:
            path: 기존 .pptx 파일 절대 경로.
            keep_last_slide: True (기본) — 마지막 슬라이드(통상 closing)를
                보존하며 add_xxx 호출은 그 **직전**에 삽입. PPTMaker 표준 PPT 가정.
                False — add_xxx 가 끝에 append (사용자가 closing 없는 PPT를
                작업하거나 마지막 위치에 새 슬라이드를 두고 싶을 때).

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
                    return cls(app, prs, blank_mode=keep_last_slide)
            except Exception:
                continue
        prs = app.Presentations.Open(str(target), ReadOnly=MSO_FALSE, WithWindow=MSO_TRUE)
        return cls(app, prs, blank_mode=keep_last_slide)

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

    # ---------- 부분 수정 (PPT가 이미 만들어진 후 수정·추가) ----------
    # 핵심 원칙: 콘텐츠 SSOT는 .pptx 자체. 매번 전체 재생성 X.
    # 이미 만들어진 PPT의 일부만 수정할 때 사용. PowerPoint에서 사용자가
    # 수동 편집한 내용은 그대로 보존된다.

    def replace_text_in_slide(self, slide_index: int, find: str, replace: str) -> int:
        """슬라이드 N의 모든 텍스트박스에서 `find` 문자열을 `replace`로 치환.

        Args:
            slide_index: 1-based 슬라이드 번호.
            find: 찾을 문자열 (정확히 일치 — substring match).
            replace: 교체할 문자열.

        Returns:
            치환된 텍스트박스 개수.

        Example:
            >>> session.replace_text_in_slide(5, "+18%", "+22%")  # KPI 카드 숫자 변경
            1
        """
        if slide_index < 1 or slide_index > self.slide_count:
            raise ValueError(
                f"slide_index out of range: {slide_index} (slide_count={self.slide_count})"
            )
        slide = self._prs.Slides(slide_index)
        count = 0
        for shape in slide.Shapes:
            try:
                if shape.HasTextFrame != MSO_TRUE:
                    continue
                tr = shape.TextFrame.TextRange
                original = tr.Text or ""
                if find in original:
                    tr.Text = original.replace(find, replace)
                    count += 1
            except Exception:
                continue
        self.focus(slide_index)
        return count

    def replace_text_global(self, find: str, replace: str) -> int:
        """전체 슬라이드를 대상으로 텍스트 치환. 치환된 텍스트박스 총 개수 반환."""
        total = 0
        for i in range(1, self.slide_count + 1):
            total += self.replace_text_in_slide(i, find, replace)
        return total

    def delete_slide(self, slide_index: int) -> None:
        """슬라이드 N 삭제. 이후 슬라이드들은 자동 앞당겨짐.

        Args:
            slide_index: 1-based 슬라이드 번호.
        """
        if slide_index < 1 or slide_index > self.slide_count:
            raise ValueError(
                f"slide_index out of range: {slide_index} (slide_count={self.slide_count})"
            )
        self._prs.Slides(slide_index).Delete()

    def move_slide(self, from_index: int, to_index: int) -> None:
        """슬라이드 순서 변경. from에서 to 위치로 이동.

        Args:
            from_index: 옮길 슬라이드 (1-based).
            to_index: 목적 위치 (1-based).
        """
        if from_index < 1 or from_index > self.slide_count:
            raise ValueError(f"from_index out of range: {from_index}")
        if to_index < 1 or to_index > self.slide_count:
            raise ValueError(f"to_index out of range: {to_index}")
        if from_index == to_index:
            return
        self._prs.Slides(from_index).MoveTo(to_index)
        self.focus(to_index)

    def get_slide_text(self, slide_index: int) -> str:
        """슬라이드 N의 모든 텍스트박스 텍스트를 줄바꿈으로 연결해서 반환.

        Claude가 "이 슬라이드에 어떤 내용이 있지?"를 자연어로 답하거나
        특정 텍스트를 정확한 형태로 찾기 위해 사용.
        """
        if slide_index < 1 or slide_index > self.slide_count:
            raise ValueError(
                f"slide_index out of range: {slide_index} (slide_count={self.slide_count})"
            )
        slide = self._prs.Slides(slide_index)
        parts: list[str] = []
        for shape in slide.Shapes:
            try:
                if shape.HasTextFrame != MSO_TRUE:
                    continue
                text = (shape.TextFrame.TextRange.Text or "").strip()
                if text:
                    parts.append(text)
            except Exception:
                continue
        return "\n".join(parts)

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

    def register_layout(
        self,
        name: str,
        *,
        theme: str | None = None,
        slide_index: int | None = None,
        placeholders: list[str] | None = None,
        description: str = "",
        overwrite: bool = False,
    ) -> dict:
        """현재 프레젠테이션의 한 슬라이드를 user_layouts 에 커스텀 템플릿으로 등록.

        Phase 2 진입점 — Claude가 사용자와 대화로 시안 슬라이드를 만든 뒤
        이 메서드 한 번 호출로:
          1. user_layouts/<theme>_user.pptx 자동 생성 (없으면)
          2. 지정한 슬라이드를 거기에 append
          3. 슬라이드 노트에 `pptmaker:custom name=... placeholders=...` 메타 추가
          4. manifest.json 갱신

        Args:
            name: 템플릿 이름 (예: "milestone_4step"). 호출자가 향후 add_custom(name, ...)으로 사용.
            theme: 어느 테마용? None이면 DEFAULT_THEME ('theme1').
            slide_index: 어느 슬라이드를 등록할지 (1-based). None이면 마지막 슬라이드.
            placeholders: {{key}} 키 목록. None이면 슬라이드 텍스트에서 자동 감지.
            description: 사람이 읽는 설명 (선택).
            overwrite: 같은 name 이 이미 있으면 덮어쓸지.

        Returns:
            등록 결과 dict (name, theme, source_file, slide_index_in_user_pptx, placeholders, description).
        """
        from pptmaker import custom_templates, themes

        theme_name = theme or themes.DEFAULT_THEME

        # 1. 어느 슬라이드를 등록할지 결정.
        # 기본값: blank_mode면 closing 직전(=가장 최근 본문 슬라이드), 아니면 마지막.
        if slide_index is not None:
            idx = slide_index
        elif getattr(self, "blank_mode", False) and self.slide_count >= 2:
            idx = self.slide_count - 1  # closing 직전 (방금 추가한 시안)
        else:
            idx = self.slide_count
        if idx < 1 or idx > self.slide_count:
            raise ValueError(f"slide_index out of range: {idx} (slide_count={self.slide_count})")
        source_slide = self._prs.Slides(idx)

        # 2. user_layouts/<theme>_user.pptx 준비 (없으면 자동 생성)
        user_pptx_path = custom_templates.create_user_pptx_if_missing(theme_name)

        # 3. placeholders 자동 감지 (명시 안 됐을 때)
        if placeholders is None:
            placeholders = custom_templates.detect_placeholders_in_com_slide(source_slide)

        # 4. 슬라이드를 user_layouts/<theme>_user.pptx 끝에 추가 (PowerPoint COM으로)
        source_slide.Copy()

        user_prs = self._app.Presentations.Open(
            str(user_pptx_path.resolve()),
            ReadOnly=MSO_FALSE,
            WithWindow=MSO_FALSE,
        )
        try:
            # user_prs 끝에 paste
            insert_at = user_prs.Slides.Count + 1
            user_prs.Slides.Paste(insert_at)
            new_slide = user_prs.Slides(insert_at)

            # 5. 노트에 메타 추가
            meta_line = custom_templates.build_meta_line(name, placeholders, description)
            try:
                # NotesPage 의 첫 placeholder가 노트 텍스트
                notes_shape = new_slide.NotesPage.Shapes(2)  # 통상 1=슬라이드 썸네일, 2=노트
                notes_shape.TextFrame.TextRange.Text = meta_line
            except Exception:
                # fallback: 모든 노트 shape 시도
                for sh in new_slide.NotesPage.Shapes:
                    try:
                        if sh.HasTextFrame == MSO_TRUE and sh.PlaceholderFormat.Type == 2:  # ppPlaceholderBody
                            sh.TextFrame.TextRange.Text = meta_line
                            break
                    except Exception:
                        continue

            user_prs.Save()
            slide_index_in_user_pptx = insert_at
        finally:
            user_prs.Close()

        # 6. manifest 갱신
        tmpl = custom_templates.CustomTemplate(
            name=name,
            theme=theme_name,
            source_file=user_pptx_path.name,
            source_slide_index=slide_index_in_user_pptx,
            placeholders=list(placeholders),
            description=description,
        )
        custom_templates.add_to_manifest(tmpl, overwrite=overwrite)

        return {
            "name": name,
            "theme": theme_name,
            "source_file": user_pptx_path.name,
            "source_slide_index": slide_index_in_user_pptx,
            "placeholders": list(placeholders),
            "description": description,
            "user_pptx_path": str(user_pptx_path),
        }

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
