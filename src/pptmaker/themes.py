"""테마 레지스트리 — 테마 이름과 템플릿 경로의 단일 진입점.

PPTMaker는 2개의 사내 표준 테마를 지원한다:
- **theme1**: 네이비 헤더 바 + 흰색 Bold 타이틀 (formal/corporate)
- **theme2**: 의료기기 사진 풀-블리드 표지 + 사진 띠 헤더 + 흰색 Bold 타이틀 (visual/branded)

두 테마 모두 동일한 14종 본문 레이아웃을 지원하고, chrome만 다르다.
CLAUDE.md §3 듀얼 테마 아키텍처 참조.

⚠️ 핵심 원칙: layouts/ 의 .pptx 파일들은 절대 직접 열지 않는다.
모든 작업은 사본을 만들어 새 경로에서 수행 (LiveSession.new() 사용).
"""
from __future__ import annotations

from datetime import datetime
from pathlib import Path
from typing import Literal

LAYOUTS_DIR = Path(__file__).resolve().parents[2] / "layouts"

ThemeName = Literal["theme1", "theme2"]
VALID_THEMES: frozenset[str] = frozenset({"theme1", "theme2"})
DEFAULT_THEME: str = "theme1"

THEME_DESCRIPTIONS: dict[str, str] = {
    "theme1": "네이비 헤더 바 + 흰색 Bold 타이틀. Formal/Corporate.",
    "theme2": "사진 풀-블리드 표지 + 사진 띠 헤더 + 흰색 Bold 타이틀. Visual/Branded.",
}


def template_path(theme: str | None = None) -> Path:
    """테마 이름으로 템플릿 .pptx 경로를 반환.

    Args:
        theme: 'theme1' | 'theme2'. None 이면 DEFAULT_THEME ('theme1') 사용.

    Raises:
        ValueError: 알 수 없는 테마 이름.
        FileNotFoundError: 템플릿 파일이 layouts/ 에 존재하지 않을 때.
    """
    name = theme or DEFAULT_THEME
    if name not in VALID_THEMES:
        raise ValueError(
            f"Unknown theme: {name!r}. Valid themes: {sorted(VALID_THEMES)}"
        )
    p = LAYOUTS_DIR / f"{name}.pptx"
    if not p.exists():
        raise FileNotFoundError(f"Template missing: {p}")
    return p


def list_themes() -> dict[str, str]:
    """사용 가능한 테마 이름 → 설명 매핑."""
    return dict(THEME_DESCRIPTIONS)


def default_working_path(theme: str | None = None, *, name_hint: str | None = None) -> Path:
    """새 PPT 작업을 위한 기본 사본 경로.

    `~/Documents/PPTMaker/<name>_<timestamp>.pptx` 형식.
    사용자가 working_path를 명시 안 했을 때 자동 사용.
    """
    base = Path.home() / "Documents" / "PPTMaker"
    base.mkdir(parents=True, exist_ok=True)
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    prefix = name_hint or (theme or DEFAULT_THEME)
    # 파일명에 안전한 문자만 (공백 → _)
    safe = "".join(c if c.isalnum() or c in "_-." else "_" for c in prefix)
    return base / f"{safe}_{stamp}.pptx"


def assert_not_template(path: Path | str) -> None:
    """경로가 layouts/ 안이면 ValueError — 템플릿 직접 작업 차단.

    layouts/theme*.pptx 는 SSOT로 보호되어야 하며,
    모든 사용자 작업은 별도 경로의 사본에서 수행되어야 한다.
    """
    p = Path(path).resolve()
    try:
        p.relative_to(LAYOUTS_DIR.resolve())
    except ValueError:
        return  # layouts/ 바깥 — OK
    raise ValueError(
        f"layouts/ 템플릿 영역의 파일은 직접 작업할 수 없습니다: {p}\n"
        f"  → 다른 경로(예: ~/Documents/PPTMaker/내작업.pptx) 를 지정하거나 "
        f"working_path를 비워서 자동 사본 생성을 사용하세요."
    )
