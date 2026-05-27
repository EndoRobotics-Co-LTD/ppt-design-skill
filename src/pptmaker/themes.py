"""테마 레지스트리 — 테마 이름과 템플릿 경로의 단일 진입점.

PPTMaker는 2개의 사내 표준 테마를 지원한다:
- **theme1**: 네이비 헤더 바 + 흰색 Bold 타이틀 (formal/corporate)
- **theme2**: 의료기기 사진 풀-블리드 표지 + 사진 띠 헤더 + 흰색 Bold 타이틀 (visual/branded)

두 테마 모두 동일한 14종 본문 레이아웃을 지원하고, chrome만 다르다.
CLAUDE.md §3 듀얼 테마 아키텍처 참조.
"""
from __future__ import annotations

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
