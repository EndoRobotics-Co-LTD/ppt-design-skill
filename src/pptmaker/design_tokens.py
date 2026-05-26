"""design_tokens.json을 읽어 코드에서 안전하게 쓰기 위한 dataclass 래퍼.

이 모듈이 단일 진실 공급원(SSOT). 슬라이드 빌더는 모두 여기서 토큰을 가져온다.
"""
from __future__ import annotations

import json
from dataclasses import dataclass, field
from pathlib import Path
from typing import Final

ROOT = Path(__file__).resolve().parents[2]
TOKENS_PATH: Final = ROOT / "design_tokens.json"


@dataclass(frozen=True)
class SlideSize:
    width_cm: float
    height_cm: float
    aspect: float


@dataclass(frozen=True)
class ColorPalette:
    """Office 기본 테마를 EndoRobotics 표준으로 채택한 컬러."""
    text_dark: str = "#000000"     # dk1
    text_light: str = "#FFFFFF"    # lt1
    dark: str = "#0E2841"          # dk2 (네이비)
    light: str = "#E8E8E8"         # lt2 (그레이)
    accent1: str = "#156082"       # 주 브랜드 (틸 블루)
    accent2: str = "#E97132"
    accent3: str = "#196B24"
    accent4: str = "#0F9ED5"
    accent5: str = "#A02B93"
    accent6: str = "#4EA72E"
    hyperlink: str = "#467886"
    followed_hyperlink: str = "#96607D"


@dataclass(frozen=True)
class Fonts:
    major: str = "맑은 고딕"  # 제목
    minor: str = "맑은 고딕"  # 본문


@dataclass(frozen=True)
class DesignTokens:
    slide_size: SlideSize
    colors: ColorPalette
    fonts: Fonts
    layouts: tuple[str, ...] = field(default_factory=tuple)
    reference_files: tuple[str, ...] = field(default_factory=tuple)


def _load() -> DesignTokens:
    if not TOKENS_PATH.exists():
        # 분석 스크립트를 한 번도 실행하지 않은 경우 — 기본값(2026.02 기준)으로 동작
        return DesignTokens(
            slide_size=SlideSize(33.87, 19.05, 1.7778),
            colors=ColorPalette(),
            fonts=Fonts(),
        )
    raw = json.loads(TOKENS_PATH.read_text(encoding="utf-8"))
    tpl = raw["templates"][0]
    size = tpl["slide_size"]
    colors_raw = tpl["theme_colors"]
    layouts = tuple(layout["name"] for layout in tpl["layouts"])
    references = tuple(t["file"] for t in raw["templates"])

    def hex_or_default(key: str, default: str) -> str:
        v = colors_raw.get(key, default)
        return v if isinstance(v, str) and v.startswith("#") else default

    return DesignTokens(
        slide_size=SlideSize(size["width_cm"], size["height_cm"], size["aspect"]),
        colors=ColorPalette(
            dark=hex_or_default("dk2", "#0E2841"),
            light=hex_or_default("lt2", "#E8E8E8"),
            accent1=hex_or_default("accent1", "#156082"),
            accent2=hex_or_default("accent2", "#E97132"),
            accent3=hex_or_default("accent3", "#196B24"),
            accent4=hex_or_default("accent4", "#0F9ED5"),
            accent5=hex_or_default("accent5", "#A02B93"),
            accent6=hex_or_default("accent6", "#4EA72E"),
            hyperlink=hex_or_default("hlink", "#467886"),
            followed_hyperlink=hex_or_default("folHlink", "#96607D"),
        ),
        fonts=Fonts(),  # 폰트는 항상 맑은 고딕으로 고정
        layouts=layouts,
        reference_files=references,
    )


TOKENS: Final[DesignTokens] = _load()
