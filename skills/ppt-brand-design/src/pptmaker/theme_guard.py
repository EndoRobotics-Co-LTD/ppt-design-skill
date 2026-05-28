"""테마 일관성 가드.

슬라이드 빌더가 생성한 결과물이 EndoRobotics 디자인 토큰을 위반하지 않는지 검증.
브랜드 색상 외의 색이 들어가거나, 폰트가 맑은 고딕이 아니면 경고를 띄운다.

이건 강제(rewrite)가 아니라 감사(audit) 도구 — 사용자 자유 편집 슬라이드는
허용해야 하므로 caller가 결과를 보고 결정한다 (medium-control 정책).
"""
from __future__ import annotations

from dataclasses import dataclass

from pptmaker.design_tokens import TOKENS


ALLOWED_FONTS = {TOKENS.fonts.major, TOKENS.fonts.minor}
ALLOWED_COLORS_HEX = {
    TOKENS.colors.text_dark.upper(),
    TOKENS.colors.text_light.upper(),
    TOKENS.colors.dark.upper(),
    TOKENS.colors.light.upper(),
    TOKENS.colors.accent1.upper(),
    TOKENS.colors.accent2.upper(),
    TOKENS.colors.accent3.upper(),
    TOKENS.colors.accent4.upper(),
    TOKENS.colors.accent5.upper(),
    TOKENS.colors.accent6.upper(),
}


@dataclass(frozen=True)
class Violation:
    slide_index: int
    where: str       # e.g. "title", "body", "shape:Picture 3"
    kind: str        # "font" | "color" | "size"
    detail: str


def audit_font(name: str, *, slide_index: int, where: str) -> Violation | None:
    if name and name not in ALLOWED_FONTS:
        return Violation(slide_index, where, "font", f"non-brand font: {name!r}")
    return None


def audit_color(hex_rgb: str, *, slide_index: int, where: str) -> Violation | None:
    if not hex_rgb:
        return None
    norm = hex_rgb.lstrip("#").upper()
    if len(norm) != 6:
        return None
    if f"#{norm}" not in ALLOWED_COLORS_HEX:
        return Violation(slide_index, where, "color", f"non-brand color: #{norm}")
    return None
