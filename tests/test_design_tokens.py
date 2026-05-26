"""디자인 토큰 로드 / 가드 동작 스모크 테스트.

런처 없이 실행:
    python -m pytest tests/  (또는)  python tests/test_design_tokens.py
"""
from __future__ import annotations

import sys
from pathlib import Path

# src/ 레이아웃이라 PYTHONPATH 추가
ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))

from pptmaker import TOKENS
from pptmaker.theme_guard import audit_color, audit_font


def test_tokens_loaded():
    assert TOKENS.slide_size.aspect == 1.7778
    assert TOKENS.colors.accent1.upper() == "#156082"
    assert TOKENS.fonts.major == "맑은 고딕"
    print("[OK] tokens loaded:", TOKENS.colors.accent1, TOKENS.fonts.major)


def test_audit_allows_brand():
    assert audit_color("#156082", slide_index=1, where="title") is None
    assert audit_font("맑은 고딕", slide_index=1, where="body") is None
    print("[OK] brand color/font passes audit")


def test_audit_flags_off_brand():
    v = audit_color("#FF00FF", slide_index=2, where="body")
    assert v is not None and v.kind == "color"
    v2 = audit_font("Arial", slide_index=2, where="body")
    assert v2 is not None and v2.kind == "font"
    print("[OK] off-brand flagged:", v.detail, "|", v2.detail)


if __name__ == "__main__":
    test_tokens_loaded()
    test_audit_allows_brand()
    test_audit_flags_off_brand()
    print("\nAll smoke tests passed.")
