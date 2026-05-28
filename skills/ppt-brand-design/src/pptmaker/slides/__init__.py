"""표준 본문 레이아웃 12종.

각 모듈은 `add_to_live(session, ...)` 함수를 제공한다.
세션은 `pptmaker.live_session.LiveSession` 인스턴스.

그룹 A — 공통 (모든 PPT): agenda, section_divider, single_column, two_column, highlight_quote
그룹 B — 데이터 보고: kpi_cards, data_table, chart, comparison
그룹 C — 시각 자료: text_image, full_image, process_flow
"""
from pptmaker.slides import (
    agenda,
    chart,
    closing,
    comparison,
    cover,
    data_table,
    full_image,
    highlight_quote,
    kpi_cards,
    process_flow,
    section_divider,
    single_column,
    text_image,
    two_column,
)

__all__ = [
    "agenda",
    "chart",
    "closing",
    "comparison",
    "cover",
    "data_table",
    "full_image",
    "highlight_quote",
    "kpi_cards",
    "process_flow",
    "section_divider",
    "single_column",
    "text_image",
    "two_column",
]
