"""A1. 목차 (Agenda) — Reference 2단 컬럼 + 가로 구분선 디자인.

Type/Role/Space/Layout 스케일만 사용.
"""
from __future__ import annotations

from pptmaker.slides import _common as C


def add_to_live(session, items: list[str], *, title: str = "목차", page: int | None = None):
    """목차 슬라이드 추가.

    Reference 표준 디자인: 항목을 2단 컬럼에 분배, 각 항목 = "번호 + 제목 + 아래 가로 구분선".
    항목 수가 적으면 좌측 컬럼 우선으로 채워짐.
    """
    slide, idx = C.add_chrome_slide(session, title=title)

    n = len(items)
    if n == 0:
        session.focus(idx)
        return idx

    # 2단 컬럼 분배: 왼쪽이 우선 (ceil(n/2)개), 나머지 오른쪽
    left_count = (n + 1) // 2
    left_items = items[:left_count]
    right_items = items[left_count:]

    # 좌표 계산 — Reference 디자인 비례
    gutter = C.Layout.COLUMN_GUTTER
    inner_w = C.SLIDE_W - 2 * C.Layout.MARGIN_X - gutter
    col_w = inner_w / 2
    left_x = C.Layout.MARGIN_X
    right_x = left_x + col_w + gutter

    # 행 높이 결정 — 최대 5행 가정해서 균등 분배 (한 컬럼당)
    available_h = C.BODY_H - C.Space.LG
    rows_per_col = max(1, max(left_count, len(right_items)))
    row_h = min(64.0, available_h / rows_per_col)
    start_y = C.BODY_TOP + C.Space.LG

    # 좌측 컬럼
    for i, item in enumerate(left_items):
        y = start_y + i * row_h
        C.draw_toc_item(slide, left_x, y, col_w, i + 1, item, row_height=row_h)
    # 우측 컬럼
    for i, item in enumerate(right_items):
        y = start_y + i * row_h
        C.draw_toc_item(slide, right_x, y, col_w, left_count + i + 1, item, row_height=row_h)

    session.focus(idx)
    return idx
