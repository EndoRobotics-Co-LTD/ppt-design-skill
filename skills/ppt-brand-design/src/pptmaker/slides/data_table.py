"""B2. 표 — Type/Role 스케일만 사용."""
from __future__ import annotations

from pptmaker.slides import _common as C


def add_to_live(
    session,
    title: str,
    rows: list[list[str]],
    *,
    section_no: int | None = None,
    page: int | None = None,
):
    if not rows or not rows[0]:
        raise ValueError("표 데이터가 비어있습니다.")

    slide, idx = C.add_chrome_slide(session, title=C.title_with_section_no(section_no, title))

    n_rows = len(rows)
    n_cols = max(len(r) for r in rows)

    table_w = C.SLIDE_W - 2 * C.Layout.MARGIN_X
    row_h = 36.0
    table_h = row_h * n_rows
    if table_h > C.BODY_H - C.Space.XL:
        row_h = (C.BODY_H - C.Space.XL) / n_rows
        table_h = row_h * n_rows
    top = C.BODY_TOP + (C.BODY_H - table_h) / 2

    table_shape = slide.Shapes.AddTable(n_rows, n_cols, C.Layout.MARGIN_X, top, table_w, table_h)
    table = table_shape.Table

    for r in range(n_rows):
        for c in range(n_cols):
            text = rows[r][c] if c < len(rows[r]) else ""
            cell = table.Cell(r + 1, c + 1)
            tr = cell.Shape.TextFrame.TextRange
            tr.Text = str(text)
            tr.Font.Name = C.FONT_MINOR
            tr.Font.NameFarEast = C.FONT_MINOR
            tr.Font.Size = C.Type.CAPTION
            tr.ParagraphFormat.Alignment = C.PP_ALIGN_LEFT
            if r == 0:
                cell.Shape.Fill.ForeColor.RGB = C.Role.BRAND_PRIMARY
                tr.Font.Color.RGB = C.Role.TEXT_INVERSE
                tr.Font.Bold = C.MSO_TRUE
            else:
                if r % 2 == 0:
                    cell.Shape.Fill.ForeColor.RGB = C.Role.BG_SUBTLE
                else:
                    cell.Shape.Fill.ForeColor.RGB = C.Role.TEXT_INVERSE
                tr.Font.Color.RGB = C.Role.TEXT_PRIMARY

    session.focus(idx)
    return idx
