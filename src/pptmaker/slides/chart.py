"""B3. 차트 — Type/Role 스케일만 사용."""
from __future__ import annotations

from pptmaker.slides import _common as C

XL_COLUMN_CLUSTERED = 51
XL_LINE = 4
XL_PIE = 5
XL_BAR_CLUSTERED = 57

CHART_TYPES = {
    "column": XL_COLUMN_CLUSTERED,
    "line": XL_LINE,
    "pie": XL_PIE,
    "bar": XL_BAR_CLUSTERED,
}


def add_to_live(
    session,
    title: str,
    categories: list[str],
    series: list[dict],
    *,
    chart_type: str = "column",
    note: str | None = None,
    page: int | None = None,
):
    if not categories or not series:
        raise ValueError("categories와 series가 모두 필요합니다.")

    slide, idx = C.add_chrome_slide(session, title=title)
    chart_kind = CHART_TYPES.get(chart_type, XL_COLUMN_CLUSTERED)

    chart_w = C.SLIDE_W - 2 * C.Layout.MARGIN_X
    chart_h = C.BODY_H - 40 if note else C.BODY_H - 10
    chart_top = C.BODY_TOP

    chart_shape = slide.Shapes.AddChart2(-1, chart_kind, C.Layout.MARGIN_X, chart_top, chart_w, chart_h, True)
    chart = chart_shape.Chart

    chart_data = chart.ChartData
    chart_data.Activate()
    wb = chart_data.Workbook
    ws = wb.Worksheets(1)
    ws.Range("A1:Z100").ClearContents()

    for j, s in enumerate(series):
        ws.Cells(1, 2 + j).Value = s.get("name", f"Series {j+1}")
    for i, cat in enumerate(categories):
        ws.Cells(2 + i, 1).Value = cat
    for j, s in enumerate(series):
        values = s.get("values", [])
        for i, v in enumerate(values):
            ws.Cells(2 + i, 2 + j).Value = v

    n_rows = len(categories) + 1
    n_cols = len(series) + 1
    last_col = chr(ord("A") + n_cols - 1)
    chart.SetSourceData(f"=Sheet1!$A$1:${last_col}${n_rows}")

    # 시리즈 컬러 — 표준 액센트 시퀀스
    try:
        for j in range(len(series)):
            srs = chart.FullSeriesCollection(j + 1)
            try:
                color = C.Role.ACCENT_SERIES[j % len(C.Role.ACCENT_SERIES)]
                srs.Format.Fill.ForeColor.RGB = color
                srs.Format.Line.ForeColor.RGB = color
            except Exception:
                pass
    except Exception:
        pass

    try:
        chart.HasTitle = C.MSO_FALSE
    except Exception:
        pass

    # 데이터 워크북 닫기
    try:
        wb.Close(SaveChanges=False)
    except Exception:
        try:
            chart_data.BreakLink()
        except Exception:
            pass

    if note:
        C.add_textbox(
            slide, C.Layout.MARGIN_X, C.BODY_BOTTOM - 24, chart_w, 20,
            text=note, font=C.FONT_MINOR, size=C.Type.CAPTION,
            color=C.Role.TEXT_SECONDARY, align=C.PP_ALIGN_LEFT,
        )

    session.focus(idx)
    return idx
