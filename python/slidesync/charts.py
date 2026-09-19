"""``chart:`` directive, replace a native PowerPoint chart's data from a range."""

from __future__ import annotations

from typing import List

from pptx.chart.data import CategoryChartData, XyChartData
from pptx.enum.chart import XL_CHART_TYPE

from .excel_source import Cell
from .formatting import format_value

_XY_TYPES = {
    XL_CHART_TYPE.XY_SCATTER,
    XL_CHART_TYPE.XY_SCATTER_LINES,
    XL_CHART_TYPE.XY_SCATTER_LINES_NO_MARKERS,
    XL_CHART_TYPE.XY_SCATTER_SMOOTH,
    XL_CHART_TYPE.XY_SCATTER_SMOOTH_NO_MARKERS,
}


def _to_number(value):
    if value is None or value == "":
        return None
    if isinstance(value, (int, float)):
        return float(value)
    try:
        return float(str(value).replace(",", "").replace("%", ""))
    except ValueError:
        return None


def _transpose(rows: List[List[Cell]]) -> List[List[Cell]]:
    width = max(len(r) for r in rows)
    padded = [r + [Cell(None, "General")] * (width - len(r)) for r in rows]
    return [list(col) for col in zip(*padded)]


def update_chart(graphic_frame, rows: List[List[Cell]], orient: str = "cols") -> str:
    """Replace the chart data.

    Layout (``orient='cols'``, default): first column = categories, first row =
    series names, body = values. ``orient='rows'`` transposes the range first.
    """
    if len(rows) < 2 or len(rows[0]) < 2:
        raise ValueError("chart range needs a header row/column plus at least one data cell")
    if orient == "rows":
        rows = _transpose(rows)

    chart = graphic_frame.chart
    header, body = rows[0], rows[1:]
    categories = [format_value(r[0].value, None, r[0].number_format) for r in body]
    series_names = [format_value(h.value, None, h.number_format) or f"Series {i}" for i, h in enumerate(header[1:], 1)]

    if chart.chart_type in _XY_TYPES:
        data = XyChartData()
        for s_idx, name in enumerate(series_names, start=1):
            series = data.add_series(name)
            for r in body:
                x = _to_number(r[0].value)
                y = _to_number(r[s_idx].value) if s_idx < len(r) else None
                if x is not None and y is not None:
                    series.add_data_point(x, y)
    else:
        data = CategoryChartData()
        data.categories = categories
        for s_idx, name in enumerate(series_names, start=1):
            values = [_to_number(r[s_idx].value) if s_idx < len(r) else None for r in body]
            fmt = next((r[s_idx].number_format for r in body if s_idx < len(r) and r[s_idx].number_format != "General"), None)
            data.add_series(name, values, number_format=fmt)

    chart.replace_data(data)
    return f"{len(series_names)} series x {len(body)} points"
