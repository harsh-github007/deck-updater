"""``table:`` and ``picture:`` directives."""

from __future__ import annotations

import copy
import io
from typing import List

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt  # noqa: E402

from pptx.util import Emu  # noqa: E402

from .excel_source import Cell  # noqa: E402
from .formatting import format_value  # noqa: E402


# --------------------------------------------------------------------------- #
# Native PowerPoint tables
# --------------------------------------------------------------------------- #

def fill_table(graphic_frame, rows: List[List[Cell]], header: bool = True) -> str:
    """Write *rows* into the table shape, resizing it to match.

    Existing cell formatting is kept: extra rows/columns are cloned from the
    last existing row/column, so the template's styling carries over.
    """
    table = graphic_frame.table
    need_rows = len(rows)
    need_cols = max((len(r) for r in rows), default=0)
    if need_rows == 0 or need_cols == 0:
        raise ValueError("range is empty")

    _resize_columns(table, need_cols)
    _resize_rows(table, need_rows, keep_header=header)

    for r_idx, row in enumerate(rows):
        for c_idx in range(need_cols):
            cell = row[c_idx] if c_idx < len(row) else Cell(None, "General")
            _set_cell_text(table.cell(r_idx, c_idx), format_value(cell.value, None, cell.number_format))
    return f"{need_rows}x{need_cols}"


def _resize_columns(table, need: int) -> None:
    tbl = table._tbl
    grid = tbl.tblGrid
    cols = grid.gridCol_lst
    original_total = sum(int(c.w) for c in cols)
    while len(cols) > need:
        grid.remove(cols[-1])
        for tr in tbl.tr_lst:
            tr.remove(tr.tc_lst[-1])
        cols = grid.gridCol_lst
    while len(cols) < need:
        grid.append(copy.deepcopy(cols[-1]))
        for tr in tbl.tr_lst:
            tr.append(copy.deepcopy(tr.tc_lst[-1]))
        cols = grid.gridCol_lst
    # keep the template's overall width: spread it evenly over the new column count
    cols = grid.gridCol_lst
    if len(cols) != need or sum(int(c.w) for c in cols) != original_total:
        each = original_total // len(cols)
        for c in cols:
            c.w = Emu(each)
    table._graphic_frame.width = Emu(sum(int(c.w) for c in grid.gridCol_lst))


def _resize_rows(table, need: int, keep_header: bool) -> None:
    tbl = table._tbl
    trs = tbl.tr_lst
    # never delete the header row or the last body row template while shrinking
    min_keep = 2 if keep_header and need >= 2 else 1
    while len(trs) > max(need, min_keep):
        tbl.remove(trs[-1])
        trs = tbl.tr_lst
    while len(trs) < need:
        tbl.append(copy.deepcopy(trs[-1]))
        trs = tbl.tr_lst
    if len(trs) > need:  # need == 1 with header template of 2 rows
        while len(tbl.tr_lst) > need:
            tbl.remove(tbl.tr_lst[-1])
    total = sum(int(tr.h) for tr in tbl.tr_lst)
    table._graphic_frame.height = Emu(total)


def _set_cell_text(cell, text: str) -> None:
    """Replace the cell's text while keeping the first run's formatting."""
    tf = cell.text_frame
    paragraphs = tf.paragraphs
    first = paragraphs[0]
    runs = first.runs
    if runs:
        runs[0].text = text
        for extra in runs[1:]:
            extra._r.getparent().remove(extra._r)
    else:
        first.add_run().text = text
    for extra_p in paragraphs[1:]:
        extra_p._p.getparent().remove(extra_p._p)


# --------------------------------------------------------------------------- #
# Range rendered as a picture
# --------------------------------------------------------------------------- #

def render_range_png(rows: List[List[Cell]], header: bool = True, dpi: int = 200) -> bytes:
    """Draw the range as a simple table image and return PNG bytes."""
    text = [[format_value(c.value, None, c.number_format) for c in row] for row in rows]
    n_rows, n_cols = len(text), max(len(r) for r in text)
    fig_w = max(2.0, 1.1 * n_cols)
    fig_h = max(0.6, 0.32 * (n_rows + 0.5))
    fig, ax = plt.subplots(figsize=(fig_w, fig_h))
    ax.axis("off")
    tbl = ax.table(cellText=text, loc="center", cellLoc="center")
    tbl.auto_set_font_size(False)
    tbl.set_fontsize(9)
    tbl.scale(1, 1.3)
    for (r, c), cell in tbl.get_celld().items():
        cell.set_edgecolor("#BFBFBF")
        if header and r == 0:
            cell.set_facecolor("#404040")
            cell.get_text().set_color("white")
            cell.get_text().set_weight("bold")
        elif r % 2 == 0:
            cell.set_facecolor("#F2F2F2")
    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=dpi, bbox_inches="tight", pad_inches=0.05, transparent=False)
    plt.close(fig)
    return buf.getvalue()


def replace_picture(slide, shape, png: bytes):
    """Swap *shape* for a new picture with the same name, position, size and z-order.

    The new picture is added at slide level first (python-pptx offers no other
    entry point), then moved into whichever element actually holds the old
    shape, which is a group's shape tree when the picture is grouped, not the
    slide's. If anything fails after the new element exists, it is removed
    again so a failed directive cannot leave a stray picture behind.
    """
    left, top, width, height = shape.left, shape.top, shape.width, shape.height
    name = shape.name
    old_el = shape._element
    parent = old_el.getparent()
    if parent is None:
        raise ValueError("picture is not attached to a slide")
    index = list(parent).index(old_el)

    new_pic = slide.shapes.add_picture(io.BytesIO(png), left, top, width=width, height=height)
    new_el = new_pic._element
    try:
        new_pic.name = name
        new_el.getparent().remove(new_el)
        parent.insert(index, new_el)
        parent.remove(old_el)
    except Exception:
        # Undo the partial mutation: drop the new element wherever it ended up
        # and leave the original picture in place.
        holder = new_el.getparent()
        if holder is not None:
            holder.remove(new_el)
        if old_el.getparent() is None:
            parent.insert(index, old_el)
        raise
    return new_pic
