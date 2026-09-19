#!/usr/bin/env python3
"""Populate the stock evaluation workbook for one company, and recalculate it.

The recalculation is the part that matters. Deck Updater reads a workbook's
*cached* formula results, because a browser cannot evaluate Excel formulas. Any
library that writes xlsx, openpyxl included, drops those cached results when
it saves, so a workbook edited in Python looks completely empty to the deck:
every derived figure, the fair value and the BUY/HOLD/SELL signal come back
blank. Writing the inputs is easy; making Excel's own answers exist in the file
afterwards is the whole job.

LibreOffice is used as the calculation engine: it opens the workbook, evaluates
every formula and writes the results back into the file.

    python3 update_workbook.py model.xlsx inputs.json -o AAPL_model.xlsx

`inputs.json` maps input-sheet cells to values, e.g.

    {
      "B6":  "AAPL",
      "B7":  "Apple Inc.",
      "B10": 245.50,
      "B11": 14840,
      "B17": [383285, 391035, 416000, 441000, 468000]
    }

A scalar sets one cell. A list of five fills the FY-4..Current FY columns
(B..F) of that row, which is how the historical block is laid out.
"""
from __future__ import annotations

import json
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

import openpyxl

INPUT_SHEET = "01_Input_Template"
OUTPUT_SHEET = "03_Output_Decision"
SOFFICE = "/mnt/skills/public/pptx/scripts/office/soffice.py"

# What the output sheet holds, for the verification step
CHECKS = {5: "Company", 6: "Ticker", 7: "Current price", 8: "Blended fair value",
          9: "Upside / downside", 10: "Buy below", 11: "Overall score", 12: "Decision",
          15: "Data quality"}
FLAG_ROWS = range(48, 59)      # tblDataQualityFlags on the output sheet
SCORE_ROWS = range(20, 26)     # tblScoreDashboard


HISTORY_ROWS = range(19, 31)     # the five-year grid, columns B..F


def apply_inputs(path: Path, values: dict) -> None:
    """Write the input cells. Formulas elsewhere are preserved as formulas."""
    wb = openpyxl.load_workbook(path)
    if INPUT_SHEET not in wb.sheetnames:
        raise SystemExit(f"{path.name} has no '{INPUT_SHEET}' sheet. Is this the right workbook?")
    ws = wb[INPUT_SHEET]

    # Clear the historical grid first. The template ships with demo figures, and
    # a row the input file happens not to mention would otherwise survive into a
    # different company's model: a stale balance sheet that still calculates.
    if values.get("_reset_history", True):
        for row in HISTORY_ROWS:
            for col in range(2, 7):
                # Assign through .value: openpyxl's cell(value=None) is a no-op
                # and would leave the template's demo figures in place.
                ws.cell(row=row, column=col).value = None

    for cell, value in values.items():
        if cell.startswith("_"):
            continue                 # a comment key (source notes, date accessed)
        if isinstance(value, list):
            # a historical row: FY-4 .. Current FY across columns B..F
            row = int("".join(ch for ch in cell if ch.isdigit()))
            if len(value) != 5:
                raise SystemExit(f"{cell}: expected 5 yearly values, got {len(value)}")
            for i, v in enumerate(value):
                ws.cell(row=row, column=2 + i, value=v)
        else:
            ws[cell] = value
    wb.save(path)


def recalculate(path: Path) -> None:
    """Make Excel's own answers exist in the file again.

    openpyxl has just removed every cached formula result. LibreOffice opens the
    workbook, evaluates the formulas and writes the values back, which is what
    lets a browser read the model's output without evaluating anything itself.
    """
    with tempfile.TemporaryDirectory() as tmp:
        r = subprocess.run(
            [sys.executable, SOFFICE, "--headless", "--convert-to", "xlsx", "--outdir", tmp, str(path)],
            capture_output=True, text=True, timeout=300)
        out = Path(tmp) / path.name
        if not out.exists():
            raise SystemExit("LibreOffice did not produce a workbook:\n" + (r.stderr or r.stdout))
        shutil.copy(out, path)


def verify(path: Path) -> tuple[dict, list, list]:
    """Read the output sheet back. Blank values here mean the recalc did not take."""
    ws = openpyxl.load_workbook(path, data_only=True)[OUTPUT_SHEET]
    results = {label: ws.cell(row=r, column=2).value for r, label in CHECKS.items()}
    missing = [k for k, v in results.items() if v is None]
    if missing:
        raise SystemExit("recalculation failed; still empty: " + ", ".join(missing))

    scores = [(ws.cell(row=r, column=1).value, ws.cell(row=r, column=2).value,
               ws.cell(row=r, column=3).value) for r in SCORE_ROWS]
    # Only the checks that actually tripped; an all-clear needs no commentary.
    flags = [(ws.cell(row=r, column=2).value, ws.cell(row=r, column=1).value,
              ws.cell(row=r, column=3).value) for r in FLAG_ROWS
             if ws.cell(row=r, column=2).value not in (None, "OK")]
    return results, scores, flags


def main() -> int:
    if len(sys.argv) < 3:
        print(__doc__)
        return 2
    src, spec = Path(sys.argv[1]), Path(sys.argv[2])
    dest = Path(sys.argv[sys.argv.index("-o") + 1]) if "-o" in sys.argv else \
        src.with_name(src.stem + "_updated.xlsx")

    shutil.copy(src, dest)                       # never edit the original
    apply_inputs(dest, json.loads(spec.read_text()))
    recalculate(dest)
    results, scores, flags = verify(dest)

    print(f"{dest.name}")
    for label, value in results.items():
        if isinstance(value, float):
            value = f"{value:,.2f}"
        print(f"  {label:20} {value}")

    print("\n  score modules")
    for label, value, word in scores:
        print(f"    {label:24} {value:6.1f}  {word}")

    if flags:
        print("\n  data-quality flags")
        for status, label, detail in flags:
            print(f"    [{status}] {label}\n             {detail}")

    print("\nDrop this workbook and the deck template into Deck Updater to build the deck.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
