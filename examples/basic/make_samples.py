"""Build a sample Excel source and a marked-up PowerPoint template.

Run:  python examples/basic/make_samples.py [output_folder]

Produces ``sales.xlsx``, ``template.pptx`` and ``config.json`` that exercise
every feature in docs/SPEC.md. Requires openpyxl and python-pptx.
"""

from __future__ import annotations

import datetime as dt
import json
import sys
from pathlib import Path

from openpyxl import Workbook
from openpyxl.workbook.defined_name import DefinedName
from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE
from pptx.util import Inches, Pt


def make_workbook(path: Path) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Summary"
    ws["A1"], ws["B1"] = "Report month", dt.date(2026, 8, 31)
    ws["B1"].number_format = "mmm yyyy"
    ws["A2"], ws["B2"] = "Total revenue", 1_284_530.25
    ws["B2"].number_format = "$#,##0"
    ws["A3"], ws["B3"] = "Growth vs LY", 0.0834
    ws["B3"].number_format = "0.0%"
    ws["A4"], ws["B4"] = "Top region", "North"
    wb.defined_names["TotalRevenue"] = DefinedName("TotalRevenue", attr_text="Summary!$B$2")

    data = wb.create_sheet("Regions")
    data.append(["Region", "Q1", "Q2", "Q3"])
    for row in [("North", 320, 355, 402), ("South", 210, 198, 240), ("East", 180, 205, 231), ("West", 150, 172, 190)]:
        data.append(list(row))
    for r in data.iter_rows(min_row=2, min_col=2):
        for c in r:
            c.number_format = "#,##0"
    wb.save(path)


def make_template(path: Path) -> None:
    """A marked-up deck laid out the way a real monthly review would be.

    Geometry matters here beyond looks: the table's row height is what a filled
    table inherits when rows are cloned, so a template row of 0.5in becomes a
    2.5in table at five rows. Rows are sized so the finished table lands inside
    the slide rather than trailing off it.
    """
    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
    blank = prs.slide_layouts[6]
    MARGIN, TOP, BOTTOM = Inches(0.7), Inches(0.55), Inches(6.95)

    def title_on(slide, text, size=28, name="Heading"):
        tb = slide.shapes.add_textbox(MARGIN, TOP, Inches(12), Inches(0.9))
        tb.name = name
        para = tb.text_frame.paragraphs[0]
        para.text = text
        para.runs[0].font.size = Pt(size)
        return tb

    # Slide 1, text tokens
    s1 = prs.slides.add_slide(blank)
    title_on(s1, "Sales update: {{Summary!B1}}", size=40, name="Title")
    body = s1.shapes.add_textbox(MARGIN, Inches(2.35), Inches(12), Inches(4.2))
    body.name = "Body"
    tf = body.text_frame
    tf.word_wrap = True
    tf.text = "Revenue reached {{TotalRevenue}} ({{Summary!B3|0.0%}} vs last year)."
    p2 = tf.add_paragraph()
    # deliberately split a token across two runs
    r1 = p2.add_run(); r1.text = "Best performing region: {{Summ"
    r2 = p2.add_run(); r2.text = "ary!B4}}."; r2.font.bold = True
    p3 = tf.add_paragraph()
    p3.text = "Raw month value: {{Summary!B1|date:dd MMM yyyy}} and a bad token {{Nowhere!Z9}}"
    for para in tf.paragraphs:
        for run in para.runs:
            run.font.size = Pt(28)
        para.space_after = Pt(26)

    # Slide 2, table + chart, side by side and filling the slide
    s2 = prs.slides.add_slide(blank)
    title_on(s2, "Regional performance, {{Summary!B1}}")
    COL_W, CONTENT_TOP, CONTENT_H = Inches(5.6), Inches(1.75), Inches(4.4)
    # 2 template rows at 0.88in each -> a 5-row table is 4.4in tall
    tbl_frame = s2.shapes.add_table(2, 2, MARGIN, CONTENT_TOP, COL_W, Inches(1.76))
    tbl = tbl_frame.table
    tbl.cell(0, 0).text, tbl.cell(0, 1).text = "Header", "Header"
    tbl.cell(1, 0).text, tbl.cell(1, 1).text = "x", "x"
    tbl_frame.name = "table:Regions!A1:D5"

    chart_data = CategoryChartData()
    chart_data.categories = ["a", "b"]
    chart_data.add_series("s", (1, 2))
    gf = s2.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, Inches(6.95), CONTENT_TOP,
                             Inches(5.7), CONTENT_H, chart_data)
    gf.name = "chart:Regions!A1:D5"
    gf.chart.has_legend = True

    # Slide 3, a range rendered as a picture
    s3 = prs.slides.add_slide(blank)
    title_on(s3, "Regional detail, {{Summary!B4}} leads")
    import io
    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt
    fig, ax = plt.subplots(figsize=(3, 1)); ax.text(0.5, 0.5, "placeholder", ha="center"); ax.axis("off")
    buf = io.BytesIO(); fig.savefig(buf, format="png"); plt.close(fig)
    pic = s3.shapes.add_picture(io.BytesIO(buf.getvalue()), MARGIN, CONTENT_TOP,
                                width=Inches(7.4), height=CONTENT_H)
    pic.name = "picture:Regions!A1:D5"
    cap = s3.shapes.add_textbox(MARGIN, Inches(6.35), Inches(9), Inches(0.6))
    cap.name = "Caption"
    cap.text_frame.text = "Table rendered as image, top region {{Summary!B4}}"

    prs.save(path)


def main() -> None:
    out = Path(sys.argv[1]) if len(sys.argv) > 1 else Path(__file__).parent / "sample"
    out.mkdir(parents=True, exist_ok=True)
    make_workbook(out / "sales.xlsx")
    make_template(out / "template.pptx")
    cfg = {
        "sources": [{"alias": "sales", "path": "sales.xlsx"}],
        "decks": [{"template": "template.pptx", "output": "output/sales_update.pptx"}],
        "log": "output/log.csv",
    }
    (out / "config.json").write_text(json.dumps(cfg, indent=2), encoding="utf-8")
    print(f"Sample files written to {out}")


if __name__ == "__main__":
    main()
