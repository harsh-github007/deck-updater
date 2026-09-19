// Builds the client-facing investment view deck for the v2 model.
//
// Every figure on these slides is a Deck Updater token or shape directive, so
// dropping a different company's copy of the workbook regenerates the whole
// deck. Nothing here is typed by hand at report time.
//
// Slide 5 is the one that earns the deck its credibility: it shows what the
// model distrusts about its own inputs. A rating that arrives with its caveats
// already on a slide is worth far more than one that looks unanimous.
const P = require("pptxgenjs");

const INK = "0F1F33", MUTED = "5A6B7D", ACCENT = "0B6E4F", WARN = "A8620B",
      RULE = "E2E7EC", PANEL = "F4F7F9", WHITE = "FFFFFF";
const FONT = "Arial";
const M = 0.62;                       // page margin
const W = 13.33, H = 7.5;

const p = new P();
p.layout = "LAYOUT_WIDE";
p.author = "Deck Updater";
p.title = "Investment View";

const foot = (s) =>
  s.addText(
    [{ text: "Educational template, not financial advice. ", options: { color: MUTED } },
     { text: "Rules-based output of the assumptions on the input sheet · ", options: { color: MUTED } },
     { text: "{{03_Output_Decision!B6}}", options: { color: MUTED, bold: true } }],
    { x: M, y: H - 0.55, w: W - 2 * M, h: 0.3, fontSize: 9, fontFace: FONT, isTextBox: true, margin: 0 });

const head = (s, title, sub) => {
  s.addText(title, { x: M, y: 0.45, w: W - 2 * M, h: 0.5, fontSize: 28, bold: true,
    color: INK, fontFace: FONT, isTextBox: true, margin: 0 });
  s.addText(sub, { x: M, y: 1.0, w: W - 2 * M, h: 0.34, fontSize: 13, color: MUTED,
    fontFace: FONT, isTextBox: true, margin: 0 });
  s.addShape(p.ShapeType.rect, { x: M, y: 1.45, w: W - 2 * M, h: 0.02, fill: { color: RULE } });
};

const label = (s, x, y, text, color) =>
  s.addText(text, { x, y, w: 4.4, h: 0.3, fontSize: 11, bold: true, color: color || ACCENT,
    fontFace: FONT, isTextBox: true, margin: 0 });

/* ------------------------------------------------------- 1. investment view */
{
  const s = p.addSlide();
  s.addText("INVESTMENT VIEW", { x: M, y: 0.6, w: 6, h: 0.3, fontSize: 11, bold: true,
    color: ACCENT, charSpacing: 2, fontFace: FONT, isTextBox: true, margin: 0 });

  // Two lines of headroom: a legal name like "Space Exploration Technologies
  // Corp." wraps, and at 40pt in a 0.9" box it lands on top of the subtitle.
  s.addText("{{03_Output_Decision!B5}}", { x: M, y: 0.95, w: 7.9, h: 1.25, fontSize: 34, bold: true,
    color: INK, fontFace: FONT, isTextBox: true, margin: 0, valign: "top" });
  s.addText("{{03_Output_Decision!B6}} · {{01_Input_Template!B9}}",
    { x: M, y: 2.28, w: 7.9, h: 0.4, fontSize: 15, color: MUTED, fontFace: FONT, isTextBox: true, margin: 0 });

  // decision badge: the one line a client reads first
  s.addShape(p.ShapeType.roundRect, { x: 9.0, y: 0.95, w: 3.7, h: 1.6, rectRadius: 0.12,
    fill: { color: INK } });
  s.addText("RECOMMENDATION", { x: 9.0, y: 1.12, w: 3.7, h: 0.25, fontSize: 9, bold: true,
    color: "9FB3C8", align: "center", charSpacing: 1.5, fontFace: FONT, isTextBox: true, margin: 0 });
  s.addText("{{03_Output_Decision!B12}}", { x: 9.0, y: 1.42, w: 3.7, h: 0.85, fontSize: 40, bold: true,
    color: WHITE, align: "center", fontFace: FONT, isTextBox: true, margin: 0 });
  // REVIEW means the model withheld a fair value; the chip says so on the cover
  s.addText("Data quality {{03_Output_Decision!B15|0}}/100 · REVIEW means a blocking flag stands",
    { x: 9.0, y: 2.62, w: 3.7, h: 0.5, fontSize: 8.5, color: MUTED, align: "center",
      fontFace: FONT, isTextBox: true, margin: 0 });

  const tiles = [
    ["Current price",       "{{03_Output_Decision!B7|$#,##0.00}}"],
    ["Blended fair value",  "{{03_Output_Decision!B8|$#,##0.00}}"],
    ["Upside to fair value","{{03_Output_Decision!B9|0.0%}}"],
    ["Buy below",           "{{03_Output_Decision!B10|$#,##0.00}}"],
  ];
  tiles.forEach(([k, v], i) => {
    const x = M + i * 3.05;
    s.addShape(p.ShapeType.rect, { x, y: 3.35, w: 2.85, h: 1.5, fill: { color: PANEL } });
    s.addText(k, { x: x + 0.22, y: 3.55, w: 2.4, h: 0.3, fontSize: 10, color: MUTED,
      fontFace: FONT, isTextBox: true, margin: 0 });
    s.addText(v, { x: x + 0.22, y: 3.9, w: 2.45, h: 0.6, fontSize: 24, bold: true, color: INK,
      fontFace: FONT, isTextBox: true, margin: 0 });
  });

  s.addShape(p.ShapeType.rect, { x: M, y: 5.25, w: 0.06, h: 1.1, fill: { color: ACCENT } });
  s.addText("{{03_Output_Decision!B67}}", { x: M + 0.25, y: 5.25, w: W - 2 * M - 0.25, h: 1.1,
    fontSize: 15, color: INK, fontFace: FONT, isTextBox: true, margin: 0, valign: "top" });
  foot(s);
}

/* ------------------------------------------------------------- 2. valuation */
{
  const s = p.addSlide();
  head(s, "What the shares are worth",
       "Two independent valuations, averaged, against the price in the market today.");

  s.addTable(
    [[{ text: "Metric" }, { text: "Value" }],
     [{ text: "-" }, { text: "-" }]],
    { x: M, y: 1.75, w: 6.1, colW: [3.8, 2.3], rowH: 0.36, fontSize: 11.5, fontFace: FONT,
      color: INK, border: { type: "solid", color: RULE, pt: 1 }, align: "left",
      objectName: "table:03_Output_Decision!A5:B15;header=0" });

  label(s, 7.15, 1.75, "How the fair value is built");
  s.addChart(p.ChartType.bar,
    [{ name: "$ / share", labels: ["A", "B", "C", "D", "E"], values: [1, 1, 1, 1, 1] }],
    { x: 7.0, y: 2.1, w: 5.7, h: 3.5, barDir: "bar", chartColors: [ACCENT],
      showLegend: false, showValue: true, dataLabelPosition: "outEnd",
      dataLabelColor: INK, dataLabelFontSize: 9.5, dataLabelFontFace: FONT,
      catAxisLabelColor: INK, catAxisLabelFontSize: 10, catAxisLabelFontFace: FONT,
      valAxisLabelColor: MUTED, valAxisLabelFontSize: 9,
      catAxisOrderReverse: true,
      valGridLine: { color: RULE, size: 1 }, catGridLine: { style: "none" },
      objectName: "chart:03_Output_Decision!A29:B34" });

  s.addText("Lynch PEG  {{03_Output_Decision!B13|0.00}}x     ·     Overall score  {{03_Output_Decision!B11|0}}/100     ·     Data quality  {{03_Output_Decision!B15|0}}/100",
    { x: M, y: 5.8, w: W - 2 * M, h: 0.4, fontSize: 12, bold: true, color: INK,
      fontFace: FONT, isTextBox: true, margin: 0 });
  foot(s);
}

/* ---------------------------------------------------------- 3. quality score */
{
  const s = p.addSlide();
  head(s, "Why we reached that view",
       "Six scored modules, weighted into one overall score. Higher is better; 100 is the cap.");

  s.addChart(p.ChartType.bar,
    [{ name: "Score", labels: ["A", "B", "C", "D", "E", "F"], values: [1, 1, 1, 1, 1, 1] }],
    { x: M, y: 1.75, w: 7.3, h: 4.0, barDir: "bar", chartColors: [ACCENT],
      showLegend: false, showValue: true, dataLabelPosition: "outEnd",
      dataLabelColor: INK, dataLabelFontSize: 10, dataLabelFontFace: FONT,
      catAxisLabelColor: INK, catAxisLabelFontSize: 11, catAxisLabelFontFace: FONT,
      valAxisLabelColor: MUTED, valAxisLabelFontSize: 10, valAxisMaxVal: 100, valAxisMinVal: 0,
      catAxisOrderReverse: true,
      valGridLine: { color: RULE, size: 1 }, catGridLine: { style: "none" },
      objectName: "chart:03_Output_Decision!A19:B25" });

  s.addShape(p.ShapeType.roundRect, { x: 8.3, y: 1.75, w: 4.4, h: 1.5, rectRadius: 0.1,
    fill: { color: PANEL } });
  s.addText("Overall score", { x: 8.55, y: 1.95, w: 3.9, h: 0.3, fontSize: 11, color: MUTED,
    fontFace: FONT, isTextBox: true, margin: 0 });
  s.addText("{{03_Output_Decision!B11|0}} / 100", { x: 8.55, y: 2.3, w: 3.9, h: 0.75, fontSize: 30,
    bold: true, color: INK, fontFace: FONT, isTextBox: true, margin: 0 });

  label(s, 8.55, 3.5, "Weighting");
  s.addText(
    [{ text: "Quality {{01_Input_Template!B49|0%}}", options: { bullet: true, breakLine: true } },
     { text: "Valuation {{01_Input_Template!B50|0%}}", options: { bullet: true, breakLine: true } },
     { text: "Financial strength {{01_Input_Template!B51|0%}}", options: { bullet: true } }],
    { x: 8.55, y: 3.85, w: 3.9, h: 1.3, fontSize: 11, color: INK, fontFace: FONT,
      isTextBox: true, margin: 0, paraSpaceAfter: 6 });

  s.addText("A module with no usable input scores zero and raises a flag. It is never averaged away.",
    { x: 8.55, y: 5.15, w: 4.0, h: 0.6, fontSize: 10, color: MUTED, fontFace: FONT,
      isTextBox: true, margin: 0 });
  foot(s);
}

/* --------------------------------------------------------- 4. five-year view */
{
  const s = p.addSlide();
  head(s, "The five years behind the number",
       "Base-case forecast that feeds the discounted cash flow. Revenue and free cash flow in $mm.");

  s.addTable(
    [["Metric", "Y0", "Y1", "Y2", "Y3", "Y4", "Y5"],
     ["-", "-", "-", "-", "-", "-", "-"]],
    { x: M, y: 1.72, w: W - 2 * M, colW: [2.75, 1.55, 1.55, 1.55, 1.55, 1.55, 1.59],
      rowH: 0.3, fontSize: 10.5, fontFace: FONT, color: INK,
      border: { type: "solid", color: RULE, pt: 1 }, align: "left",
      objectName: "table:02_Derived_Values!A53:G57" });

  s.addText(
    "Year-5 revenue {{03_Output_Decision!B63|$#,##0}}mm   ·   Year-5 EPS {{03_Output_Decision!B64|$#,##0.00}}   ·   Year-5 FCF {{03_Output_Decision!B65|$#,##0}}mm",
    { x: M, y: 3.32, w: W - 2 * M, h: 0.35, fontSize: 11, color: INK, fontFace: FONT,
      isTextBox: true, margin: 0 });

  s.addChart(p.ChartType.line,
    [{ name: "Revenue", labels: ["Y0", "Y1", "Y2", "Y3", "Y4", "Y5"], values: [1, 1, 1, 1, 1, 1] }],
    { x: M, y: 3.75, w: W - 2 * M, h: 2.55, chartColors: [ACCENT], lineSize: 3, lineSmooth: false,
      showLegend: false, showValue: false,
      catAxisLabelColor: INK, catAxisLabelFontSize: 10, catAxisLabelFontFace: FONT,
      valAxisLabelColor: MUTED, valAxisLabelFontSize: 10,
      valGridLine: { color: RULE, size: 1 }, catGridLine: { style: "none" },
      objectName: "chart:02_Derived_Values!A53:G54;orient=rows" });
  foot(s);
}

/* -------------------------------------------------- 5. what the model checked */
{
  const s = p.addSlide();
  head(s, "What we checked before rating it",
       "The model tests its own inputs. A blocking flag withholds the fair value entirely and returns REVIEW.");

  s.addTable(
    [["Check", "Status", "Detail"],
     ["-", "-", "-"]],
    { x: M, y: 1.72, w: W - 2 * M, colW: [3.9, 1.35, 6.84],
      rowH: 0.3, fontSize: 10, fontFace: FONT, color: INK,
      border: { type: "solid", color: RULE, pt: 1 }, align: "left", valign: "middle",
      objectName: "table:03_Output_Decision!A47:C58" });

  s.addShape(p.ShapeType.rect, { x: M, y: 5.75, w: 0.06, h: 0.85, fill: { color: WARN } });
  s.addText(
    "Data quality {{03_Output_Decision!B15|0}}/100. Warnings reduce this score; a blocking flag forces REVIEW and withholds the fair value. A BUY additionally requires a data-quality score of 60 or better.",
    { x: M + 0.25, y: 5.75, w: W - 2 * M - 0.25, h: 0.85, fontSize: 11, color: INK,
      fontFace: FONT, isTextBox: true, margin: 0, valign: "top" });
  foot(s);
}

/* ------------------------------------------------------ 6. basis and limits */
{
  const s = p.addSlide();
  head(s, "What this rests on",
       "Every assumption against what the company's own history actually did.");

  label(s, M, 1.72, "Assumptions against history");
  s.addTable(
    [["Assumption", "Used", "History", "Status"],
     ["-", "-", "-", "-"]],
    { x: M, y: 2.06, w: 6.3, colW: [2.5, 1.25, 1.3, 1.25],
      rowH: 0.31, fontSize: 10.5, fontFace: FONT, color: INK,
      border: { type: "solid", color: RULE, pt: 1 }, align: "left",
      objectName: "table:03_Output_Decision!A38:D43" });

  // A native table rather than a rendered image of the range: the workbook's own
  // styling does not match the deck, and a 16-row range rendered to a fixed box
  // gets clipped on the left. Editable here, and consistent with every other slide.
  label(s, 7.2, 1.72, "Metrics behind the score");
  s.addTable(
    [["Metric", "Value", "Units"],
     ["-", "-", "-"]],
    { x: 7.2, y: 2.06, w: 5.51, colW: [3.05, 1.36, 1.1],
      rowH: 0.265, fontSize: 9.5, fontFace: FONT, color: INK,
      border: { type: "solid", color: RULE, pt: 1 }, align: "left",
      objectName: "table:02_Derived_Values!A4:C19" });

  label(s, M, 4.15, "Fixed assumptions");
  s.addText(
    [{ text: "Discount rate {{01_Input_Template!B34|0.0%}}", options: { bullet: true, breakLine: true } },
     { text: "Terminal growth {{01_Input_Template!B35|0.0%}}", options: { bullet: true, breakLine: true } },
     { text: "Margin of safety required {{01_Input_Template!B43|0%}}", options: { bullet: true, breakLine: true } },
     { text: "Cash base method {{01_Input_Template!B42}}", options: { bullet: true, breakLine: true } },
     { text: "Exit P/E method {{01_Input_Template!B37}}", options: { bullet: true } }],
    { x: M, y: 4.5, w: 6.3, h: 1.6, fontSize: 11, color: INK, fontFace: FONT,
      isTextBox: true, margin: 0, paraSpaceAfter: 5 });

  s.addText("A discounted cash flow is only as good as its growth and discount-rate assumptions; small changes to either move fair value materially. Analyst note: {{01_Input_Template!B13}}",
    { x: M, y: 6.1, w: 6.3, h: 0.75, fontSize: 9.5, color: MUTED, fontFace: FONT,
      isTextBox: true, margin: 0, valign: "top" });
  foot(s);
}

p.writeFile({ fileName: "Investment_View_Template.pptx" })
 .then((f) => console.log("wrote", f));
