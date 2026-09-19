// The generic sample deck bundled with Deck Updater's "Try the sample" button.
// Deliberately plain business content, the point is to demonstrate the tool, so
// nothing here should read as belonging to a particular industry.
//
// Laid out so content runs from ~20% to ~88% of the slide height: a template
// that leaves the bottom third empty makes the preview look broken.
const P = require("pptxgenjs");

const INK = "1B2733", MUTED = "5C6B7A", ACCENT = "1F6F5C",
      RULE = "E4E8EC", PANEL = "F5F7F9";
const FONT = "Arial";
const M = 0.62, W = 13.33, H = 7.5;
const BLANK_PNG = "image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==";

const p = new P();
p.layout = "LAYOUT_WIDE";
p.title = "Monthly business review";

const foot = (s, n) => {
  s.addShape(p.ShapeType.rect, { x: M, y: 6.42, w: W - 2 * M, h: 0.015, fill: { color: RULE } });
  s.addText("Monthly business review · {{Summary!B1|date:MMMM yyyy}}",
    { x: M, y: 6.55, w: 8, h: 0.3, fontSize: 9.5, color: MUTED, fontFace: FONT, isTextBox: true, margin: 0 });
  s.addText(String(n), { x: W - M - 1, y: 6.55, w: 1, h: 0.3, fontSize: 9.5, color: MUTED,
    align: "right", fontFace: FONT, isTextBox: true, margin: 0 });
};
const head = (s, t, sub) => {
  s.addText(t, { x: M, y: 0.52, w: W - 2 * M, h: 0.56, fontSize: 30, bold: true, color: INK,
    fontFace: FONT, isTextBox: true, margin: 0 });
  s.addText(sub, { x: M, y: 1.12, w: W - 2 * M, h: 0.34, fontSize: 13, color: MUTED,
    fontFace: FONT, isTextBox: true, margin: 0 });
  s.addShape(p.ShapeType.rect, { x: M, y: 1.58, w: W - 2 * M, h: 0.02, fill: { color: RULE } });
};

/* ------------------------------------------------------------------ slide 1 */
{
  const s = p.addSlide();
  head(s, "Sales update: {{Summary!B1|date:MMM yyyy}}",
       "Revenue, growth against last year, and where the quarter was won.");

  const tiles = [
    ["Total revenue",  "{{TotalRevenue|$#,##0}}"],
    ["Growth vs last year", "{{Summary!B3|0.0%}}"],
    ["Top region",     "{{Summary!B4}}"],
  ];
  tiles.forEach(([k, v], i) => {
    const x = M + i * 4.06;
    s.addShape(p.ShapeType.rect, { x, y: 1.95, w: 3.85, h: 1.55, fill: { color: PANEL } });
    s.addText(k, { x: x + 0.25, y: 2.16, w: 3.3, h: 0.3, fontSize: 10.5, color: MUTED,
      fontFace: FONT, isTextBox: true, margin: 0 });
    s.addText(v, { x: x + 0.25, y: 2.52, w: 3.4, h: 0.7, fontSize: 28, bold: true, color: INK,
      fontFace: FONT, isTextBox: true, margin: 0 });
  });

  s.addShape(p.ShapeType.rect, { x: M, y: 3.9, w: 0.06, h: 2.2, fill: { color: ACCENT } });
  s.addText(
    [{ text: "Revenue reached {{TotalRevenue|$#,##0}}, up {{Summary!B3|0.0%}} on the same month last year.", options: { breakLine: true } },
     { text: "{{Summary!B4}} remained the strongest region across all three quarters.", options: { breakLine: true } },
     { text: "Reporting period closed {{Summary!B1|date:d MMMM yyyy}}.", options: { breakLine: true } },
     { text: "One deliberately broken reference, so the error handling is visible: {{Nowhere!Z9}}", options: {} }],
    { x: M + 0.3, y: 3.9, w: W - 2 * M - 0.3, h: 2.2, fontSize: 15, color: INK, lineSpacingMultiple: 1.5,
      fontFace: FONT, isTextBox: true, margin: 0, valign: "top" });
  foot(s, 1);
}

/* ------------------------------------------------------------------ slide 2 */
{
  const s = p.addSlide();
  head(s, "Regional performance", "The same range, filled as a table and charted beside it.");

  s.addTable(
    [[{ text: "Region" }, { text: "Q1" }, { text: "Q2" }, { text: "Q3" }],
     [{ text: "-" }, { text: "-" }, { text: "-" }, { text: "-" }]],
    { x: M, y: 1.95, w: 5.6, colW: [2.0, 1.2, 1.2, 1.2], rowH: 0.42, fontSize: 12.5,
      fontFace: FONT, color: INK, border: { type: "solid", color: RULE, pt: 1 },
      objectName: "table:Regions!A1:D5" });

  s.addText("Every figure in this table comes from the workbook. Add a region and the table grows to fit it.",
    { x: M, y: 4.5, w: 5.6, h: 0.8, fontSize: 11.5, color: MUTED, fontFace: FONT,
      isTextBox: true, margin: 0 });

  s.addChart(p.ChartType.bar,
    [{ name: "Q1", labels: ["A", "B", "C", "D"], values: [1, 1, 1, 1] },
     { name: "Q2", labels: ["A", "B", "C", "D"], values: [1, 1, 1, 1] },
     { name: "Q3", labels: ["A", "B", "C", "D"], values: [1, 1, 1, 1] }],
    { x: 6.55, y: 1.9, w: 6.16, h: 4.3, barDir: "col", chartColors: [ACCENT, "3E8E7E", "7FB8AC"],
      showLegend: true, legendPos: "b", legendColor: MUTED, legendFontSize: 10,
      catAxisLabelColor: INK, catAxisLabelFontSize: 11, catAxisLabelFontFace: FONT,
      valAxisLabelColor: MUTED, valAxisLabelFontSize: 10,
      valGridLine: { color: RULE, size: 1 }, catGridLine: { style: "none" },
      objectName: "chart:Regions!A1:D5" });
  foot(s, 2);
}

/* ------------------------------------------------------------------ slide 3 */
{
  const s = p.addSlide();
  head(s, "The same range, rendered as an image",
       "Useful when the workbook's own number formatting has to survive intact.");

  s.addImage({ x: M, y: 1.95, w: 6.6, h: 3.1, data: BLANK_PNG,
    objectName: "picture:Regions!A1:D5" });

  s.addShape(p.ShapeType.rect, { x: 7.55, y: 1.95, w: 5.16, h: 3.1, fill: { color: PANEL } });
  s.addText("Three ways to place the same data", { x: 7.85, y: 2.2, w: 4.6, h: 0.3, fontSize: 11.5,
    bold: true, color: ACCENT, fontFace: FONT, isTextBox: true, margin: 0 });
  s.addText(
    [{ text: "A native table you can restyle in PowerPoint afterwards.", options: { bullet: true, breakLine: true } },
     { text: "A native chart, so the data stays editable.", options: { bullet: true, breakLine: true } },
     { text: "An image, when the formatting matters more than editability.", options: { bullet: true } }],
    { x: 7.85, y: 2.6, w: 4.6, h: 2.2, fontSize: 11.5, color: INK, fontFace: FONT,
      isTextBox: true, margin: 0, paraSpaceAfter: 10 });

  s.addText("Top region this month: {{Summary!B4}} · reporting period {{Summary!B1|date:MMM yyyy}}",
    { x: M, y: 5.35, w: W - 2 * M, h: 0.5, fontSize: 14, color: INK, fontFace: FONT,
      isTextBox: true, margin: 0 });
  foot(s, 3);
}

p.writeFile({ fileName: "sample_deck_new.pptx" }).then(f => console.log("wrote", f));
