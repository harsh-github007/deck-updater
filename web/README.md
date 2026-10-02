# Deck Updater: browser edition

A single HTML file. Open it and it runs: no install, no server, no account.
Files never leave the machine, the workbook and the deck are read in the
page, and the updated deck is written there too.

```bash
open index.html          # or double-click it
```

To host it, this folder is the whole site. There is no build step:

```bash
vercel --prod            # uses ../vercel.json
```

`vendor/` holds JSZip, SheetJS and jsPDF, so the page fetches no third-party
JavaScript and works offline. It falls back to cdnjs only if that folder is
missing, which is what keeps the single `index.html` usable on its own.

## What it does

Implements the same [specification](../docs/SPEC.md) as the Python and Excel
editions: `{{Sheet!A1}}` text tokens, `table:` / `chart:` / `picture:` shape
directives, multiple sources, multiple decks, a per-action log, plus two
things the other editions don't have:

* **A preview.** Each slide is rebuilt from the template's real shape geometry
  and plays back as it fills in, so you see every value land before you commit
  to the download. Unresolved references stay visible in red.
* **PDF export.** One page per slide, drawn from that preview. Values, layout
  and existing images are right; the template's fonts, theme colours and
  effects are not reproduced, and a chart that carries no directive is drawn
  as a labelled outline rather than silently omitted. Every updated chart is
  drawn as a signed bar chart whatever type the deck uses. It is a read-only
  circulation copy, not a substitute for PowerPoint's own export.

## What it doesn't do

* `picture:` ranges are drawn with canvas, a clean neutral table, not a copy
  of Excel's own formatting. Use the VBA edition for a pixel-perfect paste.
* Charts are updated through their cached values, which is what the slide
  displays. The worksheet embedded behind the chart is left alone, so
  PowerPoint's "Edit Data" still shows the original numbers.
* Formulas are not recalculated; saved (cached) results are read. Save the
  workbook in Excel after formulas change.

## Multiple decks from one workbook

Drop several templates at once. To let the workbook name the outputs, add a
sheet called `Decks` (or `Setup`, `SlideSync`, `Config`):

| Template             | Output                        | Skip |
|----------------------|-------------------------------|------|
| regional-review.pptx | Region review - Aug 2026.pptx |      |
| board-pack.pptx      | Board pack Aug 2026.pptx      | Y    |

Headings are matched loosely, `Template` / `Deck` / `File Name` for the first
column, `Output` / `Save as` for the second, so an existing setup sheet
usually works unchanged.

## Tokens on the master

A token on a slide layout or the slide master is resolved once and appears on
every slide that inherits it. Useful for a reporting date or a classification
line in the footer. It appears once in the log, as a `layout` or `master` row.

## Checking a run

Both checks from the Python edition are here, and neither uses the code that
wrote the deck.

**Read-back** runs automatically. The finished `.pptx` is re-opened from its own
bytes and every logged value is checked against what the file actually says, text present, table dimensions, chart series and point counts, picture
replaced, failed references still showing their token.

**Baseline** is the third drop zone. Drop a log CSV from a run you signed off
and this run is diffed against it: a value that moved is `CHANGED`, a token that
stopped resolving or a shape that was renamed is `MISSING` or `MISMATCH`, shown
both in the summary and as a tag on the affected row.

Findings download as `deck-updater-qc.csv`.

## Frontend materials and motion

The landing page uses an Apple-inspired neutral palette, frosted navigation and
controls, and opaque workbook/slide content. Reduced transparency, increased
contrast, reduced motion, and browsers without backdrop filters get simpler
materials. Navigation selection changes immediately; the illustration updates
only when requested. The working sample advances when processing actually
finishes, without an artificial animation delay.

Motion 14.0.0 is vendored from the npm package's `dist/motion.js` as
`vendor/motion-14.0.0.min.js`, with its MIT license alongside it. It supplies the
brief, non-bouncing workspace transition. The app still runs as static files
with no build step or added external script requests.

Design references: [Emil Kowalski's Apple Design skill](https://github.com/emilkowalski/skills/blob/main/skills/apple-design/SKILL.md),
[Apple's materials guidance](https://developer.apple.com/design/human-interface-guidelines/materials),
and [ThreeUI's navigation and layered-paper studies](https://threeui.com).
ThreeUI informed the composition; its React package is not installed in this
plain-JavaScript application.
