# SlideSync specification

This document defines the behaviour that both implementations (`python/` and `vba/`)
must follow. A PowerPoint template is marked up with **tokens** (in text) and
**shape directives** (in shape names). SlideSync reads values from one or more
Excel workbooks and writes an updated copy of the deck.

## 1. Sources

A *source* is an Excel workbook (`.xlsx`, `.xlsm`, `.xlsb`, `.xls`) identified by
an **alias**. The first source listed is the *default* source. Values are read
from the workbook's saved (cached) cell values; formulas are not recalculated.

## 2. Range references

A *reference* points at a cell or range in a source:

| Form                              | Meaning                                              |
|-----------------------------------|------------------------------------------------------|
| `Sheet1!B3`                       | single cell on `Sheet1` of the default source         |
| `'Q3 Data'!A1:D12`                | range; quote the sheet name when it contains spaces   |
| `TotalRevenue`                    | workbook-level defined name in the default source     |
| `budget:Sheet1!B3`                | same, but in the source whose alias is `budget`       |

## 3. Text tokens

A token is `{{reference}}` or `{{reference|format}}` placed anywhere in a text
frame or a PowerPoint table cell. The whole token is replaced by the cell value.

* A token that spans several formatting runs is still recognised; the replaced
  text takes the formatting of the run where the token begins.
* `format` is optional:
  * `#,##0`, `0.0%`, `0.00`, `$#,##0`: Excel-style number formats (a practical
    subset: thousands separator, fixed decimals, percent, currency prefix).
  * `date:<pattern>` for date formatting, e.g. `date:MMM yyyy`.
  * When omitted, numbers are rendered using the cell's own number format when
    the implementation can read it, otherwise as a plain number.
* An empty cell produces an empty string.
* A reference that cannot be resolved leaves the token text unchanged and is
  written to the log with status `ERROR`.

## 4. Shape directives

A directive is stored in the **shape name** (Home ▸ Select ▸ Selection Pane in
PowerPoint). The format is `kind:reference` with optional `;key=value` options.

| Directive                                     | Effect                                                                  |
|-----------------------------------------------|--------------------------------------------------------------------------|
| `table:Sheet1!A1:D10`                         | Fill the PowerPoint table with the range. Rows/columns are added or removed to match. First row is treated as a header. |
| `table:Sheet1!A1:D10;header=0`                | Same, but the range has no header row.                                   |
| `picture:Sheet1!A1:D10`                       | Render the range as an image and replace the picture, keeping position and size. |
| `chart:Sheet1!A1:D10`                         | Replace the chart's data. Column A = categories, header row = series names, remaining cells = values. |
| `chart:Sheet1!A1:D10;orient=rows`             | Same, but categories run along the header row and each following row is a series. |

Shapes whose names do not match a directive are left untouched.

## 5. Processing order and output

For each deck listed:

1. Open the template (never modified in place).
2. For every slide and every shape (including shapes inside groups):
   * apply shape directives,
   * replace text tokens in text frames and table cells.
3. Save to the configured output path.

## 6. Log

Every action is appended to a log with the columns:

`Deck | Slide | Shape | Directive/Token | Resolved reference | Action | Status | Message`

`Action` is one of `text`, `table`, `picture`, `chart`. `Status` is `OK`,
`SKIPPED` or `ERROR`. Errors never abort the run; the deck is still saved.

## Verification

Implementations may offer two post-run checks. Neither may use the code path
that produced the deck.

1. **Read-back.** Re-open the produced file and confirm each logged action
   against the saved content: text present in the shape, table dimensions,
   chart series and point counts, picture replaced, failed references still
   showing their token.
2. **Baseline.** Compare the run log with an approved log, keyed on
   (deck, slide, shape, directive, action). A differing message is `CHANGED`;
   a differing status, an unmatched key on either side, is `MISMATCH` or
   `MISSING`.

Findings are written as CSV with columns Check, Deck, Slide, Shape, Directive,
Status, Detail. A run with any `MISMATCH` or `MISSING` exits non-zero.
