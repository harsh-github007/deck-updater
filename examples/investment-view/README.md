# Investment view: a worked example

A client-facing deck regenerated per company from one valuation model.

| File | Role |
|------|------|
| `Lynch_Buffett_Stock_Evaluation.xlsx` | The model. Fill the amber input cells; the derived sheet computes CAGRs, PEG, owner earnings, a DCF, six score modules and eleven data-quality checks; the output sheet resolves BUY / HOLD / SELL / REVIEW. |
| `Investment_View_Template.pptx` | The deck, marked up once. Drop it in with a company's workbook to regenerate every figure. |
| `build_model.py` | Rebuilds the workbook from scratch. Every formula lives here, so the model is reviewable as source rather than as a binary. |
| `build_deck.js` | Rebuilds the deck template with pptxgenjs. |
| `update_workbook.py` | Writes one company's inputs into a copy of the model and recalculates it. |

## Two scoring bugs worth knowing about

The first version had two scoring modules that could not return anything but
100. Both are worth understanding before trusting any scored model.

**Buffett valuation** was `50 + (earnings-multiple fair value - EPS growth) * 200`,
clamped at 100. It was meant to read *upside against the margin of safety*, and
instead read a dollar fair value, so any figure above about $0.60 pinned the
module at 100. A company 50% overvalued scored exactly as well as a bargain.

**Growth quality** was `AVERAGE(three CAGRs) * 1000`, clamped at 100, so 10%
average growth already reached full marks and nothing above it could be
distinguished. Worse, a CAGR that failed to compute returned `""`, because a
negative base year makes the fractional power undefined, and `AVERAGE` then
skipped it silently. A missing number scored the same as a perfect one.

Both scored 100 for every company tested, so a third of the composite was
constant while the output still looked like a considered rating.

Now each component score sits in its own row on `02_Derived_Values`
(`tblScoreDetail`), a metric that cannot be computed scores zero and raises a
flag, and the composite moves. Across the three worked examples it ranges from
29 to 75.

## Data-quality checks

`tblDataQuality` runs eleven checks and separates **warnings**, which dock the
data-quality score, from **blocking** flags, which force the decision to
`REVIEW` and withhold the blended fair value, the upside and the buy-below
price entirely. A fair value the model already knows it cannot support is the
one number that must never reach a slide.

The checks cover each CAGR's computability, terminal growth against the
discount rate, score weights summing to 100%, price and share count present, a
positive normalised cash base, forecast growth against the matching historical
CAGR in both directions, a quoted P/E against price divided by current-FY EPS,
and reported cash-flow stability.

## Two methods worth knowing about

**FCF Base Method** decides what the DCF compounds. `Reported` uses the current
year alone, which one heavy-capex year can distort beyond recognition: running
Amazon's FY2025 through it gives a $0.71/share cash base and a $12.78 DCF fair
value. `OwnerEarnings`, the default, uses operating cash flow less maintenance
capex. `Average3` uses a three-year average.

**Exit P/E Method** decides the earnings cross-check. A fixed 18x multiple
prices every compounder as though it were a mature business. `GrowthImplied`,
the default, sets the exit multiple to `PEG anchor * forecast EPS growth %`,
bounded by a floor and a cap. The assumption-versus-history check then flags a
forecast growth rate left sitting at its default.

## What is marked up

| Slide | Directive or token | Pulls |
|-------|--------------------|-------|
| 1 | `{{03_Output_Decision!B12}}` | the recommendation itself |
| 1 | `{{03_Output_Decision!B7\|$#,##0.00}}` and siblings | price, fair value, upside, buy-below |
| 2 | `table:03_Output_Decision!A5:B15;header=0` | the whole decision summary |
| 2 | `chart:03_Output_Decision!A29:B34` | the valuation bridge |
| 3 | `chart:03_Output_Decision!A19:B25` | six score modules |
| 4 | `table:02_Derived_Values!A53:G57` | the forecast rows |
| 4 | `chart:02_Derived_Values!A53:G54;orient=rows` | five-year revenue |
| 5 | `table:03_Output_Decision!A47:C58` | every data-quality check and its detail |
| 6 | `table:03_Output_Decision!A38:D43` | each assumption against history |
| 6 | `table:02_Derived_Values!A4:C19` | the metrics behind the score |

Forty-two values in total, none of them typed by hand at report time.

## Updating the model for a different company

```bash
python3 update_workbook.py Lynch_Buffett_Stock_Evaluation.xlsx amzn_inputs.json -o AMZN_model.xlsx
```

Deck Updater reads a workbook's **cached** formula results, because a browser
cannot evaluate Excel formulas. Every library that writes xlsx, openpyxl
included, drops those cached results on save, so a workbook edited in Python
looks empty to the deck: fair value, upside and the signal all come back blank.
The script runs LibreOffice as a calculation engine afterwards and verifies the
output sheet is populated before handing the file back. It then prints the six
module scores and any flag that tripped.

`inputs.json` maps input-sheet cells to values. A list of five fills the
FY-4..Current FY columns of that row, and keys beginning with `_` are treated
as comments. The historical grid is cleared before the inputs are applied, so a
row the file happens not to mention cannot survive from the template into a
different company's model.

Three worked input files ship here: `demo_inputs.json` (a healthy compounder,
HOLD), `amzn_inputs.json` (Amazon FY2021-FY2025 as reported, SELL with four
warnings), and `spcx_inputs.json` (one reported year and losses, REVIEW with
the fair value withheld).

Nothing here is investment advice. The decision is a rules-based output of the
assumptions on the input sheet, not a view on any company. The shipped workbook
carries demo values, and the flags exist to be read before anything leaves your
desk.
