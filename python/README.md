# SlideSync. Python edition

Cross-platform implementation using `python-pptx`, `openpyxl` and `matplotlib`.
No Office installation is needed.

```bash
cd python
pip install -e .
slidesync run path/to/config.json
```

`config.json`:

```json
{
  "sources": [{ "alias": "sales", "path": "sales.xlsx" }],
  "decks":   [{ "template": "template.pptx", "output": "output/sales_update.pptx" }],
  "log": "output/log.csv"
}
```

Relative paths are resolved against the config file's folder. The first source is
the default for tokens without an `alias:` prefix.

Try it on the bundled sample:

```bash
python ../examples/make_samples.py ../examples/sample
slidesync run ../examples/sample/config.json
```

Run the tests with `pip install -e .[dev] && pytest`.

## Notes

* Values come from the workbook's saved (cached) results, so save the source in
  Excel before running if formulas changed.
* `picture:` renders the range with matplotlib as a clean, neutral table image.
  For a pixel-perfect copy of Excel's formatting use the Excel edition.
* `chart:` supports category charts (bar, column, line, pie, area) and XY
  scatter via `python-pptx`'s `replace_data`.

## Verifying a run

The log records what the engine believes it wrote. Two checks make it a control
rather than a record:

```bash
slidesync run config.json --verify              # read the decks back and check them
slidesync run config.json --baseline approved.csv --report qc.csv
slidesync verify config.json --baseline approved.csv   # check a run that already happened
```

**Read-back** re-opens each produced deck with python-pptx and asks whether the
slide actually says what the log claims: text present, table the recorded size,
chart the recorded series and point counts, picture actually replaced, and every
failed reference still showing its token. None of the code that wrote the deck
takes part, so an engine bug cannot hide itself.

**Baseline** diffs this run's log against a log you have signed off. Values that
moved are reported as `CHANGED`, that is the month doing its job. A token that
no longer resolves, a shape that has been renamed, a table that changed shape:
`MISSING` or `MISMATCH`, and the command exits non-zero.

A regenerated "ideal log" is deliberately not offered. Read from the same
workbook by the same reader, it would agree with the run log by construction and
prove nothing.

| Status | Meaning |
|---------|---------|
| `OK` | The slide matches the log |
| `CHANGED` | Value differs from the baseline, review it |
| `MISMATCH` | The deck does not match what was recorded |
| `MISSING` | Expected content is not there |
| `SKIPPED` | Nothing verifiable (empty value, unparsable record) |
