# SlideSync. Excel/VBA edition

Runs entirely inside Excel and drives PowerPoint through late-bound COM, so it
works on any Windows Office 2016+ install without extra references.

## Files

| Path                | Purpose                                                        |
|---------------------|----------------------------------------------------------------|
| `DeckUpdater.xlsx`    | UI workbook: Home, Instructions, Setup and Log sheets          |
| `src/modMain.bas`   | `RunSlideSync` entry point, deck loop, shape walker            |
| `src/modSetup.bas`  | Constants, Setup-sheet named ranges, path helpers              |
| `src/modSources.bas`| Opens source workbooks, resolves `alias:Sheet!A1` references    |
| `src/modTokens.bas` | `{{…}}` token replacement and value formatting                 |
| `src/modDirectives.bas` | `table:` / `picture:` / `chart:` shape directives          |
| `src/modLog.bas`    | Writes the Log sheet                                           |
| `src/modUI.bas`     | `InstallButtons`: draws the navigation buttons once           |
| `build_workbook.py` | Regenerates `DeckUpdater.xlsx` (only needed if you change the UI) |

## Install (one time)

1. Open `DeckUpdater.xlsx` in Excel and press **Alt+F11** to open the VBA editor.
2. **File ▸ Import File…** and select every `.bas` file in `src/` (you can
   multi-select).
3. Back in Excel, press **Alt+F8**, run `InstallButtons`.
4. **File ▸ Save As ▸ Excel Macro-Enabled Workbook (`.xlsm`)**.

## Use

1. On **Setup**, list your Excel source workbooks (alias, folder, file) and your
   PowerPoint templates (folder, file, save-as).
2. Mark up the templates as described on the **Instructions** sheet or in
   [`../docs/SPEC.md`](../docs/SPEC.md).
3. Press **Update decks** on the Home sheet. Each template is opened read-only,
   filled in, and saved under the configured name. The **Log** sheet lists every
   value written and every reference that failed.

## Notes

* The template file is never modified; a copy is saved.
* If a source workbook is already open in Excel, that copy is used.
* `picture:` uses Excel's own range-to-picture rendering, so the result looks
  exactly like the sheet (fonts, fills, borders).
* `chart:` works on native PowerPoint charts. Charts pasted as Excel OLE
  objects are out of scope; convert them to PowerPoint charts first.
