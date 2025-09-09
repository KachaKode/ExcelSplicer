# Range Paster GUI - README

A desktop tool for quickly copying worksheet ranges from one or more Excel source workbooks into a single base workbook, with smart wildcards, row lookups, and repeatable workflows that you can save and reload.

Built with Tkinter and OpenPyXL in pure Python.

---

## Table of contents

* Purpose
* Key features
* How it works
* Installation
* Launching the app
* UI walkthrough
* Wildcards and how they resolve
* Start and End Row references
* Titles for sources
* Saving and loading a workspace
* Output log and debugging
* Example workflows
* Workspace JSON structure
* Tips and constraints
* Troubleshooting
* FAQ
* Roadmap
* License

---

## Purpose

When you need to aggregate tables or blocks from several Excel files into one destination sheet, doing it by hand is slow and error prone. This tool lets you define:

* A base workbook and one or more base cells, called tracks.
* A list of source workbooks and the ranges to pull from each.
* Optional wildcards for ranges, so the tool finds the true bounds automatically.
* Optional row references, so the tool can align source rows based on values in the base file.

You can run the process with one click, log everything that happened, and save the setup for later.

---

## Key features

* Multiple tracks. Each track is a base cell where pastes begin. Tracks advance horizontally as data is pasted, so you can build wide summary sheets.
* Multiple sources. Each source points to a workbook, a range, and a track.
* Titles for sources. Add a friendly title for logging and clarity.
* Sheet-aware ranges. Use either `A1:C10` or `Sheet1!A1:C10`.
* Wildcards in ranges. Use `?` to say "find it" rather than hardcoding. Examples:

  * `N?:B?` - start at column N and a row determined by a reference or scan, end at column B and a row determined by reference or scan.
  * `A23:B?` - end row is discovered automatically.
  * `A23:?56` - end column is discovered automatically.
  * `A23:??` - both end column and end row discovered automatically.
* Reference-based row resolution. Use values from the base sheet to find matching rows in a source sheet:

  * Start Row Ref - points to a cell in the base workbook whose value should be found in the source to anchor the start row.
  * End Row Ref - similar anchor for the end row.
* Smart scanning when references are absent. If a wildcard is present and no reference is provided, the tool finds the first or last non-empty row or column using the source worksheet data.
* Safe paste behavior. Copy values only, preserve destination formatting.
* Full, sheet-aware logging. The app logs every step, including the fully resolved source range in A1 notation, and the paste destination range.
* Workspace persistence. Save and reload all settings to JSON.

---

## How it works

High level flow:

1. You select a base workbook and define one or more tracks. Each track is a base cell, for example `Sheet1!C13`, plus optional row reference cells to help resolve wildcards.
2. You add one or more sources. Each source has a title, a file path, a source range string, and the track it should paste into.
3. For each source, the app resolves the source range into concrete boundaries:

   * Sheet is parsed from `Sheet!Ref` if present.
   * Columns and rows are parsed from the left and right tokens of the range.
   * Wildcards are resolved using references or data scan, described below.
   * The resolved A1 range is logged.
4. The app copies values from the resolved source range and pastes them into the base workbook at the current position of the chosen track.
5. The track cursor advances to the right by the width of what was pasted.
6. When all sources are processed, the base workbook is saved in place.

All steps are logged to the Output Log window.

---

## Installation

Requirements:

* Python 3.9 or newer recommended.
* OpenPyXL for Excel read and write.

Install OpenPyXL:

```bash
pip install openpyxl
```

Tkinter ships with most standard Python distributions. If your Python does not include Tkinter, install it using your OS package manager.

---

## Launching the app

Run the script:

```bash
python range_paster_gui.py
```

The main window will open.

---

## UI walkthrough

### Base File

* Pick the Excel workbook that you want to paste into.
* This workbook is saved in place when processing completes.

### Base Cells (Tracks)

* Add one or more tracks. Each track has:

  * Base Cell, for example `A1` or `Sheet1!C13`. This sets the starting paste position for that track.
  * Start Row Ref, optional. A cell reference in the base workbook, for example `Sheet1!N21`. The value found in that cell is used to locate the start row inside a source sheet if a wildcard row is used.
  * End Row Ref, optional. Same idea for the end row.
* You can remove a track, but the app requires at least one.

### Source Files

Add one row per source:

* Title. A friendly name for logs, optional.
* Source File. The workbook to read from.
* Source Range. A range with optional sheet and wildcards. Examples:

  * `Sheet2!A15:G25`
  * `N?:B?`
  * `A23:?56`
  * `A23:??`
  * `B7` for a single cell
* Track (Base Cell). Which track this source should paste into.

### Buttons

* Process Ranges. Runs the process, logs actions, and saves the base file.
* Clear All. Resets the form and log.
* Save Workspace. Saves your setup to JSON.
* Load Workspace. Restores a saved setup.

### Output Log

Shows detailed progress and debug information. It also receives the module logger output so helper functions can report how wildcards were resolved.

---

## Wildcards and how they resolve

A range is parsed as `left:right`.

* `left` and `right` each contain a column token and an optional row token.

  * Examples: `A23`, `N?`, `?56`, `??`, `B`
* `?` inside a column or row means the tool must resolve it.

Resolution rules:

1. Start column

   * If specified as a letter, use it directly.
   * If `?`, default to column A.

2. Start row

   * If an explicit number exists, use it.
   * If `?`, try Start Row Ref:

     * Read the value from the referenced base cell.
     * Search that value in the source sheet across non-empty columns.
     * If found, use that row.
     * If not found, fall back to the first non-empty row on the start column.

3. End column

   * If an explicit letter exists, use it.
   * If `?`, compute temporarily using the entire sheet height, then refine after the end row is known. Final pass uses `last_nonempty_col_in_row_range` across the resolved row span.

4. End row

   * If an explicit number exists, use it.
   * If `?`, try End Row Ref first using the same lookup method as Start Row Ref.
   * If there is no End Row Ref or the lookup fails, compute the last non-empty row across the span of columns `[start_col..end_col]`, starting at `start_row`.

After both sides are resolved, the app logs something like:

```
[calc] resolved source range -> Sheet2!B12:H34
```

---

## Start and End Row references

Set these in each track if you plan to use wildcard rows:

* Start Row Ref helps resolve a `?` start row by anchoring to a value from the base workbook that should also appear in the source sheet.
* End Row Ref helps resolve a `?` end row in the same way.

If either reference is missing and a wildcard row is used, the app falls back to data scans to find first or last non-empty rows.

References can be sheet qualified, for example `Sheet1!N21`, or relative to the active sheet if you omit `Sheet!`.

---

## Titles for sources

Each source has an optional Title field. It appears in logs like:

```
Processing source 2 Q2 Sales (sales_q2.xlsx) on track #1 [Sheet1!C13]: A23:??
```

This helps you scan the Output Log quickly.

---

## Saving and loading a workspace

* Save Workspace creates a JSON file with:

  * Base file path
  * Track definitions
  * Source rows, including titles
* Load Workspace restores everything, then rebinds source rows to the correct track label where possible.

You can edit the JSON by hand if needed. See the structure section below.

---

## Output log and debugging

The app writes both high level progress and detailed debug lines. Examples you will see:

* Which files were opened.
* How tokens were split.
* What the resolved rows and columns are after wildcard logic.
* The final, sheet-aware A1 reference of each source range.
* Where the paste landed on the base sheet.
* Cursor advancement per track.

There is a module level logger hook:

```python
LOGGER = None

def set_logger(fn):
    global LOGGER
    LOGGER = fn

def log_debug(msg: str):
    if LOGGER:
        LOGGER(msg)
    else:
        print(msg)
```

The GUI calls `set_logger(self.log)` during startup. All helper functions call `log_debug`, which writes to the Output Log.

---

## Example workflows

### Example 1 - Simple copy without wildcards

* Base file: `report.xlsx`
* Track: `Sheet1!C13`
* Source: `sales_q1.xlsx`, range `Sheet2!A15:G25`, track `Sheet1!C13`

Run Process Ranges.

Log will show the resolved range and the paste destination on `Sheet1`.

### Example 2 - Use Start Row Ref to align a changing table

* Base file: `report.xlsx`
* Track:

  * Base Cell: `Sheet1!C13`
  * Start Row Ref: `Sheet1!N21` - this cell contains a customer ID that also appears in the source.
  * End Row Ref: blank
* Source:

  * File: `customers.xlsx`
  * Range: `N?:B?` - start row and end row are wildcards
  * Track: the track above

Process Ranges.

* Start row resolved using the value in `N21` on the base sheet, searched inside `customers.xlsx`.
* End row resolved by scanning down across columns from start to end columns.

---

## Workspace JSON structure

Saved files contain something like:

```json
{
  "base_file_path": "/path/to/base.xlsx",
  "tracks": [
    {
      "base_cell": "Sheet1!C13",
      "start_ref": "Sheet1!N21",
      "end_ref": "Sheet1!A250"
    }
  ],
  "sources": [
    {
      "title": "Q1 Sales",
      "file_path": "/path/to/sales_q1.xlsx",
      "range": "Sheet2!A15:G25",
      "track_label": "Sheet1!C13"
    },
    {
      "title": "Customers block",
      "file_path": "/path/to/customers.xlsx",
      "range": "N?:B?",
      "track_label": "Sheet1!C13"
    }
  ]
}
```

Notes:

* `track_label` must match the display label of a track, which is the literal Base Cell text you entered, for example `Sheet1!C13`. The loader tries to match labels and defaults to the first track if it cannot.

---

## Tips and constraints

* The tool copies values only. It does not copy styles or formulas. `data_only=True` is used when loading sources, so evaluated values are read.
* Worksheets are addressed by name. Keep names stable across runs.
* The app writes changes directly into the base file you selected.
* The script scans up to `ws.max_column` and `ws.max_row` on the source sheet when it needs to resolve wildcards.
* Ranges must be well formed. If you get a parsing error, check for typos in the range string.

---

## Troubleshooting

**Error: Sheet "X" not found in workbook.**
Check that the sheet name is correct and exists in the file.

**Error: Invalid start column in "left".**
The left side of the range must start with a valid column token. For example `A23`, `N?`, or `B`.

**Crash when using "N?" on the left.**
Fixed in this version. The code does not call `parse_cell` on wildcard tokens. It uses `split_col_row` first and resolves the row before converting to coordinates.

**The tool pasted into the wrong place.**
Check the selected track for the source. Each paste advances the track cursor to the right by the width of what was pasted.

**Start or End Row Ref did not match anything.**
The app falls back to scanning. The Output Log will show both the lookup attempt and the fallback path.

**Cannot open file or file not found.**
Verify that the file paths are correct. On Windows, verify that the file is not open in Excel with write locks that could block saving the base workbook.

---

## FAQ

**How do I include a sheet in a range?**
Use `SheetName!A1:C10`. If you omit the sheet, the active sheet of that workbook is used.

**Can I use a single cell range?**
Yes. Enter `B7` or `Sheet2!B7`.

**What happens to formulas in the source?**
The tool reads values only from sources, not formulas.

**What does the app do when both start row and end row are `?`?**
It tries references first. If references are missing or not found, it uses first and last non-empty row scans across the resolved column span.

---

## Roadmap

* Optional paste orientation control, for example advance vertically instead of horizontally.
* Option to copy styles.
* Preview of resolved ranges before running the copy.
* Error markers inline in the UI next to misconfigured fields.
* Batch mode command line entry point to run workspaces without opening the GUI.

---

## License

This project is provided as is. You may adapt it for your own workflows. If you redistribute modified versions, please preserve credit in comments.

---

