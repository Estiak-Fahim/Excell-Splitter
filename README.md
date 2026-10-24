# Excel Splitter

A Windows desktop app that takes an Excel file, lets you pick a column, and
generates one Excel file per unique value in that column.

```
Customer_Data.xlsx  --split by "Territory"-->  Dhaka.xlsx, Chittagong.xlsx, Sylhet.xlsx, ...
```

Built with Python 3.12, PySide6, pandas, and openpyxl; packaged as a
standalone `ExcelSplitter.exe` with PyInstaller (no Python required to run
the finished app).

---

## Contents

- [Quick start](#quick-start)
- [Building ExcelSplitter.exe](#building-excelsplitterexe)
- [Project structure](#project-structure)
- [How it works](#how-it-works)
- [Known limitations](#known-limitations)
- [Running the tests](#running-the-tests)
- [Troubleshooting](#troubleshooting)

---

## Quick start

Requires Python 3.12+ on Windows.

```bat
python -m venv venv
venv\Scripts\activate
pip install -r requirements.txt
python main.py
```

That's it -- the app window should open. Try it on the included sample file
at `sample_data\Customer_Data.xlsx`.

The workflow is: **select a file -> pick the column to split by -> pick
which columns to keep -> pick an output folder -> Generate Files.** Every
other setting has a sensible default, so you can go from opening the app to
generated files in about five clicks.

---

## Building ExcelSplitter.exe

Once the Quick Start steps above have been run at least once (so `venv\`
exists with dependencies installed):

```bat
build.bat
```

This runs PyInstaller against `ExcelSplitter.spec` and produces
`dist\ExcelSplitter.exe` -- a single file that runs on a clean Windows
machine with no Python installation required. Launching it opens the app
directly with no console window.

If you'd rather build manually instead of via `build.bat`:

```bat
pip install -r requirements.txt
pyinstaller ExcelSplitter.spec --noconfirm
```

**Build on Windows.** PyInstaller does not cross-compile -- to produce a
Windows `.exe` you need to actually run PyInstaller *on* Windows (a real
machine, or a Windows VM/CI runner). Building on macOS/Linux will produce a
macOS/Linux binary instead.

**onefile vs. onedir:** the spec file defaults to a single-file build
(simplest to distribute). If you'd prefer faster startup at the cost of
shipping a folder instead of one file, open `ExcelSplitter.spec` and change
`onefile = True` to `onefile = False`, then rebuild.

---

## Project structure

```
ExcelSplitter/
├── main.py                    Application entry point
├── requirements.txt           Runtime + build dependencies
├── requirements-dev.txt       Extra dependency (pytest) for running tests
├── build.bat                  Windows build script
├── ExcelSplitter.spec         PyInstaller configuration
├── README.md
│
├── app/
│   ├── models.py               Typed data structures (SplitConfig, ProcessingResult, ...)
│   ├── exceptions.py           User-facing error types
│   ├── filename_utils.py       Filename sanitization, collision handling
│   ├── excel_reader.py         Reading workbooks (sheets, data, dedup headers)
│   ├── excel_processor.py      Core split/grouping logic + run_split() orchestration
│   ├── excel_writer.py         Writing output files (formatting preservation + fast path)
│   ├── settings.py             Small local JSON settings store
│   ├── logging_config.py       Rotating debug log setup
│   │
│   ├── theme.py                 Qt stylesheet (visual design)
│   ├── widgets.py                Drop zone, checkable column list, unique-values table
│   ├── completion_dialog.py      "Files created" summary dialog
│   ├── os_utils.py               Open folder / reveal file in Explorer
│   ├── workers.py                QThread workers (keeps the GUI responsive)
│   └── main_window.py            Main window: layout + wiring
│
├── tests/                      pytest suite (see "Running the tests" below)
├── sample_data/
│   ├── generate_sample_data.py  Regenerates the sample workbook
│   └── Customer_Data.xlsx       Ready-to-use sample file
└── resources/icons/app_icon.ico
```

The `app/` package is split deliberately: everything through
`logging_config.py` has **no dependency on PySide6 or any GUI code** --
it's plain Python + pandas/openpyxl that can be imported, tested, or
reused (e.g. from a script) independently of the GUI. Only the second
group of files touches Qt. If you want to change how splitting works, look
in `excel_processor.py` / `excel_writer.py`; if you want to change how it
looks, look in `theme.py` / `main_window.py`.

---

## How it works

1. **Reading a file** (`excel_reader.py`): cell values are read directly
   through openpyxl (`.xlsx`/`.xlsm`) or xlrd (`.xls`) rather than via
   `pandas.read_excel()` directly. This is deliberate: pandas' own Excel
   parser will silently convert a text-formatted column into numbers
   whenever every value in it happens to look "cleanly numeric" (e.g. a
   Phone or Zip Code column with no row that breaks the pattern), which
   quietly **drops leading zeros**. Reading cell-by-cell and trusting each
   engine's reported cell type keeps text as text and numbers as numbers,
   exactly as authored in the source file. Duplicate column headers
   (`Phone`, `Phone`) are renamed to stay unique internally (`Phone`,
   `Phone_2`) while the original header text is preserved for display.

2. **Grouping** (`excel_processor.analyze_split_column`): values are
   trimmed of leading/trailing whitespace for the purpose of *grouping*
   only (`"Dhaka"`, `" Dhaka"`, `"Dhaka "` all become one group) --
   the underlying row data is never modified. Blank/NaN/whitespace-only
   values become their own group, handled per the "Blank Values" setting.

3. **Choosing output columns** (`excel_processor.determine_output_columns`):
   always resolved by filtering the *original* column list, so output
   column order matches the source file. The split column is added or
   removed from that set based on "Split Column in Output" specifically --
   this setting always wins over that column's own checkbox in the
   "Columns to Include" list, so excluding it from the output never
   affects which rows go to which file.

4. **Writing files** (`excel_writer.py`): two paths, chosen automatically.
   - For `.xlsx`/`.xlsm` sources, `extract_source_formatting()` reads the
     source header's font/fill/border/alignment, per-column widths and
     number formats, freeze panes, and autofilter **once**, and
     `write_group_file()` (openpyxl) reapplies them to every generated
     file.
   - For `.xls` sources (and as a safety fallback if reading source styles
     fails for any reason), `write_group_file_fast()` (xlsxwriter) skips
     style reconstruction -- there's nothing reliable to copy from a
     legacy source, and xlsxwriter can't read an existing file's styles
     anyway (it's write-only) -- but still produces a clean, bold-header,
     frozen-header-row, auto-filtered result, and is noticeably faster on
     very large sheets.

5. **Filenames** (`filename_utils.py`): three separate steps. First,
   `sanitize_filename()` makes one value into a safe Windows filename
   (invalid characters, reserved device names like `CON`/`AUX`, length,
   trailing dots/spaces). Second, `deduplicate_filename()` handles two
   *different* unique values sanitizing to the same string within one run
   (e.g. `"Dhaka/Zone"` and `"Dhaka:Zone"` both becoming `Dhaka_Zone`).
   Third, `resolve_output_path()` handles a file left over from a
   *previous* run, per the Overwrite / Skip / Create new filename setting.

6. **Threading** (`workers.py`): loading a file and running a split both
   happen on a background `QThread`, so the window never freezes -- this
   matters most on large files, where reading tens of thousands of rows or
   writing dozens of output files can take a few seconds.

---

## Known limitations

These are deliberate, documented trade-offs rather than bugs -- each one
is also surfaced in the app itself (as a warning on the completion screen)
when it applies to the file you're processing.

- **Legacy `.xls` formatting isn't preserved.** openpyxl (used to read
  styles) only supports the modern `.xlsx`/`.xlsm` format. `.xls` output
  files get clean default formatting instead of a copy of the source
  styling. Convert the source file to `.xlsx` in Excel first if you need
  full formatting preservation.
- **VBA macros are never copied into split output files**, even from an
  `.xlsm` source. Excel Splitter's output files are new, filtered subsets
  of the data (not the original workbook modified in place), and macros
  generally reference specific ranges/sheet structures that wouldn't
  correspond to a split-out subset anyway. The original `.xlsm` file
  itself is never modified.
- **Per-column, not per-cell, style copying.** A generated column's number
  format is taken from that column's first data row in the source file.
  If a single source column genuinely mixes different number formats
  row-to-row (e.g. some cells formatted as currency, others as plain
  numbers, in the same column), only the dominant/first-seen format is
  reproduced.
- **A cell's value type follows the source cell's own type.** Genuinely
  mixed-type columns (some rows numeric, some text, in the *same* column)
  are grouped/written using each cell's own type as read; a numeric
  `5000` and a text `"5000"` are treated as distinct unique values if the
  split column itself is mixed like this.
- **Formulas are not preserved.** Source cells are read for their current
  value (openpyxl is loaded with `data_only=True`), not their formula
  text, since a formula referencing other rows generally wouldn't make
  sense copied into a filtered subset of the data.

---

## Running the tests

```bat
pip install -r requirements.txt -r requirements-dev.txt
pytest
```

`tests/` covers, across `test_filename_utils.py`, `test_excel_reader.py`,
`test_excel_processor.py`, `test_excel_writer.py`, and
`test_integration.py`: splitting by a text column, splitting by a numeric
column, excluding/including the split column, selecting specific columns,
selecting all columns, blank split values, duplicate values, filename
sanitization (invalid characters, reserved names, length, collisions),
existing-output-file handling (overwrite/skip/create new), a 10,000-row
dataset, and multiple worksheets. All test fixtures are generated on the
fly (no checked-in binary test files), so the suite is fully
self-contained.

The GUI itself (`main_window.py`, `workers.py`, `widgets.py`,
`completion_dialog.py`) isn't covered by these automated tests -- PySide6
doesn't have a meaningful way to unit-test window/widget behavior without
a display. It's written to depend only on the modules above, all of which
*are* fully tested, so the GUI layer's job is mostly plumbing: reading
widget state into a `SplitConfig`, and displaying a `ProcessingResult`.

---

## Troubleshooting

**"This file appears to be open in another program"** -- close the file in
Excel and try again. Excel locks files for writing, and sometimes for
reading too depending on how it was opened.

**The built `.exe` won't start / closes immediately** -- run it from a
Command Prompt instead of double-clicking, so you can see any error
output: `dist\ExcelSplitter.exe`. If you see a `ModuleNotFoundError`, the
module needs to be added to `hiddenimports` in `ExcelSplitter.spec`
(pandas' and openpyxl's more dynamic imports are already collected, but a
future version bump could introduce a new one).

**A detailed log is always kept** at
`%LOCALAPPDATA%\ExcelSplitter\logs\excelsplitter.log`, regardless of what
the on-screen error message says -- useful for diagnosing anything that
doesn't have an obvious cause from the UI alone.

**Settings** (last output folder, window size, filename format, existing-
file preference) are stored at `%APPDATA%\ExcelSplitter\settings.json` --
delete this file to reset the app to defaults.
