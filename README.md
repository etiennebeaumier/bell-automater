# Bell Automater

Bell Automater parses BCECN new-issue pricing PDFs and writes the extracted values into a master Excel workbook.

It runs as a standalone desktop application (no Python installation required) or from the command line.

## What the Tool Does

- Detects the sending bank from the PDF filename or first-page content
- Parses BCECN PDFs from TD, Scotiabank, CIBC, NBCM, BMO, Desjardins, Mizuho, and BofA
- Extracts CAD and USD spreads and yields for 3Y, 5Y, 7Y, 10Y, and 30Y tenors
- Extracts CAD and USD NC5 and NC10 spread/coupon fields when present
- Appends one row per PDF into the `Pricing` sheet during processing
- Runs end-of-run post-processing (non-dry-run only): removes duplicate `Pricing` rows by `(date, bank)` case-insensitively (keep newest), reorders rows by bank/date
- Supports dry-run mode to preview parsed data without writing

## Desktop Application

The primary way to use the tool is the standalone desktop app, built with CustomTkinter and packaged with PyInstaller.

### Features

- **Upload PDFs** tab: select one or more local PDFs, process or preview them
- **Dry-run mode**: preview parsed spreads, yields, and hybrid data in the results panel without writing to the workbook
- **Dark / Light / System theme** toggle
- **Persistent settings**: workbook path, default PDF source folder, and preferences saved to `~/.bcecn_pricing/config.json`
- **Quick startup mode**: when workbook + preferred PDF folder are valid and parseable PDFs are present, processes them immediately and exits without opening the full interface

### Quick Startup Flow

At app launch (`python3 app.py` or packaged app):

1. Validate saved workbook path (`Pricing` sheet required).
2. Validate saved preferred PDF source folder.
3. Check for at least one parseable PDF in that folder.

If all checks pass, PDFs are processed immediately and the app exits.

If any check fails, the app opens the full GUI.

### Running the App

From source:

```bash
python3 app.py
```

Or use the pre-built executable (no Python needed):

- **macOS**: `dist/BCECN Pricing Tool.app`
- **Windows**: `dist/BCECN Pricing Tool.exe`

### Building the Executable

macOS:

```bash
./build_mac.sh
```

Windows:

```bat
build_win.bat
```

Both scripts use PyInstaller to produce a standalone single-file executable.

## Repository Layout

```text
app.py                                 Desktop app entry point
ui/                                    CustomTkinter GUI package
  app_window.py                        Main window, tab layout, theme
  tab_pdf.py                           Upload PDFs tab
  settings_panel.py                    Sidebar: workbook path, PDF source folder, dry-run, theme
  results_panel.py                     Scrollable log widget
  workers.py                           Background threading workers
config.py                             JSON config manager
main.py                               CLI entry point, orchestration
excel_writer.py                        Workbook append and dedupe
parsers/                               Bank-specific PDF parsers
build_mac.sh                           macOS build script
build_win.bat                          Windows build script
requirements.txt                       Python dependencies
```

## Requirements

### Desktop app (pre-built)

No requirements. Run the executable directly.

### Running from source

- Python 3.10+
- Dependencies from `requirements.txt`:
  - `pdfplumber`
  - `openpyxl`
  - `customtkinter`
  - `pyinstaller` (only needed for building)

### Workbook

A master workbook containing a `Pricing` sheet.

## Setup (from source)

### 1. Create and activate a virtual environment

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

### 2. Launch the app

```bash
python3 app.py
```

Settings (workbook path, default PDF source folder, etc.) are configured in the GUI and persisted automatically to `~/.bcecn_pricing/config.json`.

## Testing

Run automated tests:

```bash
python3 -m unittest discover -s tests -v
```

Current test coverage includes workbook deduplication behavior and CLI post-processing orchestration.

## CLI Usage

The CLI is still available for scripting and automation.

### Process one local PDF

```bash
python3 main.py --pdf "samples/BCECN 03.02.26.pdf"
```

### Process all PDFs in a directory

```bash
python3 main.py --dir samples
```

### Dry run

```bash
python3 main.py --pdf "samples/BCECN 03.02.26.pdf" --dry-run
```

### Preflight checks

```bash
python3 main.py --check
```

### Interactive terminal mode

```bash
python3 main.py
```

### CLI Flags

```text
--pdf PDF          Path to one PDF file
--dir DIR          Directory containing PDF files
--master MASTER    Path to the master workbook
--dry-run          Parse only; do not write the workbook
--check            Run preflight checks first
--interactive      Force the interactive menu
```

## Workbook Format

- Column A: date
- Column B: bank name
- Columns C through AD: parsed CAD and USD spreads, yields, and NC values
- For non-dry-run runs, duplicate `Pricing` rows are removed once at end-of-run and valid rows are sorted by bank/date

### Duplicate Handling

- Duplicate key: `(Date, Bank)` with bank matched case-insensitively
- Keep policy: newest row in workbook order is preserved per duplicate key
- Post-dedupe ordering: valid rows are sorted by bank (case-insensitive) then date (earliest to latest)
- Rows with missing date or bank are not part of deduplication and are kept at the bottom in their original relative order

## Supported Banks

- TD
- Scotiabank
- CIBC
- NBCM
- BMO
- Desjardins
- Mizuho
- BofA

Bank detection works by filename hints first, then first-page PDF text. If detection fails, include the bank name in the filename.

## Troubleshooting

### Workbook validation fails

- Confirm the workbook path is correct in the settings panel
- Confirm the workbook contains a `Pricing` sheet
- Use dry-run mode to test parsing without writing

### Bank detection fails

- Include the bank name in the filename
- Confirm the first page contains extractable text
- Test with dry-run mode to isolate parsing from workbook writes
