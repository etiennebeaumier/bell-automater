# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Commands

```bash
# Set up environment
python3 -m venv .venv && source .venv/bin/activate && pip install -r requirements.txt

# Run desktop GUI
python3 app.py

# CLI usage
python3 main.py --pdf "samples/BCECN 03.02.26.pdf"   # one PDF
python3 main.py --dir samples                          # directory of PDFs
python3 main.py --pdf "..." --dry-run                  # preview without writing
python3 main.py --check                                # preflight only

# Run tests
python3 -m unittest discover -s tests -v

# Run a single test file
python3 -m unittest tests.test_main -v

# Build standalone executable
./build_mac.sh   # macOS
build_win.bat    # Windows
```

## Architecture

The tool has two entry points:
- **`app.py`** — Desktop GUI (CustomTkinter). On startup, if the workbook and PDF folder are configured and parseable PDFs are present, processes them immediately and exits. Otherwise opens the full GUI.
- **`main.py`** — CLI orchestrator. Bank detection → parser dispatch → `excel_writer`. Also runs preflight validation.

**Data flow:**

```
PDF file
  → detect_bank()              (main.py: filename hints first, then PDF content)
  → BANK_PARSERS[key]()        (parsers/*.py: returns a dict with date, bank, and metric keys)
  → append_row()               (excel_writer.py: writes one row to Pricing sheet)
  → deduplicate_pricing_rows() (end-of-batch: dedupes by (date, bank), keeps newest, sorts)
```

**`parsers/`** — One module per bank (td, scotiabank, cibc, nbcm, bmo, desjardins, mizuho). Each exports a single `parse_<bank>_pdf(path) -> dict` function. The dict must contain `date` (datetime), `bank` (str), and metric keys matching `COLUMN_MAP` in `excel_writer.py`.

**`excel_writer.py`** — All workbook I/O lives here:
- `append_row()` — writes one row to the `Pricing` sheet
- `deduplicate_pricing_rows()` — dedupes by `(date, bank)` case-insensitively, keeps newest, sorts by bank then date

**`COLUMN_MAP`** (excel_writer.py) is the canonical mapping between column numbers (1-indexed) and parser dict keys. Columns 3–22 are CAD/USD spread+yield for 3Y/5Y/7Y/10Y/30Y; columns 23–30 are NC5/NC10 spread/coupon.

**`ui/`** — CustomTkinter GUI package:
- `app_window.py` — main window, tab layout, theme switching
- `tab_pdf.py` — "Upload PDFs" tab; calls `workers.py` in a background thread
- `settings_panel.py` — workbook path, PDF source folder, dry-run toggle, theme
- `results_panel.py` — scrollable log widget
- `workers.py` — background threading so the GUI stays responsive during processing

**`config.py`** — Persists settings to `~/.bcecn_pricing/config.json`.

## Adding a New Bank Parser

1. Create `parsers/<bank>.py` with `parse_<bank>_pdf(path: str) -> dict`.
2. Return keys matching `COLUMN_MAP` in `excel_writer.py` (use `None` for missing tenors).
3. Import and register in `main.py`: add to `BANK_PARSERS` dict.
4. Add filename/content detection rules to `detect_bank()` in `main.py`.

## Workbook Requirements

The master workbook must contain a `Pricing` sheet. Column A = date, column B = bank, columns C–AD = metrics per `COLUMN_MAP`.
