# Bell Automater

Automates Bell Canada (BCECN) PDF parsing and Excel updates for internal pricing workflows.

The tool can:
- Parse BCECN indicative new issue PDFs from `TD`, `Scotiabank`, and `CIBC`
- Append parsed values into `Pricing` in a master workbook
- Rebuild yield/spread charts in `Summary Charts`
- Fetch BCECN PDF attachments from Outlook (Exchange/O365)
- Run as CLI, interactive menu, or Streamlit GUI

## Requirements

- Python `3.10+`
- A workbook with required sheets:
  - `Pricing`
  - `Summary Charts`
- Dependencies in `requirements.txt`:
  - `pdfplumber`
  - `openpyxl`
  - `exchangelib`
  - `streamlit`

## Setup

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

Create environment file:

```bash
cp .env.example .env
```

Then set your values in `.env` (especially for Outlook fetch mode):

```env
OUTLOOK_EMAIL=your.email@bell.ca
OUTLOOK_PASSWORD=your_password_here
OUTLOOK_SERVER=outlook.office365.com
OUTLOOK_DAYS=7
BCECN_SENDER=
MASTER_FILE=data/Master File.xlsx
```

## Usage

### 1) Interactive mode (default)

```bash
python3 main.py
```

If no operation flags are provided, the app opens an interactive menu.

### 2) Process one local PDF

```bash
python3 main.py --pdf "samples/BCECN 03.02.26.pdf"
```

### 3) Process all PDFs in a directory

```bash
python3 main.py --dir samples
```

### 4) Dry-run (parse + preview only, no Excel writes)

```bash
python3 main.py --pdf "samples/BCECN 03.02.26.pdf" --dry-run
```

### 5) Fetch from Outlook and process

```bash
python3 main.py --fetch --check
```

`--check` runs preflight diagnostics first (recommended for first run).

### 6) Launch GUI

```bash
python3 main.py --gui
```

This starts Streamlit and opens the app defined in `gui.py`.

## CLI Flags

```text
--pdf PDF            Path to a single PDF file
--dir DIR            Directory of PDF files
--master MASTER      Path to master workbook
--fetch              Fetch PDFs from Outlook
--email EMAIL        Outlook email
--password PASSWORD  Outlook password
--days DAYS          Days back to search emails
--server SERVER      Outlook Exchange server
--sender SENDER      Optional sender filter
--dry-run            Parse only, do not write workbook
--check              Run preflight checks
--interactive        Force interactive menu
--gui                Launch Streamlit GUI
```

## Project Layout

```text
main.py              CLI entry point + preflight + orchestration
gui.py               Streamlit interface
email_fetcher.py     Outlook/Exchange fetch logic
excel_writer.py      Workbook append + chart generation
parsers/             Bank-specific PDF parsers
data/Master File.xlsx
samples/             Example PDFs
```

## Notes

- Bank detection is automatic from filename/content.
- If detection fails, ensure PDF content includes bank identifiers (TD/Scotiabank/CIBC) or include bank name in filename.
- Preflight verifies dependencies, workbook presence/sheets, and fetch configuration.
