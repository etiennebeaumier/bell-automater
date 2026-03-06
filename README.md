# Bell Automater

Bell Automater parses BCECN new-issue pricing PDFs, writes the extracted values into a master Excel workbook, and rebuilds the summary charts used in Bell pricing workflows.

It supports three local operating modes:

- CLI for one-off and batch runs
- Interactive terminal mode for guided use
- Streamlit GUI for browser-based use

The repository also includes Power Automate, Azure Function, and Office Script assets for a cloud-hosted version of the same workflow.

## What the Tool Does

- Detects the sending bank from the PDF filename or first-page content
- Parses BCECN PDFs from TD, Scotiabank, CIBC, NBCM, and BMO
- Extracts CAD and USD spreads and yields for 3Y, 5Y, 7Y, 10Y, and 30Y tenors
- Extracts CAD and USD NC5 and NC10 spread/coupon fields when present
- Appends one row per PDF into the `Pricing` sheet
- Rebuilds the four curve charts in `Summary Charts`
- Optionally fetches matching BCECN PDF attachments from Outlook / Exchange Online

## Repository Layout

```text
main.py                                CLI entry point, preflight checks, orchestration
gui.py                                 Streamlit interface
email_fetcher.py                       Outlook / Exchange Online fetch logic
excel_writer.py                        Workbook append + chart generation
parsers/                               Bank-specific PDF parsers
data/                                  Default workbook location
samples/                               Example PDFs
power_automate/flow_design.md          Power Automate flow design
power_automate/azure_function/         Azure Function parser endpoint
power_automate/office_script/          Office Script for Excel Online
requirements.txt                       Local Python dependencies
```

## Requirements

### Local usage

- Python 3.10+
- A master workbook containing these sheets:
  - `Pricing`
  - `Summary Charts`
- A Microsoft 365 mailbox if you want to use Outlook fetch mode

### Python packages

Installed from `requirements.txt`:

- `pdfplumber`
- `openpyxl`
- `exchangelib`
- `msal`
- `streamlit`

### Cloud automation usage

Only required for the Power Automate path:

- Azure Function App
- Office Script in Excel Online
- Power Automate flow
- SharePoint or OneDrive workbook storage

## Setup

### 1. Create and activate a virtual environment

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

### 2. Create `.env`

```bash
cp .env.example .env
```

### 3. Fill in the environment values

```env
OUTLOOK_EMAIL=your.email@bell.ca
OUTLOOK_SERVER=outlook.office365.com
OUTLOOK_DAYS=7
BCECN_SENDER=
MASTER_FILE=data/Master File.xlsx
```

Variable reference:

- `OUTLOOK_EMAIL`: mailbox used for Outlook fetch mode
- `OUTLOOK_SERVER`: Exchange Web Services host, usually `outlook.office365.com`
- `OUTLOOK_DAYS`: default search window for mailbox fetches
- `BCECN_SENDER`: optional sender filter to narrow mailbox results
- `MASTER_FILE`: default workbook path used by CLI and GUI

## Workbook Expectations

The Python writer and the Excel Online Office Script both assume the workbook already contains:

- `Pricing`
- `Summary Charts`

Write behavior:

- Column A: date
- Column B: bank name
- Columns C through AD: parsed CAD and USD spreads, yields, and NC values

Chart behavior:

- Four line charts are rebuilt on each non-dry-run write
- Charts cover CAD spread, CAD yield, USD spread, and USD yield curves
- Charts use only the most recent row per bank

## Running the Tool

### Interactive terminal mode

```bash
python3 main.py
```

If no operation flag is provided, the app opens an interactive menu with these actions:

- Process one PDF
- Process all PDFs in a directory
- Fetch PDFs from Outlook and process them
- Run preflight checks only

### Process one local PDF

```bash
python3 main.py --pdf "samples/BCECN 03.02.26.pdf"
```

### Process all PDFs in a directory

```bash
python3 main.py --dir samples
```

### Dry run

Dry run parses the file and prints a structured preview without writing to Excel.

```bash
python3 main.py --pdf "samples/BCECN 03.02.26.pdf" --dry-run
```

### Preflight only

Preflight checks validate dependencies, workbook accessibility, required sheet names, and fetch settings.

```bash
python3 main.py --check
```

If combined with another command, preflight runs first and stops execution if validation fails.

### Outlook fetch mode

```bash
python3 main.py --fetch --check
```

Fetch mode uses Microsoft device-code authentication:

- You start the command locally
- The app shows a Microsoft sign-in prompt/code
- You complete sign-in in the browser
- The app searches inbox emails with `BCECN` in the subject
- Matching PDF attachments are downloaded to a temporary folder
- Each PDF is parsed and then previewed or written like any local file

This flow does not require storing an Outlook password in `.env`.

### Streamlit GUI

```bash
python3 main.py --gui
```

The GUI supports:

- Uploading one or more local PDFs
- Dry-run preview mode
- Live workbook writes
- Outlook fetch using the same Microsoft device-code login flow

## CLI Flags

```text
--pdf PDF          Path to one PDF file
--dir DIR          Directory containing PDF files
--master MASTER    Path to the master workbook
--fetch            Fetch BCECN PDFs from Outlook
--email EMAIL      Outlook email address
--days DAYS        Days back to search emails
--server SERVER    Exchange server hostname
--sender SENDER    Optional sender email filter
--dry-run          Parse only; do not write the workbook
--check            Run preflight checks first
--interactive      Force the interactive menu
--gui              Launch the Streamlit GUI
```

## Supported Banks and Detection

Supported parser coverage:

- TD
- Scotiabank
- CIBC
- NBCM
- BMO

Bank detection works in this order:

- Filename-based hints
- First-page PDF text

If detection fails, include the bank name in the filename or confirm the first page contains recognizable bank identifiers.

## Examples

Process one PDF into the default workbook:

```bash
python3 main.py --pdf "samples/BCECN 03.02.26.pdf" --check
```

Process a directory into a specific workbook:

```bash
python3 main.py --dir samples --master "data/Master File.xlsx" --check
```

Fetch Outlook attachments from the last 5 days with a sender filter:

```bash
python3 main.py --fetch --days 5 --sender "capitalmarkets@bank.com" --check
```

## Power Automate and Azure Assets

The `power_automate/` folder contains a cloud-hosted variant of the workflow:

- `flow_design.md`: step-by-step Power Automate design
- `azure_function/`: HTTP endpoint that parses incoming PDF payloads
- `office_script/`: Excel Online script that appends rows and rebuilds charts

Use this path when the process should run automatically on mailbox arrival rather than from a local machine.

## Troubleshooting

### Preflight fails on workbook validation

- Confirm the workbook path is correct
- Confirm the workbook contains `Pricing` and `Summary Charts`
- Use `--dry-run` to test parsing without writing to the workbook

### Outlook fetch returns no PDFs

- Confirm the mailbox contains emails with `BCECN` in the subject
- Increase the search range with `OUTLOOK_DAYS` or `--days`
- Remove or loosen the `BCECN_SENDER` or `--sender` filter

### Bank detection fails

- Include the bank name in the filename
- Confirm the first page contains extractable text
- Test with `--dry-run` to isolate parsing from workbook writes

### GUI live mode is blocked

- The workbook path is invalid, unreadable, or missing required sheets
- Fix the path in the sidebar or switch to dry-run mode

## Notes

- Charts are rebuilt on every non-dry-run local write.
- Outlook-fetched files are stored temporarily during processing.
- The Power Automate Office Script mirrors the same column mapping as the local Python writer.
