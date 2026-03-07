# Bell Automater

Bell Automater parses BCECN new-issue pricing PDFs, writes the extracted values into a master Excel workbook, and rebuilds the summary charts used in Bell pricing workflows.

It runs as a standalone desktop application (no Python installation required) or from the command line.

The repository also includes Power Automate, Azure Function, and Office Script assets for a cloud-hosted version of the same workflow.

## What the Tool Does

- Detects the sending bank from the PDF filename or first-page content
- Parses BCECN PDFs from TD, Scotiabank, CIBC, NBCM, and BMO
- Extracts CAD and USD spreads and yields for 3Y, 5Y, 7Y, 10Y, and 30Y tenors
- Extracts CAD and USD NC5 and NC10 spread/coupon fields when present
- Appends one row per PDF into the `Pricing` sheet
- Rebuilds the four curve charts in `Summary Charts`
- Optionally fetches matching BCECN PDF attachments from Outlook / Exchange Online
- Supports dry-run mode to preview parsed data without writing

## Desktop Application

The primary way to use the tool is the standalone desktop app, built with CustomTkinter and packaged with PyInstaller.

### Features

- **Upload PDFs** tab: select one or more local PDFs, process or preview them
- **Fetch from Outlook** tab: pull BCECN attachments from an Outlook mailbox using OAuth2 device-code authentication
- **Dry-run mode**: preview parsed spreads, yields, and hybrid data in the results panel without writing to the workbook
- **Dark / Light / System theme** toggle
- **Persistent settings**: workbook path, email, server, and preferences saved to `~/.bcecn_pricing/config.json`
- **Auto browser open**: during Outlook authentication, the sign-in URL opens automatically and the device code is copied to clipboard

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
  tab_outlook.py                       Fetch from Outlook tab
  settings_panel.py                    Sidebar: workbook path, dry-run, theme
  results_panel.py                     Scrollable log widget
  workers.py                           Background threading workers
config.py                             JSON config manager
main.py                               CLI entry point, orchestration
email_fetcher.py                       Outlook / Exchange Online fetch logic
excel_writer.py                        Workbook append + chart generation
parsers/                               Bank-specific PDF parsers
build_mac.sh                           macOS build script
build_win.bat                          Windows build script
power_automate/flow_design.md          Power Automate flow design
power_automate/azure_function/         Azure Function parser endpoint
power_automate/office_script/          Office Script for Excel Online
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
  - `exchangelib`
  - `msal`
  - `customtkinter`
  - `pyinstaller` (only needed for building)

### Workbook

A master workbook containing these sheets:

- `Pricing`
- `Summary Charts`

### Outlook fetch

A Microsoft 365 mailbox. Authentication uses the OAuth2 device-code flow — no password storage required.

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

Settings (workbook path, email, server, etc.) are configured in the GUI and persisted automatically to `~/.bcecn_pricing/config.json`.

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

### Outlook fetch

```bash
python3 main.py --fetch --check
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
--fetch            Fetch BCECN PDFs from Outlook
--email EMAIL      Outlook email address
--days DAYS        Days back to search emails
--server SERVER    Exchange server hostname
--sender SENDER    Optional sender email filter
--dry-run          Parse only; do not write the workbook
--check            Run preflight checks first
--interactive      Force the interactive menu
```

## Workbook Format

- Column A: date
- Column B: bank name
- Columns C through AD: parsed CAD and USD spreads, yields, and NC values
- Four line charts are rebuilt on each non-dry-run write (CAD spread, CAD yield, USD spread, USD yield)
- Charts use only the most recent row per bank

## Supported Banks

- TD
- Scotiabank
- CIBC
- NBCM
- BMO

Bank detection works by filename hints first, then first-page PDF text. If detection fails, include the bank name in the filename.

## Power Automate and Azure Assets

The `power_automate/` folder contains a cloud-hosted variant of the workflow:

- `flow_design.md`: step-by-step Power Automate design
- `azure_function/`: HTTP endpoint that parses incoming PDF payloads
- `office_script/`: Excel Online script that appends rows and rebuilds charts

Use this path when the process should run automatically on mailbox arrival rather than from a local machine.

## Troubleshooting

### Workbook validation fails

- Confirm the workbook path is correct in the settings panel
- Confirm the workbook contains `Pricing` and `Summary Charts`
- Use dry-run mode to test parsing without writing

### Outlook fetch returns no PDFs

- Confirm the mailbox contains emails with `BCECN` in the subject
- Increase the search range with the Days Back setting or `--days`
- Remove or loosen the sender filter

### Bank detection fails

- Include the bank name in the filename
- Confirm the first page contains extractable text
- Test with dry-run mode to isolate parsing from workbook writes
