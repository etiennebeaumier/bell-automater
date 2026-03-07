#!/usr/bin/env python3
"""Bell Canada (BCECN) yield spread automater.

Usage:
    # Process a local PDF file:
    python main.py --pdf "BCECN 03.02.26.pdf"

    # Fetch from Outlook and process:
    python main.py --fetch --email etienne.beaumier@bell.ca

    # Or configure .env and just run:
    python main.py --fetch

    # Process all PDFs in a directory:
    python main.py --dir ./pdfs/
"""

import argparse
import importlib.util
import os
import sys
from datetime import datetime
from parsers.td import parse_td_pdf
from parsers.scotiabank import parse_scotiabank_pdf
from parsers.cibc import parse_cibc_pdf
from parsers.nbcm import parse_nbcm_pdf
from parsers.bmo import parse_bmo_pdf

MASTER_FILE = os.path.join("data", "Master File.xlsx")
REQUIRED_SHEETS = ("Pricing", "Summary Charts")

BANK_PARSERS = {
    "td": parse_td_pdf,
    "scotiabank": parse_scotiabank_pdf,
    "cibc": parse_cibc_pdf,
    "nbcm": parse_nbcm_pdf,
    "bmo": parse_bmo_pdf,
}


def load_env_file(path: str = ".env") -> None:
    """Load KEY=VALUE pairs from a local .env file into os.environ."""
    if not os.path.exists(path):
        return

    with open(path, "r", encoding="utf-8") as f:
        for raw_line in f:
            line = raw_line.strip()
            if not line or line.startswith("#") or "=" not in line:
                continue
            key, value = line.split("=", 1)
            key = key.strip()
            value = value.strip()
            if not key:
                continue

            # Strip wrapping quotes for values like PASSWORD="abc123".
            if (value.startswith('"') and value.endswith('"')) or (
                value.startswith("'") and value.endswith("'")
            ):
                value = value[1:-1]

            os.environ.setdefault(key, value)


def env_int(name: str, default: int) -> int:
    """Read an integer from environment with a safe fallback."""
    raw = os.environ.get(name)
    if raw is None:
        return default
    try:
        return int(raw)
    except ValueError:
        print(f"Warning: invalid {name}={raw!r}; using {default}")
        return default


def _module_available(module_name: str) -> bool:
    """Return True if a Python module can be imported."""
    return importlib.util.find_spec(module_name) is not None


def run_preflight(
    master_file: str,
    require_workbook: bool = True,
    check_fetch: bool = False,
    email: str | None = None,
    days: int | None = None,
    server: str | None = None,
    sender: str | None = None,
    verbose: bool = True,
) -> bool:
    """Validate environment, dependencies, workbook, and optional fetch config."""
    failures = 0
    warnings = 0

    def report(level: str, message: str) -> None:
        nonlocal failures, warnings
        if level == "FAIL":
            failures += 1
        elif level == "WARN":
            warnings += 1
        if verbose or level == "FAIL":
            print(f"[{level}] {message}")

    if verbose:
        print("Running preflight checks...")

    if os.path.exists(".env"):
        report("PASS", "Found .env file.")
    else:
        report("WARN", "No .env file found (CLI flags still work). Copy .env.example to .env for easier setup.")

    if _module_available("pdfplumber"):
        report("PASS", "Dependency check: pdfplumber available.")
    else:
        report("FAIL", "Missing dependency: pdfplumber. Install with `pip install pdfplumber`.")

    openpyxl_available = _module_available("openpyxl")
    if openpyxl_available:
        report("PASS", "Dependency check: openpyxl available.")
    elif require_workbook:
        report("FAIL", "Missing dependency: openpyxl. Install with `pip install openpyxl`.")
    else:
        report("WARN", "openpyxl is missing. Install with `pip install openpyxl` before non-dry-run writes.")

    if check_fetch:
        if _module_available("exchangelib"):
            report("PASS", "Dependency check: exchangelib available for Outlook fetch.")
        else:
            report("FAIL", "Missing dependency: exchangelib. Install with `pip install exchangelib`.")

    if require_workbook:
        if not os.path.exists(master_file):
            report("FAIL", f"Master workbook not found: {master_file}")
        elif not openpyxl_available:
            report("FAIL", "Cannot validate workbook sheets because openpyxl is unavailable.")
        else:
            try:
                from openpyxl import load_workbook

                wb = load_workbook(master_file, read_only=True, data_only=True)
                missing_sheets = [s for s in REQUIRED_SHEETS if s not in wb.sheetnames]
                wb.close()

                if missing_sheets:
                    report(
                        "FAIL",
                        f"Workbook is missing required sheets: {', '.join(missing_sheets)}",
                    )
                else:
                    report("PASS", f"Workbook check passed: found {', '.join(REQUIRED_SHEETS)} sheets.")
            except Exception as exc:
                report("FAIL", f"Failed to open workbook '{master_file}': {exc}")
    else:
        report("PASS", "Workbook validation skipped (dry-run mode).")

    if check_fetch:
        if email:
            report("PASS", f"Fetch config: email set to {email}.")
        else:
            report("FAIL", "Fetch config: missing Outlook email. Use --email or OUTLOOK_EMAIL.")

        if server:
            report("PASS", f"Fetch config: server set to {server}.")
        else:
            report("FAIL", "Fetch config: missing Outlook server. Use --server or OUTLOOK_SERVER.")

        if days is None or days < 1:
            report("FAIL", "Fetch config: days must be >= 1.")
        else:
            report("PASS", f"Fetch config: searching last {days} day(s).")

        if sender:
            report("PASS", f"Fetch config: sender filter enabled ({sender}).")
        else:
            report("WARN", "Fetch config: no sender filter set; search scope may be broad.")

    if verbose or failures:
        if failures:
            print(f"Preflight failed with {failures} error(s) and {warnings} warning(s).")
        else:
            print(f"Preflight passed with {warnings} warning(s).")

    return failures == 0


def detect_bank(pdf_path: str) -> str:
    """Detect which bank sent the PDF based on filename, then PDF content."""
    if not os.path.isfile(pdf_path):
        raise FileNotFoundError(f"PDF file not found: {pdf_path}")

    filename = pdf_path.lower()

    # Filename-based detection
    if "scotiabank" in filename or "bns" in filename:
        return "scotiabank"
    if "cibc" in filename:
        return "cibc"
    if "nbcm" in filename or "national bank" in filename:
        return "nbcm"
    if "bmo" in filename:
        return "bmo"
    if "td" in filename:
        return "td"

    # Content-based detection as fallback
    import pdfplumber

    with pdfplumber.open(pdf_path) as pdf:
        if not pdf.pages:
            raise ValueError(f"PDF appears empty: {pdf_path}")
        text = (pdf.pages[0].extract_text() or "").lower()

    if "scotiabank" in text or "bns-internal" in text:
        return "scotiabank"
    if "cibc capital markets" in text or "cibc" in text:
        return "cibc"
    if "national bank" in text or "nbcm" in text:
        return "nbcm"
    if "bmo nesbitt burns" in text or "bmo capital markets" in text:
        return "bmo"
    if "td securities" in text:
        return "td"

    raise ValueError(f"Could not detect bank for: {pdf_path}")


def _format_preview_value(key: str, value) -> str:
    """Format parsed values for dry-run preview output."""
    if value is None:
        return "-"
    if isinstance(value, datetime):
        return value.strftime("%Y-%m-%d")
    if isinstance(value, (int, float)):
        if "yield" in key or "coupon" in key:
            return f"{value:.3%}"
        return f"{value:.2f}".rstrip("0").rstrip(".")
    return str(value)


def print_dry_run_preview(pdf_path: str, data: dict) -> None:
    """Print a user-friendly preview of parsed values without writing Excel."""
    ordered_keys = []
    for currency in ("cad", "usd"):
        for metric in ("spread", "yield"):
            for tenor in ("3y", "5y", "7y", "10y", "30y"):
                ordered_keys.append(f"{currency}_{metric}_{tenor}")
    for currency in ("cad", "usd"):
        for nc in ("nc5", "nc10"):
            ordered_keys.append(f"{currency}_{nc}_spread")
            ordered_keys.append(f"{currency}_{nc}_coupon")

    metric_keys = [k for k in data.keys() if k not in {"date", "bank"}]
    extras = sorted(k for k in metric_keys if k not in ordered_keys)
    preview_keys = [k for k in ordered_keys if k in data] + extras

    date_val = data.get("date")
    if isinstance(date_val, datetime):
        date_text = date_val.strftime("%Y-%m-%d")
    else:
        date_text = str(date_val)

    print("\n" + "=" * 72)
    print(f"DRY RUN PREVIEW: {pdf_path}")
    print(f"Bank: {data.get('bank', 'Unknown')} | Date: {date_text}")
    print("-" * 72)
    for key in preview_keys:
        print(f"{key:<20} {_format_preview_value(key, data.get(key))}")
    print("=" * 72)


def process_pdf(pdf_path: str, master_file: str = MASTER_FILE, dry_run: bool = False):
    """Parse a single PDF and write its data to the Master File."""
    bank_key = detect_bank(pdf_path)
    parser = BANK_PARSERS.get(bank_key)
    if not parser:
        raise ValueError(f"No parser available for bank: {bank_key}")

    print(f"Parsing {pdf_path} with {bank_key.upper()} parser...")
    data = parser(pdf_path)

    if dry_run:
        print_dry_run_preview(pdf_path, data)
        return data

    from excel_writer import append_row, update_charts

    print(f"Writing to {master_file}...")
    append_row(master_file, data)
    update_charts(master_file)
    print("Done.")
    return data


def process_many_pdfs(pdf_paths: list[str], master_file: str, dry_run: bool = False) -> bool:
    """Process a list of PDFs and print a concise success/failure summary."""
    success = 0
    failed = 0

    for pdf_path in pdf_paths:
        try:
            process_pdf(pdf_path, master_file, dry_run=dry_run)
            success += 1
        except Exception as exc:
            failed += 1
            print(f"Error processing {pdf_path}: {exc}")

    print(f"\nRun summary: {success} succeeded, {failed} failed.")
    return failed == 0


def prompt_with_default(prompt: str, default: str | None = None, required: bool = False) -> str:
    """Prompt user for input with optional default and required validation."""
    while True:
        suffix = f" [{default}]" if default else ""
        value = input(f"{prompt}{suffix}: ").strip()
        if value:
            return value
        if default:
            return default
        if not required:
            return ""
        print("This value is required.")


def prompt_yes_no(prompt: str, default: bool = False) -> bool:
    """Prompt for yes/no input."""
    hint = "Y/n" if default else "y/N"
    while True:
        response = input(f"{prompt} ({hint}): ").strip().lower()
        if not response:
            return default
        if response in {"y", "yes"}:
            return True
        if response in {"n", "no"}:
            return False
        print("Please answer with y or n.")


def prompt_int(prompt: str, default: int, minimum: int = 1) -> int:
    """Prompt for an integer with minimum bound."""
    while True:
        raw = input(f"{prompt} [{default}]: ").strip()
        if not raw:
            return default
        try:
            value = int(raw)
        except ValueError:
            print("Please enter a valid integer.")
            continue
        if value < minimum:
            print(f"Please enter a value >= {minimum}.")
            continue
        return value


def interactive_mode(
    default_master: str,
    default_email: str,
    default_days: int,
    default_server: str,
    default_sender: str | None,
) -> None:
    """Run interactive menu mode for no-argument usage."""
    print("Interactive mode")

    while True:
        print("\nChoose an action:")
        print("1) Process one PDF")
        print("2) Process all PDFs in a directory")
        print("3) Fetch PDFs from Outlook and process")
        print("4) Run preflight checks")
        print("5) Exit")

        choice = input("Selection [1-5]: ").strip().lower()
        if choice in {"5", "q", "quit", "exit"}:
            return

        if choice == "1":
            pdf_path = prompt_with_default("PDF path", required=True)
            dry_run = prompt_yes_no("Dry run (preview only)", default=True)
            master_file = prompt_with_default("Master workbook path", default_master, required=not dry_run)
            if not run_preflight(
                master_file=master_file,
                require_workbook=not dry_run,
                check_fetch=False,
                verbose=True,
            ):
                if not prompt_yes_no("Preflight failed. Continue anyway", default=False):
                    continue
            process_many_pdfs([pdf_path], master_file, dry_run=dry_run)
            continue

        if choice == "2":
            pdf_dir = prompt_with_default("PDF directory", ".", required=True)
            dry_run = prompt_yes_no("Dry run (preview only)", default=True)
            master_file = prompt_with_default("Master workbook path", default_master, required=not dry_run)
            if not os.path.isdir(pdf_dir):
                print(f"Directory not found: {pdf_dir}")
                continue
            pdf_files = sorted(
                os.path.join(pdf_dir, name)
                for name in os.listdir(pdf_dir)
                if name.lower().endswith(".pdf")
            )
            if not pdf_files:
                print(f"No PDF files found in {pdf_dir}")
                continue
            if not run_preflight(
                master_file=master_file,
                require_workbook=not dry_run,
                check_fetch=False,
                verbose=True,
            ):
                if not prompt_yes_no("Preflight failed. Continue anyway", default=False):
                    continue
            process_many_pdfs(pdf_files, master_file, dry_run=dry_run)
            continue

        if choice == "3":
            from email_fetcher import connect_outlook, fetch_bcecn_pdfs

            email = prompt_with_default("Outlook email", default_email, required=True)
            server = prompt_with_default("Exchange server", default_server, required=True)
            days = prompt_int("Days back to search", default_days, minimum=1)
            sender = prompt_with_default("Sender filter (optional)", default_sender or "")
            dry_run = prompt_yes_no("Dry run (preview only)", default=True)
            master_file = prompt_with_default("Master workbook path", default_master, required=not dry_run)

            if not run_preflight(
                master_file=master_file,
                require_workbook=not dry_run,
                check_fetch=True,
                email=email,
                days=days,
                server=server,
                sender=sender,
                verbose=True,
            ):
                if not prompt_yes_no("Preflight failed. Continue anyway", default=False):
                    continue

            try:
                print(f"Connecting to Outlook as {email}...")
                account = connect_outlook(email, server=server)
                pdfs = fetch_bcecn_pdfs(account, sender_filter=sender or None, days_back=days)
            except Exception as exc:
                print(f"Fetch failed: {exc}")
                continue

            if not pdfs:
                print("No PDFs downloaded.")
                continue

            process_many_pdfs(pdfs, master_file, dry_run=dry_run)
            continue

        if choice == "4":
            check_fetch = prompt_yes_no("Include Outlook fetch checks", default=False)
            dry_run = prompt_yes_no("Assume dry-run mode (skip workbook validation)", default=False)
            master_file = prompt_with_default("Master workbook path", default_master, required=not dry_run)

            email = default_email if check_fetch else None
            days = default_days if check_fetch else None
            server = default_server if check_fetch else None
            sender = default_sender if check_fetch else None

            run_preflight(
                master_file=master_file,
                require_workbook=not dry_run,
                check_fetch=check_fetch,
                email=email,
                days=days,
                server=server,
                sender=sender,
                verbose=True,
            )
            continue

        print("Invalid selection. Please choose 1-5.")


def main():
    load_env_file(".env")

    default_master = os.environ.get("MASTER_FILE", MASTER_FILE)
    default_email = os.environ.get("OUTLOOK_EMAIL", "etienne.beaumier@bell.ca")
    default_days = env_int("OUTLOOK_DAYS", 7)
    default_server = os.environ.get("OUTLOOK_SERVER", "outlook.office365.com")
    default_sender = os.environ.get("BCECN_SENDER")

    parser = argparse.ArgumentParser(description="BCECN Yield Spread Automater")
    parser.add_argument("--pdf", help="Path to a single PDF file to process")
    parser.add_argument("--dir", help="Directory containing PDF files to process")
    parser.add_argument("--master", default=default_master, help="Path to Master File.xlsx")
    parser.add_argument("--fetch", action="store_true", help="Fetch PDFs from Outlook")
    parser.add_argument("--email", default=default_email, help="Outlook email")
    parser.add_argument("--days", type=int, default=default_days, help="Days back to search emails")
    parser.add_argument(
        "--server",
        default=default_server,
        help="Outlook Exchange server",
    )
    parser.add_argument(
        "--sender",
        default=default_sender,
        help="Optional sender email filter for BCECN messages",
    )
    parser.add_argument("--dry-run", action="store_true", help="Parse and preview data without writing to Excel")
    parser.add_argument("--check", action="store_true", help="Run preflight checks before processing")
    parser.add_argument("--interactive", action="store_true", help="Launch interactive mode menu")
    args = parser.parse_args()

    operation_requested = bool(args.fetch or args.pdf or args.dir)

    if args.interactive or (not operation_requested and not args.check):
        interactive_mode(
            default_master=default_master,
            default_email=default_email,
            default_days=default_days,
            default_server=default_server,
            default_sender=default_sender,
        )
        return

    if args.check:
        preflight_ok = run_preflight(
            master_file=args.master,
            require_workbook=not args.dry_run,
            check_fetch=args.fetch,
            email=args.email if args.fetch else None,
            days=args.days if args.fetch else None,
            server=args.server if args.fetch else None,
            sender=args.sender if args.fetch else None,
            verbose=True,
        )
        if not preflight_ok:
            sys.exit(1)
        if not operation_requested:
            return

    if operation_requested and not args.check:
        preflight_ok = run_preflight(
            master_file=args.master,
            require_workbook=not args.dry_run,
            check_fetch=args.fetch,
            email=args.email if args.fetch else None,
            days=args.days if args.fetch else None,
            server=args.server if args.fetch else None,
            sender=args.sender if args.fetch else None,
            verbose=False,
        )
        if not preflight_ok:
            print("Run with --check for full diagnostics.")
            sys.exit(1)

    if args.fetch:
        from email_fetcher import connect_outlook, fetch_bcecn_pdfs

        print(f"Connecting to Outlook as {args.email}...")
        account = connect_outlook(args.email, server=args.server)
        pdfs = fetch_bcecn_pdfs(account, sender_filter=args.sender, days_back=args.days)
        if not pdfs:
            return
        if not process_many_pdfs(pdfs, args.master, dry_run=args.dry_run):
            sys.exit(1)

    elif args.pdf:
        if not process_many_pdfs([args.pdf], args.master, dry_run=args.dry_run):
            sys.exit(1)

    elif args.dir:
        pdf_dir = args.dir
        if not os.path.isdir(pdf_dir):
            print(f"Directory not found: {pdf_dir}")
            sys.exit(1)
        pdf_files = sorted(
            os.path.join(pdf_dir, f)
            for f in os.listdir(pdf_dir)
            if f.lower().endswith(".pdf")
        )
        if not pdf_files:
            print(f"No PDF files found in {pdf_dir}")
            sys.exit(1)
        if not process_many_pdfs(pdf_files, args.master, dry_run=args.dry_run):
            sys.exit(1)

    else:
        interactive_mode(
            default_master=default_master,
            default_email=default_email,
            default_days=default_days,
            default_server=default_server,
            default_sender=default_sender,
        )


if __name__ == "__main__":
    main()
