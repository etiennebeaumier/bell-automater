"""Background threading workers for long-running operations."""

import os
import re
import tempfile
import threading
import webbrowser

TENORS = ["3y", "5y", "7y", "10y", "30y"]
HYBRIDS = ["nc5", "nc10"]


def _format_dry_run(data: dict) -> str:
    """Format parsed data as a compact text preview for the results log."""
    lines = []
    date_str = data["date"].strftime("%Y-%m-%d")
    lines.append(f"{data['bank']} - {date_str} - DRY RUN PREVIEW")
    for ccy in ("CAD", "USD"):
        c = ccy.lower()
        parts = []
        for tenor in TENORS:
            spread = data.get(f"{c}_spread_{tenor}")
            yld = data.get(f"{c}_yield_{tenor}")
            s = f"{spread:.1f}" if spread is not None else "-"
            y = f"{yld:.3%}" if yld is not None else "-"
            parts.append(f"{tenor.upper()}: {s}bps / {y}")
        lines.append(f"  {ccy}: {' | '.join(parts)}")
        hybrid_parts = []
        for h in HYBRIDS:
            spread = data.get(f"{c}_{h}_spread")
            coupon = data.get(f"{c}_{h}_coupon")
            if spread is not None or coupon is not None:
                s = f"{spread:.1f}" if spread is not None else "-"
                cp = f"{coupon:.3%}" if coupon is not None else "-"
                hybrid_parts.append(f"{h.upper()}: {s}bps / {cp}")
        if hybrid_parts:
            lines.append(f"  {ccy} Hybrid: {' | '.join(hybrid_parts)}")
    return "\n".join(lines)


def _open_device_code_url(message: str, app=None):
    """Extract the device code, copy it to clipboard, and open the login URL."""
    code_match = re.search(r"enter the code\s+(\S+)", message, re.IGNORECASE)
    url_match = re.search(r"(https://\S+)", message)
    if code_match and app:
        code = code_match.group(1)
        app.after(0, lambda: (_clipboard_copy(app, code)))
    if url_match:
        webbrowser.open(url_match.group(1))


def _clipboard_copy(app, text):
    app.clipboard_clear()
    app.clipboard_append(text)


class PdfProcessWorker(threading.Thread):
    """Process a list of PDF files in a background thread."""

    def __init__(self, app, pdf_paths, master_file, dry_run, on_progress=None, on_result=None, on_complete=None, on_error=None):
        super().__init__(daemon=True)
        self.app = app
        self.pdf_paths = pdf_paths
        self.master_file = master_file
        self.dry_run = dry_run
        self.on_progress = on_progress
        self.on_result = on_result
        self.on_complete = on_complete
        self.on_error = on_error

    def run(self):
        from main import detect_bank, BANK_PARSERS
        from excel_writer import append_row, update_charts

        total = len(self.pdf_paths)
        success = 0
        failed = 0

        for idx, pdf_path in enumerate(self.pdf_paths):
            try:
                bank_key = detect_bank(pdf_path)
                parser = BANK_PARSERS[bank_key]
                data = parser(pdf_path)
                date_str = data["date"].strftime("%Y-%m-%d")

                if self.dry_run:
                    msg = _format_dry_run(data)
                else:
                    append_row(self.master_file, data)
                    update_charts(self.master_file)
                    msg = f"{data['bank']} - {date_str} - Written to workbook"

                success += 1
                if self.on_result:
                    self.app.after(0, self.on_result, msg, True)
            except Exception as exc:
                failed += 1
                msg = f"{os.path.basename(pdf_path)} - Error: {exc}"
                if self.on_result:
                    self.app.after(0, self.on_result, msg, False)

            if self.on_progress:
                self.app.after(0, self.on_progress, (idx + 1) / total)

        summary = f"Done: {success} succeeded, {failed} failed"
        if self.on_complete:
            self.app.after(0, self.on_complete, summary)


class OutlookFetchWorker(threading.Thread):
    """Fetch PDFs from Outlook and process them in a background thread."""

    def __init__(self, app, email, server, days, sender, master_file, dry_run,
                 on_auth_status=None, on_progress=None, on_result=None, on_complete=None, on_error=None):
        super().__init__(daemon=True)
        self.app = app
        self.email = email
        self.server = server
        self.days = days
        self.sender = sender
        self.master_file = master_file
        self.dry_run = dry_run
        self.on_auth_status = on_auth_status
        self.on_progress = on_progress
        self.on_result = on_result
        self.on_complete = on_complete
        self.on_error = on_error

    def run(self):
        try:
            from email_fetcher import connect_outlook, fetch_bcecn_pdfs
            from main import detect_bank, BANK_PARSERS
            from excel_writer import append_row, update_charts

            def status_callback(msg):
                _open_device_code_url(msg, app=self.app)
                if self.on_auth_status:
                    self.app.after(0, self.on_auth_status, msg)

            account = connect_outlook(self.email, server=self.server, status_callback=status_callback)

            if self.on_auth_status:
                self.app.after(0, self.on_auth_status, "Searching for BCECN emails...")

            pdfs = fetch_bcecn_pdfs(account, sender_filter=self.sender or None, days_back=self.days)

            if not pdfs:
                if self.on_complete:
                    self.app.after(0, self.on_complete, "No BCECN PDFs found in the selected range.")
                return

            total = len(pdfs)
            success = 0
            failed = 0

            for idx, pdf_path in enumerate(pdfs):
                try:
                    bank_key = detect_bank(pdf_path)
                    parser = BANK_PARSERS[bank_key]
                    data = parser(pdf_path)
                    date_str = data["date"].strftime("%Y-%m-%d")

                    if self.dry_run:
                        msg = _format_dry_run(data)
                    else:
                        append_row(self.master_file, data)
                        update_charts(self.master_file)
                        msg = f"{data['bank']} - {date_str} - Written to workbook"

                    success += 1
                    if self.on_result:
                        self.app.after(0, self.on_result, msg, True)
                except Exception as exc:
                    failed += 1
                    msg = f"{os.path.basename(pdf_path)} - Error: {exc}"
                    if self.on_result:
                        self.app.after(0, self.on_result, msg, False)

                if self.on_progress:
                    self.app.after(0, self.on_progress, (idx + 1) / total)

            summary = f"Done: {success} succeeded, {failed} failed (from {total} PDFs fetched)"
            if self.on_complete:
                self.app.after(0, self.on_complete, summary)

        except Exception as exc:
            if self.on_error:
                self.app.after(0, self.on_error, str(exc))
