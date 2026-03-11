#!/usr/bin/env python3
"""BCECN Pricing Tool desktop entry point.

Startup behavior:
- If workbook + preferred PDF source folder are valid and contain parseable PDFs,
  show a compact quick-run prompt for chart year range.
- Otherwise, launch the full GUI.
"""

import sys
import os
from dataclasses import dataclass
from datetime import datetime

# Ensure the project root is on sys.path (needed when running from PyInstaller bundle)
if getattr(sys, "frozen", False):
    os.chdir(os.path.dirname(sys.executable))
    sys.path.insert(0, os.path.dirname(sys.executable))
else:
    sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from config import AppConfig, get_default_pdf_dir
from ui.app_window import AppWindow

REQUIRED_SHEETS = {"Pricing", "Summary Charts"}


@dataclass
class QuickRunContext:
    """Validated inputs required for headless quick processing."""

    master_file: str
    pdf_dir: str
    pdf_paths: list[str]


@dataclass
class QuickRunSelection:
    """User selection returned by the quick-run popup."""

    action: str
    avg_start_year: int | None = None
    avg_end_year: int | None = None


def _launch_gui(config: AppConfig) -> None:
    """Launch the full desktop interface."""
    app = AppWindow(config)
    app.mainloop()


def _is_workbook_ready(path: str) -> bool:
    """Return True when workbook exists and includes required sheets."""
    if not path or not os.path.isfile(path):
        return False
    try:
        from openpyxl import load_workbook

        wb = load_workbook(path, read_only=True)
        sheets = set(wb.sheetnames)
        wb.close()
        return REQUIRED_SHEETS.issubset(sheets)
    except Exception:
        return False


def _collect_parseable_pdfs(pdf_dir: str) -> list[str]:
    """Return PDFs in `pdf_dir` that pass bank detection.

    Bank detection is used as a lightweight parseability pre-check for quick-run.
    """
    if not pdf_dir or not os.path.isdir(pdf_dir):
        return []

    pdf_files = sorted(
        os.path.join(pdf_dir, name)
        for name in os.listdir(pdf_dir)
        if name.lower().endswith(".pdf") and os.path.isfile(os.path.join(pdf_dir, name))
    )
    if not pdf_files:
        return []

    from main import detect_bank

    parseable: list[str] = []
    for pdf_path in pdf_files:
        try:
            detect_bank(pdf_path)
            parseable.append(pdf_path)
        except Exception:
            continue
    return parseable


def _parse_year(raw: object, fallback: int) -> int:
    """Parse an integer year, falling back when invalid."""
    try:
        return int(raw)
    except (TypeError, ValueError):
        return fallback


def _load_available_years(master_file: str) -> list[int]:
    """Load available years from Pricing column A, always including current year."""
    current_year = datetime.now().year
    if not master_file or not os.path.isfile(master_file):
        return [current_year]

    years = {current_year}
    try:
        from openpyxl import load_workbook

        wb = load_workbook(master_file, read_only=True, data_only=True)
        if "Pricing" in wb.sheetnames:
            ws = wb["Pricing"]
            for row in range(2, ws.max_row + 1):
                date_val = ws.cell(row=row, column=1).value
                if hasattr(date_val, "year"):
                    years.add(int(date_val.year))
                    continue
                if isinstance(date_val, str):
                    raw = date_val.strip()
                    if not raw:
                        continue
                    parsed = None
                    for fmt in ("%Y-%m-%d", "%Y/%m/%d", "%m/%d/%Y", "%d/%m/%Y"):
                        try:
                            parsed = datetime.strptime(raw, fmt)
                            break
                        except ValueError:
                            continue
                    if parsed is None:
                        try:
                            parsed = datetime.fromisoformat(raw)
                        except ValueError:
                            parsed = None
                    if parsed is not None:
                        years.add(parsed.year)
        wb.close()
    except Exception:
        return [current_year]

    return sorted(years)


def _collect_quick_run_context(config: AppConfig) -> QuickRunContext | None:
    """Build quick-run context, or return None when preconditions fail.

    Preconditions:
    - workbook path is valid and ready
    - preferred PDF source directory exists
    - directory contains at least one parseable PDF
    """
    master_file = (config.get("master_file") or "").strip()
    if not _is_workbook_ready(master_file):
        return None

    preferred_pdf_dir = (config.get("pdf_source_dir") or "").strip() or get_default_pdf_dir()
    parseable_pdfs = _collect_parseable_pdfs(preferred_pdf_dir)
    if not parseable_pdfs:
        return None

    return QuickRunContext(
        master_file=master_file,
        pdf_dir=preferred_pdf_dir,
        pdf_paths=parseable_pdfs,
    )


def _show_quick_run_dialog(
    pdf_count: int,
    pdf_dir: str,
    years: list[int],
    default_start_year: int,
    default_end_year: int,
) -> QuickRunSelection:
    """Show compact startup dialog and return user action + selected year bounds."""
    import tkinter as tk
    from tkinter import messagebox, ttk

    result = QuickRunSelection(action="open_gui")
    root = tk.Tk()
    root.title("Quick PDF Processing")
    root.resizable(False, False)
    try:
        root.attributes("-topmost", True)
    except Exception:
        pass

    frame = ttk.Frame(root, padding=12)
    frame.grid(row=0, column=0, sticky="nsew")
    frame.columnconfigure(1, weight=1)

    info = (
        f"Found {pdf_count} parseable PDF file(s) in:\n"
        f"{pdf_dir}\n\n"
        "Select the year range for average spread charts."
    )
    ttk.Label(frame, text=info, justify="left").grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 10))

    year_values = [str(year) for year in years]
    start_var = tk.StringVar(value=str(default_start_year))
    end_var = tk.StringVar(value=str(default_end_year))

    ttk.Label(frame, text="Start Year").grid(row=1, column=0, sticky="w")
    start_combo = ttk.Combobox(frame, textvariable=start_var, values=year_values, state="readonly", width=10)
    start_combo.grid(row=1, column=1, sticky="ew", pady=(0, 6))

    ttk.Label(frame, text="End Year").grid(row=2, column=0, sticky="w")
    end_combo = ttk.Combobox(frame, textvariable=end_var, values=year_values, state="readonly", width=10)
    end_combo.grid(row=2, column=1, sticky="ew", pady=(0, 10))

    btns = ttk.Frame(frame)
    btns.grid(row=3, column=0, columnspan=2, sticky="e")

    def _on_ok():
        try:
            start_year = int(start_var.get())
            end_year = int(end_var.get())
        except ValueError:
            messagebox.showerror("Invalid Year", "Start Year and End Year must be valid numbers.")
            return
        result.action = "run"
        result.avg_start_year = start_year
        result.avg_end_year = end_year
        root.destroy()

    def _on_open_gui():
        result.action = "open_gui"
        root.destroy()

    ttk.Button(btns, text="Open GUI", command=_on_open_gui).grid(row=0, column=0, padx=(0, 8))
    ttk.Button(btns, text="OK", command=_on_ok).grid(row=0, column=1)

    root.protocol("WM_DELETE_WINDOW", _on_open_gui)
    root.mainloop()
    return result


def main():
    """Run startup flow: quick-run when possible, otherwise launch full GUI."""
    config = AppConfig()
    quick_context = _collect_quick_run_context(config)

    if not quick_context:
        _launch_gui(config)
        return

    years = _load_available_years(quick_context.master_file)
    start_pref = _parse_year(config.get("avg_start_year"), years[0])
    end_pref = _parse_year(config.get("avg_end_year"), years[-1])
    start_year = start_pref if start_pref in years else years[0]
    end_year = end_pref if end_pref in years else years[-1]

    selection = _show_quick_run_dialog(
        pdf_count=len(quick_context.pdf_paths),
        pdf_dir=quick_context.pdf_dir,
        years=years,
        default_start_year=start_year,
        default_end_year=end_year,
    )

    if selection.action != "run":
        _launch_gui(config)
        return

    config["avg_start_year"] = selection.avg_start_year
    config["avg_end_year"] = selection.avg_end_year
    config.save()

    from main import process_many_pdfs

    process_many_pdfs(
        quick_context.pdf_paths,
        quick_context.master_file,
        dry_run=False,
        avg_start_year=selection.avg_start_year,
        avg_end_year=selection.avg_end_year,
    )


if __name__ == "__main__":
    main()
