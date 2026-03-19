"""Parser for Desjardins Capital Markets Bell Canada indicative new issue PDFs.

Desjardins provides CAD-only pricing (no USD rates).
"""

import re
from datetime import datetime


def parse_desjardins_pdf(pdf_path: str) -> dict:
    import pdfplumber

    with pdfplumber.open(pdf_path) as pdf:
        text = pdf.pages[0].extract_text()

    lines = text.split("\n")
    date = _extract_date(lines)

    senior_lines, hybrid_lines = _split_sections(lines)

    cad_spreads = _extract_bps_row(senior_lines, "New Issue Spread")
    cad_yields = _extract_pct_row(senior_lines, "New Issue Yield")

    hybrid_spreads = _extract_bps_row(hybrid_lines, "New Issue Spread")
    hybrid_yields = _extract_pct_row(hybrid_lines, "New Issue Yield")

    result = {"date": date, "bank": "Desjardins"}

    tenors = ["3y", "5y", "7y", "10y", "30y"]
    for i, tenor in enumerate(tenors):
        result[f"cad_spread_{tenor}"] = cad_spreads[i]
        result[f"cad_yield_{tenor}"] = cad_yields[i]

    if len(hybrid_spreads) >= 2:
        result["cad_nc5_spread"] = hybrid_spreads[0]
        result["cad_nc10_spread"] = hybrid_spreads[1]
    if len(hybrid_yields) >= 2:
        result["cad_nc5_coupon"] = hybrid_yields[0]
        result["cad_nc10_coupon"] = hybrid_yields[1]

    return result


def _extract_date(lines: list[str]) -> datetime:
    for line in lines[:10]:
        match = re.match(r"^(\w+ \d{1,2},?\s*\d{4})$", line.strip())
        if match:
            date_str = match.group(1).replace(",", "").strip()
            return datetime.strptime(date_str, "%B %d %Y")
    raise ValueError("Could not find date in Desjardins PDF")


def _split_sections(lines: list[str]) -> tuple[list[str], list[str]]:
    """Split into Senior Unsecured and Hybrid sections, excluding Peers."""
    senior_start = None
    hybrid_start = None
    peers_start = None

    for i, line in enumerate(lines):
        stripped = line.strip()
        if "Senior Unsecured" in stripped and senior_start is None:
            senior_start = i
        elif stripped.startswith("Hybrid") and hybrid_start is None:
            hybrid_start = i
        elif "Peers" in stripped and peers_start is None:
            peers_start = i
            break

    if senior_start is None or hybrid_start is None:
        raise ValueError("Could not find Senior/Hybrid sections in Desjardins PDF")

    end = peers_start if peers_start is not None else len(lines)
    return lines[senior_start:hybrid_start], lines[hybrid_start:end]


def _extract_bps_row(lines: list[str], label: str) -> list[float]:
    """Extract a row with 'bps' values, handling optional ~ prefix."""
    for line in lines:
        if label in line:
            values = re.findall(r"~?(\d+(?:\.\d+)?)\s*bps", line)
            if values:
                return [float(v) for v in values]
    raise ValueError(f"Could not find '{label}' row in Desjardins PDF")


def _extract_pct_row(lines: list[str], label: str) -> list[float]:
    """Extract a row with percentage values, returned as decimals."""
    for line in lines:
        if label in line:
            pcts = re.findall(r"(\d+\.\d+)\s*%", line)
            if pcts:
                return [float(p) / 100 for p in pcts]
    raise ValueError(f"Could not find '{label}' row in Desjardins PDF")
