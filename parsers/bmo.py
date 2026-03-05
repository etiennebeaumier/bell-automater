"""Parser for BMO Capital Markets Bell Canada indicative new issue PDFs."""

import re
from datetime import datetime


def parse_bmo_pdf(pdf_path: str) -> dict:
    import pdfplumber

    with pdfplumber.open(pdf_path) as pdf:
        text = pdf.pages[0].extract_text()

    lines = text.split("\n")
    date = _extract_date(lines)

    cad_lines, usd_lines = _split_sections(lines)

    # Both sections have 7 values: 3Y, 5Y, 7Y, 10Y, 30Y, NC5, NC10
    cad_spreads = _extract_bps_row(cad_lines, "New Issue Spread")
    cad_yields = _extract_pct_row(cad_lines, "CAD Coupon")

    usd_spreads = _extract_bps_row(usd_lines, "New Issue Spread")
    usd_yields = _extract_pct_row(usd_lines, "USD Coupon")

    result = {"date": date, "bank": "BMO"}

    tenors = ["3y", "5y", "7y", "10y", "30y"]
    for i, tenor in enumerate(tenors):
        result[f"cad_spread_{tenor}"] = cad_spreads[i]
        result[f"cad_yield_{tenor}"] = cad_yields[i]
        result[f"usd_spread_{tenor}"] = usd_spreads[i]
        result[f"usd_yield_{tenor}"] = usd_yields[i]

    # NC5 and NC10 at indices 5 and 6
    if len(cad_spreads) > 6:
        result["cad_nc5_spread"] = cad_spreads[5]
        result["cad_nc10_spread"] = cad_spreads[6]
    if len(cad_yields) > 6:
        result["cad_nc5_coupon"] = cad_yields[5]
        result["cad_nc10_coupon"] = cad_yields[6]

    if len(usd_spreads) > 6:
        result["usd_nc5_spread"] = usd_spreads[5]
        result["usd_nc10_spread"] = usd_spreads[6]
    if len(usd_yields) > 6:
        result["usd_nc5_coupon"] = usd_yields[5]
        result["usd_nc10_coupon"] = usd_yields[6]

    return result


def _extract_date(lines: list[str]) -> datetime:
    for line in lines:
        match = re.search(r"as (?:at|of)\s+(\w+ \d{1,2},?\s*\d{4})", line, re.IGNORECASE)
        if match:
            date_str = match.group(1).replace(",", "").strip()
            return datetime.strptime(date_str, "%B %d %Y")
    raise ValueError("Could not find date in BMO PDF")


def _split_sections(lines: list[str]) -> tuple[list[str], list[str]]:
    """Split into CAD and USD sections at the second 'Bell Canada' line."""
    bell_indices = [i for i, line in enumerate(lines) if line.strip() == "Bell Canada"]
    if len(bell_indices) < 2:
        raise ValueError("Could not find CAD/USD sections in BMO PDF")
    # CAD section: from first "Bell Canada" to second "Bell Canada"
    # USD section: from second "Bell Canada" to "Disclaimer"
    disclaimer_idx = len(lines)
    for i, line in enumerate(lines):
        if line.strip().startswith("Disclaimer"):
            disclaimer_idx = i
            break
    return lines[bell_indices[0]:bell_indices[1]], lines[bell_indices[1]:disclaimer_idx]


def _extract_bps_row(lines: list[str], label: str) -> list[float]:
    """Extract a row with 'bps' values."""
    for line in lines:
        if label in line:
            values = re.findall(r"(\d+(?:\.\d+)?)\s*bps", line)
            if values:
                return [float(v) for v in values]
    raise ValueError(f"Could not find '{label}' row in BMO PDF")


def _extract_pct_row(lines: list[str], label: str) -> list[float]:
    """Extract a row with percentage values, returned as decimals."""
    for line in lines:
        if label in line:
            pcts = re.findall(r"(\d+\.\d+)%", line)
            if pcts:
                return [float(p) / 100 for p in pcts]
    raise ValueError(f"Could not find '{label}' row in BMO PDF")
