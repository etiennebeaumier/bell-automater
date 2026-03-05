"""Parser for Scotiabank Bell Canada indicative new issue PDFs."""

import re
from datetime import datetime


def parse_scotiabank_pdf(pdf_path: str) -> dict:
    import pdfplumber

    with pdfplumber.open(pdf_path) as pdf:
        text = pdf.pages[0].extract_text()

    lines = text.split("\n")
    date = _extract_date(lines)

    # Split into CAD and USD sections
    cad_lines, usd_lines = _split_sections(lines)

    # CAD section: "Indicative Spread" and "Indicative Re-offer Yield"
    cad_spreads = _extract_spread_row(cad_lines)
    cad_yields = _extract_yield_row(cad_lines)

    # USD section: same row names
    usd_spreads = _extract_spread_row(usd_lines)
    usd_yields = _extract_yield_row(usd_lines)

    result = {"date": date, "bank": "Scotiabank"}

    # Scotiabank CAD has: floating 3Y, then fixed 3Y, 5Y, 7Y, 10Y, 30Y, NC5, NC10
    # The first value (+75) is the floating rate 3Y — skip it.
    # Fixed rate values start at index 1.
    tenors = ["3y", "5y", "7y", "10y", "30y"]
    for i, tenor in enumerate(tenors):
        result[f"cad_spread_{tenor}"] = cad_spreads[i + 1]  # skip floating 3Y
        result[f"usd_spread_{tenor}"] = usd_spreads[i + 1]  # skip floating 3Y

    # Yields: no floating rate 3Y yield, so indices 0-4 are 3Y-30Y
    for i, tenor in enumerate(tenors):
        result[f"cad_yield_{tenor}"] = cad_yields[i]
        result[f"usd_yield_{tenor}"] = usd_yields[i]

    # NC5 and NC10 — spreads at indices 6 and 7 (after floating 3Y skip)
    if len(cad_spreads) > 7:
        result["cad_nc5_spread"] = cad_spreads[6]
        result["cad_nc10_spread"] = cad_spreads[7]
    if len(cad_yields) > 6:
        result["cad_nc5_coupon"] = cad_yields[5]
        result["cad_nc10_coupon"] = cad_yields[6]

    if len(usd_spreads) > 7:
        result["usd_nc5_spread"] = usd_spreads[6]
        result["usd_nc10_spread"] = usd_spreads[7]
    if len(usd_yields) > 6:
        result["usd_nc5_coupon"] = usd_yields[5]
        result["usd_nc10_coupon"] = usd_yields[6]

    return result


def _extract_date(lines: list[str]) -> datetime:
    for line in lines:
        match = re.search(r"Pricing as of (\w+ \d{1,2},?\s*\d{4})", line)
        if match:
            date_str = match.group(1).replace(",", "").strip()
            return datetime.strptime(date_str, "%B %d %Y")
    raise ValueError("Could not find date in Scotiabank PDF")


def _split_sections(lines: list[str]) -> tuple[list[str], list[str]]:
    """Split into CAD (before 'US$ NEW ISSUE') and USD (after) sections."""
    split_idx = None
    for i, line in enumerate(lines):
        if "US$" in line and "NEW ISSUE" in line:
            split_idx = i
            break
    if split_idx is None:
        raise ValueError("Could not find US$ section in Scotiabank PDF")
    return lines[:split_idx], lines[split_idx:]


def _extract_spread_row(lines: list[str]) -> list[float]:
    """Extract 'Indicative Spread' row values (numbers with + prefix)."""
    for line in lines:
        if line.strip().startswith("Indicative Spread"):
            values = re.findall(r"[+](\d+(?:\.\d+)?)", line)
            return [float(v) for v in values]
    raise ValueError("Could not find 'Indicative Spread' row")


def _extract_yield_row(lines: list[str]) -> list[float]:
    """Extract 'Indicative Re-offer Yield' row percentages as decimals."""
    for line in lines:
        if "Re-offer Yield" in line or "Re-Offer Yield" in line:
            if "Swapped" not in line:
                pcts = re.findall(r"(\d+\.\d+)%", line)
                return [float(p) / 100 for p in pcts]
    raise ValueError("Could not find 'Re-offer Yield' row")
