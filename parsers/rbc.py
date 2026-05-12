"""Parser for RBC Capital Markets Bell Canada (BCECN) indicative new issue PDFs."""

import re
from datetime import datetime


def parse_rbc_pdf(pdf_path: str) -> dict:
    """Parse an RBC Capital Markets BCECN PDF and return structured data.

    Returns a dict with keys:
        date, bank,
        cad_spread_3y..30y, cad_yield_3y..30y,
        usd_spread_3y..30y, usd_yield_3y..30y,
        cad_nc5_spread, cad_nc5_coupon, cad_nc10_spread, cad_nc10_coupon,
        usd_nc5_spread, usd_nc5_coupon, usd_nc10_spread, usd_nc10_coupon,
    """
    import pdfplumber

    with pdfplumber.open(pdf_path) as pdf:
        text = pdf.pages[0].extract_text()

    lines = text.split("\n")
    date = _extract_date(lines)
    cad_lines, usd_lines = _split_cad_usd(lines)

    cad_spreads = _find_credit_spread(cad_lines)
    cad_yields = _find_reoffer_yield(cad_lines)
    usd_spreads = _find_credit_spread(usd_lines)
    usd_yields = _find_reoffer_yield(usd_lines)

    result = {"date": date, "bank": "RBC"}

    tenors = ["3y", "5y", "7y", "10y", "30y"]
    for i, tenor in enumerate(tenors):
        result[f"cad_spread_{tenor}"] = cad_spreads[i] if i < len(cad_spreads) else None
        result[f"cad_yield_{tenor}"] = cad_yields[i] if i < len(cad_yields) else None
        result[f"usd_spread_{tenor}"] = usd_spreads[i] if i < len(usd_spreads) else None
        result[f"usd_yield_{tenor}"] = usd_yields[i] if i < len(usd_yields) else None

    # Indices 5 and 6 are NC5 and NC10
    if len(cad_spreads) > 5:
        result["cad_nc5_spread"] = cad_spreads[5]
        result["cad_nc5_coupon"] = cad_yields[5] if len(cad_yields) > 5 else None
    if len(cad_spreads) > 6:
        result["cad_nc10_spread"] = cad_spreads[6]
        result["cad_nc10_coupon"] = cad_yields[6] if len(cad_yields) > 6 else None
    if len(usd_spreads) > 5:
        result["usd_nc5_spread"] = usd_spreads[5]
        result["usd_nc5_coupon"] = usd_yields[5] if len(usd_yields) > 5 else None
    if len(usd_spreads) > 6:
        result["usd_nc10_spread"] = usd_spreads[6]
        result["usd_nc10_coupon"] = usd_yields[6] if len(usd_yields) > 6 else None

    return result


def _extract_date(lines: list[str]) -> datetime:
    for line in lines:
        match = re.search(r"as at:\s*(\w+ \d{1,2},\s*\d{4})", line, re.IGNORECASE)
        if match:
            return datetime.strptime(match.group(1).strip(), "%B %d, %Y")
    raise ValueError("Could not find date in RBC PDF")


def _split_cad_usd(lines: list[str]) -> tuple[list[str], list[str]]:
    """Split page lines into CAD and USD sections using the column header rows."""
    # Find lines containing the column headers ("3-Year" and "5-Year")
    header_indices = [i for i, l in enumerate(lines) if "3-Year" in l and "5-Year" in l]
    if len(header_indices) < 2:
        raise ValueError("Could not locate CAD/USD section headers in RBC PDF")

    cad_start = header_indices[0] + 1
    usd_start = header_indices[1] + 1

    # EUR section starts at the third header (if present) or at "EUR Midswap"
    if len(header_indices) >= 3:
        usd_end = header_indices[2]
    else:
        usd_end = next(
            (i for i in range(usd_start, len(lines)) if "EUR Midswap" in lines[i]),
            len(lines),
        )

    return lines[cad_start : header_indices[1]], lines[usd_start:usd_end]


def _find_credit_spread(lines: list[str]) -> list[float]:
    for line in lines:
        if "Credit Spread (bps)" in line:
            after_label = line.split("Credit Spread (bps)", 1)[-1]
            return [float(x) for x in re.findall(r"\d+(?:\.\d+)?", after_label)]
    return []


def _find_reoffer_yield(lines: list[str]) -> list[float]:
    for line in lines:
        if line.strip().startswith("Re-Offer Yield"):
            return [float(x) / 100 for x in re.findall(r"(\d+\.\d+)%", line)]
    return []
