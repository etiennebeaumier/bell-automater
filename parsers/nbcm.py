"""Parser for NBCM (National Bank Capital Markets) Bell Canada indicative new issue PDFs."""

import re
from datetime import datetime


def parse_nbcm_pdf(pdf_path: str) -> dict:
    import pdfplumber

    with pdfplumber.open(pdf_path) as pdf:
        text = pdf.pages[0].extract_text()

    lines = text.split("\n")
    date = _extract_date(lines)

    cad_lines, usd_lines, hybrid_lines = _split_sections(lines)

    # CAD: 6 values (2Y, 3Y, 5Y, 7Y, 10Y, 30Y) — skip 2Y at index 0
    cad_spreads = _extract_bps_row(cad_lines, "Reoffer Spread")
    cad_yields = _extract_pct_row(cad_lines, "Coupon")

    # USD: 5 values (3Y, 5Y, 7Y, 10Y, 30Y)
    usd_spreads = _extract_bps_row(usd_lines, "Reoffer Spread")
    usd_yields = _extract_pct_row(usd_lines, "Coupon")

    result = {"date": date, "bank": "NBCM"}

    tenors = ["3y", "5y", "7y", "10y", "30y"]

    # CAD has a 2Y column first, so fixed tenors start at index 1
    for i, tenor in enumerate(tenors):
        result[f"cad_spread_{tenor}"] = cad_spreads[i + 1]
        result[f"cad_yield_{tenor}"] = cad_yields[i + 1]

    # USD has no 2Y, indices 0-4 map directly to 3Y-30Y
    for i, tenor in enumerate(tenors):
        result[f"usd_spread_{tenor}"] = usd_spreads[i]
        result[f"usd_yield_{tenor}"] = usd_yields[i]

    # Hybrid: 4 values ordered as USD NC5, USD NC10, CAD NC5, CAD NC10
    hybrid_spreads = _extract_bps_row(hybrid_lines, "Reoffer Spread")
    hybrid_coupons = _extract_pct_row(hybrid_lines, "Coupon")

    if len(hybrid_spreads) >= 4:
        result["usd_nc5_spread"] = hybrid_spreads[0]
        result["usd_nc10_spread"] = hybrid_spreads[1]
        result["cad_nc5_spread"] = hybrid_spreads[2]
        result["cad_nc10_spread"] = hybrid_spreads[3]

    if len(hybrid_coupons) >= 4:
        result["usd_nc5_coupon"] = hybrid_coupons[0]
        result["usd_nc10_coupon"] = hybrid_coupons[1]
        result["cad_nc5_coupon"] = hybrid_coupons[2]
        result["cad_nc10_coupon"] = hybrid_coupons[3]

    return result


def _extract_date(lines: list[str]) -> datetime:
    for line in lines:
        match = re.search(r"(\w+ \d{1,2},?\s*\d{4})", line)
        if match:
            date_str = match.group(1).replace(",", "").strip()
            return datetime.strptime(date_str, "%B %d %Y")
    raise ValueError("Could not find date in NBCM PDF")


def _split_sections(lines: list[str]) -> tuple[list[str], list[str], list[str]]:
    """Split into C$ Pricing, US$ Pricing, and Hybrid Pricing sections."""
    usd_idx = None
    hybrid_idx = None
    secondary_idx = None

    for i, line in enumerate(lines):
        stripped = line.strip()
        if stripped.startswith("US$ Pricing"):
            if usd_idx is None:
                usd_idx = i
        elif stripped.startswith("Hybrid Pricing"):
            if hybrid_idx is None:
                hybrid_idx = i
        elif "Secondary Trading" in stripped:
            if secondary_idx is None:
                secondary_idx = i

    cad_end = usd_idx or hybrid_idx or secondary_idx or len(lines)
    usd_end = hybrid_idx or secondary_idx or len(lines)
    hybrid_end = secondary_idx or len(lines)

    cad_lines = lines[:cad_end]
    usd_lines = lines[usd_idx:usd_end] if usd_idx else []
    hybrid_lines = lines[hybrid_idx:hybrid_end] if hybrid_idx else []

    return cad_lines, usd_lines, hybrid_lines


def _extract_bps_row(lines: list[str], label: str) -> list[float]:
    """Extract a row with 'bps' values (e.g., '65 bps 75 bps ...')."""
    for line in lines:
        if line.strip().startswith(label) and "CORRA" not in line:
            values = re.findall(r"(\d+(?:\.\d+)?)\s*bps", line)
            if values:
                return [float(v) for v in values]
    raise ValueError(f"Could not find '{label}' row in NBCM PDF")


def _extract_pct_row(lines: list[str], label: str) -> list[float]:
    """Extract a row with percentage values, returned as decimals."""
    for line in lines:
        if line.strip().startswith(label) and "Swapped" not in line:
            pcts = re.findall(r"(\d+\.\d+)%", line)
            if pcts:
                return [float(p) / 100 for p in pcts]
    raise ValueError(f"Could not find '{label}' row in NBCM PDF")
