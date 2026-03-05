"""Parser for CIBC Bell Canada indicative new issue PDFs."""

import re
from datetime import datetime


def parse_cibc_pdf(pdf_path: str) -> dict:
    import pdfplumber

    with pdfplumber.open(pdf_path) as pdf:
        text = pdf.pages[0].extract_text()

    lines = text.split("\n")
    date = _extract_date(lines)

    # Split into sections: C$ Pricing, US$ Pricing, C$ Hybrid, US$ Hybrid
    cad_lines, usd_lines, cad_hybrid_lines, usd_hybrid_lines = _split_sections(lines)

    # CAD: "Spread to GoC Curve" and "Re-Offer Yield"
    cad_spreads = _extract_bps_row(cad_lines, "Spread to GoC Curve")
    cad_yields = _extract_pct_row(cad_lines, "Re-Offer Yield")

    # USD: "Spread to UST Benchmark" and "Re-Offer Yield"
    usd_spreads = _extract_bps_row(usd_lines, "Spread to UST Benchmark")
    usd_yields = _extract_pct_row(usd_lines, "Re-Offer Yield")

    result = {"date": date, "bank": "CIBC"}

    tenors = ["3y", "5y", "7y", "10y", "30y"]
    for i, tenor in enumerate(tenors):
        result[f"cad_spread_{tenor}"] = cad_spreads[i]
        result[f"cad_yield_{tenor}"] = cad_yields[i]
        result[f"usd_spread_{tenor}"] = usd_spreads[i]
        result[f"usd_yield_{tenor}"] = usd_yields[i]

    # C$ Hybrid: "Hybrid Spread" and "Hybrid Coupon"
    if cad_hybrid_lines:
        cad_h_spreads = _extract_bps_row(cad_hybrid_lines, "Hybrid Spread")
        cad_h_coupons = _extract_pct_row(cad_hybrid_lines, "Hybrid Coupon")
        if len(cad_h_spreads) >= 2:
            result["cad_nc5_spread"] = cad_h_spreads[0]
            result["cad_nc10_spread"] = cad_h_spreads[1]
        if len(cad_h_coupons) >= 2:
            result["cad_nc5_coupon"] = cad_h_coupons[0]
            result["cad_nc10_coupon"] = cad_h_coupons[1]

    # US$ Hybrid: "Hybrid Spread" and "Hybrid Coupon"
    if usd_hybrid_lines:
        usd_h_spreads = _extract_bps_row(usd_hybrid_lines, "Hybrid Spread")
        usd_h_coupons = _extract_pct_row(usd_hybrid_lines, "Hybrid Coupon")
        if len(usd_h_spreads) >= 2:
            result["usd_nc5_spread"] = usd_h_spreads[0]
            result["usd_nc10_spread"] = usd_h_spreads[1]
        if len(usd_h_coupons) >= 2:
            result["usd_nc5_coupon"] = usd_h_coupons[0]
            result["usd_nc10_coupon"] = usd_h_coupons[1]

    return result


def _extract_date(lines: list[str]) -> datetime:
    for line in lines:
        match = re.search(r"(?:SPREADS|Spreads)\s*-\s*(\w+ \d{1,2},?\s*\d{4})", line)
        if match:
            date_str = match.group(1).replace(",", "").strip()
            return datetime.strptime(date_str, "%B %d %Y")
    raise ValueError("Could not find date in CIBC PDF")


def _split_sections(lines: list[str]) -> tuple[list[str], list[str], list[str], list[str]]:
    """Split into C$ Pricing, US$ Pricing, C$ Hybrid, US$ Hybrid sections."""
    usd_idx = None
    cad_hybrid_idx = None
    usd_hybrid_idx = None

    for i, line in enumerate(lines):
        stripped = line.strip()
        if "US$ Pricing" in stripped:
            if usd_idx is None:
                usd_idx = i
        elif "C$ Hybrid" in stripped:
            if cad_hybrid_idx is None:
                cad_hybrid_idx = i
        elif "US$ Hybrid" in stripped:
            if usd_hybrid_idx is None:
                usd_hybrid_idx = i

    cad_lines = lines[:usd_idx] if usd_idx else lines
    usd_end = cad_hybrid_idx or usd_hybrid_idx or len(lines)
    usd_lines = lines[usd_idx:usd_end] if usd_idx else []
    cad_hybrid_lines = lines[cad_hybrid_idx:usd_hybrid_idx] if cad_hybrid_idx else []
    usd_hybrid_lines = lines[usd_hybrid_idx:] if usd_hybrid_idx else []

    return cad_lines, usd_lines, cad_hybrid_lines, usd_hybrid_lines


def _extract_bps_row(lines: list[str], label: str) -> list[float]:
    """Extract a row with 'bps' values (e.g., '62 bps 77 bps ...')."""
    for line in lines:
        if label in line:
            values = re.findall(r"(\d+(?:\.\d+)?)\s*bps", line)
            if values:
                return [float(v) for v in values]
    raise ValueError(f"Could not find '{label}' row")


def _extract_pct_row(lines: list[str], label: str) -> list[float]:
    """Extract a row with percentage values, returned as decimals."""
    for line in lines:
        if label in line:
            if "Swapped" in line:
                continue
            pcts = re.findall(r"(\d+\.\d+)%", line)
            if pcts:
                return [float(p) / 100 for p in pcts]
    raise ValueError(f"Could not find '{label}' row")
