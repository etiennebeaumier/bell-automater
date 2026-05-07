"""Parser for Mizuho Bell Canada indicative new issue PDFs.

Mizuho provides USD-only senior pricing (no CAD rates) and hybrid (NC5/NC10) indications.
"""

import re
from datetime import datetime


def parse_mizuho_pdf(pdf_path: str) -> dict:
    import pdfplumber

    with pdfplumber.open(pdf_path) as pdf:
        pages_text = [page.extract_text() or "" for page in pdf.pages]

    date = _extract_date(pages_text)

    # Senior USD pricing is on the page containing "Indicative New Issue USD Pricing"
    senior_lines = _find_page_lines(pages_text, "Indicative New Issue USD Pricing")
    usd_section = _slice_section(senior_lines, "Indicative New Issue USD Pricing", "Indicative New Issue EUR Pricing")

    usd_spreads = _extract_bp_row(usd_section, "Reoffer Spread")
    usd_yields = _extract_pct_row(usd_section, "Reoffer Yield")

    # Hybrid NC5/NC10 from the coupon step-ups table on the junior sub page
    hybrid_lines = _find_page_lines(pages_text, "coupon step-ups")
    step_up_section = _slice_section(hybrid_lines, "coupon step-ups", "coupon floor")

    nc5_spread, nc10_spread = _extract_hybrid_spreads(step_up_section)
    nc5_yield, nc10_yield = _extract_hybrid_yields(step_up_section)

    result = {"date": date, "bank": "Mizuho"}

    tenors = ["3y", "5y", "7y", "10y", "30y"]
    for i, tenor in enumerate(tenors):
        result[f"usd_spread_{tenor}"] = usd_spreads[i] if i < len(usd_spreads) else None
        result[f"usd_yield_{tenor}"] = usd_yields[i] if i < len(usd_yields) else None

    if nc5_spread is not None:
        result["usd_nc5_spread"] = nc5_spread
    if nc10_spread is not None:
        result["usd_nc10_spread"] = nc10_spread
    if nc5_yield is not None:
        result["usd_nc5_coupon"] = nc5_yield
    if nc10_yield is not None:
        result["usd_nc10_coupon"] = nc10_yield

    return result


def _extract_date(pages_text: list[str]) -> datetime:
    for text in pages_text:
        # "May 6, 2026" on its own line (cover page)
        match = re.search(r"^(\w+ \d{1,2},\s*\d{4})$", text, re.MULTILINE)
        if match:
            try:
                return datetime.strptime(match.group(1).strip(), "%B %d, %Y")
            except ValueError:
                pass
    # Fallback: "As of 5/6/2026" or "As of 05/06/2026"
    for text in pages_text:
        match = re.search(r"As of\s+(\d{1,2}/\d{1,2}/\d{4})", text, re.IGNORECASE)
        if match:
            return datetime.strptime(match.group(1), "%m/%d/%Y")
    raise ValueError("Could not find date in Mizuho PDF")


def _find_page_lines(pages_text: list[str], marker: str) -> list[str]:
    for text in pages_text:
        if marker.lower() in text.lower():
            return text.split("\n")
    return []


def _slice_section(lines: list[str], start_marker: str, end_marker: str) -> list[str]:
    start = None
    for i, line in enumerate(lines):
        if start_marker.lower() in line.lower():
            start = i
            break
    if start is None:
        return lines
    for i in range(start + 1, len(lines)):
        if end_marker.lower() in lines[i].lower():
            return lines[start:i]
    return lines[start:]


def _extract_bp_row(lines: list[str], label: str) -> list[float]:
    for line in lines:
        if label in line:
            values = re.findall(r"\+?(\d+(?:\.\d+)?)\s*bp\b", line)
            if values:
                return [float(v) for v in values]
    raise ValueError(f"Could not find '{label}' row in Mizuho PDF")


def _extract_pct_row(lines: list[str], label: str) -> list[float]:
    for line in lines:
        if label in line and "%" in line:
            pcts = re.findall(r"(\d+\.\d+)%", line)
            if pcts:
                return [float(p) / 100 for p in pcts]
    raise ValueError(f"Could not find '{label}' row in Mizuho PDF")


def _extract_hybrid_spreads(lines: list[str]) -> tuple[float | None, float | None]:
    for line in lines:
        if "Reoffer Spread" in line:
            values = re.findall(r"(\d+)\s*bps\s*area", line)
            if len(values) >= 2:
                return float(values[0]), float(values[-1])
    return None, None


def _extract_hybrid_yields(lines: list[str]) -> tuple[float | None, float | None]:
    for line in lines:
        if "Reoffer Yield to Call" in line:
            values = re.findall(r"(\d+\.\d+)%\s*area", line)
            if len(values) >= 2:
                return float(values[0]) / 100, float(values[-1]) / 100
    return None, None
