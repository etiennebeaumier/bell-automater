"""Parser for TD Securities Bell Canada (BCECN) indicative new issue PDFs."""

import re
from datetime import datetime


def parse_td_pdf(pdf_path: str) -> dict:
    """Parse a TD Securities BCECN PDF and return structured data.

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

    # Extract date from header: "As of: March 02, 2026"
    date = _extract_date(lines)

    # Split into USD, CAD, EUR sections by finding "Tenor" lines
    sections = _split_sections(lines)
    usd_section = sections[0]  # First section is USD
    cad_section = sections[1]  # Second section is CAD

    # Parse each section
    usd_data = _parse_section(usd_section, currency="USD")
    cad_data = _parse_section(cad_section, currency="CAD")

    result = {"date": date, "bank": "TD"}

    # Tenors: index 0=2Y, 1=3Y, 2=5Y, 3=7Y, 4=10Y, 5=30Y, 6=NC5, 7=NC10
    tenors = ["3y", "5y", "7y", "10y", "30y"]
    tenor_indices = [1, 2, 3, 4, 5]  # Skip 2Y (index 0) as Master File starts at 3Y

    for tenor, idx in zip(tenors, tenor_indices):
        result[f"cad_spread_{tenor}"] = cad_data["spread"][idx]
        result[f"cad_yield_{tenor}"] = cad_data["yield"][idx]
        result[f"usd_spread_{tenor}"] = usd_data["spread"][idx]
        result[f"usd_yield_{tenor}"] = usd_data["yield"][idx]

    # NC5 and NC10 (indices 6 and 7 in the data arrays)
    if len(usd_data["spread"]) > 7:
        result["usd_nc5_spread"] = usd_data["spread"][6]
        result["usd_nc5_coupon"] = usd_data["yield"][6]
        result["usd_nc10_spread"] = usd_data["spread"][7]
        result["usd_nc10_coupon"] = usd_data["yield"][7]

    if len(cad_data["spread"]) > 7:
        result["cad_nc5_spread"] = cad_data["spread"][6]
        result["cad_nc5_coupon"] = cad_data["yield"][6]
        result["cad_nc10_spread"] = cad_data["spread"][7]
        result["cad_nc10_coupon"] = cad_data["yield"][7]

    return result


def _extract_date(lines: list[str]) -> datetime:
    for line in lines:
        match = re.search(r"As of:\s*(\w+ \d{1,2},\s*\d{4})", line)
        if match:
            return datetime.strptime(match.group(1).replace("  ", " "), "%B %d, %Y")
    raise ValueError("Could not find date in PDF")


def _split_sections(lines: list[str]) -> list[list[str]]:
    """Split text into sections based on 'Tenor' header lines."""
    sections = []
    current = []
    for line in lines:
        if line.strip().startswith("Tenor"):
            if current:
                sections.append(current)
            current = [line]
        elif current:
            current.append(line)
    if current:
        sections.append(current)
    return sections


def _parse_section(lines: list[str], currency: str) -> dict:
    """Parse a section (USD or CAD) and extract spread and reoffer yield rows."""
    spread_row = None
    yield_row = None

    for line in lines:
        stripped = line.strip()
        if currency == "USD" and "New Issue Spread vs. UST" in stripped and not stripped.startswith("All-in"):
            spread_row = line
        elif currency == "CAD" and "New Issue Spread vs. GOC" in stripped and not stripped.startswith("All-in"):
            spread_row = line
        elif stripped.startswith("Reoffer Yield"):
            yield_row = line

    if not spread_row or not yield_row:
        raise ValueError(f"Could not find spread/yield rows for {currency}")

    spreads = _extract_numbers(spread_row)
    yields = _extract_percentages(yield_row)

    return {"spread": spreads, "yield": yields}


def _extract_numbers(line: str) -> list[float]:
    """Extract all numeric values from a line (integers and decimals, not percentages)."""
    # Remove the label part before the numbers
    # Find where the numbers start (after the label text)
    values = re.findall(r"(?<!\d[.])(\d+(?:\.\d+)?)(?!%)", line)
    # Filter: only keep values that look like spread values (not part of label)
    # The label might contain numbers, so we find the pattern after known text
    parts = re.split(r"\)\s*", line, maxsplit=1)
    if len(parts) > 1:
        num_part = parts[1]
    else:
        # Try splitting after the label
        num_part = line
    return [float(x) for x in re.findall(r"(\d+(?:\.\d+)?)", num_part)]


def _extract_percentages(line: str) -> list[float]:
    """Extract all percentage values from a line, returned as decimals (e.g., 3.36% -> 0.0336)."""
    pcts = re.findall(r"(\d+\.\d+)%", line)
    return [float(p) / 100 for p in pcts]
