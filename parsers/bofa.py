"""Parser for Bank of America Bell Canada indicative new issue PDFs."""

import re
from datetime import datetime


def parse_bofa_pdf(pdf_path: str) -> dict:
    import pdfplumber

    with pdfplumber.open(pdf_path) as pdf:
        pages = list(pdf.pages)
        pages_text = [p.extract_text() or "" for p in pages]

    date = _extract_date(pages_text)

    pricing_idx = next(
        i for i, t in enumerate(pages_text) if "USD Senior Unsecured New Issue Pricing" in t
    )
    pricing_lines = pages_text[pricing_idx].split("\n")

    usd_spreads, usd_yields = _parse_senior_section(pricing_lines, "USD Senior Unsecured")
    cad_spreads, cad_yields = _parse_senior_section(pricing_lines, "CAD Senior Unsecured")

    result = {"date": date, "bank": "BofA"}

    for tenor in ["3y", "5y", "7y", "10y", "30y"]:
        result[f"usd_spread_{tenor}"] = usd_spreads.get(tenor)
        result[f"usd_yield_{tenor}"] = usd_yields.get(tenor)
        result[f"cad_spread_{tenor}"] = cad_spreads.get(tenor)
        result[f"cad_yield_{tenor}"] = cad_yields.get(tenor)

    hybrid_idx = next(
        (i for i, t in enumerate(pages_text) if "Hybrid New Issue Indicative Pricing" in t), None
    )
    if hybrid_idx is not None:
        _parse_hybrid(pages[hybrid_idx], pages_text[hybrid_idx], result)

    return result


def _extract_date(pages_text: list[str]) -> datetime:
    for text in pages_text:
        m = re.search(r"As of (\w+ \d{1,2})(?:st|nd|rd|th),\s*(\d{4})", text)
        if m:
            return datetime.strptime(f"{m.group(1)}, {m.group(2)}", "%B %d, %Y")
    raise ValueError("Could not find date in BofA PDF")


def _parse_senior_section(
    lines: list[str], section_header: str
) -> tuple[dict, dict]:
    """Return (spreads, yields) dicts keyed by tenor (3y/5y/7y/10y/30y)."""
    year_to_key = {"3": "3y", "5": "5y", "7": "7y", "10": "10y", "30": "30y"}

    start = next((i for i, l in enumerate(lines) if section_header in l), None)
    if start is None:
        return {}, {}

    end = len(lines)
    for i in range(start + 2, len(lines)):
        if "Senior Unsecured" in lines[i] or "Relative Value" in lines[i]:
            end = i
            break

    section = lines[start:end]
    spreads: dict = {}
    yields: dict = {}

    i = 0
    while i < len(section):
        m = re.match(r"^(\d+)-Year$", section[i].strip())
        if m and m.group(1) in year_to_key:
            key = year_to_key[m.group(1)]
            block, j = [], i + 1
            while j < len(section):
                nxt = section[j].strip()
                if re.match(r"^\d+-Year$", nxt):
                    break
                if nxt == "U" and j + 1 < len(section) and re.match(
                    r"^\d+-Year$", section[j + 1].strip()
                ):
                    break
                block.append(nxt)
                j += 1
            sp, yld = _parse_tenor_block(block)
            if sp is not None:
                spreads[key] = sp
            if yld is not None:
                yields[key] = yld
        i += 1

    return spreads, yields


def _parse_tenor_block(block: list[str]) -> tuple[float | None, float | None]:
    """Extract (spread, yield) from lines following a tenor name."""
    spread = None
    spread_idx = None

    for i, line in enumerate(block):
        if not line or line == "-":
            continue
        if re.match(r"^\+\d", line) and re.search(r"b\s*p\s*s", line):
            if spread is None:
                spread = _parse_spread(line)
                spread_idx = i

    yld = None
    if spread_idx is not None:
        for line in block[spread_idx + 1 :]:
            if re.match(r"^\d+\.\d+%", line):
                yld = _parse_yield(line)
                break

    return spread, yld


def _parse_spread(line: str) -> float | None:
    line = re.sub(r"b\s+p\s+s", "bps", line)
    m = re.search(r"\+(\d+(?:\.\d+)?)\s*-\s*(\d+(?:\.\d+)?)\s*bps", line)
    if m:
        return (float(m.group(1)) + float(m.group(2))) / 2
    m = re.search(r"\+(\d+(?:\.\d+)?)\s*bps", line)
    if m:
        return float(m.group(1))
    return None


def _parse_yield(line: str) -> float | None:
    m = re.search(r"(\d+\.\d+)%\s*-\s*(\d+\.\d+)%", line)
    if m:
        return (float(m.group(1)) + float(m.group(2))) / 2 / 100
    m = re.search(r"(\d+\.\d+)%", line)
    if m:
        return float(m.group(1)) / 100
    return None


def _parse_hybrid(page: object, page_text: str, result: dict) -> None:
    """Extract NC5/NC10 spread/coupon for CAD and USD (With Coupon Floors variant)."""
    # CAD: the clean "Term: 30NC5 30NC10" table in extract_text() is the With Coupon Floors table
    cad_m = re.search(
        r"Term:\s+30NC5\s+30NC10\b.*?"
        r"Sr\.\s+Unsecured\s+Spread:\s+([^\n]+)\n.*?"
        r"Re-Offer\s+Yield:\s+([^\n]+)",
        page_text,
        re.DOTALL,
    )
    if cad_m:
        cad_sp = _extract_spread_list(cad_m.group(1).strip())
        cad_yld = re.findall(r"(\d+\.\d+)%", cad_m.group(2).strip())
        if len(cad_sp) >= 1:
            result["cad_nc5_spread"] = cad_sp[0]
        if len(cad_yld) >= 1:
            result["cad_nc5_coupon"] = float(cad_yld[0]) / 100
        if len(cad_sp) >= 2:
            result["cad_nc10_spread"] = cad_sp[1]
        if len(cad_yld) >= 2:
            result["cad_nc10_coupon"] = float(cad_yld[1]) / 100

    # USD: the garbled tables require word-position extraction.
    # The With Coupon Floors columns occupy x < 400 in the USD hybrid section (y 260-400).
    words = page.extract_words()
    usd_rows: dict = {}
    for w in words:
        if 260 <= w["top"] <= 400:
            y_key = round(w["top"] / 5) * 5
            usd_rows.setdefault(y_key, []).append(w)

    for y_key in sorted(usd_rows.keys()):
        row_words = sorted(usd_rows[y_key], key=lambda w: w["x0"])
        compact = "".join(w["text"] for w in row_words if w["x0"] < 400)
        # Sr. Unsecured Spread row: first row with two "+X bps" patterns
        if not result.get("usd_nc5_spread"):
            bps_vals = re.findall(r"\+(\d+)bps", compact)
            if len(bps_vals) >= 2:
                result["usd_nc5_spread"] = float(bps_vals[0])
                result["usd_nc10_spread"] = float(bps_vals[1])
        # Re-Offer Yield row: row containing two percentages > 5%
        if not result.get("usd_nc5_coupon"):
            pcts = re.findall(r"(\d+\.\d+)%", compact)
            if len(pcts) >= 2 and all(float(p) > 5.0 for p in pcts[:2]):
                result["usd_nc5_coupon"] = float(pcts[0]) / 100
                result["usd_nc10_coupon"] = float(pcts[1]) / 100


def _extract_spread_list(line: str) -> list[float]:
    values = []
    for m in re.finditer(r"\+(\d+(?:\.\d+)?)\s*(?:-\s*(\d+(?:\.\d+)?))?\s*bps", line):
        lo = float(m.group(1))
        hi = float(m.group(2)) if m.group(2) else lo
        values.append((lo + hi) / 2)
    return values
