"""Write parsed PDF data into the Master File Excel workbook."""

from copy import copy
from datetime import date, datetime, timedelta
from openpyxl import load_workbook
from openpyxl.chart import LineChart, Reference
from openpyxl.chart.series import SeriesLabel
from openpyxl.styles import Alignment

# Column mapping: Master File column letters -> parsed data keys
# Pricing sheet columns (1-indexed):
#  A=1 Date, B=2 Bank,
#  C=3..L=12: CAD 3Y-30Y (spread, yield alternating)
#  M=13..V=22: USD 3Y-30Y (spread, yield alternating)
#  W=23..Z=26: CAD NC5/NC10 (spread, coupon)
#  AA=27..AD=30: USD NC5/NC10 (spread, coupon)

COLUMN_MAP = {
    3: "cad_spread_3y", 4: "cad_yield_3y",
    5: "cad_spread_5y", 6: "cad_yield_5y",
    7: "cad_spread_7y", 8: "cad_yield_7y",
    9: "cad_spread_10y", 10: "cad_yield_10y",
    11: "cad_spread_30y", 12: "cad_yield_30y",
    13: "usd_spread_3y", 14: "usd_yield_3y",
    15: "usd_spread_5y", 16: "usd_yield_5y",
    17: "usd_spread_7y", 18: "usd_yield_7y",
    19: "usd_spread_10y", 20: "usd_yield_10y",
    21: "usd_spread_30y", 22: "usd_yield_30y",
    23: "cad_nc5_spread", 24: "cad_nc5_coupon",
    25: "cad_nc10_spread", 26: "cad_nc10_coupon",
    27: "usd_nc5_spread", 28: "usd_nc5_coupon",
    29: "usd_nc10_spread", 30: "usd_nc10_coupon",
}


def _major_unit(span: float, is_pct: bool) -> float:
    """Choose a readable y-axis major unit based on data range."""
    if is_pct:
        # Values are decimals (0.0450 = 4.50%); target denser gridlines.
        candidates = [0.00025, 0.0005, 0.001, 0.002, 0.0025, 0.005, 0.01]
    else:
        # Spread values in bps; include tighter steps for analysis precision.
        candidates = [0.5, 1.0, 2.0, 2.5, 5.0, 10.0, 20.0, 25.0, 50.0]

    target = max(span / 12.0, candidates[0])
    for step in candidates:
        if step >= target:
            return step
    return candidates[-1]


def _is_macro_enabled(path: str) -> bool:
    return path.lower().endswith(".xlsm")


def append_row(master_file_path: str, data: dict):
    """Append a new row to the Pricing sheet of the Master File."""
    wb = load_workbook(master_file_path, keep_vba=_is_macro_enabled(master_file_path))
    ws = wb["Pricing"]

    next_row = ws.max_row + 1

    # Find actual next empty row (max_row can overcount with formatting)
    for r in range(2, ws.max_row + 2):
        if ws.cell(row=r, column=1).value is None:
            next_row = r
            break

    center = Alignment(horizontal="center", vertical="center")

    cell = ws.cell(row=next_row, column=1, value=data["date"])
    cell.number_format = "YYYY-MM-DD"
    cell.alignment = center

    cell = ws.cell(row=next_row, column=2, value=data["bank"])
    cell.alignment = center

    # Columns that hold yield/coupon percentages (even-numbered in senior, coupon in hybrid)
    pct_columns = {4, 6, 8, 10, 12, 14, 16, 18, 20, 22, 24, 26, 28, 30}

    for col, key in COLUMN_MAP.items():
        value = data.get(key)
        if value is not None:
            cell = ws.cell(row=next_row, column=col, value=value)
            cell.alignment = center
            if col in pct_columns:
                cell.number_format = "0.000%"

    wb.save(master_file_path)
    print(f"Wrote {data['bank']} data for {data['date'].strftime('%Y-%m-%d')} to row {next_row}")
    return wb


def _normalize_date(value):
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    if isinstance(value, str):
        raw = value.strip()
        if not raw:
            return None
        for fmt in ("%Y-%m-%d", "%Y/%m/%d", "%m/%d/%Y", "%d/%m/%Y"):
            try:
                return datetime.strptime(raw, fmt).date()
            except ValueError:
                continue
        try:
            return datetime.fromisoformat(raw).date()
        except ValueError:
            return None
    return None


def _normalize_bank(value):
    if value is None:
        return None
    bank = str(value).strip()
    return bank or None


def _numeric_value(value):
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    return None


def _collect_pricing_rows(ws_data):
    rows = []
    for r in range(2, ws_data.max_row + 1):
        date_val = _normalize_date(ws_data.cell(row=r, column=1).value)
        bank_val = _normalize_bank(ws_data.cell(row=r, column=2).value)
        if not date_val or not bank_val:
            continue
        rows.append({"row": r, "date": date_val, "bank": bank_val})
    return rows


def _row_dedup_key(row):
    return row["date"], row["bank"].casefold()


def _dedupe_rows_by_date_bank(rows):
    """Keep only the newest row for each (date, bank) key."""
    seen = set()
    deduped_reverse = []
    for row in reversed(rows):
        key = _row_dedup_key(row)
        if key in seen:
            continue
        seen.add(key)
        deduped_reverse.append(row)
    return list(reversed(deduped_reverse))


def _snapshot_cell_payload(cell):
    return {
        "value": cell.value,
        "style": copy(cell._style),
        "comment": copy(cell.comment),
        "hyperlink": copy(cell.hyperlink),
    }


def _restore_cell_payload(cell, payload):
    cell.value = payload["value"]
    cell._style = copy(payload["style"])
    cell.comment = copy(payload["comment"])
    cell.hyperlink = copy(payload["hyperlink"])


def _collect_pricing_row_payloads(ws_data):
    max_col = ws_data.max_column
    rows = []
    for r in range(2, ws_data.max_row + 1):
        cells = [_snapshot_cell_payload(ws_data.cell(row=r, column=c)) for c in range(1, max_col + 1)]
        date_val = _normalize_date(cells[0]["value"]) if cells else None
        bank_val = _normalize_bank(cells[1]["value"]) if len(cells) > 1 else None
        key = (date_val, bank_val.casefold()) if date_val and bank_val else None
        rows.append(
            {
                "row": r,
                "date": date_val,
                "bank": bank_val,
                "dedup_key": key,
                "cells": cells,
            }
        )
    return rows


def _dedupe_pricing_row_payloads(rows):
    seen = set()
    deduped_reverse = []
    removed = 0

    for row in reversed(rows):
        key = row["dedup_key"]
        if key is None:
            deduped_reverse.append(row)
            continue
        if key in seen:
            removed += 1
            continue
        seen.add(key)
        deduped_reverse.append(row)

    return list(reversed(deduped_reverse)), removed


def _sort_pricing_row_payloads(rows):
    valid_rows = [row for row in rows if row["dedup_key"] is not None]
    invalid_rows = [row for row in rows if row["dedup_key"] is None]
    sorted_valid = sorted(valid_rows, key=lambda row: (row["bank"].casefold(), row["date"], row["row"]))
    return sorted_valid + invalid_rows


def _rewrite_pricing_rows(ws_data, rows):
    existing_rows = ws_data.max_row - 1
    if existing_rows > 0:
        ws_data.delete_rows(2, existing_rows)

    for target_row, row_payload in enumerate(rows, start=2):
        for col, cell_payload in enumerate(row_payload["cells"], start=1):
            cell = ws_data.cell(row=target_row, column=col)
            _restore_cell_payload(cell, cell_payload)


def deduplicate_pricing_rows(master_file_path: str) -> int:
    """Deduplicate and reorder Pricing rows.

    1) Remove duplicate rows by (date, bank case-insensitive), keeping newest.
    2) Order valid rows by bank (case-insensitive), then by date ascending.
    3) Keep rows with missing date or bank at the bottom in original order.

    Returns the number of duplicates removed.
    """
    wb = load_workbook(master_file_path, keep_vba=_is_macro_enabled(master_file_path))
    ws_data = wb["Pricing"]

    all_rows = _collect_pricing_row_payloads(ws_data)
    deduped_rows, removed = _dedupe_pricing_row_payloads(all_rows)
    sorted_rows = _sort_pricing_row_payloads(deduped_rows)
    order_changed = [row["row"] for row in sorted_rows] != [row["row"] for row in deduped_rows]

    if removed or order_changed:
        _rewrite_pricing_rows(ws_data, sorted_rows)
        wb.save(master_file_path)

    return removed


def _latest_per_bank(rows):
    latest = {}
    for row in rows:
        key = row["bank"].casefold()
        existing = latest.get(key)
        if not existing or row["date"] > existing["date"]:
            latest[key] = row
    return sorted(latest.values(), key=lambda item: (item["date"], item["bank"].casefold()))


def _coerce_year(value):
    try:
        return int(value)
    except (TypeError, ValueError):
        return None


def _resolve_year_range(rows, avg_start_year=None, avg_end_year=None):
    """Resolve an inclusive year window for average spread charts.

    The UI passes `avg_start_year` / `avg_end_year` as optional values.
    Missing values default to the min/max year present in `rows` (or current
    year when no rows are available). If bounds are reversed, they are swapped.
    """
    years = sorted({row["date"].year for row in rows})
    current_year = datetime.now().year

    if years:
        start = _coerce_year(avg_start_year)
        end = _coerce_year(avg_end_year)
        if start is None:
            start = years[0]
        if end is None:
            end = years[-1]
    else:
        start = _coerce_year(avg_start_year) or current_year
        end = _coerce_year(avg_end_year) or current_year

    if start > end:
        start, end = end, start
    return start, end


def _build_standard_curve_chart(ws, ws_data, rows, cfg, tenors, center):
    sr = cfg["table_start_row"]
    sc = cfg["table_start_col"]

    ws.cell(row=sr, column=sc, value="Bank").alignment = center
    for j, tenor in enumerate(tenors):
        ws.cell(row=sr, column=sc + 1 + j, value=tenor).alignment = center

    for i, row_info in enumerate(rows):
        label = f"{row_info['bank']} ({row_info['date'].strftime('%Y-%m-%d')})"
        ws.cell(row=sr + 1 + i, column=sc, value=label).alignment = center
        for j, pricing_col in enumerate(cfg["cols"]):
            value = ws_data.cell(row=row_info["row"], column=pricing_col).value
            cell = ws.cell(row=sr + 1 + i, column=sc + 1 + j, value=value)
            cell.alignment = center
            if cfg["is_pct"] and value is not None:
                cell.number_format = "0.00%"

    chart = LineChart()
    chart.title = cfg["title"]
    chart.title.overlay = False
    chart.width = 24.0
    chart.height = 14.0
    chart.style = 2

    chart.x_axis.title = "Tenor"
    chart.x_axis.title.overlay = False
    chart.x_axis.tickLblPos = "low"
    chart.x_axis.delete = False

    chart.y_axis.title = cfg["y_label"]
    chart.y_axis.title.overlay = False
    chart.y_axis.delete = False
    chart.y_axis.numFmt = "0.00%" if cfg["is_pct"] else "0"

    y_values = []
    for row_info in rows:
        for pricing_col in cfg["cols"]:
            value = _numeric_value(ws_data.cell(row=row_info["row"], column=pricing_col).value)
            if value is not None:
                y_values.append(value)

    if y_values:
        y_min = min(y_values)
        y_max = max(y_values)
        span = max(y_max - y_min, 1e-9)
        major = _major_unit(span, cfg["is_pct"])
        chart.y_axis.majorUnit = major
        chart.y_axis.minorUnit = major / 5
        chart.y_axis.scaling.min = max(0.0, y_min - major)
        chart.y_axis.scaling.max = y_max + major

    categories = Reference(ws, min_col=sc + 1, max_col=sc + len(tenors), min_row=sr)
    for i in range(len(rows)):
        values = Reference(ws, min_col=sc + 1, max_col=sc + len(tenors), min_row=sr + 1 + i)
        chart.add_data(values, from_rows=True, titles_from_data=False)
        series = chart.series[-1]
        series.tx = SeriesLabel(v=ws.cell(row=sr + 1 + i, column=sc).value)
        series.smooth = False
        series.graphicalProperties.line.width = 22000
        series.marker.symbol = "circle"
        series.marker.size = 7

    chart.set_categories(categories)
    chart.legend.position = "r"
    chart.legend.overlay = False
    ws.add_chart(chart, cfg["chart_anchor"])


def _iso_week_start(value_date):
    """Return the Monday date for `value_date`'s ISO week."""
    return value_date - timedelta(days=value_date.isoweekday() - 1)


def _aggregate_weekly_average_spreads(ws_data, rows, spread_cols):
    """Aggregate spread columns into weekly equal-bank-weight averages.

    For each ISO week and tenor column:
    1. Collect all numeric observations per bank within the week.
    2. Compute each bank's mean for that tenor/week.
    3. Compute the simple average across banks.

    Weeks where all requested tenors are missing are excluded.
    """
    weekly_values = {}
    for row in rows:
        week_start = _iso_week_start(row["date"])
        per_week = weekly_values.setdefault(week_start, {})
        bank_key = row["bank"].casefold()
        per_bank = per_week.setdefault(bank_key, {})

        for col in spread_cols:
            value = _numeric_value(ws_data.cell(row=row["row"], column=col).value)
            if value is None:
                continue
            per_bank.setdefault(col, []).append(value)

    weekly_points = []
    for week_start in sorted(weekly_values):
        per_week = weekly_values[week_start]
        tenor_values = []
        has_any_value = False

        for col in spread_cols:
            bank_means = []
            for per_bank in per_week.values():
                values = per_bank.get(col, [])
                if values:
                    bank_means.append(sum(values) / len(values))

            if bank_means:
                tenor_mean = sum(bank_means) / len(bank_means)
                has_any_value = True
            else:
                tenor_mean = None
            tenor_values.append(tenor_mean)

        if has_any_value:
            weekly_points.append({"week_start": week_start, "values": tenor_values})

    return weekly_points


def _build_average_spread_time_series_chart(ws, weekly_points, cfg, year_start, year_end, center):
    """Render one weekly average spread time-series chart and backing table."""
    sr = cfg["table_start_row"]
    sc = cfg["table_start_col"]
    tenors = cfg["tenors"]

    ws.cell(row=sr, column=sc, value="Week Start").alignment = center
    for j, tenor in enumerate(tenors):
        ws.cell(row=sr, column=sc + 1 + j, value=tenor).alignment = center

    for i, point in enumerate(weekly_points, start=1):
        date_cell = ws.cell(row=sr + i, column=sc, value=point["week_start"])
        date_cell.number_format = "YYYY-MM-DD"
        date_cell.alignment = center
        for j, value in enumerate(point["values"]):
            ws.cell(row=sr + i, column=sc + 1 + j, value=value).alignment = center

    if weekly_points:
        data_end_row = sr + len(weekly_points)
    else:
        data_end_row = sr + 1
        ws.cell(row=data_end_row, column=sc, value=None).alignment = center
        for j in range(len(tenors)):
            ws.cell(row=data_end_row, column=sc + 1 + j, value=None).alignment = center

    chart = LineChart()
    chart.title = f"{cfg['title_prefix']} ({year_start}-{year_end})"
    chart.title.overlay = False
    chart.width = 24.0
    chart.height = 14.0
    chart.style = 2

    chart.x_axis.title = "Week Start (Monday)"
    chart.x_axis.title.overlay = False
    chart.x_axis.tickLblPos = "low"
    chart.x_axis.delete = False

    chart.y_axis.title = "Spread (bps)"
    chart.y_axis.title.overlay = False
    chart.y_axis.delete = False
    chart.y_axis.numFmt = "0"

    y_values = [value for point in weekly_points for value in point["values"] if value is not None]
    if y_values:
        y_min = min(y_values)
        y_max = max(y_values)
        span = max(y_max - y_min, 1e-9)
        major = _major_unit(span, is_pct=False)
        chart.y_axis.majorUnit = major
        chart.y_axis.minorUnit = major / 5
        chart.y_axis.scaling.min = max(0.0, y_min - major)
        chart.y_axis.scaling.max = y_max + major

    categories = Reference(ws, min_col=sc, max_col=sc, min_row=sr + 1, max_row=data_end_row)
    for j, tenor in enumerate(tenors):
        values = Reference(
            ws,
            min_col=sc + 1 + j,
            max_col=sc + 1 + j,
            min_row=sr + 1,
            max_row=data_end_row,
        )
        chart.add_data(values, from_rows=False, titles_from_data=False)
        series = chart.series[-1]
        series.tx = SeriesLabel(v=tenor)
        series.smooth = False
        series.graphicalProperties.line.width = 22000
        series.marker.symbol = "circle"
        series.marker.size = 6

    chart.set_categories(categories)
    chart.legend.position = "r"
    chart.legend.overlay = False
    ws.add_chart(chart, cfg["chart_anchor"])


def update_charts(master_file_path: str, avg_start_year=None, avg_end_year=None):
    """Create/update all Summary Charts outputs (6 charts total).

    Core charts (unchanged behavior):
    - CAD spread curve
    - CAD yield curve
    - USD spread curve
    - USD yield curve
    Each core chart uses the most recent available row per bank.

    Average charts (weekly time-series):
    - CAD Average Spread Through Time
    - USD Average Spread Through Time
    Each average chart plots ISO-week (Monday) categories and 4 tenor series
    (3Y, 5Y, 10Y, 30Y), bounded by the inclusive `avg_start_year` /
    `avg_end_year` filter.

    Input rows are deduplicated by (date, bank case-insensitive), keeping
    the newest row for each key.
    """
    wb = load_workbook(master_file_path, keep_vba=_is_macro_enabled(master_file_path))
    ws_data = wb["Pricing"]
    ws = wb["Summary Charts"]

    # Clear sheet completely
    for row in ws.iter_rows():
        for cell in row:
            cell.value = None
    ws._charts = []

    all_rows = _collect_pricing_rows(ws_data)
    deduped_rows = _dedupe_rows_by_date_bank(all_rows)
    rows = _latest_per_bank(deduped_rows)

    tenors = ["3Y", "5Y", "7Y", "10Y", "30Y"]

    # Data tables are placed far below the visible chart area (row 100+)
    # so they don't clutter the view. Charts are anchored in the visible area.
    chart_configs = [
        {
            "title": "Bell Canada - CAD New Issue Spread Curve (bps)",
            "cols": [3, 5, 7, 9, 11],
            "y_label": "Spread (bps)",
            "is_pct": False,
            "chart_anchor": "A1",
            "table_start_row": 140,
            "table_start_col": 1,
        },
        {
            "title": "Bell Canada - CAD Re-Offer Yield Curve",
            "cols": [4, 6, 8, 10, 12],
            "y_label": "Yield (%)",
            "is_pct": True,
            "chart_anchor": "Q1",
            "table_start_row": 140,
            "table_start_col": 10,
        },
        {
            "title": "Bell Canada - USD New Issue Spread Curve (bps)",
            "cols": [13, 15, 17, 19, 21],
            "y_label": "Spread (bps)",
            "is_pct": False,
            "chart_anchor": "A33",
            "table_start_row": 170,
            "table_start_col": 1,
        },
        {
            "title": "Bell Canada - USD Re-Offer Yield Curve",
            "cols": [14, 16, 18, 20, 22],
            "y_label": "Yield (%)",
            "is_pct": True,
            "chart_anchor": "Q33",
            "table_start_row": 170,
            "table_start_col": 10,
        },
    ]

    avg_tenors = ["3Y", "5Y", "10Y", "30Y"]
    avg_chart_configs = [
        {
            "title_prefix": "Bell Canada - CAD Average Spread Through Time",
            "spread_cols": [3, 5, 9, 11],
            "tenors": avg_tenors,
            "chart_anchor": "A65",
            "table_start_row": 200,
            "table_start_col": 1,
        },
        {
            "title_prefix": "Bell Canada - USD Average Spread Through Time",
            "spread_cols": [13, 15, 19, 21],
            "tenors": avg_tenors,
            "chart_anchor": "Q65",
            "table_start_row": 200,
            "table_start_col": 10,
        },
    ]

    center = Alignment(horizontal="center", vertical="center")

    for cfg in chart_configs:
        _build_standard_curve_chart(ws, ws_data, rows, cfg, tenors, center)

    year_start, year_end = _resolve_year_range(
        deduped_rows,
        avg_start_year=avg_start_year,
        avg_end_year=avg_end_year,
    )
    average_rows = [row for row in deduped_rows if year_start <= row["date"].year <= year_end]

    for cfg in avg_chart_configs:
        weekly_points = _aggregate_weekly_average_spreads(ws_data, average_rows, cfg["spread_cols"])
        _build_average_spread_time_series_chart(
            ws,
            weekly_points=weekly_points,
            cfg=cfg,
            year_start=year_start,
            year_end=year_end,
            center=center,
        )

    wb.save(master_file_path)
    print(f"Updated {len(chart_configs) + len(avg_chart_configs)} yield curve charts in Summary Charts tab")
