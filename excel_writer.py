"""Write parsed PDF data into the Master File Excel workbook."""

from datetime import datetime
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


def update_charts(master_file_path: str):
    """Create/update yield curve charts on the Summary Charts sheet.

    Standard finance yield curve format:
    - X axis: Tenor (3Y, 5Y, 7Y, 10Y, 30Y)
    - Y axis: Spread (bps) or Yield (%)
    - One line per bank/date combination
    """
    wb = load_workbook(master_file_path, keep_vba=_is_macro_enabled(master_file_path))
    ws_data = wb["Pricing"]
    ws = wb["Summary Charts"]

    # Clear sheet completely
    for row in ws.iter_rows():
        for cell in row:
            cell.value = None
    ws._charts = []

    # Collect data rows from Pricing (only contiguous rows starting from row 2)
    all_rows = []
    for r in range(2, ws_data.max_row + 1):
        date_val = ws_data.cell(row=r, column=1).value
        bank_val = ws_data.cell(row=r, column=2).value
        if date_val and bank_val:
            all_rows.append(r)
        else:
            break  # stop at first empty row to skip stray data

    # Keep only the most recent row per bank
    latest_per_bank = {}
    for r in all_rows:
        date_val = ws_data.cell(row=r, column=1).value
        bank_val = ws_data.cell(row=r, column=2).value
        if bank_val not in latest_per_bank or date_val > latest_per_bank[bank_val][1]:
            latest_per_bank[bank_val] = (r, date_val)
    rows = [v[0] for v in sorted(latest_per_bank.values(), key=lambda x: x[1])]

    if not rows:
        wb.save(master_file_path)
        return

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

    center = Alignment(horizontal="center", vertical="center")

    for cfg in chart_configs:
        sr = cfg["table_start_row"]
        sc = cfg["table_start_col"]

        # Write tenor headers: row sr, cols sc+1 .. sc+5
        ws.cell(row=sr, column=sc, value="Bank").alignment = center
        for j, tenor in enumerate(tenors):
            ws.cell(row=sr, column=sc + 1 + j, value=tenor).alignment = center

        # Write one row per bank/date
        for i, data_row in enumerate(rows):
            date_val = ws_data.cell(row=data_row, column=1).value
            bank_val = ws_data.cell(row=data_row, column=2).value
            if isinstance(date_val, datetime):
                label = f"{bank_val} ({date_val.strftime('%Y-%m-%d')})"
            else:
                label = f"{bank_val}"

            ws.cell(row=sr + 1 + i, column=sc, value=label).alignment = center
            for j, pricing_col in enumerate(cfg["cols"]):
                val = ws_data.cell(row=data_row, column=pricing_col).value
                cell = ws.cell(row=sr + 1 + i, column=sc + 1 + j, value=val)
                cell.alignment = center
                if cfg["is_pct"] and val is not None:
                    cell.number_format = "0.00%"

        num_series = len(rows)

        # Build the chart
        chart = LineChart()
        chart.title = cfg["title"]
        chart.width = 24.0
        chart.height = 14.0
        chart.style = 2

        # X axis = Tenor
        chart.x_axis.title = "Tenor"
        chart.x_axis.tickLblPos = "low"
        chart.x_axis.delete = False

        # Y axis
        chart.y_axis.title = cfg["y_label"]
        chart.y_axis.delete = False
        chart.y_axis.numFmt = "0.00%" if cfg["is_pct"] else "0"
        # Set a finance-friendly y-axis scale with tighter grid spacing.
        y_values = []
        for data_row in rows:
            for pricing_col in cfg["cols"]:
                val = ws_data.cell(row=data_row, column=pricing_col).value
                if isinstance(val, (int, float)):
                    y_values.append(float(val))

        if y_values:
            y_min = min(y_values)
            y_max = max(y_values)
            span = max(y_max - y_min, 1e-9)
            major = _major_unit(span, cfg["is_pct"])
            chart.y_axis.majorUnit = major
            chart.y_axis.minorUnit = major / 5

            lower = max(0.0, y_min - major)
            upper = y_max + major
            chart.y_axis.scaling.min = lower
            chart.y_axis.scaling.max = upper

        # Categories = tenor labels
        categories = Reference(ws, min_col=sc + 1, max_col=sc + 5, min_row=sr)

        # One series per bank/date row
        for i in range(num_series):
            values = Reference(ws, min_col=sc + 1, max_col=sc + 5, min_row=sr + 1 + i)
            chart.add_data(values, from_rows=True, titles_from_data=False)
            series = chart.series[-1]
            series.tx = SeriesLabel(v=ws.cell(row=sr + 1 + i, column=sc).value)
            series.smooth = False
            series.graphicalProperties.line.width = 22000  # ~1.75pt
            series.marker.symbol = "circle"
            series.marker.size = 7

        chart.set_categories(categories)

        # Put legend on the right for clearer bank/date labels.
        chart.legend.position = "r"
        chart.legend.overlay = False

        ws.add_chart(chart, cfg["chart_anchor"])

    wb.save(master_file_path)
    print(f"Updated {len(chart_configs)} yield curve charts in Summary Charts tab")
