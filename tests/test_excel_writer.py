from datetime import date, datetime
from tempfile import TemporaryDirectory
import unittest

from openpyxl import Workbook, load_workbook

import excel_writer


def _spread_values(
    cad_3y=None,
    cad_5y=None,
    cad_7y=None,
    cad_10y=None,
    cad_30y=None,
    usd_3y=None,
    usd_5y=None,
    usd_7y=None,
    usd_10y=None,
    usd_30y=None,
):
    values = {
        3: cad_3y,
        5: cad_5y,
        7: cad_7y,
        9: cad_10y,
        11: cad_30y,
        13: usd_3y,
        15: usd_5y,
        17: usd_7y,
        19: usd_10y,
        21: usd_30y,
    }
    return {col: value for col, value in values.items() if value is not None}


def _to_date(value):
    if isinstance(value, datetime):
        return value.date()
    return value


def _read_weekly_table(ws, table_start_row, table_start_col):
    points = []
    row = table_start_row + 1
    while True:
        week_start = ws.cell(row=row, column=table_start_col).value
        tenor_values = [ws.cell(row=row, column=table_start_col + i).value for i in range(1, 5)]
        if week_start is None and all(value is None for value in tenor_values):
            break
        points.append((_to_date(week_start), tenor_values))
        row += 1
    return points


def _chart_title_text(chart):
    return chart.title.tx.rich.p[0].r[0].t


class UpdateChartsWeeklyAverageTests(unittest.TestCase):
    def _create_sample_workbook(self, path):
        wb = Workbook()
        ws_pricing = wb.active
        ws_pricing.title = "Pricing"
        wb.create_sheet("Summary Charts")

        ws_pricing.cell(row=1, column=1, value="Date")
        ws_pricing.cell(row=1, column=2, value="Bank")

        rows = [
            (
                date(2023, 12, 31),
                "LegacyBank",
                _spread_values(70, 80, 90, 100, 120, 170, 180, 190, 200, 220),
            ),
            (
                date(2024, 1, 2),
                "BankA",
                _spread_values(100, 110, 120, 130, 150, 200, 210, 220, 230, 250),
            ),
            (
                date(2024, 1, 2),
                "banka",
                _spread_values(999, 999, 999, 999, 999, 999, 999, 999, 999, 999),
            ),
            (
                date(2024, 1, 3),
                "BankA",
                _spread_values(120, 130, 140, 150, 170, 220, 230, 240, 250, 270),
            ),
            (
                date(2024, 1, 4),
                "BankB",
                _spread_values(80, 90, 100, 110, 130, 180, 190, 200, 210, 230),
            ),
            (
                date(2024, 1, 9),
                "BankA",
                _spread_values(140, 150, 160, 170, 190, 240, 250, 260, 270, 290),
            ),
            (
                date(2024, 1, 10),
                "BankB",
                _spread_values(100, 110, 120, 130, 150, 200, 210, 220, 230, 250),
            ),
            (date(2024, 2, 6), "SparseBank", {}),
            (
                date(2025, 1, 7),
                "BankA",
                _spread_values(160, 170, 180, 190, 210, 260, 270, 280, 290, 310),
            ),
            (
                date(2026, 1, 6),
                "FutureBank",
                _spread_values(180, 190, 200, 210, 230, 280, 290, 300, 310, 330),
            ),
        ]

        for row_number, (row_date, bank, spread_map) in enumerate(rows, start=2):
            ws_pricing.cell(row=row_number, column=1, value=row_date)
            ws_pricing.cell(row=row_number, column=2, value=bank)
            for col, value in spread_map.items():
                ws_pricing.cell(row=row_number, column=col, value=value)

        wb.save(path)

    def test_iso_week_start_uses_monday(self):
        self.assertEqual(excel_writer._iso_week_start(date(2024, 1, 3)), date(2024, 1, 1))
        self.assertEqual(excel_writer._iso_week_start(date(2023, 12, 31)), date(2023, 12, 25))

    def test_weekly_aggregation_dedup_equal_weight_and_sparse_omit(self):
        with TemporaryDirectory() as temp_dir:
            workbook_path = f"{temp_dir}/master.xlsx"
            self._create_sample_workbook(workbook_path)

            excel_writer.update_charts(workbook_path, avg_start_year=2024, avg_end_year=2024)

            wb = load_workbook(workbook_path)
            ws = wb["Summary Charts"]
            cad_points = _read_weekly_table(ws, table_start_row=200, table_start_col=1)
            usd_points = _read_weekly_table(ws, table_start_row=200, table_start_col=10)

            self.assertEqual([week for week, _ in cad_points], [date(2024, 1, 1), date(2024, 1, 8)])
            self.assertEqual([week for week, _ in usd_points], [date(2024, 1, 1), date(2024, 1, 8)])

            expected_cad = [
                [95.0, 105.0, 125.0, 145.0],
                [120.0, 130.0, 150.0, 170.0],
            ]
            expected_usd = [
                [195.0, 205.0, 225.0, 245.0],
                [220.0, 230.0, 250.0, 270.0],
            ]

            for actual_values, expected_values in zip([values for _, values in cad_points], expected_cad):
                for actual, expected in zip(actual_values, expected_values):
                    self.assertAlmostEqual(actual, expected)

            for actual_values, expected_values in zip([values for _, values in usd_points], expected_usd):
                for actual, expected in zip(actual_values, expected_values):
                    self.assertAlmostEqual(actual, expected)

    def test_swapped_year_range_is_inclusive_and_charts_regressions_hold(self):
        with TemporaryDirectory() as temp_dir:
            workbook_path = f"{temp_dir}/master.xlsx"
            self._create_sample_workbook(workbook_path)

            excel_writer.update_charts(workbook_path, avg_start_year=2025, avg_end_year=2024)

            wb = load_workbook(workbook_path)
            ws = wb["Summary Charts"]
            cad_points = _read_weekly_table(ws, table_start_row=200, table_start_col=1)
            usd_points = _read_weekly_table(ws, table_start_row=200, table_start_col=10)

            expected_weeks = [date(2024, 1, 1), date(2024, 1, 8), date(2025, 1, 6)]
            self.assertEqual([week for week, _ in cad_points], expected_weeks)
            self.assertEqual([week for week, _ in usd_points], expected_weeks)

            self.assertEqual(len(ws._charts), 6)
            self.assertEqual(len(ws._charts[4].series), 4)
            self.assertEqual(len(ws._charts[5].series), 4)

            expected_core_titles = [
                "Bell Canada - CAD New Issue Spread Curve (bps)",
                "Bell Canada - CAD Re-Offer Yield Curve",
                "Bell Canada - USD New Issue Spread Curve (bps)",
                "Bell Canada - USD Re-Offer Yield Curve",
            ]
            self.assertEqual([_chart_title_text(chart) for chart in ws._charts[:4]], expected_core_titles)
            self.assertEqual(
                _chart_title_text(ws._charts[4]),
                "Bell Canada - CAD Average Spread Through Time (2024-2025)",
            )
            self.assertEqual(
                _chart_title_text(ws._charts[5]),
                "Bell Canada - USD Average Spread Through Time (2024-2025)",
            )


if __name__ == "__main__":
    unittest.main()
