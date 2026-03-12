from datetime import date, datetime
from tempfile import TemporaryDirectory
import unittest

from openpyxl import Workbook, load_workbook
from openpyxl.chart.axis import DateAxis
from openpyxl.styles import PatternFill

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


def _axis_format_code(axis):
    if axis.numFmt is None:
        return None
    return axis.numFmt.formatCode


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
                [319.75, 327.25, 342.25, 357.25],
                [120.0, 130.0, 150.0, 170.0],
            ]
            expected_usd = [
                [394.75, 402.25, 417.25, 432.25],
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
            self.assertIsInstance(ws._charts[4].x_axis, DateAxis)
            self.assertIsInstance(ws._charts[5].x_axis, DateAxis)
            self.assertEqual(
                _axis_format_code(ws._charts[4].x_axis),
                excel_writer.TIME_SERIES_AXIS_DATE_FORMAT,
            )
            self.assertEqual(
                _axis_format_code(ws._charts[5].x_axis),
                excel_writer.TIME_SERIES_AXIS_DATE_FORMAT,
            )

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


class DeduplicatePricingRowsTests(unittest.TestCase):
    def _create_workbook(self, path, rows):
        wb = Workbook()
        ws = wb.active
        ws.title = "Pricing"
        ws.cell(row=1, column=1, value="Date")
        ws.cell(row=1, column=2, value="Bank")
        ws.cell(row=1, column=3, value="CAD 3Y Spread")
        for row_number, row_values in enumerate(rows, start=2):
            ws.cell(row=row_number, column=1, value=row_values["date"])
            ws.cell(row=row_number, column=2, value=row_values["bank"])
            ws.cell(row=row_number, column=3, value=row_values["spread"])
        wb.save(path)

    def _read_pricing_rows(self, path):
        wb = load_workbook(path)
        ws = wb["Pricing"]
        rows = []
        for row_number in range(2, ws.max_row + 1):
            rows.append(
                (
                    _to_date(ws.cell(row=row_number, column=1).value),
                    ws.cell(row=row_number, column=2).value,
                    ws.cell(row=row_number, column=3).value,
                )
            )
        return rows

    def test_duplicate_date_bank_keeps_newest_case_insensitive(self):
        with TemporaryDirectory() as temp_dir:
            workbook_path = f"{temp_dir}/master.xlsx"
            self._create_workbook(
                workbook_path,
                rows=[
                    {"date": date(2024, 1, 2), "bank": "BankA", "spread": 100},
                    {"date": date(2024, 1, 2), "bank": "banka", "spread": 999},
                    {"date": date(2024, 1, 3), "bank": "BankB", "spread": 120},
                    {"date": date(2024, 1, 3), "bank": "BANKB", "spread": 130},
                ],
            )

            removed = excel_writer.deduplicate_pricing_rows(workbook_path)
            self.assertEqual(removed, 2)

            self.assertEqual(
                self._read_pricing_rows(workbook_path),
                [
                    (date(2024, 1, 2), "banka", 999),
                    (date(2024, 1, 3), "BANKB", 130),
                ],
            )

    def test_non_duplicate_rows_are_sorted_by_bank_then_date(self):
        with TemporaryDirectory() as temp_dir:
            workbook_path = f"{temp_dir}/master.xlsx"
            rows = [
                {"date": date(2024, 1, 2), "bank": "BankB", "spread": 90},
                {"date": date(2024, 1, 3), "bank": "BankA", "spread": 110},
                {"date": date(2024, 1, 2), "bank": "BankA", "spread": 100},
            ]
            self._create_workbook(workbook_path, rows=rows)

            removed = excel_writer.deduplicate_pricing_rows(workbook_path)

            self.assertEqual(removed, 0)
            self.assertEqual(
                self._read_pricing_rows(workbook_path),
                [
                    (date(2024, 1, 2), "BankA", 100),
                    (date(2024, 1, 3), "BankA", 110),
                    (date(2024, 1, 2), "BankB", 90),
                ],
            )

    def test_rows_missing_date_or_bank_are_untouched(self):
        with TemporaryDirectory() as temp_dir:
            workbook_path = f"{temp_dir}/master.xlsx"
            self._create_workbook(
                workbook_path,
                rows=[
                    {"date": date(2024, 1, 2), "bank": "BankA", "spread": 100},
                    {"date": date(2024, 1, 2), "bank": "banka", "spread": 999},
                    {"date": None, "bank": "NoDate", "spread": 10},
                    {"date": date(2024, 1, 4), "bank": None, "spread": 11},
                ],
            )

            removed = excel_writer.deduplicate_pricing_rows(workbook_path)
            self.assertEqual(removed, 1)
            self.assertEqual(
                self._read_pricing_rows(workbook_path),
                [
                    (date(2024, 1, 2), "banka", 999),
                    (None, "NoDate", 10),
                    (date(2024, 1, 4), None, 11),
                ],
            )

    def test_rows_missing_date_or_bank_move_to_bottom_in_original_order(self):
        with TemporaryDirectory() as temp_dir:
            workbook_path = f"{temp_dir}/master.xlsx"
            self._create_workbook(
                workbook_path,
                rows=[
                    {"date": None, "bank": "NoDate", "spread": 10},
                    {"date": date(2024, 1, 2), "bank": "BankB", "spread": 90},
                    {"date": date(2024, 1, 2), "bank": "BankA", "spread": 100},
                    {"date": date(2024, 1, 4), "bank": None, "spread": 11},
                ],
            )

            removed = excel_writer.deduplicate_pricing_rows(workbook_path)
            self.assertEqual(removed, 0)
            self.assertEqual(
                self._read_pricing_rows(workbook_path),
                [
                    (date(2024, 1, 2), "BankA", 100),
                    (date(2024, 1, 2), "BankB", 90),
                    (None, "NoDate", 10),
                    (date(2024, 1, 4), None, 11),
                ],
            )

    def test_reorder_preserves_cell_style_metadata(self):
        with TemporaryDirectory() as temp_dir:
            workbook_path = f"{temp_dir}/master.xlsx"
            self._create_workbook(
                workbook_path,
                rows=[
                    {"date": date(2024, 1, 2), "bank": "BankB", "spread": 90},
                    {"date": date(2024, 1, 2), "bank": "BankA", "spread": 100},
                ],
            )

            wb = load_workbook(workbook_path)
            ws = wb["Pricing"]
            ws.cell(row=2, column=3).fill = PatternFill(fill_type="solid", fgColor="00FF0000")
            ws.cell(row=2, column=3).number_format = "0.000"
            ws.cell(row=3, column=3).fill = PatternFill(fill_type="solid", fgColor="0000FF00")
            ws.cell(row=3, column=3).number_format = "0.00"
            wb.save(workbook_path)

            removed = excel_writer.deduplicate_pricing_rows(workbook_path)
            self.assertEqual(removed, 0)

            wb = load_workbook(workbook_path)
            ws = wb["Pricing"]

            self.assertEqual(ws.cell(row=2, column=2).value, "BankA")
            self.assertEqual(ws.cell(row=2, column=3).fill.fgColor.rgb, "0000FF00")
            self.assertEqual(ws.cell(row=2, column=3).number_format, "0.00")

            self.assertEqual(ws.cell(row=3, column=2).value, "BankB")
            self.assertEqual(ws.cell(row=3, column=3).fill.fgColor.rgb, "00FF0000")
            self.assertEqual(ws.cell(row=3, column=3).number_format, "0.000")


if __name__ == "__main__":
    unittest.main()
