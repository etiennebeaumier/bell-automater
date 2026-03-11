import unittest
import sys
from unittest.mock import patch

import main


class ProcessManyPdfsTests(unittest.TestCase):
    @patch("excel_writer.update_charts")
    @patch("excel_writer.deduplicate_pricing_rows", return_value=2)
    @patch("main.process_pdf")
    def test_non_dry_run_post_processing_runs_once(
        self,
        mock_process_pdf,
        mock_deduplicate_pricing_rows,
        mock_update_charts,
    ):
        mock_process_pdf.return_value = {"bank": "TD"}

        ok = main.process_many_pdfs(
            ["a.pdf", "b.pdf"],
            master_file="/tmp/master.xlsx",
            dry_run=False,
        )

        self.assertTrue(ok)
        self.assertEqual(mock_process_pdf.call_count, 2)
        mock_deduplicate_pricing_rows.assert_called_once_with("/tmp/master.xlsx")
        mock_update_charts.assert_called_once_with("/tmp/master.xlsx")

    @patch("excel_writer.update_charts")
    @patch("excel_writer.deduplicate_pricing_rows")
    @patch("main.process_pdf")
    def test_dry_run_never_calls_post_processing(
        self,
        mock_process_pdf,
        mock_deduplicate_pricing_rows,
        mock_update_charts,
    ):
        mock_process_pdf.return_value = {"bank": "TD"}

        ok = main.process_many_pdfs(
            ["a.pdf", "b.pdf"],
            master_file="/tmp/master.xlsx",
            dry_run=True,
        )

        self.assertTrue(ok)
        self.assertEqual(mock_process_pdf.call_count, 2)
        mock_deduplicate_pricing_rows.assert_not_called()
        mock_update_charts.assert_not_called()

    @patch("excel_writer.update_charts")
    @patch("excel_writer.deduplicate_pricing_rows", return_value=1)
    @patch("main.process_pdf")
    def test_non_dry_run_with_year_bounds_passes_to_update_charts(
        self,
        mock_process_pdf,
        mock_deduplicate_pricing_rows,
        mock_update_charts,
    ):
        mock_process_pdf.return_value = {"bank": "TD"}

        ok = main.process_many_pdfs(
            ["a.pdf"],
            master_file="/tmp/master.xlsx",
            dry_run=False,
            avg_start_year=2024,
            avg_end_year=2025,
        )

        self.assertTrue(ok)
        self.assertEqual(mock_process_pdf.call_count, 1)
        mock_deduplicate_pricing_rows.assert_called_once_with("/tmp/master.xlsx")
        mock_update_charts.assert_called_once_with(
            "/tmp/master.xlsx",
            avg_start_year=2024,
            avg_end_year=2025,
        )


class MainCliTests(unittest.TestCase):
    @patch("main.load_env_file")
    def test_removed_fetch_flag_is_rejected(self, mock_load_env):
        mock_load_env.return_value = None
        with patch.object(sys, "argv", ["main.py", "--fetch"]):
            with self.assertRaises(SystemExit) as ctx:
                main.main()
        self.assertNotEqual(ctx.exception.code, 0)

    @patch("main.process_many_pdfs", return_value=True)
    @patch("main.run_preflight", return_value=True)
    @patch("main.load_env_file")
    def test_cli_pdf_mode_runs_processing(
        self,
        mock_load_env,
        mock_run_preflight,
        mock_process_many_pdfs,
    ):
        mock_load_env.return_value = None
        with patch.object(
            sys,
            "argv",
            ["main.py", "--pdf", "a.pdf", "--master", "/tmp/master.xlsx"],
        ):
            main.main()

        mock_run_preflight.assert_called_once_with(
            master_file="/tmp/master.xlsx",
            require_workbook=True,
            verbose=False,
        )
        mock_process_many_pdfs.assert_called_once_with(
            ["a.pdf"],
            "/tmp/master.xlsx",
            dry_run=False,
        )

    @patch("main.process_many_pdfs", return_value=True)
    @patch("main.run_preflight", return_value=True)
    @patch("main.load_env_file")
    @patch("os.listdir", return_value=["a.pdf", "ignore.txt"])
    @patch("os.path.isdir", return_value=True)
    def test_cli_dir_mode_runs_processing(
        self,
        _mock_isdir,
        _mock_listdir,
        mock_load_env,
        mock_run_preflight,
        mock_process_many_pdfs,
    ):
        mock_load_env.return_value = None
        with patch.object(
            sys,
            "argv",
            ["main.py", "--dir", "/tmp/pdfs", "--master", "/tmp/master.xlsx"],
        ):
            main.main()

        mock_run_preflight.assert_called_once_with(
            master_file="/tmp/master.xlsx",
            require_workbook=True,
            verbose=False,
        )
        mock_process_many_pdfs.assert_called_once_with(
            ["/tmp/pdfs/a.pdf"],
            "/tmp/master.xlsx",
            dry_run=False,
        )

    @patch("main.process_many_pdfs", return_value=True)
    @patch("main.run_preflight", return_value=True)
    @patch("main.load_env_file")
    def test_cli_check_mode_without_operation_only_runs_preflight(
        self,
        mock_load_env,
        mock_run_preflight,
        mock_process_many_pdfs,
    ):
        mock_load_env.return_value = None
        with patch.object(
            sys,
            "argv",
            ["main.py", "--check", "--master", "/tmp/master.xlsx"],
        ):
            main.main()

        mock_run_preflight.assert_called_once_with(
            master_file="/tmp/master.xlsx",
            require_workbook=True,
            verbose=True,
        )
        mock_process_many_pdfs.assert_not_called()


if __name__ == "__main__":
    unittest.main()
