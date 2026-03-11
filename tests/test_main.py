import unittest
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


if __name__ == "__main__":
    unittest.main()
