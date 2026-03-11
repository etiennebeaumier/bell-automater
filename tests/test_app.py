import unittest
from unittest.mock import patch

import app


class FakeConfig:
    def __init__(self, data):
        self._data = dict(data)
        self.saved = False

    def get(self, key, default=None):
        return self._data.get(key, default)

    def __setitem__(self, key, value):
        self._data[key] = value

    def save(self):
        self.saved = True


class AppStartupModeTests(unittest.TestCase):
    @patch("main.process_many_pdfs")
    @patch("app._launch_gui")
    @patch("app._show_quick_run_dialog")
    @patch("app._load_available_years")
    @patch("app._collect_quick_run_context")
    @patch("app.AppConfig")
    def test_quick_run_processes_and_skips_gui(
        self,
        mock_app_config,
        mock_collect_quick_run_context,
        mock_load_available_years,
        mock_show_quick_run_dialog,
        mock_launch_gui,
        mock_process_many_pdfs,
    ):
        fake_cfg = FakeConfig({"avg_start_year": 2025, "avg_end_year": 2025})
        mock_app_config.return_value = fake_cfg

        context = app.QuickRunContext(
            master_file="/tmp/master.xlsx",
            pdf_dir="/tmp/pdfs",
            pdf_paths=["/tmp/pdfs/a.pdf", "/tmp/pdfs/b.pdf"],
        )
        mock_collect_quick_run_context.return_value = context
        mock_load_available_years.return_value = [2024, 2025, 2026]
        mock_show_quick_run_dialog.return_value = app.QuickRunSelection(
            action="run",
            avg_start_year=2024,
            avg_end_year=2026,
        )

        app.main()

        mock_launch_gui.assert_not_called()
        mock_process_many_pdfs.assert_called_once_with(
            ["/tmp/pdfs/a.pdf", "/tmp/pdfs/b.pdf"],
            "/tmp/master.xlsx",
            dry_run=False,
            avg_start_year=2024,
            avg_end_year=2026,
        )
        self.assertEqual(fake_cfg.get("avg_start_year"), 2024)
        self.assertEqual(fake_cfg.get("avg_end_year"), 2026)
        self.assertTrue(fake_cfg.saved)

    @patch("main.process_many_pdfs")
    @patch("app._launch_gui")
    @patch("app._collect_quick_run_context")
    @patch("app.AppConfig")
    def test_invalid_quick_context_opens_gui(
        self,
        mock_app_config,
        mock_collect_quick_run_context,
        mock_launch_gui,
        mock_process_many_pdfs,
    ):
        fake_cfg = FakeConfig({})
        mock_app_config.return_value = fake_cfg
        mock_collect_quick_run_context.return_value = None

        app.main()

        mock_launch_gui.assert_called_once_with(fake_cfg)
        mock_process_many_pdfs.assert_not_called()

    @patch("main.process_many_pdfs")
    @patch("app._launch_gui")
    @patch("app._show_quick_run_dialog")
    @patch("app._load_available_years")
    @patch("app._collect_quick_run_context")
    @patch("app.AppConfig")
    def test_open_gui_option_from_quick_prompt(
        self,
        mock_app_config,
        mock_collect_quick_run_context,
        mock_load_available_years,
        mock_show_quick_run_dialog,
        mock_launch_gui,
        mock_process_many_pdfs,
    ):
        fake_cfg = FakeConfig({"avg_start_year": 2025, "avg_end_year": 2025})
        mock_app_config.return_value = fake_cfg
        mock_collect_quick_run_context.return_value = app.QuickRunContext(
            master_file="/tmp/master.xlsx",
            pdf_dir="/tmp/pdfs",
            pdf_paths=["/tmp/pdfs/a.pdf"],
        )
        mock_load_available_years.return_value = [2025]
        mock_show_quick_run_dialog.return_value = app.QuickRunSelection(action="open_gui")

        app.main()

        mock_process_many_pdfs.assert_not_called()
        mock_launch_gui.assert_called_once_with(fake_cfg)
        self.assertFalse(fake_cfg.saved)

    @patch("main.detect_bank")
    @patch("os.path.isfile")
    @patch("os.listdir")
    @patch("os.path.isdir")
    def test_collect_parseable_pdfs_skips_unparseable_files(
        self,
        mock_isdir,
        mock_listdir,
        mock_isfile,
        mock_detect_bank,
    ):
        mock_isdir.return_value = True
        mock_listdir.return_value = ["good.pdf", "bad.pdf", "note.txt"]
        mock_isfile.side_effect = lambda p: p.endswith(".pdf")

        def detect_side_effect(path):
            if path.endswith("good.pdf"):
                return "td"
            raise ValueError("Unsupported PDF")

        mock_detect_bank.side_effect = detect_side_effect

        result = app._collect_parseable_pdfs("/tmp/pdfs")

        self.assertEqual(result, ["/tmp/pdfs/good.pdf"])

    @patch("app._collect_parseable_pdfs", return_value=[])
    @patch("app._is_workbook_ready", return_value=True)
    def test_collect_quick_run_context_returns_none_when_no_parseable_pdfs(
        self,
        _mock_is_workbook_ready,
        _mock_collect_parseable_pdfs,
    ):
        cfg = FakeConfig(
            {
                "master_file": "/tmp/master.xlsx",
                "pdf_source_dir": "/tmp/pdfs",
            }
        )

        context = app._collect_quick_run_context(cfg)
        self.assertIsNone(context)


if __name__ == "__main__":
    unittest.main()
