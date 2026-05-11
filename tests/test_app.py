import unittest
from unittest.mock import patch

import app


class FakeConfig:
    def __init__(self, data):
        self._data = dict(data)

    def get(self, key, default=None):
        return self._data.get(key, default)

    def __setitem__(self, key, value):
        self._data[key] = value

    def save(self):
        pass


class AppStartupModeTests(unittest.TestCase):
    @patch("main.process_many_pdfs")
    @patch("app._launch_gui")
    @patch("app._collect_parseable_pdfs", return_value=["/tmp/pdfs/a.pdf"])
    @patch("app._is_workbook_ready", return_value=True)
    @patch("app.AppConfig")
    def test_quick_run_processes_and_skips_gui(
        self,
        mock_app_config,
        _mock_is_workbook_ready,
        _mock_collect_parseable_pdfs,
        mock_launch_gui,
        mock_process_many_pdfs,
    ):
        fake_cfg = FakeConfig({"master_file": "/tmp/master.xlsx", "pdf_source_dir": "/tmp/pdfs"})
        mock_app_config.return_value = fake_cfg

        app.main()

        mock_launch_gui.assert_not_called()
        mock_process_many_pdfs.assert_called_once_with(
            ["/tmp/pdfs/a.pdf"],
            "/tmp/master.xlsx",
            dry_run=False,
        )

    @patch("main.process_many_pdfs")
    @patch("app._launch_gui")
    @patch("app._collect_parseable_pdfs", return_value=[])
    @patch("app._is_workbook_ready", return_value=True)
    @patch("app.AppConfig")
    def test_no_parseable_pdfs_opens_gui(
        self,
        mock_app_config,
        _mock_is_workbook_ready,
        _mock_collect_parseable_pdfs,
        mock_launch_gui,
        mock_process_many_pdfs,
    ):
        fake_cfg = FakeConfig({"master_file": "/tmp/master.xlsx", "pdf_source_dir": "/tmp/pdfs"})
        mock_app_config.return_value = fake_cfg

        app.main()

        mock_launch_gui.assert_called_once_with(fake_cfg)
        mock_process_many_pdfs.assert_not_called()

    @patch("main.process_many_pdfs")
    @patch("app._launch_gui")
    @patch("app._is_workbook_ready", return_value=False)
    @patch("app.AppConfig")
    def test_invalid_workbook_opens_gui(
        self,
        mock_app_config,
        _mock_is_workbook_ready,
        mock_launch_gui,
        mock_process_many_pdfs,
    ):
        fake_cfg = FakeConfig({"master_file": "/tmp/missing.xlsx"})
        mock_app_config.return_value = fake_cfg

        app.main()

        mock_launch_gui.assert_called_once_with(fake_cfg)
        mock_process_many_pdfs.assert_not_called()

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


if __name__ == "__main__":
    unittest.main()
