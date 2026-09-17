"""
test_file_loading.py

Covers spreadsheet_processor.discover_files, load_workbook_safe, and
select_worksheet: finding the right .xlsx files in a folder and opening
them safely, before any label/value logic is involved.
"""

import tempfile
import unittest
from pathlib import Path

from spreadsheet_processor import discover_files, load_workbook_safe, select_worksheet

TEST_DATA_DIR = Path(__file__).resolve().parent.parent / "test_data"


class TestDiscoverFiles(unittest.TestCase):
    def test_finds_all_valid_fixtures_sorted_by_name(self):
        files = discover_files(TEST_DATA_DIR / "valid")
        self.assertEqual(
            [f.name for f in files],
            ["weather_01.xlsx", "weather_02.xlsx", "weather_03.xlsx"],
        )

    def test_missing_folder_raises(self):
        with self.assertRaises(FileNotFoundError):
            discover_files(TEST_DATA_DIR / "does_not_exist")

    def test_empty_folder_returns_empty_list(self):
        with tempfile.TemporaryDirectory() as tmp:
            self.assertEqual(discover_files(Path(tmp)), [])

    def test_excel_lock_files_are_excluded(self):
        with tempfile.TemporaryDirectory() as tmp:
            tmp_path = Path(tmp)
            (tmp_path / "data.xlsx").write_bytes(b"")
            (tmp_path / "~$data.xlsx").write_bytes(b"")

            files = discover_files(tmp_path)

            self.assertEqual([f.name for f in files], ["data.xlsx"])

    def test_non_xlsx_files_are_ignored(self):
        with tempfile.TemporaryDirectory() as tmp:
            tmp_path = Path(tmp)
            (tmp_path / "data.xlsx").write_bytes(b"")
            (tmp_path / "notes.txt").write_bytes(b"")

            files = discover_files(tmp_path)

            self.assertEqual([f.name for f in files], ["data.xlsx"])

    def test_recursive_discovery_finds_nested_files(self):
        with tempfile.TemporaryDirectory() as tmp:
            tmp_path = Path(tmp)
            nested = tmp_path / "subfolder"
            nested.mkdir()
            (tmp_path / "top.xlsx").write_bytes(b"")
            (nested / "nested.xlsx").write_bytes(b"")

            non_recursive = discover_files(tmp_path, recursive=False)
            recursive = discover_files(tmp_path, recursive=True)

            self.assertEqual([f.name for f in non_recursive], ["top.xlsx"])
            self.assertEqual(
                sorted(f.name for f in recursive), ["nested.xlsx", "top.xlsx"]
            )


class TestLoadWorkbookSafe(unittest.TestCase):
    def test_valid_file_opens_without_error(self):
        workbook, error = load_workbook_safe(TEST_DATA_DIR / "valid" / "weather_01.xlsx")
        self.addCleanup(workbook.close)
        self.assertIsNotNone(workbook)
        self.assertIsNone(error)

    def test_missing_file_reports_error_instead_of_raising(self):
        workbook, error = load_workbook_safe(TEST_DATA_DIR / "valid" / "does_not_exist.xlsx")
        self.assertIsNone(workbook)
        self.assertIsNotNone(error)

    def test_non_xlsx_file_reports_error_instead_of_raising(self):
        with tempfile.TemporaryDirectory() as tmp:
            bad_file = Path(tmp) / "not_really_excel.xlsx"
            bad_file.write_text("this is not a workbook")

            workbook, error = load_workbook_safe(bad_file)

            self.assertIsNone(workbook)
            self.assertIsNotNone(error)


class TestSelectWorksheet(unittest.TestCase):
    def setUp(self):
        self.workbook, _ = load_workbook_safe(TEST_DATA_DIR / "valid" / "weather_01.xlsx")
        self.addCleanup(self.workbook.close)

    def test_no_sheet_name_uses_first_sheet(self):
        worksheet, error = select_worksheet(self.workbook, sheet_name=None)
        self.assertIsNone(error)
        self.assertEqual(worksheet.title, "Weather")

    def test_matching_sheet_name_is_selected(self):
        worksheet, error = select_worksheet(self.workbook, sheet_name="Weather")
        self.assertIsNone(error)
        self.assertEqual(worksheet.title, "Weather")

    def test_unknown_sheet_name_reports_error_instead_of_raising(self):
        worksheet, error = select_worksheet(self.workbook, sheet_name="DoesNotExist")
        self.assertIsNone(worksheet)
        self.assertIsNotNone(error)


if __name__ == "__main__":
    unittest.main()
