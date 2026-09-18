"""
test_sheets.py - reading, selecting and ordering.

Run against the fixtures in test_data/, plus workbooks built on the fly where
a fixture would only ever be used once.
"""

import tempfile
import unittest
from datetime import date
from pathlib import Path

import openpyxl

import sheets

DATA = Path(__file__).resolve().parent.parent / "test_data"


def _book(path, rows, title="Sheet1"):
    workbook = openpyxl.Workbook()
    worksheet = workbook.active
    worksheet.title = title
    for row in rows:
        worksheet.append(list(row))
    workbook.save(path)
    return path


def _sheet(rows):
    tmp = Path(tempfile.mkdtemp()) / "t.xlsx"
    _book(tmp, rows)
    sheet, error = sheets.open_sheet(tmp)
    assert error is None, error
    return sheet


class TestFindFiles(unittest.TestCase):
    def test_lists_xlsx_sorted_by_name(self):
        names = [p.name for p in sheets.find_files(DATA / "valid")]
        self.assertEqual(names, [f"weather_{i:02d}.xlsx" for i in range(1, 11)])

    def test_skips_excel_lock_files(self):
        with tempfile.TemporaryDirectory() as tmp:
            folder = Path(tmp)
            _book(folder / "real.xlsx", [("TEMP", 1)])
            (folder / "~$real.xlsx").write_bytes(b"not a workbook")
            self.assertEqual([p.name for p in sheets.find_files(folder)], ["real.xlsx"])

    def test_missing_folder_raises(self):
        with self.assertRaises(FileNotFoundError):
            sheets.find_files(DATA / "nope")


class TestOpenSheet(unittest.TestCase):
    def test_unreadable_file_reports_instead_of_raising(self):
        with tempfile.TemporaryDirectory() as tmp:
            bad = Path(tmp) / "broken.xlsx"
            bad.write_bytes(b"definitely not a workbook")
            sheet, error = sheets.open_sheet(bad)
            self.assertIsNone(sheet)
            self.assertIn("could not open", error)

    def test_named_sheet_that_is_absent_is_an_error(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = _book(Path(tmp) / "a.xlsx", [("TEMP", 1)], title="Data")
            self.assertIsNotNone(sheets.open_sheet(path, "Data")[0])
            self.assertIn("not found", sheets.open_sheet(path, "Missing")[1])


class TestSelectionPlanning(unittest.TestCase):
    def test_plain_word_is_a_label(self):
        picked, _ = sheets.plan_selections(["TEMP"], None)
        self.assertEqual((picked[0].mode, picked[0].row_label), ("label", "TEMP"))

    def test_cell_address_is_a_cell(self):
        picked, _ = sheets.plan_selections(["b2"], None)
        self.assertEqual((picked[0].mode, picked[0].address), ("cell", "B2"))

    def test_label_that_looks_like_a_cell_is_still_a_label(self):
        sheet = _sheet([("Ticker", "Close"), ("T5", 12.5), ("CO2", 400)])
        picked, _ = sheets.plan_selections(["T5", "CO2"], sheet)
        self.assertEqual([s.mode for s in picked], ["label", "label"])
        value, _, error = sheets.read_selection(sheet, picked[0])
        self.assertIsNone(error)
        self.assertEqual(value, 12.5)

    def test_prefixes_force_the_interpretation(self):
        sheet = _sheet([("T5", 12.5)])
        forced_cell, _ = sheets.plan_selections(["cell:T5"], sheet)
        forced_label, _ = sheets.plan_selections(["label:B99"], sheet)
        self.assertEqual(forced_cell[0].mode, "cell")
        self.assertEqual(forced_label[0].mode, "label")

    def test_rename_sets_the_legend_name(self):
        picked, _ = sheets.plan_selections(["TEMP=Outside temp"], None)
        self.assertEqual(picked[0].name, "Outside temp")
        self.assertEqual(picked[0].row_label, "TEMP")

    def test_row_and_column_selection(self):
        picked, _ = sheets.plan_selections(["Bob:Exam"], None)
        self.assertEqual(picked[0].mode, "grid")
        self.assertEqual((picked[0].row_label, picked[0].col_label), ("Bob", "Exam"))

    def test_wildcards_expand_to_real_series(self):
        sheet, _ = sheets.open_sheet(DATA / "grid" / "session_01.xlsx")
        by_row, _ = sheets.plan_selections(["Bob:*"], sheet)
        by_col, _ = sheets.plan_selections(["*:Exam"], sheet)
        self.assertEqual([s.name for s in by_row], ["Bob - Quiz", "Bob - Exam", "Bob - Project"])
        self.assertEqual([s.name for s in by_col], ["Alice", "Bob", "Carol"])


class TestReadSelection(unittest.TestCase):
    def test_label_takes_the_value_to_its_right(self):
        sheet = _sheet([("Property", "Value"), ("TEMP", 72)])
        value, cell, error = sheets.read_selection(sheet, sheets.Selection("TEMP", "label", row_label="TEMP"))
        self.assertEqual((value, cell, error), (72, "B2", None))

    def test_missing_label_is_reported(self):
        sheet = _sheet([("TEMP", 72)])
        _, _, error = sheets.read_selection(sheet, sheets.Selection("X", "label", row_label="X"))
        self.assertIn("not found", error)

    def test_duplicate_label_warns_but_still_returns_a_value(self):
        sheet = _sheet([("TEMP", 1), ("TEMP", 2)])
        value, _, error = sheets.read_selection(sheet, sheets.Selection("TEMP", "label", row_label="TEMP"))
        self.assertEqual(value, 1)
        self.assertIn("appears 2 times", error)

    def test_grid_reads_the_intersection(self):
        sheet, _ = sheets.open_sheet(DATA / "grid" / "session_01.xlsx")
        selection = sheets.Selection("Bob - Exam", "grid", row_label="Bob", col_label="Exam")
        value, cell, error = sheets.read_selection(sheet, selection)
        self.assertEqual((value, cell, error), (68, "C3", None))

    def test_labels_are_found_wherever_the_table_sits(self):
        for path in sheets.find_files(DATA / "shifted_labels"):
            sheet, _ = sheets.open_sheet(path)
            value, _, error = sheets.read_selection(
                sheet, sheets.Selection("TEMP", "label", row_label="TEMP")
            )
            self.assertIsNone(error, f"{path.name}: {error}")
            self.assertIsInstance(value, (int, float))


class TestToNumber(unittest.TestCase):
    def test_accepts_numbers_and_numeric_text(self):
        self.assertEqual(sheets.to_number(72)[0], 72.0)
        self.assertEqual(sheets.to_number("  73.5 ")[0], 73.5)
        self.assertEqual(sheets.to_number(-4)[0], -4.0)

    def test_rejects_blank_text_and_booleans(self):
        for raw in (None, "", "   ", "N/A", True, False):
            value, error = sheets.to_number(raw)
            self.assertIsNone(value, f"{raw!r} should not be a number")
            self.assertTrue(error)


class TestOrdering(unittest.TestCase):
    def test_filename_order(self):
        key, label, warning = sheets.order_of(Path("a/weather_02.xlsx"), None, "filename")
        self.assertEqual((label, warning), ("weather_02.xlsx", None))
        self.assertEqual(key[1], "weather_02.xlsx")

    def test_date_in_filename(self):
        key, label, warning = sheets.order_of(Path("w_2026-01-15.xlsx"), None, "date_in_filename")
        self.assertEqual((label, warning), ("2026-01-15", None))
        self.assertEqual(key[1].date(), date(2026, 1, 15))

    def test_date_in_filename_falls_back_with_a_warning(self):
        _, label, warning = sheets.order_of(Path("nodate.xlsx"), None, "date_in_filename")
        self.assertEqual(label, "nodate.xlsx")
        self.assertIn("no date", warning)

    def test_date_in_sheet(self):
        sheet = _sheet([("DATE", date(2026, 3, 4)), ("TEMP", 70)])
        key, label, warning = sheets.order_of(Path("x.xlsx"), sheet, "date_in_sheet")
        self.assertEqual((label, warning), ("2026-03-04", None))
        self.assertEqual(key[0], 0)

    def test_custom_order_sorts_by_position_but_labels_by_filename(self):
        order = ["c.xlsx", "a.xlsx"]
        first, first_label, _ = sheets.order_of(Path("c.xlsx"), None, "custom", order)
        second, second_label, _ = sheets.order_of(Path("a.xlsx"), None, "custom", order)
        unlisted, unlisted_label, warning = sheets.order_of(Path("z.xlsx"), None, "custom", order)
        self.assertLess(first, second)
        self.assertLess(second, unlisted)
        self.assertEqual([first_label, second_label, unlisted_label], ["c.xlsx", "a.xlsx", "z.xlsx"])
        self.assertIn("custom order", warning)

    def test_mixed_orderings_still_sort_without_type_errors(self):
        dated, _, _ = sheets.order_of(Path("w_2026-01-15.xlsx"), None, "date_in_filename")
        named, _, _ = sheets.order_of(Path("nodate.xlsx"), None, "date_in_filename")
        sorted([dated, named])


if __name__ == "__main__":
    unittest.main()
