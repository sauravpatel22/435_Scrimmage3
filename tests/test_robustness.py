"""
test_robustness.py - the "someone is going to try to break this" tests.

Run against test_data/edge_cases, multi_sheet and broken, built by
test_data/make_edge_cases.py. The standard throughout: never raise, never
invent a number, never lose the other files because one file was bad.
"""

import tempfile
import unittest
import warnings
from pathlib import Path

import chart
import sheets

DATA = Path(__file__).resolve().parent.parent / "test_data"
EDGE, BROKEN, MULTI = DATA / "edge_cases", DATA / "broken", DATA / "multi_sheet"


def _run(folder, series, **kwargs):
    return sheets.extract(folder, list(series), **kwargs)


def _draw_all(result):
    """Every chart type must survive whatever the extraction produced."""
    with tempfile.TemporaryDirectory() as tmp:
        for kind in chart.CHART_TYPES:
            chart.draw(result, Path(tmp) / f"{kind}.png", kind=kind)


class TestFixturesExist(unittest.TestCase):
    def test_generator_output_is_present(self):
        for folder, least in ((EDGE, 25), (BROKEN, 4), (MULTI, 2)):
            with self.subTest(folder=folder.name):
                self.assertGreaterEqual(
                    len(list(folder.glob("*.xlsx"))), least,
                    "missing fixtures - run: python3 test_data/make_edge_cases.py",
                )


class TestHostileValues(unittest.TestCase):
    def test_unplottable_numbers_are_refused_not_drawn(self):
        result, _ = _run(EDGE, ["TEMP"])
        self.assertTrue(all(abs(p.value) < 1e300 for p in result.points))
        self.assertTrue(any("too large" in w for w in result.warnings))
        _draw_all(result)

    def test_infinity_and_nan_are_not_numbers(self):
        for raw in (float("inf"), float("-inf"), float("nan")):
            value, error = sheets.to_number(raw)
            self.assertIsNone(value)
            self.assertTrue(error)

    def test_text_that_is_really_a_number_still_counts(self):
        sheet, _ = sheets.open_sheet(EDGE / "numbers_as_text.xlsx")
        selection = sheets.Selection("TEMP", "label", row_label="TEMP")
        raw, _, _ = sheets.read_selection(sheet, selection)
        self.assertEqual(sheets.to_number(raw)[0], 72.0)

    def test_units_and_currency_are_refused_with_a_reason(self):
        result, _ = _run(EDGE, ["TEMP"])
        self.assertTrue(any("72F" in w for w in result.warnings))

    def test_zero_and_negatives_survive(self):
        result, _ = _run(EDGE, ["TEMP", "HUMIDITY"])
        values = [p.value for p in result.points]
        self.assertIn(-40.0, values)
        self.assertIn(0.0, values)


class TestHostileLabels(unittest.TestCase):
    def test_names_that_look_like_cell_addresses(self):
        result, _ = _run(EDGE, ["T5", "CO2", "Q1"])
        self.assertEqual(sorted(p.value for p in result.points), [12.5, 88.0, 400.0])

    def test_label_containing_a_colon(self):
        result, _ = _run(EDGE, ["Time: AM"])
        self.assertEqual([p.value for p in result.points], [6.0])

    def test_label_containing_an_equals_sign(self):
        result, _ = _run(EDGE, ["A=B"])
        self.assertEqual([p.value for p in result.points], [42.0])

    def test_label_containing_an_asterisk(self):
        result, _ = _run(EDGE, ["TEMP*"])
        self.assertEqual([p.value for p in result.points], [73.0])

    def test_case_padding_and_nonbreaking_spaces_all_match(self):
        result, _ = _run(EDGE, ["  temp  "])
        values = [p.value for p in result.points]
        self.assertIn(75.0, values)   # case_variation.xlsx
        self.assertIn(76.0, values)   # padded_label.xlsx
        self.assertIn(77.0, values)   # nonbreaking_space.xlsx

    def test_unicode_labels(self):
        result, _ = _run(EDGE, ["温度", "TEMPÉRATURE"])
        self.assertEqual(sorted(p.value for p in result.points), [21.0, 22.0])

    def test_duplicate_label_warns_and_takes_the_first(self):
        result, _ = _run(EDGE, ["TEMP"])
        self.assertTrue(any("appears 2 times" in w for w in result.warnings))
        self.assertNotIn(999.0, [p.value for p in result.points])

    def test_renaming_still_works_when_the_name_has_no_equals(self):
        result, selections = _run(EDGE, ["TEMP=Outside"])
        self.assertEqual(selections[0].name, "Outside")
        self.assertTrue(result.points)


class TestHostileShapes(unittest.TestCase):
    def test_label_with_nothing_to_its_right(self):
        result, _ = _run(EDGE, ["TEMP"])
        self.assertTrue(any("missing value" in w for w in result.warnings))

    def test_layouts_we_do_not_support_warn_rather_than_guess(self):
        for name in ("value_to_the_left.xlsx", "transposed.xlsx"):
            with self.subTest(file=name):
                sheet, _ = sheets.open_sheet(EDGE / name)
                raw, _, error = sheets.read_selection(
                    sheet, sheets.Selection("TEMP", "label", row_label="TEMP")
                )
                value, bad = sheets.to_number(raw)
                self.assertTrue(error or bad)
                self.assertIsNone(value)

    def test_empty_and_header_only_sheets(self):
        for name in ("empty_sheet.xlsx", "headers_only.xlsx"):
            with self.subTest(file=name):
                sheet, error = sheets.open_sheet(EDGE / name)
                self.assertIsNone(error)
                self.assertEqual(sheet.find("TEMP"), [])

    def test_wide_tall_hidden_and_merged_sheets_still_resolve(self):
        expected = {
            "very_wide.xlsx": 84.0,
            "very_tall.xlsx": 85.0,
            "hidden_row.xlsx": 86.0,
            "merged_cells.xlsx": 87.0,
            "blank_rows_between.xlsx": 83.0,
        }
        for name, value in expected.items():
            with self.subTest(file=name):
                sheet, _ = sheets.open_sheet(EDGE / name)
                raw, _, error = sheets.read_selection(
                    sheet, sheets.Selection("TEMP", "label", row_label="TEMP")
                )
                self.assertIsNone(error)
                self.assertEqual(sheets.to_number(raw)[0], value)

    def test_grid_with_headers_not_in_the_first_row(self):
        sheet, _ = sheets.open_sheet(EDGE / "grid_shifted_header.xlsx")
        raw, _, error = sheets.read_selection(
            sheet, sheets.Selection("Bob - Exam", "grid", row_label="Bob", col_label="Exam")
        )
        self.assertIsNone(error)
        self.assertEqual(raw, 90)

    def test_wildcard_finds_rows_that_only_exist_in_a_later_file(self):
        result, selections = _run(DATA / "grid", ["*:Exam"])
        self.assertEqual([s.name for s in selections], ["Alice", "Bob", "Carol"])
        self.assertEqual(len(result.points), 9)


class TestHostileFilesAndSheets(unittest.TestCase):
    def test_unreadable_files_are_counted_and_the_good_one_survives(self):
        result, _ = _run(BROKEN, ["TEMP"])
        self.assertEqual(result.files_scanned, 4)
        self.assertEqual(result.files_failed, 3)
        self.assertEqual([p.value for p in result.points], [95.0])
        _draw_all(result)

    def test_data_behind_a_cover_sheet_is_found(self):
        result, _ = _run(MULTI, ["TEMP", "HUMIDITY"])
        self.assertEqual(len(result.points), 4)
        self.assertTrue(any("Readings" in n for n in result.notes))

    def test_naming_the_sheet_explicitly_works(self):
        result, _ = _run(MULTI, ["TEMP"], sheet_name="Readings")
        self.assertEqual(sorted(p.value for p in result.points), [93.0, 94.0])

    def test_naming_a_sheet_that_is_not_there_is_reported(self):
        result, _ = _run(MULTI, ["TEMP"], sheet_name="Nope")
        self.assertEqual(result.files_failed, 2)
        self.assertTrue(all("not found" in w for w in result.warnings))


class TestHostileOrdering(unittest.TestCase):
    def test_every_ordering_survives_the_edge_case_folder(self):
        for ordering in sheets.ORDERINGS:
            with self.subTest(ordering=ordering):
                custom = ["date_real.xlsx"] if ordering == "custom" else None
                result, _ = _run(EDGE, ["TEMP"], ordering=ordering, custom_order=custom)
                self.assertTrue(result.points)
                _draw_all(result)

    def test_dates_in_every_form_are_ordered_or_explained(self):
        result, _ = _run(EDGE, ["TEMP"], ordering="date_in_sheet")
        labels = {p.file.name: p.x_label for p in result.points}
        self.assertEqual(labels["date_real.xlsx"], "2026-05-04")
        self.assertTrue(labels["date_datetime.xlsx"].startswith("2026-05-05"))
        self.assertTrue(any("not a date" in w for w in result.warnings))
        self.assertTrue(any("no DATE label" in w for w in result.warnings))

    def test_mixed_date_and_fallback_keys_still_sort(self):
        result, _ = _run(EDGE, ["TEMP"], ordering="date_in_sheet")
        keys, labels = chart.x_axis(result)
        self.assertEqual(len(keys), len(labels))
        self.assertEqual(keys, sorted(keys))


class TestHostileInput(unittest.TestCase):
    def test_nonsense_selections_are_handled_not_crashed(self):
        for series in ([""], [":"], ["*"], ["*:*"], ["   "], ["cell:ZZ999"], ["label:B2"]):
            with self.subTest(series=series):
                result, _ = _run(EDGE, series)
                self.assertIsInstance(result.warnings + result.notes, list)
                _draw_all(result)

    def test_a_folder_of_everything_at_once_never_raises(self):
        with warnings.catch_warnings():
            warnings.simplefilter("ignore")
            result, _ = _run(EDGE, ["TEMP", "HUMIDITY", "Alice:Exam", "B2", "温度"])
        self.assertTrue(result.points)
        self.assertEqual(result.files_failed, 0)
        _draw_all(result)


if __name__ == "__main__":
    unittest.main()
