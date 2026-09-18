"""
test_extract_and_chart.py - whole runs, against every fixture folder.

A bad file, a missing label or an unreadable value never stops the run, and
whatever did come out lines up correctly on the chart.
"""

import tempfile
import unittest
from pathlib import Path

import openpyxl

import chart
import sheets

DATA = Path(__file__).resolve().parent.parent / "test_data"


def _run(folder, series=("TEMP", "HUMIDITY"), **kwargs):
    return sheets.extract(DATA / folder, list(series), **kwargs)


class TestValidFolder(unittest.TestCase):
    def test_every_file_and_series_comes_through(self):
        result, _ = _run("valid")
        self.assertEqual(len(result.points), 20)
        self.assertEqual(result.warnings, [])
        self.assertEqual(result.files_scanned, 10)
        self.assertEqual(result.files_with_data, 10)

    def test_values_match_the_fixtures(self):
        result, _ = _run("valid", series=("TEMP",))
        temps = {p.file.name: p.value for p in result.points}
        self.assertEqual(temps["weather_01.xlsx"], 72.0)
        self.assertEqual(temps["weather_04.xlsx"], 68.0)
        self.assertEqual(len(temps), 10)


class TestAwkwardFolders(unittest.TestCase):
    def test_missing_values_are_warnings_not_crashes(self):
        result, _ = _run("missing_values")
        self.assertEqual(len(result.points), 18)
        self.assertEqual(len(result.warnings), 2)
        self.assertEqual(result.files_scanned, 10)

    def test_invalid_values_never_become_numbers(self):
        result, _ = _run("invalid_data")
        self.assertTrue(all(isinstance(p.value, float) for p in result.points))
        self.assertTrue(result.warnings)
        self.assertEqual(result.files_scanned, 10)

    def test_shifted_tables_still_resolve(self):
        result, _ = _run("shifted_labels")
        self.assertEqual(len(result.points), 20)
        self.assertEqual(result.warnings, [])

    def test_structural_noise_is_ignored(self):
        result, _ = _run("mixed_structure")
        self.assertEqual(len(result.points), 20)
        self.assertEqual(result.warnings, [])

    def test_unreadable_file_is_counted_and_skipped(self):
        with tempfile.TemporaryDirectory() as tmp:
            folder = Path(tmp)
            workbook = openpyxl.Workbook()
            workbook.active.append(["TEMP", 70])
            workbook.save(folder / "good.xlsx")
            (folder / "bad.xlsx").write_bytes(b"not a workbook")

            result, _ = sheets.extract(folder, ["TEMP"])
            self.assertEqual(len(result.points), 1)
            self.assertEqual((result.files_scanned, result.files_failed), (2, 1))

    def test_empty_folder_says_so(self):
        with tempfile.TemporaryDirectory() as tmp:
            result, _ = sheets.extract(Path(tmp), ["TEMP"])
            self.assertEqual(result.points, [])
            self.assertIn("no .xlsx files", result.warnings[0])


class TestGridFolder(unittest.TestCase):
    def test_row_by_column_series(self):
        result, _ = _run("grid", series=("Bob:Exam",))
        self.assertEqual([p.value for p in result.points], [68.0, 71.0, 74.0])

    def test_wildcard_gives_one_series_per_student(self):
        result, selections = _run("grid", series=("*:Exam",))
        self.assertEqual([s.name for s in selections], ["Alice", "Bob", "Carol"])
        self.assertEqual(len(result.points), 9)


class TestChart(unittest.TestCase):
    def test_x_axis_keeps_file_order_when_a_series_has_gaps(self):
        result, _ = _run("missing_values")
        _, labels = chart.x_axis(result)
        self.assertEqual(labels, sorted(labels))
        self.assertEqual(len(labels), 10)

    def test_missing_values_stay_as_gaps(self):
        result, _ = _run("missing_values")
        keys, _ = chart.x_axis(result)
        humidity = chart.align(result, "HUMIDITY", keys)
        self.assertEqual(len(humidity), len(keys))
        self.assertIn(None, humidity)

    def test_files_sharing_an_order_key_keep_separate_x_positions(self):
        result, _ = _run("valid", series=("TEMP",), ordering="custom",
                         custom_order=["weather_03.xlsx", "weather_01.xlsx"])
        keys, labels = chart.x_axis(result)
        self.assertEqual(len(keys), 10)
        self.assertEqual(labels[:2], ["weather_03.xlsx", "weather_01.xlsx"])
        self.assertEqual(len(set(labels)), 10)

    def test_two_files_with_the_same_date_both_survive(self):
        with tempfile.TemporaryDirectory() as tmp:
            folder = Path(tmp)
            for name in ("morning_2026-01-05.xlsx", "evening_2026-01-05.xlsx"):
                workbook = openpyxl.Workbook()
                workbook.active.append(["TEMP", 70])
                workbook.save(folder / name)

            result, _ = sheets.extract(folder, ["TEMP"], ordering="date_in_filename")
            self.assertEqual(len(result.points), 2)
            self.assertEqual(len(chart.x_axis(result)[0]), 2)

    def test_draws_each_chart_type(self):
        result, _ = _run("valid")
        with tempfile.TemporaryDirectory() as tmp:
            for kind in chart.CHART_TYPES:
                path = chart.draw(result, Path(tmp) / f"{kind}.png", kind=kind)
                self.assertTrue(path.is_file() and path.stat().st_size > 0)

    def test_no_data_means_no_chart_file(self):
        result, _ = _run("valid", series=("NOT_A_LABEL",))
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "empty.png"
            self.assertIsNone(chart.draw(result, path))
            self.assertFalse(path.exists())

    def test_right_axis_series_is_still_drawn(self):
        result, _ = _run("valid")
        with tempfile.TemporaryDirectory() as tmp:
            path = chart.draw(result, Path(tmp) / "two.png", right_axis=["HUMIDITY"])
            self.assertTrue(path.is_file())


if __name__ == "__main__":
    unittest.main()
