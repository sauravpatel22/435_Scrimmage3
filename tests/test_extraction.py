"""
test_extraction.py

End-to-end coverage of spreadsheet_processor.extract_dataset: running the
full pipeline against each test_data/ fixture folder and checking that the
"Missing / Invalid Data Should Not Stop the Program" design requirement
actually holds.
"""

import tempfile
import unittest
from pathlib import Path

from models import OrderingMethod, RunConfig, Selector, WarningLevel
from spreadsheet_processor import extract_dataset

TEST_DATA_DIR = Path(__file__).resolve().parent.parent / "test_data"


def _config(subfolder: str, output_folder: Path, **overrides) -> RunConfig:
    defaults = dict(
        input_folder=TEST_DATA_DIR / subfolder,
        output_folder=output_folder,
        selectors=[Selector.from_input("TEMP"), Selector.from_input("HUMIDITY")],
        ordering_method=OrderingMethod.FILENAME,
    )
    defaults.update(overrides)
    return RunConfig(**defaults)


class TestExtractDatasetValid(unittest.TestCase):
    def test_all_files_and_series_are_extracted_with_no_warnings(self):
        with tempfile.TemporaryDirectory() as tmp:
            result = extract_dataset(_config("valid", Path(tmp)))

            self.assertEqual(len(result.observations), 20)  # 10 files x 2 series
            self.assertEqual(result.warnings, [])
            self.assertEqual(len(result.observations_for("TEMP")), 10)
            self.assertEqual(len(result.observations_for("HUMIDITY")), 10)

    def test_values_match_known_fixture_contents(self):
        with tempfile.TemporaryDirectory() as tmp:
            result = extract_dataset(_config("valid", Path(tmp)))
            temps = {
                obs.source_file.name: obs.value
                for obs in result.observations
                if obs.series_name == "TEMP"
            }
            self.assertEqual(
                temps,
                {
                    "weather_01.xlsx": 72.0,
                    "weather_02.xlsx": 74.0,
                    "weather_03.xlsx": 71.0,
                    "weather_04.xlsx": 68.0,
                    "weather_05.xlsx": 70.0,
                    "weather_06.xlsx": 73.0,
                    "weather_07.xlsx": 76.0,
                    "weather_08.xlsx": 69.0,
                    "weather_09.xlsx": 75.0,
                    "weather_10.xlsx": 72.0,
                },
            )


class TestExtractDatasetMissingValues(unittest.TestCase):
    def test_missing_label_and_missing_value_are_warned_not_fatal(self):
        with tempfile.TemporaryDirectory() as tmp:
            result = extract_dataset(_config("missing_values", Path(tmp)))

            # Of 10 files, only weather_missing_01 (HUMIDITY blank) and
            # weather_missing_02 (TEMP label absent) have a real problem;
            # the other 8 are fully valid, so 18 of the 20 possible
            # (file, series) pairs succeed and exactly 2 warn.
            self.assertEqual(len(result.observations), 18)
            self.assertEqual(len(result.observations_for("TEMP")), 9)
            self.assertEqual(len(result.observations_for("HUMIDITY")), 9)
            self.assertEqual(len(result.warnings), 2)

            messages = [str(w) for w in result.warnings]
            self.assertTrue(any("HUMIDITY" in m and "missing" in m for m in messages))
            self.assertTrue(any("TEMP" in m and "not found" in m for m in messages))


class TestExtractDatasetInvalidData(unittest.TestCase):
    def test_non_numeric_and_blank_cells_are_warned_not_fatal(self):
        with tempfile.TemporaryDirectory() as tmp:
            result = extract_dataset(_config("invalid_data", Path(tmp)))

            # 10 files, each breaking TEMP and/or HUMIDITY in a different
            # way (boolean, blank, whitespace, "N/A", stray text, a
            # formula with no cached value, etc.). Only invalid_humidity's
            # TEMP and 7 files' HUMIDITY are clean; nothing raises.
            self.assertEqual(len(result.observations), 8)
            self.assertEqual(len(result.observations_for("TEMP")), 1)
            self.assertEqual(len(result.observations_for("HUMIDITY")), 7)
            self.assertEqual(len(result.warnings), 12)

            invalid_warning = next(
                w for w in result.warnings if w.source_file.name == "invalid_text.xlsx"
            )
            self.assertEqual(invalid_warning.level, WarningLevel.ERROR)

            # bool is a subclass of int in Python but must still be
            # rejected as a numeric value (see normalize_numeric_value).
            boolean_warning = next(
                w for w in result.warnings if w.source_file.name == "boolean_value.xlsx"
            )
            self.assertEqual(boolean_warning.level, WarningLevel.ERROR)


class TestExtractDatasetShiftedLabels(unittest.TestCase):
    def test_shifted_tables_still_resolve_correctly(self):
        with tempfile.TemporaryDirectory() as tmp:
            result = extract_dataset(_config("shifted_labels", Path(tmp)))

            # 10 files, each with a genuinely different table position or
            # row order (different starting cell, extra rows, reversed
            # label order); every one must still resolve via label search.
            self.assertEqual(len(result.observations), 20)  # 10 files x 2 series
            self.assertEqual(result.warnings, [])


class TestExtractDatasetMixedStructure(unittest.TestCase):
    def test_extra_columns_and_reordered_rows_do_not_break_extraction(self):
        with tempfile.TemporaryDirectory() as tmp:
            result = extract_dataset(_config("mixed_structure", Path(tmp)))

            # 10 files, each with different harmless structural noise
            # (extra rows/columns, a title row, mixed label case, stray
            # whitespace, a blank row, a leading ID column).
            self.assertEqual(len(result.observations), 20)  # 10 files x 2 series
            self.assertEqual(result.warnings, [])


class TestExtractDatasetEmptyFolder(unittest.TestCase):
    def test_empty_input_folder_produces_an_actionable_warning(self):
        with tempfile.TemporaryDirectory() as empty_input, tempfile.TemporaryDirectory() as output:
            config = _config("valid", Path(output), input_folder=Path(empty_input))
            result = extract_dataset(config)

            self.assertEqual(result.observations, [])
            self.assertEqual(len(result.warnings), 1)
            self.assertEqual(result.warnings[0].level, WarningLevel.ERROR)
            self.assertIn("no .xlsx files found", result.warnings[0].message)


if __name__ == "__main__":
    unittest.main()
