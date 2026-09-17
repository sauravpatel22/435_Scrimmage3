"""
test_extraction.py

End-to-end coverage of spreadsheet_processor.extract_dataset,
to_dataframe, and export_reports: running the full pipeline against each
test_data/ fixture folder and checking that the "Missing / Invalid Data
Should Not Stop the Program" and "Automatically Produce Reusable Output"
design requirements actually hold.
"""

import tempfile
import unittest
from pathlib import Path

from models import OrderingMethod, RunConfig, Selector, WarningLevel
from spreadsheet_processor import export_reports, extract_dataset, to_dataframe

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

            self.assertEqual(len(result.observations), 6)  # 3 files x 2 series
            self.assertEqual(result.warnings, [])
            self.assertEqual(result.series_summaries["TEMP"].files_found, 3)
            self.assertEqual(result.series_summaries["HUMIDITY"].files_found, 3)

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
                {"weather_01.xlsx": 72.0, "weather_02.xlsx": 74.0, "weather_03.xlsx": 71.0},
            )


class TestExtractDatasetMissingValues(unittest.TestCase):
    def test_missing_label_and_missing_value_are_warned_not_fatal(self):
        with tempfile.TemporaryDirectory() as tmp:
            result = extract_dataset(_config("missing_values", Path(tmp)))

            # TEMP present in weather_missing_01 only; HUMIDITY present in
            # weather_missing_02 only - one good observation per series.
            self.assertEqual(len(result.observations), 2)
            self.assertGreaterEqual(len(result.warnings), 2)

            messages = [str(w) for w in result.warnings]
            self.assertTrue(any("HUMIDITY" in m and "missing" in m for m in messages))
            self.assertTrue(any("TEMP" in m and "not found" in m for m in messages))


class TestExtractDatasetInvalidData(unittest.TestCase):
    def test_non_numeric_and_blank_cells_are_warned_not_fatal(self):
        with tempfile.TemporaryDirectory() as tmp:
            result = extract_dataset(_config("invalid_data", Path(tmp)))

            # Only HUMIDITY in invalid_text.xlsx is a clean value; both
            # cells in blank_cells.xlsx and TEMP in invalid_text.xlsx fail.
            self.assertEqual(len(result.observations), 1)
            self.assertEqual(result.observations[0].series_name, "HUMIDITY")
            self.assertEqual(result.observations[0].value, 45.0)

            invalid_warning = next(
                w for w in result.warnings if w.source_file.name == "invalid_text.xlsx"
            )
            self.assertEqual(invalid_warning.level, WarningLevel.ERROR)


class TestExtractDatasetShiftedLabels(unittest.TestCase):
    def test_shifted_tables_still_resolve_correctly(self):
        with tempfile.TemporaryDirectory() as tmp:
            result = extract_dataset(_config("shifted_labels", Path(tmp)))

            self.assertEqual(len(result.observations), 4)  # 2 files x 2 series
            self.assertEqual(result.warnings, [])


class TestExtractDatasetMixedStructure(unittest.TestCase):
    def test_extra_columns_and_reordered_rows_do_not_break_extraction(self):
        with tempfile.TemporaryDirectory() as tmp:
            result = extract_dataset(_config("mixed_structure", Path(tmp)))

            self.assertEqual(len(result.observations), 4)  # 2 files x 2 series
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


class TestToDataframe(unittest.TestCase):
    def test_dataframe_has_one_row_per_observation(self):
        with tempfile.TemporaryDirectory() as tmp:
            result = extract_dataset(_config("valid", Path(tmp)))
            frame = to_dataframe(result)

            self.assertEqual(len(frame), len(result.observations))
            self.assertListEqual(
                list(frame.columns),
                ["source_file", "sheet", "series", "cell", "order_key", "value"],
            )


class TestExportReports(unittest.TestCase):
    def test_export_writes_csv_warnings_and_summary_files(self):
        with tempfile.TemporaryDirectory() as tmp:
            output_folder = Path(tmp)
            config = _config("missing_values", output_folder)
            result = extract_dataset(config)

            paths = export_reports(result, config, run_id="unittest_run")

            self.assertTrue(paths["csv"].is_file())
            self.assertTrue(paths["warnings"].is_file())
            self.assertTrue(paths["summary"].is_file())

            csv_text = paths["csv"].read_text()
            # Header row plus one row per successful observation.
            self.assertEqual(len(csv_text.strip().splitlines()), 1 + len(result.observations))

            summary_text = paths["summary"].read_text()
            self.assertIn("Run ID: unittest_run", summary_text)
            self.assertIn("Total observations:", summary_text)

    def test_repeated_runs_do_not_overwrite_each_other(self):
        with tempfile.TemporaryDirectory() as tmp:
            output_folder = Path(tmp)
            config = _config("valid", output_folder)
            result = extract_dataset(config)

            first = export_reports(result, config, run_id="run_one")
            second = export_reports(result, config, run_id="run_two")

            self.assertNotEqual(first["csv"], second["csv"])
            self.assertTrue(first["csv"].is_file())
            self.assertTrue(second["csv"].is_file())


if __name__ == "__main__":
    unittest.main()
