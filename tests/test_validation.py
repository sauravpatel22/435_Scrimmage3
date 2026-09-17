"""
test_validation.py

Covers spreadsheet_processor.normalize_numeric_value (deciding whether a
raw cell value is usable) and the validation built into models.py's
dataclasses (RunConfig, Selector), which fail fast on a configuration that
could never produce a sensible run.
"""

import unittest
from pathlib import Path

from models import OrderingMethod, RunConfig, Selector, SelectorType
from spreadsheet_processor import normalize_numeric_value


class TestNormalizeNumericValue(unittest.TestCase):
    def test_none_is_missing(self):
        value, error = normalize_numeric_value(None)
        self.assertIsNone(value)
        self.assertEqual(error, "missing value")

    def test_blank_string_is_missing(self):
        value, error = normalize_numeric_value("   ")
        self.assertIsNone(value)
        self.assertEqual(error, "missing value")

    def test_int_is_converted_to_float(self):
        value, error = normalize_numeric_value(72)
        self.assertEqual(value, 72.0)
        self.assertIsNone(error)

    def test_float_passes_through(self):
        value, error = normalize_numeric_value(45.5)
        self.assertEqual(value, 45.5)
        self.assertIsNone(error)

    def test_numeric_string_is_parsed(self):
        value, error = normalize_numeric_value("45.5")
        self.assertEqual(value, 45.5)
        self.assertIsNone(error)

    def test_non_numeric_string_is_invalid(self):
        value, error = normalize_numeric_value("not a number")
        self.assertIsNone(value)
        self.assertIn("invalid", error)

    def test_boolean_is_rejected_even_though_bool_is_an_int_subclass(self):
        value, error = normalize_numeric_value(True)
        self.assertIsNone(value)
        self.assertIn("invalid", error)

    def test_unsupported_type_is_rejected(self):
        value, error = normalize_numeric_value(["not", "a", "number"])
        self.assertIsNone(value)
        self.assertIn("unsupported value type", error)


class TestRunConfigValidation(unittest.TestCase):
    def _selector(self):
        return Selector(raw="TEMP", type=SelectorType.LABEL, series_name="TEMP")

    def test_requires_at_least_one_selector(self):
        with self.assertRaises(ValueError):
            RunConfig(
                input_folder=Path("in"),
                output_folder=Path("out"),
                selectors=[],
                ordering_method=OrderingMethod.FILENAME,
            )

    def test_custom_ordering_requires_a_custom_order_list(self):
        with self.assertRaises(ValueError):
            RunConfig(
                input_folder=Path("in"),
                output_folder=Path("out"),
                selectors=[self._selector()],
                ordering_method=OrderingMethod.CUSTOM,
                custom_order=None,
            )

    def test_custom_ordering_with_a_list_is_accepted(self):
        config = RunConfig(
            input_folder=Path("in"),
            output_folder=Path("out"),
            selectors=[self._selector()],
            ordering_method=OrderingMethod.CUSTOM,
            custom_order=["b.xlsx", "a.xlsx"],
        )
        self.assertEqual(config.custom_order, ["b.xlsx", "a.xlsx"])

    def test_non_custom_ordering_does_not_require_a_custom_order_list(self):
        config = RunConfig(
            input_folder=Path("in"),
            output_folder=Path("out"),
            selectors=[self._selector()],
            ordering_method=OrderingMethod.FILENAME,
        )
        self.assertIsNone(config.custom_order)


if __name__ == "__main__":
    unittest.main()
