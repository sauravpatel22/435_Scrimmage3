"""
test_main.py - the command line layer.

main.run() does no asking, so it is driven here with the same objects
argparse would produce.
"""

import contextlib
import io
import tempfile
import unittest
from pathlib import Path
from unittest import mock

import main

DATA = Path(__file__).resolve().parent.parent / "test_data"


def _quiet():
    """Swallow the console output these tests trigger on purpose."""
    return contextlib.redirect_stdout(io.StringIO())


def _args(**overrides):
    args = main.build_parser().parse_args([])
    for key, value in overrides.items():
        setattr(args, key, value)
    return args


class TestRun(unittest.TestCase):
    def test_end_to_end_writes_one_png(self):
        with tempfile.TemporaryDirectory() as tmp:
            result, selections, drawn = main.run(
                _args(
                    input_folder=str(DATA / "valid"),
                    series=["TEMP", "HUMIDITY"],
                    order="filename",
                    chart="line",
                    output_folder=tmp,
                )
            )
            self.assertEqual(len(result.points), 20)
            self.assertEqual(len(selections), 2)
            self.assertTrue(drawn.is_file())
            self.assertEqual(len(list((Path(tmp) / "graphs").glob("*.png"))), 1)

    def test_renaming_shows_up_in_the_series_name(self):
        with tempfile.TemporaryDirectory() as tmp:
            result, _, _ = main.run(
                _args(
                    input_folder=str(DATA / "valid"),
                    series=["TEMP=Outside temp"],
                    order="filename",
                    chart="line",
                    output_folder=tmp,
                )
            )
            self.assertEqual(result.series_names, ["Outside temp"])

    def test_nothing_found_means_no_chart_and_a_failing_exit_code(self):
        with tempfile.TemporaryDirectory() as tmp:
            with _quiet():
                code = main.main([
                    "--input-folder", str(DATA / "valid"),
                    "--series", "NOPE",
                    "--order", "filename",
                    "--chart", "line",
                    "--output-folder", tmp,
                ])
            self.assertEqual(code, 1)
            self.assertFalse(list((Path(tmp) / "graphs").glob("*.png")))

    def test_successful_cli_run_exits_zero(self):
        with tempfile.TemporaryDirectory() as tmp:
            with _quiet():
                code = main.main([
                    "--input-folder", str(DATA / "grid"),
                    "--series", "Bob:Exam",
                    "--order", "filename",
                    "--chart", "bar",
                    "--output-folder", tmp,
                ])
            self.assertEqual(code, 0)


class TestAsFolder(unittest.TestCase):
    def test_plain_absolute_and_relative_paths(self):
        self.assertEqual(main.as_folder(str(DATA / "valid")), DATA / "valid")
        self.assertTrue(main.as_folder(str(DATA / "valid") + "/").is_dir())
        self.assertTrue(main.as_folder(f"  {DATA / 'valid'}  ").is_dir())

    def test_pasted_quotes_and_escaped_spaces(self):
        with tempfile.TemporaryDirectory() as tmp:
            spaced = Path(tmp) / "My Weather Data"
            spaced.mkdir()
            for text in (
                f'"{spaced}"',
                f"'{spaced}'",
                str(spaced).replace(" ", "\\ "),
            ):
                with self.subTest(text=text):
                    self.assertEqual(main.as_folder(text), spaced)

    def test_a_name_that_really_contains_a_quote_still_works(self):
        with tempfile.TemporaryDirectory() as tmp:
            odd = Path(tmp) / "'quoted'"
            odd.mkdir()
            self.assertEqual(main.as_folder(str(odd)), odd)


class TestCustomOrderParser(unittest.TestCase):
    def setUp(self):
        self.files = [Path(f"f{i}.xlsx") for i in range(1, 4)]
        self.parse = main.custom_order_parser(self.files)

    def test_good_input_maps_numbers_to_filenames(self):
        self.assertEqual(self.parse("3, 1, 2"), ["f3.xlsx", "f1.xlsx", "f2.xlsx"])

    def test_partial_order_is_allowed(self):
        self.assertEqual(self.parse("2"), ["f2.xlsx"])

    def test_every_bad_input_is_a_clear_message(self):
        for text, expected in [
            ("99", "not between"),
            ("abc", "not a number"),
            ("2, 2", "twice"),
            ("", "at least one"),
        ]:
            with self.subTest(text=text):
                with self.assertRaises(ValueError) as caught:
                    self.parse(text)
                self.assertIn(expected, str(caught.exception))

    def test_ask_reprompts_until_the_answer_is_valid(self):
        with mock.patch("builtins.input", side_effect=["99, 1", "abc", "3, 1, 2"]):
            with _quiet():
                answer = main._ask("> ", self.parse)
        self.assertEqual(answer, ["f3.xlsx", "f1.xlsx", "f2.xlsx"])


if __name__ == "__main__":
    unittest.main()
