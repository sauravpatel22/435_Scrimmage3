"""
main.py - the command line front end.

Anything passed as a flag is not asked about, so the same code backs both the
guided walk-through and a scripted run. run() never prints or prompts, so
tests drive it directly.
"""

from __future__ import annotations

import argparse
from collections import Counter
from datetime import datetime
from pathlib import Path
from typing import Optional

import chart
import sheets

_ORDER_MENU = {
    "1": "filename",
    "2": "date_in_filename",
    "3": "date_in_sheet",
    "4": "custom",
}

_ORDER_HELP = {
    "filename": "Current order (alphabetical by filename)",
    "date_in_filename": "Date in the filename, e.g. weather_2026-01-15.xlsx",
    "date_in_sheet": "Date inside the sheet (a DATE/DATETIME/TIMESTAMP label)",
    "custom": "An order you type yourself",
}

_CHART_MENU = {"1": "line", "2": "scatter", "3": "bar"}

_WARNINGS_SHOWN = 12   # console stays readable on a 500-file run


def run(args) -> tuple[sheets.Result, list, Optional[Path]]:
    """Extract, draw, and return (result, selections used, chart path)."""
    result, selections = sheets.extract(
        folder=Path(args.input_folder),
        raw_selections=args.series,
        ordering=args.order,
        custom_order=args.custom_order,
        sheet_name=args.sheet,
        recursive=args.recursive,
    )

    run_id = datetime.now().strftime("%Y%m%d_%H%M%S")
    path = Path(args.output_folder) / "graphs" / f"{run_id}_chart.png"
    drawn = chart.draw(
        result,
        path,
        kind=args.chart,
        title=args.title or "Spreadsheet Visualizer",
        xlabel=args.xlabel or _ORDER_HELP.get(args.order, "Observation"),
        ylabel=args.ylabel or "Value",
        right_axis=args.right_axis or (),
    )
    return result, selections, drawn


def _ask(prompt: str, parse):
    """Keep asking until parse() accepts the answer. Ctrl-C/EOF still quits."""
    while True:
        try:
            return parse(input(prompt).strip())
        except (ValueError, IndexError, KeyError) as exc:
            print(f"  {exc}")
        except EOFError:
            raise SystemExit(1)


def as_folder(text: str) -> Path:
    """
    Turn typed or pasted text into a path. Dragging a folder into a terminal
    or copying its path from Finder brings quotes or backslash-escaped spaces
    with it; the literal text is tried first so a name really containing them
    still works.
    """
    literal = Path(text.strip()).expanduser()
    if literal.is_dir():
        return literal

    cleaned = text.strip()
    for quote in ("'", '"'):
        if len(cleaned) > 1 and cleaned.startswith(quote) and cleaned.endswith(quote):
            cleaned = cleaned[1:-1]
    cleaned = cleaned.replace("\\ ", " ").strip()
    candidate = Path(cleaned).expanduser()
    return candidate if candidate.is_dir() else literal


def _ask_folder(args) -> Path:
    if args.input_folder:
        return as_folder(str(args.input_folder))

    def parse(text):
        if not text:
            raise ValueError("Please type a folder path.")
        folder = as_folder(text)
        if not folder.is_dir():
            raise ValueError(f"No such folder: {folder}")
        if not sheets.find_files(folder, recursive=args.recursive):
            raise ValueError(f"No .xlsx files in {folder}")
        return folder

    return _ask("Folder containing the Excel files (full path is fine):\n> ", parse)


def _ask_series(files, sheet_name) -> list[str]:
    rows, headers = sheets.suggest_labels(files, sheet_name)
    if rows:
        print("\nLabels found:")
        for index, label in enumerate(rows, start=1):
            print(f"  {index}. {label}")
    if headers:
        print("Column headers found: " + ", ".join(headers))
        print("  (for a grid, combine them: Bob:Exam, or Bob:* for every column)")

    def parse(text):
        if not text:
            raise ValueError("Type at least one name, number, or cell.")
        chosen = []
        for token in (t.strip() for t in text.split(",")):
            if not token:
                continue
            # A bare number means the one printed above.
            if token.isdigit() and rows and 1 <= int(token) <= len(rows):
                chosen.append(rows[int(token) - 1])
            else:
                chosen.append(token)
        if not chosen:
            raise ValueError("Type at least one name, number, or cell.")
        return chosen

    print("\nWhat do you want to graph?")
    print("  a name (TEMP), a number from the list, a cell (B2),")
    print("  a row and column (Bob:Exam), or several separated by commas")
    return _ask("> ", parse)


def custom_order_parser(files):
    """Read "3, 1, 2" against a numbered file list, rejecting anything else."""
    def parse(text):
        numbers: list[int] = []
        for token in (t.strip() for t in text.split(",")):
            if not token:
                continue
            if not token.isdigit():
                raise ValueError(f"'{token}' is not a number.")
            number = int(token)
            if not 1 <= number <= len(files):
                raise ValueError(f"{number} is not between 1 and {len(files)}.")
            if number in numbers:
                raise ValueError(f"{number} is listed twice.")
            numbers.append(number)
        if not numbers:
            raise ValueError("Type at least one number.")
        return [files[n - 1].name for n in numbers]

    return parse


def _ask_order(files) -> tuple[str, Optional[list[str]]]:
    print("\nHow should the files be ordered along the x-axis?")
    for key, name in _ORDER_MENU.items():
        print(f"  {key}. {_ORDER_HELP[name]}")

    ordering = _ask("> ", lambda text: _ORDER_MENU[text or "1"])
    if ordering != "custom":
        return ordering, None

    print("\nFiles found:")
    for index, path in enumerate(files, start=1):
        print(f"  {index}. {path.name}")

    order = _ask("Type the order you want, e.g. 3, 1, 2\n> ", custom_order_parser(files))
    missing = len(files) - len(order)
    if missing:
        print(f"  ({missing} file(s) you didn't list will go last)")
    return "custom", order


def _ask_chart() -> str:
    print("\nChart type:")
    print("  1. Line    2. Scatter    3. Bar")
    return _ask("> ", lambda text: _CHART_MENU[text or "1"])


def fill_in(args) -> argparse.Namespace:
    """Ask for whatever the flags did not already answer."""
    args.input_folder = _ask_folder(args)
    files = sheets.find_files(Path(args.input_folder), recursive=args.recursive)
    if not files:
        print(f"No .xlsx files found in {args.input_folder}.")
        raise SystemExit(1)
    print(f"{len(files)} Excel file(s) found.")

    if not args.series:
        args.series = _ask_series(files, args.sheet)
    if not args.order:
        args.order, custom = _ask_order(files)
        args.custom_order = args.custom_order or custom
    if not args.chart:
        args.chart = _ask_chart()

    if args.order == "custom" and not args.custom_order:
        print("--order custom needs --custom-order too.")
        raise SystemExit(1)
    return args


def report(result, selections, drawn) -> None:
    print("\n" + "=" * 46)
    if selections:
        print("Graphing: " + ", ".join(f"{s.name} [{s.describe()}]" for s in selections[:6]))
        if len(selections) > 6:
            print(f"  ... and {len(selections) - 6} more series")
    print(
        f"{result.files_scanned} file(s) read, "
        f"{result.files_with_data} with data, "
        f"{result.files_failed} unreadable"
    )
    print(f"{len(result.points)} data point(s) across {len(result.series_names)} series")

    for note in result.notes:
        print(f"  note: {note}")

    if result.warnings:
        print(f"\n{len(result.warnings)} warning(s):")
        for message in result.warnings[:_WARNINGS_SHOWN]:
            print(f"  {message}")
        hidden = len(result.warnings) - _WARNINGS_SHOWN
        if hidden > 0:
            # Grouped, so a folder with one repeated problem is one line.
            kinds = Counter(m.split(": ", 1)[-1] for m in result.warnings[_WARNINGS_SHOWN:])
            print(f"  ... and {hidden} more:")
            for message, count in kinds.most_common(5):
                print(f"    {count} x {message}")

    if drawn:
        print(f"\nGraph saved to: {drawn}")
    else:
        print("\nNo data was extracted, so no graph was drawn.")
        print("Check the spelling of what you asked for, or try a different label.")


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Graph values across a folder of similarly structured .xlsx files."
    )
    parser.add_argument("--input-folder", help="Folder holding the .xlsx files")
    parser.add_argument("--output-folder", default="output", help="Where the graph goes")
    parser.add_argument(
        "--series", nargs="+",
        help="What to graph: TEMP, B2, Bob:Exam, Bob:*, cell:T5, TEMP=Outside",
    )
    parser.add_argument("--order", choices=sheets.ORDERINGS, help="X-axis order")
    parser.add_argument("--custom-order", nargs="+", help="Filenames, in the order you want")
    parser.add_argument("--chart", choices=chart.CHART_TYPES, help="Chart type")
    parser.add_argument("--sheet", help="Worksheet name, if not the first one")
    parser.add_argument("--recursive", action="store_true", help="Include subfolders")
    parser.add_argument("--title", help="Chart title")
    parser.add_argument("--xlabel", help="X-axis label")
    parser.add_argument("--ylabel", help="Y-axis label")
    parser.add_argument(
        "--right-axis", nargs="+", metavar="SERIES",
        help="Series to put on a second y-axis (for very different ranges)",
    )
    return parser


def main(argv=None) -> int:
    args = build_parser().parse_args(argv)
    try:
        args = fill_in(args)
    except (FileNotFoundError, ValueError) as exc:
        print(f"Error: {exc}")
        return 1
    except KeyboardInterrupt:
        print("\nCancelled.")
        return 1

    result, selections, drawn = run(args)
    report(result, selections, drawn)
    return 0 if result.points else 1


if __name__ == "__main__":
    raise SystemExit(main())
