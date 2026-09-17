"""
main.py

Command-line entry point that ties the other modules together into the
guided workflow described in the Scrimmage 2 plan and mocked up in
scrimmage2_workflow.pptx: point at a folder of .xlsx files, pick one or
more series to track, pick how to order observations, extract + validate
the data, pick a chart type, and write the graph to disk. The only output
file is a single PNG chart; any warnings from the run are printed to the
console instead of being written to disk.

This module is split into two layers on purpose:
  - run_pipeline(config): pure orchestration, no user interaction. Given a
    finished RunConfig, it calls spreadsheet_processor and visualizer and
    returns the results. This is what tests should call directly.
  - the interactive functions (prompt_for_config, main): gather a
    RunConfig from the terminal (or from argparse flags, for
    non-interactive/scripted use) and then call run_pipeline.

No new external dependency is introduced here beyond what
spreadsheet_processor.py and visualizer.py already require.
"""

from __future__ import annotations

import argparse
import sys
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import List, Optional

import openpyxl

from models import ChartType, OrderingMethod, RunConfig, Selector
from spreadsheet_processor import discover_files, extract_dataset
from visualizer import create_chart

_ORDERING_MENU = {
    "1": OrderingMethod.FILENAME,
    "2": OrderingMethod.FILE_CREATION_DATE,
    "3": OrderingMethod.CUSTOM,
}

_CHART_MENU = {
    "1": ChartType.LINE,
    "2": ChartType.SCATTER,
    "3": ChartType.BAR,
}


@dataclass
class RunOutputs:
    """Everything a completed run produced, gathered in one place so the
    interactive summary and any calling test can inspect it uniformly."""

    run_id: str
    files_processed: int
    observation_count: int
    warnings: List[str]
    chart_path: Path


# ---------------------------------------------------------------------------
# Non-interactive orchestration (what tests should call)
# ---------------------------------------------------------------------------

def run_pipeline(config: RunConfig) -> RunOutputs:
    """
    Execute one full run for an already-built RunConfig: extract data and
    generate the chart. Contains no input()/print() calls so it can be
    exercised directly by tests or by other tools, and reused unchanged
    whether config was built interactively or from CLI flags. Warnings are
    returned as plain strings for the caller to print or otherwise handle,
    rather than being written to a file here.
    """
    run_id = datetime.now().strftime("%Y%m%d_%H%M%S")

    result = extract_dataset(config)
    chart_path = create_chart(result, config, run_id)

    files_processed = len({obs.source_file for obs in result.observations})

    return RunOutputs(
        run_id=run_id,
        files_processed=files_processed,
        observation_count=len(result.observations),
        warnings=[str(warning) for warning in result.warnings],
        chart_path=chart_path,
    )


# ---------------------------------------------------------------------------
# Interactive prompts
# ---------------------------------------------------------------------------

def _detect_candidate_labels(input_folder: Path, sample_size: int = 5) -> List[str]:
    """
    Best-effort peek at the first few files to suggest labels the user
    might want to select, matching the "Available labels detected" step
    of the sample workflow. This is only a convenience prompt - it never
    blocks the user from typing a label that wasn't detected, since a
    label might legitimately appear only in some files.
    """
    labels: List[str] = []
    seen = set()

    for path in discover_files(input_folder)[:sample_size]:
        try:
            workbook = openpyxl.load_workbook(path, data_only=True, read_only=True)
        except Exception:  # noqa: BLE001 - a preview failure is not fatal
            continue
        try:
            worksheet = workbook.worksheets[0]
            for row in worksheet.iter_rows(max_col=1):
                for cell in row:
                    value = cell.value
                    if isinstance(value, str) and value.strip() and value.strip().casefold() not in (
                        "property",
                        "label",
                        "name",
                    ):
                        normalized = value.strip()
                        if normalized.casefold() not in seen:
                            seen.add(normalized.casefold())
                            labels.append(normalized)
        finally:
            workbook.close()

    return labels


def _prompt_ordering(files: List[Path]) -> tuple[OrderingMethod, Optional[List[str]]]:
    print("How should observations be ordered?")
    print("  1. Current order (filename)")
    print("  2. Date created")
    print("  3. Custom order")
    choice = input("> ").strip()
    method = _ORDERING_MENU.get(choice, OrderingMethod.FILENAME)

    if method != OrderingMethod.CUSTOM:
        return method, None

    # Matches the numbered-selection demo in the workflow slides: list
    # every discovered file with a number, then let the user type the
    # desired order back as a comma-separated list of those numbers.
    print("Detected files:")
    for index, path in enumerate(files, start=1):
        print(f"  {index}. {path.name}")
    order_text = input("Enter desired order as numbers, e.g. 3, 1, 4, 2\n> ")
    indices = [int(token.strip()) for token in order_text.split(",") if token.strip()]
    custom_order = [files[i - 1].name for i in indices]
    return method, custom_order


def prompt_for_config(args: argparse.Namespace) -> RunConfig:
    """
    Build a RunConfig by combining any CLI flags already supplied in
    `args` with interactive input() prompts for whatever is still missing.
    A fully-flagged invocation (all of --input-folder, --series, --order)
    runs with no prompts at all, which is what lets this same function
    back both the guided interactive flow and scripted/CI usage.
    """
    input_folder = Path(args.input_folder) if args.input_folder else None
    if input_folder is None:
        input_folder = Path(input("Enter folder containing Excel files:\n> ").strip())

    print("Scanning folder...")
    files = discover_files(input_folder)
    print(f"  {len(files)} Excel files found")

    if not files:
        print(f"No .xlsx files found in {input_folder}.")
        sys.exit(1)

    if args.series:
        selectors = [Selector.from_input(token) for token in args.series]
    else:
        candidates = _detect_candidate_labels(input_folder)
        if candidates:
            print("Available labels detected:")
            for index, label in enumerate(candidates, start=1):
                print(f"  {index}. {label}")
        raw = input("Select one or more labels to graph (name or number, comma-separated):\n> ")
        tokens = [token.strip() for token in raw.split(",") if token.strip()]
        # A token matching one of the numbers just printed is resolved to
        # that label's actual text; anything else is taken as a literal
        # label name or cell reference. Without this, typing the displayed
        # number (a very natural thing to do, right below a menu that DOES
        # take numbers) silently searches for a label named "1" or "2" and
        # finds nothing.
        resolved = [
            candidates[int(token) - 1]
            if token.isdigit() and candidates and 1 <= int(token) <= len(candidates)
            else token
            for token in tokens
        ]
        selectors = [Selector.from_input(token) for token in resolved]

    if args.order:
        ordering_method = OrderingMethod(args.order)
        custom_order = args.custom_order
    else:
        ordering_method, custom_order = _prompt_ordering(files)

    if args.chart:
        chart_type = ChartType(args.chart)
    else:
        print("Select graph type:")
        print("  1. Line Graph")
        print("  2. Scatter Plot")
        print("  3. Bar Graph")
        chart_type = _CHART_MENU.get(input("> ").strip(), ChartType.LINE)

    output_folder = Path(args.output_folder) if args.output_folder else Path("output")

    return RunConfig(
        input_folder=input_folder,
        output_folder=output_folder,
        selectors=selectors,
        ordering_method=ordering_method,
        custom_order=custom_order,
        chart_type=chart_type,
        sheet_name=args.sheet,
    )


def _print_summary(config: RunConfig, outputs: RunOutputs) -> None:
    print("=" * 40)
    print("             COMPLETE")
    print("=" * 40)
    print(f"{outputs.files_processed} files processed")
    print(f"{outputs.observation_count} data points extracted")
    print(f"{len(outputs.warnings)} warnings")
    if outputs.warnings:
        print("Warnings:")
        for warning in outputs.warnings:
            print(f"  {warning}")
    print(f"Graph saved to: {outputs.chart_path}")


def build_arg_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Visualize a property across a folder of similarly structured .xlsx files."
    )
    parser.add_argument("--input-folder", help="Folder containing the .xlsx files to process")
    parser.add_argument("--output-folder", default=None, help="Where to write generated output")
    parser.add_argument(
        "--series",
        nargs="+",
        help="One or more labels/cell references to extract, e.g. --series TEMP HUMIDITY",
    )
    parser.add_argument(
        "--order",
        choices=[method.value for method in OrderingMethod],
        help="Ordering method for the x-axis",
    )
    parser.add_argument(
        "--custom-order",
        nargs="+",
        help="Filenames in the desired order (required when --order custom is used non-interactively)",
    )
    parser.add_argument(
        "--chart",
        choices=[chart.value for chart in ChartType],
        help="Chart type; omit to be prompted",
    )
    parser.add_argument("--sheet", default=None, help="Worksheet name, if not the first sheet")
    return parser


def main(argv: Optional[List[str]] = None) -> int:
    parser = build_arg_parser()
    args = parser.parse_args(argv)

    try:
        config = prompt_for_config(args)
    except (ValueError, FileNotFoundError) as exc:
        print(f"Error: {exc}")
        return 1

    outputs = run_pipeline(config)
    _print_summary(config, outputs)

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
