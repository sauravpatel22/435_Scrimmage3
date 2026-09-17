"""
spreadsheet_processor.py

Core engine of the spreadsheet visualizer: finding .xlsx files in a folder,
resolving a user's Selector (label or cell) to an actual value inside each
workbook, determining a consistent x-axis ordering across files, and
extracting a combined dataset (as an ExtractionResult) for visualizer.py to
chart. Warnings collected along the way are handed back to the caller to
print - this module writes no files of its own.

This module is the direct implementation of the "Specific Design Notes"
from the Scrimmage 2 plan:
  - labels are searched for, not assumed to sit at a fixed cell (files can
    drift in layout between runs)
  - a missing/invalid value in one file is recorded as a warning and the
    run continues, it never aborts the whole batch
  - ordering can come from the filename, filesystem metadata, or an
    explicit user-supplied order

External dependency: openpyxl, used to read .xlsx files without needing
Excel installed.
"""

from __future__ import annotations

from datetime import datetime
from pathlib import Path
from typing import Any, List, Optional, Tuple

import openpyxl
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet

from models import (
    ExtractionResult,
    Observation,
    OrderingMethod,
    RunConfig,
    RunWarning,
    Selector,
    SelectorType,
    WarningLevel,
)

# Excel creates a temporary lock file (e.g. "~$weather_01.xlsx") next to any
# workbook that's currently open in the desktop app. It is not real data and
# openpyxl cannot read it, so discovery must skip it explicitly.
_EXCEL_LOCK_FILE_PREFIX = "~$"


# ---------------------------------------------------------------------------
# File discovery
# ---------------------------------------------------------------------------

def discover_files(folder: Path, recursive: bool = False) -> List[Path]:
    """
    Return every real .xlsx file directly inside `folder` (or, if
    `recursive`, anywhere below it), sorted by filename so that behavior is
    deterministic regardless of the operating system's directory listing
    order. Excel lock files are filtered out since they aren't real
    workbooks and openpyxl cannot open them.
    """
    if not folder.is_dir():
        raise FileNotFoundError(f"Input folder does not exist: {folder}")

    pattern = "**/*.xlsx" if recursive else "*.xlsx"
    candidates = [
        path
        for path in folder.glob(pattern)
        if path.is_file() and not path.name.startswith(_EXCEL_LOCK_FILE_PREFIX)
    ]
    return sorted(candidates, key=lambda p: p.name.lower())


def load_workbook_safe(path: Path) -> Tuple[Optional[openpyxl.Workbook], Optional[str]]:
    """
    Try to open a workbook, returning (workbook, None) on success or
    (None, reason) on failure. data_only=True reads each formula cell's
    last-calculated value rather than the formula text, since this tool
    has no spreadsheet engine of its own to recalculate anything.
    """
    try:
        workbook = openpyxl.load_workbook(path, data_only=True, read_only=True)
        return workbook, None
    except Exception as exc:  # noqa: BLE001 - any failure means "skip this file"
        return None, f"could not open file ({exc.__class__.__name__}: {exc})"


def select_worksheet(
    workbook: openpyxl.Workbook, sheet_name: Optional[str]
) -> Tuple[Optional[Worksheet], Optional[str]]:
    """
    Pick which worksheet to read from a workbook. If the caller asked for a
    specific sheet name and it isn't present, that file is treated as
    unusable (its structure doesn't match what the user described) rather
    than silently substituting a different sheet. With no sheet_name given,
    the first sheet is used, matching the "single worksheet detected"
    common case from the sample workflow.
    """
    if sheet_name is not None:
        if sheet_name not in workbook.sheetnames:
            return None, f"worksheet '{sheet_name}' not found"
        return workbook[sheet_name], None

    return workbook.worksheets[0], None


# ---------------------------------------------------------------------------
# Label / cell resolution
# ---------------------------------------------------------------------------

def _find_label_cell(worksheet: Worksheet, label: str) -> Optional[Tuple[int, int]]:
    """
    Scan the worksheet for a cell whose text matches `label`
    (case-insensitive, whitespace-trimmed) and return its (row, column),
    or None if it isn't present. Stops at the first match - this is the
    "search for the label again instead of relying on a fixed cell
    location" requirement, nothing more; a file with the same label
    written twice is not a case the plan called out.
    """
    target = label.strip().casefold()

    for row in worksheet.iter_rows():
        for cell in row:
            value = cell.value
            if isinstance(value, str) and value.strip().casefold() == target:
                return cell.row, cell.column

    return None


def resolve_selector(
    worksheet: Worksheet, selector: Selector
) -> Tuple[Optional[Any], Optional[str], Optional[str]]:
    """
    Resolve one Selector against one worksheet.

    Returns (raw_value, cell_address, error_message):
      - On success: (value, "B2", None)
      - On failure: (None, None, "<why it failed>")

    For a CELL selector this is a direct lookup. For a LABEL selector, the
    label's own cell only tells us *where* the name is written - the actual
    value lives in the cell immediately to its right on the same row, which
    is the layout convention described in the plan (a "Property"/"Value"
    style table). Searching by label instead of trusting a fixed address is
    what lets the tool tolerate rows being reordered or the whole table
    being moved between files.
    """
    if selector.type == SelectorType.CELL:
        address = selector.raw.upper()
        try:
            cell = worksheet[address]
        except ValueError as exc:
            return None, None, f"invalid cell reference '{selector.raw}': {exc}"
        # Use the requested address directly rather than cell.coordinate:
        # in read_only mode, a cell past the sheet's populated range comes
        # back as an EmptyCell, which has no .coordinate attribute.
        return cell.value, address, None

    match = _find_label_cell(worksheet, selector.raw)
    if match is None:
        return None, None, f"label '{selector.raw}' not found"

    row, col = match
    value_cell = worksheet.cell(row=row, column=col + 1)
    # In read_only mode, a cell past the sheet's populated range comes back
    # as an EmptyCell, which (unlike a normal Cell) has no .coordinate
    # attribute - so the address is built manually from row/col instead of
    # relying on the cell object to know its own position.
    value_coordinate = f"{get_column_letter(col + 1)}{row}"
    return value_cell.value, value_coordinate, None


# ---------------------------------------------------------------------------
# Value validation
# ---------------------------------------------------------------------------

def normalize_numeric_value(raw: Any) -> Tuple[Optional[float], Optional[str]]:
    """
    Convert whatever openpyxl handed back into a plotting-ready float.

    Returns (value, error_message): value is None whenever the cell was
    blank or could not be interpreted as a number, and error_message then
    explains why - the caller turns that into a RunWarning instead of
    treating a missing/invalid value as zero (zero would be a real,
    misleading data point).
    """
    if raw is None:
        return None, "missing value"

    if isinstance(raw, bool):
        # bool is a subclass of int in Python; treating True/False as 1/0
        # would silently misrepresent a checkbox-like cell as a real value.
        return None, f"invalid (non-numeric) value: {raw!r}"

    if isinstance(raw, (int, float)):
        return float(raw), None

    if isinstance(raw, str):
        text = raw.strip()
        if not text:
            return None, "missing value"
        try:
            return float(text), None
        except ValueError:
            return None, f"invalid (non-numeric) value: {raw!r}"

    return None, f"unsupported value type: {type(raw).__name__}"


# ---------------------------------------------------------------------------
# Ordering
# ---------------------------------------------------------------------------

def determine_order_key(path: Path, config: RunConfig) -> Tuple[Any, Optional[str]]:
    """
    Compute the value used to position one file's observations on the
    x-axis / processing order, per `config.ordering_method`. Returns
    (order_key, warning_message); when a preferred ordering signal isn't
    available (e.g. a file missing from a custom order), this falls back to
    the filename so the run can still proceed, and reports the fallback as
    a warning rather than failing.
    """
    method = config.ordering_method

    if method == OrderingMethod.FILENAME:
        return path.name, None

    if method == OrderingMethod.CUSTOM:
        try:
            index = config.custom_order.index(path.name)  # type: ignore[union-attr]
        except ValueError:
            # File wasn't in the user's explicit list; push it to the end
            # rather than dropping it, and say why.
            return path.name, f"'{path.name}' not found in custom order; placed last"
        return index, None

    if method == OrderingMethod.FILE_CREATION_DATE:
        stat = path.stat()
        # macOS exposes true creation time as st_birthtime; other platforms
        # (notably Linux) do not, and st_ctime there is actually the last
        # metadata-change time, not creation time. We prefer the accurate
        # value where it exists and warn when we must fall back.
        if hasattr(stat, "st_birthtime"):
            return datetime.fromtimestamp(stat.st_birthtime), None
        return (
            datetime.fromtimestamp(stat.st_ctime),
            "true file creation time is unavailable on this platform; "
            "used last metadata-change time instead",
        )

    raise AssertionError(f"unhandled OrderingMethod: {method}")  # pragma: no cover


# ---------------------------------------------------------------------------
# Extraction
# ---------------------------------------------------------------------------

def extract_dataset(config: RunConfig) -> ExtractionResult:
    """
    Process every .xlsx file in config.input_folder and build the combined
    ExtractionResult: one Observation per (file, selector) that resolves to
    a valid number, and one RunWarning for everything else (file failed to
    open, worksheet missing, label not found, value missing/invalid,
    ordering fallback). No exception from a single file is allowed to stop
    the batch - that is the "don't terminate the program" design
    requirement in practice.
    """
    result = ExtractionResult()
    files = discover_files(config.input_folder)

    if not files:
        result.add_warning(
            RunWarning(
                message=f"no .xlsx files found in {config.input_folder}",
                level=WarningLevel.ERROR,
            )
        )
        return result

    for path in files:
        workbook, open_error = load_workbook_safe(path)
        if workbook is None:
            result.add_warning(
                RunWarning(source_file=path, message=open_error, level=WarningLevel.ERROR)
            )
            continue

        try:
            worksheet, sheet_error = select_worksheet(workbook, config.sheet_name)
            if worksheet is None:
                result.add_warning(
                    RunWarning(source_file=path, message=sheet_error, level=WarningLevel.ERROR)
                )
                continue

            order_key, order_warning = determine_order_key(path, config)
            if order_warning:
                result.add_warning(RunWarning(source_file=path, message=order_warning))

            for selector in config.selectors:
                raw_value, cell_address, resolve_error = resolve_selector(worksheet, selector)

                if resolve_error is not None:
                    result.add_warning(
                        RunWarning(
                            source_file=path,
                            series_name=selector.series_name,
                            message=resolve_error,
                            level=WarningLevel.ERROR,
                        )
                    )
                    continue

                value, value_error = normalize_numeric_value(raw_value)
                if value_error is not None:
                    result.add_warning(
                        RunWarning(
                            source_file=path,
                            series_name=selector.series_name,
                            message=value_error,
                            level=WarningLevel.ERROR
                            if raw_value is not None
                            else WarningLevel.WARNING,
                        )
                    )
                    continue

                result.add_observation(
                    Observation(
                        source_file=path,
                        sheet_name=worksheet.title,
                        series_name=selector.series_name,
                        cell=cell_address,
                        order_key=order_key,
                        value=value,
                    )
                )
        finally:
            # read_only workbooks hold an open file handle until closed;
            # closing promptly matters here since a folder may contain
            # hundreds of files in one run.
            workbook.close()

    return result