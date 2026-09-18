"""
sheets.py - getting numbers out of a folder of .xlsx files.

Each worksheet is read into memory once (class Sheet) and every later lookup
is a dictionary hit, which is what makes it cheap to search for a label
rather than trust it to sit at a fixed address. Nothing here raises on bad
data; a problem with one file becomes a warning string and the run goes on.
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from datetime import date, datetime
from pathlib import Path
from typing import Any, Optional

import openpyxl
from openpyxl.utils import get_column_letter
from openpyxl.utils.cell import coordinate_to_tuple

ORDERINGS = ("filename", "date_in_filename", "date_in_sheet", "custom")

_CELL_RE = re.compile(r"^[A-Za-z]{1,3}[1-9][0-9]*$")       # A1, c9, AA123
_DATE_RE = re.compile(r"(\d{4})[-_]?(\d{2})[-_]?(\d{2})")  # 2026-01-31, 2026_01_31
_DATE_LABELS = ("DATE", "DATETIME", "TIMESTAMP")
_LOCK_PREFIX = "~$"          # Excel's lock file for an open workbook
_HEADER_WORDS = {"property", "label", "name", "value", "item", "metric"}
_PLOTTABLE_LIMIT = 1e300     # beyond this matplotlib cannot compute an axis
_SAMPLE_FILES = 3

# Ranking dates before numbers before text lets one sorted list hold all three
# without "'<' not supported between instances of ..." errors.
_DATE_RANK, _NUMBER_RANK, _TEXT_RANK = 0, 1, 2


@dataclass
class Point:
    """One number pulled out of one file."""

    file: Path
    series: str
    sort_key: tuple      # (rank, value, filename)
    x_label: str
    value: float


@dataclass
class Result:
    points: list[Point] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)
    notes: list[str] = field(default_factory=list)
    files_scanned: int = 0
    files_failed: int = 0

    @property
    def series_names(self) -> list[str]:
        """Series in first-seen order, so the legend is stable."""
        names: list[str] = []
        for point in self.points:
            if point.series not in names:
                names.append(point.series)
        return names

    @property
    def files_with_data(self) -> int:
        return len({point.file for point in self.points})

    def points_for(self, series: str) -> list[Point]:
        return [point for point in self.points if point.series == series]


def _norm(text: str) -> str:
    """Compare text ignoring case and any kind of surrounding or double space."""
    return " ".join(str(text).replace(" ", " ").split()).casefold()


class Sheet:
    """One worksheet, read once into plain dictionaries."""

    def __init__(self, worksheet) -> None:
        self.title = worksheet.title
        self.values: dict[tuple[int, int], Any] = {}
        self.text: dict[str, list[tuple[int, int]]] = {}

        for row in worksheet.iter_rows():
            for cell in row:
                if cell.value is None:
                    # Must come first: in read-only mode an empty cell is an
                    # EmptyCell, which has no .row or .column to ask about.
                    continue
                self.values[(cell.row, cell.column)] = cell.value
                if isinstance(cell.value, str) and cell.value.strip():
                    self.text.setdefault(_norm(cell.value), []).append(
                        (cell.row, cell.column)
                    )

    def find(self, text: str) -> list[tuple[int, int]]:
        """Every (row, col) holding this text."""
        return self.text.get(_norm(text), [])

    def value(self, row: int, col: int) -> Any:
        return self.values.get((row, col))

    def text_left_of(self, row: int, col: int) -> Optional[str]:
        """Nearest text cell to the left on the same row - a row's label."""
        best = None
        for (r, c), value in self.values.items():
            if r == row and c < col and isinstance(value, str) and value.strip():
                if best is None or c > best[0]:
                    best = (c, value.strip())
        return best[1] if best else None

    def text_above(self, row: int, col: int) -> Optional[str]:
        """Nearest text cell above in the same column - a column's header."""
        best = None
        for (r, c), value in self.values.items():
            if c == col and r < row and isinstance(value, str) and value.strip():
                if best is None or r > best[0]:
                    best = (r, value.strip())
        return best[1] if best else None


def find_files(folder: Path, recursive: bool = False) -> list[Path]:
    """Every readable .xlsx in the folder, sorted by name."""
    if not folder.is_dir():
        raise FileNotFoundError(f"Input folder does not exist: {folder}")
    pattern = "**/*.xlsx" if recursive else "*.xlsx"
    files = [
        path
        for path in folder.glob(pattern)
        if path.is_file() and not path.name.startswith(_LOCK_PREFIX)
    ]
    return sorted(files, key=lambda path: path.name.lower())


def open_sheets(path: Path, sheet_name: Optional[str] = None):
    """
    Open one file. Returns (list of Sheets, None) or (None, why it failed).
    With no sheet_name every worksheet comes back, first one first.
    """
    try:
        workbook = openpyxl.load_workbook(path, data_only=True, read_only=True)
    except Exception as exc:  # noqa: BLE001 - any failure just means "skip it"
        return None, f"could not open file ({exc.__class__.__name__}: {exc})"
    try:
        if sheet_name is not None:
            if sheet_name not in workbook.sheetnames:
                return None, f"worksheet '{sheet_name}' not found"
            return [Sheet(workbook[sheet_name])], None
        return [Sheet(worksheet) for worksheet in workbook.worksheets], None
    finally:
        workbook.close()


def open_sheet(path: Path, sheet_name: Optional[str] = None):
    """The first worksheet only."""
    opened, error = open_sheets(path, sheet_name)
    return (opened[0] if opened else None), error


@dataclass
class Selection:
    """
    One thing to graph, resolved to a single way of looking it up: "label"
    (value to the right), "grid" (row crossed with column), or "cell".
    """

    name: str
    mode: str
    row_label: str = ""
    col_label: str = ""
    address: str = ""
    typed: str = ""        # what the user typed, before any parsing
    forced: bool = False   # they said cell:/label:, so do not second-guess

    def describe(self) -> str:
        if self.mode == "label":
            return f"label '{self.row_label}' (value to its right)"
        if self.mode == "grid":
            return f"row '{self.row_label}' x column '{self.col_label}'"
        return f"cell {self.address}"


def _is_number(value: Any) -> bool:
    return isinstance(value, (int, float)) and not isinstance(value, bool)


def plan_selections(raw_inputs, samples):
    """
    Turn what the user typed into concrete Selections, settling anything
    ambiguous against a few sample sheets. Returns (selections, notes).

    Accepted forms:
      TEMP            a label
      Bob:Exam        a row label and a column header
      Bob:*           every column for that row       (expands to many series)
      *:Exam          every row for that column       (expands to many series)
      B2              a cell address
      cell:T5         force an address that reads like a label
      label:T5        force a label that reads like an address
      TEMP=Outside    any of the above, renamed in the legend
    """
    samples = [s for s in (samples if isinstance(samples, list) else [samples]) if s]
    known = lambda text: any(sheet.find(text) for sheet in samples)   # noqa: E731

    selections: list[Selection] = []
    notes: list[str] = []

    for raw in raw_inputs:
        text = str(raw).strip()
        if not text:
            continue
        original = text

        display = ""
        if "=" in text and not known(text):
            text, _, display = (part.strip() for part in text.rpartition("="))
            if not text:                       # someone typed "=X"
                text, display = display, ""

        forced = ""
        for prefix in ("cell:", "label:"):
            if text.casefold().startswith(prefix):
                forced = prefix[:-1]
                text = text[len(prefix):].strip()

        if forced == "cell":
            selections.append(
                Selection(display or text, "cell", address=text.upper(), forced=True)
            )
            continue
        if forced == "label":
            selections.append(
                Selection(display or text, "label", row_label=text, forced=True)
            )
            continue

        # A real label beats any other reading, so a sheet with a row called
        # "T5", "CO2" or "Time: AM" behaves the way the user expects.
        if known(text):
            selections.append(Selection(display or text, "label", row_label=text))
            continue

        if ":" in text:
            row_label, _, col_label = (part.strip() for part in text.partition(":"))
            if row_label and col_label:
                expanded, expand_notes = _expand_grid(row_label, col_label, display, samples)
                for selection in expanded:
                    selection.typed = original
                selections.extend(expanded)
                notes.extend(expand_notes)
                continue

        if _CELL_RE.match(text):
            selections.append(
                Selection(display or text, "cell", address=text.upper(), typed=original)
            )
            if samples:
                notes.append(
                    f"'{text}' read as a cell address; type label:{text} if it is a label"
                )
            continue

        selections.append(
            Selection(display or text, "label", row_label=text, typed=original)
        )

    return selections, notes


def _expand_grid(row_label: str, col_label: str, display: str, samples):
    """Turn Bob:Exam, Bob:* and *:Exam into concrete Selections."""
    if "*" not in (row_label, col_label):
        name = display or f"{row_label} - {col_label}"
        return [Selection(name, "grid", row_label=row_label, col_label=col_label)], []

    if row_label == "*" and col_label == "*":
        return [], ["'*:*' is not supported; name either the row or the column"]

    if not samples:
        return [], [f"cannot expand '{row_label}:{col_label}' without a readable file"]

    # Merged across samples so a row that only appears in a later file counts.
    selections: list[Selection] = []
    for sheet in samples:
        for found in _expand_one(row_label, col_label, sheet):
            if not any(s.name == found.name for s in selections):
                selections.append(found)

    if not selections:
        return [], [f"nothing numeric found to expand '{row_label}:{col_label}'"]
    return selections, [f"'{row_label}:{col_label}' expanded to {len(selections)} series"]


def _expand_one(row_label: str, col_label: str, sheet: Sheet) -> list[Selection]:
    """One sheet's worth of wildcard expansion."""
    selections: list[Selection] = []

    if row_label == "*" and col_label != "*":
        matches = sheet.find(col_label)
        if not matches:
            return []
        _, col = matches[0]
        for (r, c), value in sorted(sheet.values.items()):
            if c != col or not _is_number(value):
                continue
            name = sheet.text_left_of(r, c)
            if name:
                selections.append(
                    Selection(name, "grid", row_label=name, col_label=col_label)
                )
    elif col_label == "*" and row_label != "*":
        matches = sheet.find(row_label)
        if not matches:
            return []
        row, label_col = matches[0]
        for (r, c), value in sorted(sheet.values.items()):
            if r != row or c <= label_col or not _is_number(value):
                continue
            header = sheet.text_above(r, c)
            if header:
                selections.append(
                    Selection(
                        f"{row_label} - {header}", "grid",
                        row_label=row_label, col_label=header,
                    )
                )
    return selections


def read_selection(sheet: Sheet, selection: Selection):
    """Returns (raw value, cell address, error). Only one of value/error is set."""
    # The samples that planned this selection may not have represented every
    # file, so one dictionary lookup rescues a file they missed.
    if selection.typed and not selection.forced and sheet.find(selection.typed):
        selection = Selection(selection.name, "label", row_label=selection.typed)

    if selection.mode == "cell":
        try:
            row, col = coordinate_to_tuple(selection.address)
        except Exception:  # noqa: BLE001 - openpyxl raises several types here
            return None, "", f"'{selection.address}' is not a cell address"
        return sheet.value(row, col), selection.address, None

    if selection.mode == "grid":
        rows = sheet.find(selection.row_label)
        if not rows:
            return None, "", f"row '{selection.row_label}' not found"
        cols = sheet.find(selection.col_label)
        if not cols:
            return None, "", f"column '{selection.col_label}' not found"
        row = rows[0][0]
        above = [c for (r, c) in cols if r < row]      # a header sits above its data
        col = above[0] if above else cols[0][1]
        return sheet.value(row, col), f"{get_column_letter(col)}{row}", None

    matches = sheet.find(selection.row_label)
    if not matches:
        return None, "", f"label '{selection.row_label}' not found"
    row, col = matches[0]
    if len(matches) > 1:
        where = ", ".join(f"{get_column_letter(c)}{r}" for r, c in matches[:3])
        return (
            sheet.value(row, col + 1),
            f"{get_column_letter(col + 1)}{row}",
            f"label '{selection.row_label}' appears {len(matches)} times ({where}); used the first",
        )
    return sheet.value(row, col + 1), f"{get_column_letter(col + 1)}{row}", None


def to_number(raw: Any):
    """Returns (value, error). Blank and non-numeric cells are errors, never 0."""
    if raw is None:
        return None, "missing value"
    if isinstance(raw, bool):
        # bool is an int in Python, so True would otherwise plot as 1.
        return None, f"invalid (non-numeric) value: {raw!r}"
    if isinstance(raw, (int, float)):
        return _plottable(float(raw), raw)
    if isinstance(raw, str):
        text = raw.strip()
        if not text:
            return None, "missing value"
        try:
            return _plottable(float(text), raw)
        except (ValueError, OverflowError):
            return None, f"invalid (non-numeric) value: {raw!r}"
    return None, f"unsupported value type: {type(raw).__name__}"


def _plottable(value: float, raw: Any):
    """Refuse floats matplotlib cannot draw, which would cost the whole chart."""
    if value != value or value in (float("inf"), float("-inf")):
        return None, f"invalid (non-numeric) value: {raw!r}"
    if abs(value) > _PLOTTABLE_LIMIT:
        return None, f"value too large to plot: {raw!r}"
    return value, None


def _date_from_filename(path: Path) -> Optional[date]:
    match = _DATE_RE.search(path.stem)
    if not match:
        return None
    try:
        return date(*(int(part) for part in match.groups()))
    except ValueError:
        return None


def _date_from_sheet(sheet: Optional[Sheet]) -> Any:
    if sheet is None:
        return None
    for candidate in _DATE_LABELS:
        matches = sheet.find(candidate)
        if matches:
            row, col = matches[0]
            return sheet.value(row, col + 1)
    return None


def _as_x(value: Any) -> tuple[tuple, str]:
    """Build a (sort_key, label) pair for whatever the ordering worked out."""
    if isinstance(value, datetime):
        stamp = "%Y-%m-%d %H:%M" if value.time() != datetime.min.time() else "%Y-%m-%d"
        return (_DATE_RANK, value), value.strftime(stamp)
    if isinstance(value, date):
        return (_DATE_RANK, datetime.combine(value, datetime.min.time())), value.isoformat()
    if _is_number(value):
        return (_NUMBER_RANK, float(value)), str(value)
    return (_TEXT_RANK, str(value)), str(value)


def order_of(path: Path, sheet: Optional[Sheet], ordering: str, custom_order=None):
    """
    Where this file sits on the x-axis. Returns (sort_key, x_label, warning).
    The label is separate from the key: a custom order sorts by position but
    is still labelled with filenames.
    """
    if ordering == "filename":
        return (_TEXT_RANK, path.name), path.name, None

    if ordering == "custom":
        order = list(custom_order or [])
        if path.name in order:
            return (_NUMBER_RANK, float(order.index(path.name))), path.name, None
        return (
            (_NUMBER_RANK, float(len(order))), path.name,
            "not in the custom order; placed last",
        )

    if ordering == "date_in_filename":
        parsed = _date_from_filename(path)
        if parsed is not None:
            return _as_x(parsed) + (None,)
        return (_TEXT_RANK, path.name), path.name, "no date in the filename; sorted by name"

    if ordering == "date_in_sheet":
        raw = _date_from_sheet(sheet)
        if raw is None:
            return (_TEXT_RANK, path.name), path.name, "no DATE label in the sheet; sorted by name"
        if isinstance(raw, (date, datetime)):
            return _as_x(raw) + (None,)
        key, label = _as_x(raw)
        return key, label, f"DATE value {raw!r} is not a date; sorted by its raw value"

    raise ValueError(f"unknown ordering: {ordering}")


def extract(
    folder: Path,
    raw_selections,
    ordering: str = "filename",
    custom_order=None,
    sheet_name: Optional[str] = None,
    recursive: bool = False,
) -> tuple[Result, list[Selection]]:
    """
    Read every file in the folder and pull out every selection. Returns
    (Result, the selections used), since "Bob:*" only becomes a list of real
    series once a file has been read.
    """
    result = Result()
    files = find_files(folder, recursive=recursive)
    if not files:
        result.warnings.append(f"no .xlsx files found in {folder}")
        return result, []

    # Sampled from across the folder: the first file alphabetically is often
    # the odd one out, and a wildcard needs to see every row that exists.
    samples: list[Sheet] = []
    for index in sorted({0, len(files) // 2, len(files) - 1}):
        opened, _ = open_sheets(files[index], sheet_name)
        if opened:
            samples.extend(opened)
        if len(samples) >= _SAMPLE_FILES:
            break

    selections, notes = plan_selections(raw_selections, samples)
    result.notes.extend(notes)
    if not selections:
        result.warnings.append("nothing to graph")
        return result, []

    for path in files:
        opened, error = open_sheets(path, sheet_name)
        result.files_scanned += 1
        if not opened:
            result.files_failed += 1
            result.warnings.append(f"{path.name}: {error}")
            continue

        sheet = opened[0]
        sort_key, x_label, order_warning = order_of(path, sheet, ordering, custom_order)
        if order_warning:
            result.warnings.append(f"{path.name}: {order_warning}")
        # The filename keeps files apart when the ordering gives them the same
        # key; without it they would share an x position and overwrite it.
        sort_key = (*sort_key, path.name)

        for selection in selections:
            raw, address, error = read_selection(sheet, selection)
            if error is not None and len(opened) > 1:
                # The data may be behind a cover sheet.
                for other in opened[1:]:
                    other_raw, other_address, other_error = read_selection(other, selection)
                    if other_error is None:
                        raw, address, error = other_raw, other_address, None
                        result.notes.append(
                            f"{path.name} -> {selection.name}: found on sheet '{other.title}'"
                        )
                        break
            if error:
                result.warnings.append(f"{path.name} -> {selection.name}: {error}")
                if raw is None:
                    continue

            value, bad = to_number(raw)
            if bad:
                where = f" at {address}" if address else ""
                result.warnings.append(f"{path.name} -> {selection.name}: {bad}{where}")
                continue

            result.points.append(
                Point(
                    file=path,
                    series=selection.name,
                    sort_key=sort_key,
                    x_label=x_label,
                    value=value,
                )
            )

    return result, selections


def suggest_labels(files, sheet_name=None, sample_size=3, limit=25):
    """Labels worth offering the user. Returns (row labels, column headers)."""
    if not files:
        return [], []
    picks = {0, len(files) // 2, len(files) - 1}
    rows: list[str] = []
    headers: list[str] = []

    for index in sorted(picks)[:sample_size]:
        sheet, _ = open_sheet(files[index], sheet_name)
        if sheet is None:
            continue
        for (r, c), value in sorted(sheet.values.items()):
            if not _is_number(value):
                continue
            label = sheet.text_left_of(r, c)
            if label and label.casefold() not in _HEADER_WORDS and label not in rows:
                rows.append(label)
            header = sheet.text_above(r, c)
            if header and header.casefold() not in _HEADER_WORDS and header not in headers:
                headers.append(header)

    return rows[:limit], headers[:limit]
