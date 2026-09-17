"""
models.py

Shared data contracts for the spreadsheet visualizer. Every other module
(spreadsheet_processor.py, visualizer.py, main.py) imports the types defined
here instead of passing around raw dicts/tuples, so that the shape of a
"selector", a "run configuration", an "extracted observation", or a
"warning" is defined in exactly one place.

"""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from enum import Enum
from pathlib import Path
from typing import Any, Dict, List, Optional

# Matches a standard Excel cell reference such as "A1", "c9", or "AA123".
# Used to tell "the user typed a cell address" apart from "the user typed
# a symbolic label like TEMP or $AAPL".
_CELL_REFERENCE_PATTERN = re.compile(r"^[A-Za-z]{1,3}[1-9][0-9]*$")


class SelectorType(Enum):
    """
    How the user identified the data they want to track.

    LABEL: a symbolic name that appears as a value somewhere in the sheet
           (e.g. "TEMP", "$AAPL"). The processor must search for it -
           its position can differ between files.
    CELL:  an explicit spreadsheet coordinate (e.g. "A1", "D4"). Still
           resolved defensively (see spreadsheet_processor.py), but the
           user is naming a position rather than a name.
    """

    LABEL = "label"
    CELL = "cell"


class OrderingMethod(Enum):
    """
    How observations across many files are placed on the x-axis /
    processing order, per the "Determine a Consistent X-Axis" design
    requirement.
    """

    FILENAME = "filename"                # sort by filename text
    FILE_CREATION_DATE = "file_creation_date"  # sort by filesystem metadata
    DATE_IN_FILENAME = "date_in_filename"      # parse a date out of the filename
    DATE_IN_SHEET = "date_in_sheet"            # read a date cell from inside each file
    CUSTOM = "custom"                          # user supplies an explicit order


class ChartType(Enum):
    """Visualization requested by the user, or AUTO to let the tool decide."""

    LINE = "line"
    SCATTER = "scatter"
    BAR = "bar"
    AUTO = "auto"


class WarningLevel(Enum):
    """Severity of a RunWarning, used when printing the run summary."""

    INFO = "info"
    WARNING = "warning"
    ERROR = "error"


@dataclass(frozen=True)
class Selector:
    """
    A single series the user wants extracted and graphed, e.g. "TEMP" or
    "B2". Immutable (frozen) because a selector is decided once during the
    guided prompts and should not be silently mutated afterward.
    """

    raw: str            # exactly what the user typed, e.g. "TEMP" or "b2"
    type: SelectorType  # LABEL or CELL, decided in __post_init__
    series_name: str    # display name used in legends/output; defaults to raw

    def __post_init__(self) -> None:
        # dataclass(frozen=True) blocks plain attribute assignment, so we
        # go through object.__setattr__ to fill in defaults derived from
        # other fields.
        if not self.raw or not self.raw.strip():
            raise ValueError("Selector.raw must not be empty")

        object.__setattr__(self, "raw", self.raw.strip())

        if not self.series_name:
            object.__setattr__(self, "series_name", self.raw)

    @staticmethod
    def from_input(raw_text: str) -> "Selector":
        """
        Build a Selector from one piece of raw user input, auto-detecting
        whether it looks like a cell reference (e.g. "A1") or a symbolic
        label (e.g. "TEMP"). This is the "Labels or Cells" design note:
        the tool should figure out intent rather than force the user to
        say which kind of selector they mean.
        """
        text = raw_text.strip()
        selector_type = (
            SelectorType.CELL
            if _CELL_REFERENCE_PATTERN.match(text)
            else SelectorType.LABEL
        )
        return Selector(raw=text, type=selector_type, series_name=text)

    @property
    def is_cell_reference(self) -> bool:
        return self.type == SelectorType.CELL


@dataclass
class RunConfig:
    """
    Everything needed to execute one end-to-end run: where the input files
    live, what to extract, how to order it, and how to chart it. Built
    once by the guided prompts in main.py and then passed down to
    spreadsheet_processor.py and visualizer.py unchanged.
    """

    input_folder: Path
    output_folder: Path
    selectors: List[Selector]
    ordering_method: OrderingMethod
    sheet_name: Optional[str] = None       # None = use each file's first/only sheet
    custom_order: Optional[List[str]] = None  # filenames in user-specified order
    chart_type: ChartType = ChartType.AUTO

    def __post_init__(self) -> None:
        if not self.selectors:
            raise ValueError("RunConfig requires at least one selector")

        # A CUSTOM ordering method is meaningless without the explicit
        # order to apply, so fail fast here instead of discovering the
        # problem partway through processing every file.
        if self.ordering_method == OrderingMethod.CUSTOM and not self.custom_order:
            raise ValueError(
                "ordering_method=CUSTOM requires a non-empty custom_order list"
            )


@dataclass
class Observation:
    """
    One extracted (file, series, value) data point. A full run produces a
    list of these, one per selector per successfully-read file; missing or
    invalid values are recorded as warnings (see RunWarning) rather than as
    an Observation with a fabricated value.
    """

    source_file: Path      # which .xlsx file this came from
    sheet_name: str        # which worksheet within that file
    series_name: str       # which Selector.series_name this belongs to
    cell: Optional[str]    # resolved cell address, e.g. "B2", if known
    order_key: Any         # value used to sort/position this point on the x-axis
    value: float           # the numeric value plotted on the y-axis


@dataclass
class RunWarning:
    """
    A non-fatal problem encountered during a run (missing label, invalid
    value, unparseable date, etc.). Per the "Missing / Invalid Data Should
    Not Stop the Program" design note, these are collected and reported
    instead of raising and aborting the run.
    """

    message: str
    source_file: Optional[Path] = None
    series_name: Optional[str] = None
    level: WarningLevel = WarningLevel.WARNING

    def __str__(self) -> str:
        # Human-readable form used when writing warnings.txt / printing
        # the run summary, e.g. "weather_2026-02-14.xlsx -> HUMIDITY: missing"
        location = self.source_file.name if self.source_file else "run"
        subject = f" -> {self.series_name}" if self.series_name else ""
        return f"{location}{subject}: {self.message}"


@dataclass
class SeriesSummary:
    """Per-series counts shown in the validation summary, e.g. the pptx's
    "TEMP found in 120 / 120 files" line."""

    series_name: str
    files_found: int = 0
    files_missing: int = 0
    files_invalid: int = 0

    @property
    def files_considered(self) -> int:
        return self.files_found + self.files_missing + self.files_invalid


@dataclass
class ExtractionResult:
    """
    The aggregate output of processing an entire folder: every successful
    Observation, every RunWarning raised along the way, and a per-series
    summary. spreadsheet_processor.py builds this; visualizer.py and
    main.py's reporting step both read from it.
    """

    observations: List[Observation] = field(default_factory=list)
    warnings: List[RunWarning] = field(default_factory=list)
    series_summaries: Dict[str, SeriesSummary] = field(default_factory=dict)

    def add_observation(self, observation: Observation) -> None:
        self.observations.append(observation)
        summary = self.series_summaries.setdefault(
            observation.series_name, SeriesSummary(series_name=observation.series_name)
        )
        summary.files_found += 1

    def add_warning(self, warning: RunWarning) -> None:
        self.warnings.append(warning)

        # Only bump the missing/invalid counters when the warning is tied
        # to a specific series; run-level warnings (e.g. "file could not
        # be opened") don't count against any one series' totals.
        if warning.series_name is None:
            return

        summary = self.series_summaries.setdefault(
            warning.series_name, SeriesSummary(series_name=warning.series_name)
        )
        if warning.level == WarningLevel.ERROR:
            summary.files_invalid += 1
        else:
            summary.files_missing += 1

    def observations_for(self, series_name: str) -> List[Observation]:
        """Convenience accessor used by visualizer.py to pull one series'
        points without re-filtering the full observation list each time."""
        return [obs for obs in self.observations if obs.series_name == series_name]
