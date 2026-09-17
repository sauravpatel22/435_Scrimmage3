"""
visualizer.py

Turns an ExtractionResult (produced by spreadsheet_processor.extract_dataset)
into a single chart: a PNG image saved with matplotlib. This module works
directly off the extracted observations - there is no separate table/
dataframe step - since matplotlib only needs plain lists of x/y values to
plot a line, scatter, or bar chart.

This module only draws pictures. It does not read files or validate data -
by the time an ExtractionResult reaches here, spreadsheet_processor.py has
already decided what counts as a usable value.

External dependency: matplotlib.
"""

from __future__ import annotations

from datetime import date, datetime
from pathlib import Path
from typing import Any, List, Tuple

import matplotlib

# Use a non-interactive backend: this tool runs from the command line and
# must be able to save a chart to disk even when no display is available
# (e.g. over SSH or in a CI job). This must be set before pyplot is
# imported, since the backend can't be changed afterward.
matplotlib.use("Agg")

import matplotlib.pyplot as plt

from models import ChartType, ExtractionResult, RunConfig


def _sortable_key(order_key: Any) -> Tuple[int, Any]:
    """
    Build a key usable by sorted() even when different files produced
    order_keys of different, otherwise-incomparable types (e.g. a date on
    one file and a fallback filename string on another, per the warning
    paths in determine_order_key). Grouping by type name first keeps same-
    typed keys sorted correctly relative to each other, while still
    producing a total order across the whole dataset instead of raising
    "'<' not supported between instances of 'str' and 'datetime.date'".
    """
    return (0 if isinstance(order_key, (date, datetime)) else 1, str(type(order_key)), order_key)


def _series_points(result: ExtractionResult, series_name: str) -> Tuple[List[Any], List[float]]:
    """Return one series' (x, y) points sorted by order_key, ready to plot."""
    observations = sorted(
        result.observations_for(series_name), key=lambda obs: _sortable_key(obs.order_key)
    )
    x_values = [obs.order_key for obs in observations]
    y_values = [obs.value for obs in observations]
    return x_values, y_values


def _shared_x_domain(result: ExtractionResult) -> List[Any]:
    """
    The union of every order_key seen across all series, sorted once. Used
    only for the bar chart, where every series needs to agree on the same
    x-axis positions to be grouped correctly.
    """
    unique_keys = {obs.order_key for obs in result.observations}
    return sorted(unique_keys, key=_sortable_key)


def create_chart(result: ExtractionResult, config: RunConfig, run_id: str) -> Path:
    """
    Render the extracted dataset as a PNG chart under
    config.output_folder / "graphs" and return the path that was written.
    """
    graphs_dir = config.output_folder / "graphs"
    graphs_dir.mkdir(parents=True, exist_ok=True)
    png_path = graphs_dir / f"{run_id}_chart.png"

    chart_type = config.chart_type
    series_names = result.series_names

    figure, axes = plt.subplots(figsize=(10, 6))

    if chart_type == ChartType.BAR:
        # Grouped bar chart: one cluster of bars per x position, one bar
        # per series within each cluster.
        x_domain = _shared_x_domain(result)
        bar_width = 0.8 / max(len(series_names), 1)
        for index, series_name in enumerate(series_names):
            x_values, y_values = _series_points(result, series_name)
            positions = [x_domain.index(x) + index * bar_width for x in x_values]
            axes.bar(positions, y_values, width=bar_width, label=series_name)
        axes.set_xticks(
            [i + bar_width * (len(series_names) - 1) / 2 for i in range(len(x_domain))]
        )
        axes.set_xticklabels([str(x) for x in x_domain], rotation=45, ha="right")
    else:
        # LINE and SCATTER both plot one series at a time over its own
        # sorted points; missing values are simply absent points, which
        # naturally shows as a break in the line rather than a fabricated
        # zero or an interpolated value.
        for series_name in series_names:
            x_values, y_values = _series_points(result, series_name)
            if chart_type == ChartType.SCATTER:
                axes.scatter(x_values, y_values, label=series_name)
            else:
                axes.plot(x_values, y_values, marker="o", label=series_name)
        figure.autofmt_xdate()

    axes.set_xlabel("Observation")
    axes.set_ylabel("Value")
    axes.set_title("Spreadsheet Visualizer - Extracted Series")
    axes.legend()
    figure.tight_layout()
    figure.savefig(png_path)
    plt.close(figure)

    return png_path
