"""
visualizer.py

Turns an ExtractionResult (produced by spreadsheet_processor.extract_dataset)
into charts. Per the "Automatically Produce Reusable Output" design note,
every run saves both a static image (matplotlib, .png) and a self-contained
interactive graph (Plotly, .html) that still works with no internet
connection - the Plotly JavaScript library is embedded directly in the file
rather than loaded from a CDN.

This module only draws pictures. It does not read files or validate data -
by the time an ExtractionResult reaches here, spreadsheet_processor.py has
already decided what counts as a usable value.

External dependencies: matplotlib (static PNG) and plotly (interactive
HTML).
"""

from __future__ import annotations

from datetime import date, datetime
from pathlib import Path
from typing import Any, Dict, List, Tuple

import matplotlib

# Use a non-interactive backend: this tool runs from the command line and
# must be able to save a chart to disk even when no display is available
# (e.g. over SSH or in a CI job). This must be set before pyplot is
# imported, since the backend can't be changed afterward.
matplotlib.use("Agg")

import matplotlib.pyplot as plt
import plotly.graph_objects as go

from models import ChartType, ExtractionResult, RunConfig


def choose_chart_type(result: ExtractionResult, requested: ChartType) -> ChartType:
    """
    Resolve ChartType.AUTO to a concrete chart type. A user-requested
    concrete type is always honored unchanged.

    Heuristic: this tool exists to show how values change across a set of
    files, so a genuine trend (more than one point per series on average)
    defaults to a line chart. A run that only produced one point per
    series - e.g. comparing a single day's reading across several tickers -
    reads better as a bar chart, since there is no line to draw between a
    single point per series.
    """
    if requested != ChartType.AUTO:
        return requested

    if not result.observations:
        return ChartType.LINE

    points_per_series = len(result.observations) / max(len(result.series_summaries), 1)
    return ChartType.LINE if points_per_series > 1 else ChartType.BAR


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
    The union of every order_key seen across all series, sorted once. Using
    one shared x-axis (instead of each series drawing its own) is what
    makes a missing value show up as a visible gap in that series' line
    rather than silently compressing the x-axis for just that series.
    """
    unique_keys = {obs.order_key for obs in result.observations}
    return sorted(unique_keys, key=_sortable_key)


def create_charts(
    result: ExtractionResult, config: RunConfig, run_id: str
) -> Dict[str, Path]:
    """
    Render the extracted dataset as both a PNG (matplotlib) and a
    self-contained interactive HTML file (Plotly), saved under
    config.output_folder / "graphs". Returns the paths that were written.
    """
    graphs_dir = config.output_folder / "graphs"
    graphs_dir.mkdir(parents=True, exist_ok=True)

    chart_type = choose_chart_type(result, config.chart_type)
    series_names = list(result.series_summaries.keys())
    x_domain = _shared_x_domain(result)

    png_path = graphs_dir / f"{run_id}_chart.png"
    html_path = graphs_dir / f"{run_id}_chart.html"

    _render_matplotlib(result, series_names, chart_type, png_path)
    _render_plotly(result, series_names, chart_type, x_domain, html_path)

    return {"png": png_path, "html": html_path}


def _render_matplotlib(
    result: ExtractionResult, series_names: List[str], chart_type: ChartType, png_path: Path
) -> None:
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


def _render_plotly(
    result: ExtractionResult,
    series_names: List[str],
    chart_type: ChartType,
    x_domain: List[Any],
    html_path: Path,
) -> None:
    figure = go.Figure()

    for series_name in series_names:
        x_values, y_values = _series_points(result, series_name)
        if chart_type == ChartType.BAR:
            figure.add_trace(go.Bar(x=x_values, y=y_values, name=series_name))
        elif chart_type == ChartType.SCATTER:
            figure.add_trace(
                go.Scatter(x=x_values, y=y_values, mode="markers", name=series_name)
            )
        else:
            # A line trace with connectgaps=False (the default) leaves a
            # visible break at any x position this series has no point
            # for, instead of drawing a straight line across the gap.
            figure.add_trace(
                go.Scatter(x=x_values, y=y_values, mode="lines+markers", name=series_name)
            )

    figure.update_layout(
        title="Spreadsheet Visualizer - Extracted Series",
        xaxis_title="Observation",
        yaxis_title="Value",
        template="plotly_white",
    )

    # include_plotlyjs=True embeds the full Plotly.js library inside this
    # one HTML file so the chart still renders with no network access,
    # per this tool's "processing/output is local" requirement.
    figure.write_html(str(html_path), include_plotlyjs=True, full_html=True)
