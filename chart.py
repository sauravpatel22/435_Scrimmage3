"""
chart.py - turning a sheets.Result into a PNG.

Every series is drawn against one shared x-axis built from all the files, not
just the files that series had a value in. That is what keeps the files in
order and what leaves a missing value as a visible gap.
"""

from __future__ import annotations

import warnings
from pathlib import Path
from typing import Optional

import matplotlib

matplotlib.use("Agg")   # no display needed; must precede the pyplot import

import matplotlib.pyplot as plt  # noqa: E402
from matplotlib import font_manager  # noqa: E402
from matplotlib.ticker import MaxNLocator  # noqa: E402

# matplotlib uses the first font here that exists and does not fall through
# per glyph, so wide-coverage fonts must precede its own default or non-Latin
# labels are drawn as empty boxes.
_FONT_CANDIDATES = (
    "Arial Unicode MS", "PingFang SC", "Hiragino Sans GB", "Heiti SC",
    "Noto Sans CJK SC", "Microsoft YaHei", "SimHei", "DejaVu Sans",
)
_INSTALLED = {font.name for font in font_manager.fontManager.ttflist}
plt.rcParams["font.sans-serif"] = [
    name for name in _FONT_CANDIDATES if name in _INSTALLED
] or ["DejaVu Sans"]

CHART_TYPES = ("line", "scatter", "bar")

_MAX_TICK_LABELS = 30   # above this, thin out the x labels
_DENSE_POINTS = 60      # above this, drop the per-point markers
_DATE_RANK = 0          # matches sheets._DATE_RANK


def x_axis(result) -> tuple[list[tuple], list[str]]:
    """The shared x-axis: every position any file produced, in order."""
    seen = {point.sort_key: point.x_label for point in result.points}
    keys = sorted(seen)
    return keys, [seen[key] for key in keys]


def align(result, series: str, keys) -> list[Optional[float]]:
    """This series' values lined up with the shared axis; None where it has none."""
    by_key = {point.sort_key: point.value for point in result.points_for(series)}
    return [by_key.get(key) for key in keys]


def draw(
    result,
    path: Path,
    kind: str = "line",
    title: str = "Spreadsheet Visualizer",
    xlabel: str = "Observation",
    ylabel: str = "Value",
    right_axis=(),
) -> Optional[Path]:
    """Save the chart and return its path, or None if there was nothing to draw."""
    if not result.points:
        return None

    path.parent.mkdir(parents=True, exist_ok=True)
    keys, labels = x_axis(result)
    right = {name for name in right_axis if name in result.series_names}

    # Real dates get a date axis, so uneven gaps in time show as uneven spacing.
    dated = all(key[0] == _DATE_RANK for key in keys)
    positions = [key[1] for key in keys] if dated else list(range(len(keys)))

    figure, axes = plt.subplots(figsize=(11, 6))
    right_axes = axes.twinx() if right else None
    drawn = []

    # A marker per point turns hundreds of files into a solid band.
    dense = len(keys) > _DENSE_POINTS
    marker = "" if dense else "o"
    width = 1.0 if dense else 1.8

    if kind == "bar":
        drawn += _bars(result, axes, right_axes, right, keys, range(len(keys)))
        positions = list(range(len(keys)))
        dated = False
    else:
        for index, series in enumerate(result.series_names):
            target = right_axes if series in right else axes
            values = align(result, series, keys)
            label = f"{series} (right)" if series in right else series
            if kind == "scatter":
                pairs = [(x, y) for x, y in zip(positions, values) if y is not None]
                drawn.append(
                    target.scatter(
                        [x for x, _ in pairs], [y for _, y in pairs],
                        label=label, color=f"C{index}", s=18 if dense else 36,
                    )
                )
            else:
                # matplotlib skips NaN, leaving the gap visible.
                drawn.append(
                    target.plot(
                        positions,
                        [float("nan") if v is None else v for v in values],
                        marker=marker, linewidth=width, label=label, color=f"C{index}",
                    )[0]
                )

    if dated:
        figure.autofmt_xdate()
    else:
        axes.set_xticks(list(range(len(keys))))
        axes.set_xticklabels(labels, rotation=45, ha="right")
        axes.set_xlim(-0.6, len(keys) - 0.4)   # keeps bars sane when there are few

        if len(keys) > _MAX_TICK_LABELS:
            axes.xaxis.set_major_locator(MaxNLocator(nbins=_MAX_TICK_LABELS, integer=True))

    axes.set_xlabel(xlabel)
    axes.set_ylabel(ylabel)
    axes.set_title(title)
    axes.grid(True, alpha=0.3)
    if right_axes is not None:
        right_axes.set_ylabel("Value (right axis)")

    # One legend for both axes, so a right-axis series is not left out.
    handles = [h for h in drawn if hasattr(h, "get_label")]
    if handles:
        axes.legend(handles, [h.get_label() for h in handles], loc="best")

    with warnings.catch_warnings():
        # An undrawable glyph shows as a box in the image; that is enough.
        warnings.filterwarnings("ignore", message=".*missing from font.*")
        figure.tight_layout()
        figure.savefig(path, dpi=120)
    plt.close(figure)
    return path


def _bars(result, axes, right_axes, right, keys, positions):
    """Grouped bars: one cluster per x position, one bar per series."""
    series_names = result.series_names
    width = 0.8 / max(len(series_names), 1)
    offsets = [i * width - 0.4 + width / 2 for i in range(len(series_names))]
    drawn = []

    for index, series in enumerate(series_names):
        target = right_axes if series in right else axes
        values = align(result, series, keys)
        xs = [p + offsets[index] for p, v in zip(positions, values) if v is not None]
        ys = [v for v in values if v is not None]
        label = f"{series} (right)" if series in right else series
        drawn.append(target.bar(xs, ys, width=width, label=label, color=f"C{index}"))

    return drawn
