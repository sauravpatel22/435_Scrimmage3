# 435_Scrimmage3

## Authors

Saurav Patel & Vedant Patel

## Purpose

A local, offline Python tool that lets a user point at a folder of similarly
structured Excel (`.xlsx`) files, select one or more values by symbolic label
(e.g. `TEMP`) or cell reference (e.g. `A1`), extract those values across every
file in the folder, and generate a visualization of how they trend across the
dataset (e.g. daily weather readings, weekly stock prices, or student
performance across sessions). The tool is built to tolerate real-world
messiness in the input files: labels that shift position between files,
missing or invalid values, and extra columns/rows, without stopping the run.
Each run produces one output file: a PNG chart, rendered with matplotlib
directly from the extracted values. Warnings from the run (missing labels,
invalid values, ordering fallbacks) are printed to the console, not written
to disk.

## File structure and description so far

```text
scrimmage3/
├── README.md                  This file
├── main.py                    Entry point; CLI prompts/flags and workflow orchestration (implemented)
├── spreadsheet_processor.py   Discovery, label/cell resolution, ordering, extraction, validation (implemented)
├── visualizer.py              Chart generation - line/scatter/bar, matplotlib PNG only (implemented)
├── models.py                  Shared data models: selectors, run config, observations, warnings (implemented; stdlib only, no external dependencies)
├── requirements.txt           Python dependencies: openpyxl, matplotlib
│
├── input/
│   └── user_files/            Where a user drops their own .xlsx files to analyze (empty, gitignored contents)
│
├── output/
│   └── graphs/                The only output: one PNG chart per run
│
├── test_data/                 Synthetic .xlsx fixtures used for development and testing (10 files per folder)
│   ├── valid/                 Clean, standard-layout files (weather_01-10.xlsx)
│   ├── missing_values/        8 clean files plus 2 with a missing label or blank value
│   ├── shifted_labels/        10 genuinely different table positions/row orders, all label-searchable
│   ├── invalid_data/          10 files, each breaking TEMP/HUMIDITY a different way (blank, bool, "N/A", stray text, uncached formula, ...)
│   └── mixed_structure/       10 files with different harmless structural noise (extra rows/columns, title row, mixed case, whitespace, blank row, leading ID column)
│
└── tests/                     Unit tests (implemented, 43 tests, run against test_data/ fixtures)
    ├── __init__.py             Puts the project root on sys.path so `import models` etc. work under discovery
    ├── test_file_loading.py    discover_files, load_workbook_safe, select_worksheet
    ├── test_label_resolution.py  Selector.from_input, resolve_selector (incl. shifted labels, duplicates)
    ├── test_extraction.py      extract_dataset against every fixture folder
    └── test_validation.py      normalize_numeric_value, RunConfig/Selector validation
```

