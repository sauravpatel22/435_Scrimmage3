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
Each run produces reusable output: an interactive HTML graph, a static image,
a CSV of the extracted data, and a warnings/summary report.

## File structure and description so far

```text
scrimmage3/
├── README.md                  This file
├── main.py                    Entry point; user input/output and workflow orchestration (empty, not yet implemented)
├── spreadsheet_processor.py   Discovery, label/cell resolution, ordering, extraction, validation (empty, not yet implemented)
├── visualizer.py              Chart generation - line/scatter/bar, PNG + interactive HTML (empty, not yet implemented)
├── models.py                  Shared data models: selections, run config, observations, warnings (empty, not yet implemented)
├── requirements.txt           Python dependencies (pandas, numpy, matplotlib, etc.) (empty, not yet filled in)
│
├── input/
│   └── user_files/            Where a user drops their own .xlsx files to analyze (empty, gitignored contents)
│
├── output/                    Destination for generated run artifacts
│   ├── graphs/                Generated PNG and interactive HTML charts
│   ├── csv/                   Extracted/combined dataset as CSV
│   └── reports/                Run summaries and warnings
│
├── test_data/                 Synthetic .xlsx fixtures used for development and testing
│   ├── valid/                 Clean, standard-layout files (weather_01-03.xlsx)
│   ├── missing_values/        Files with a missing label or blank value
│   ├── shifted_labels/        Files where the label/value table has moved position
│   ├── invalid_data/          Files with non-numeric or blank measurement cells
│   └── mixed_structure/       Files with reordered rows or extra columns
│
└── tests/                     Unit tests (not yet implemented)
    ├── test_file_loading.py
    ├── test_label_resolution.py
    ├── test_extraction.py
    └── test_validation.py
```
