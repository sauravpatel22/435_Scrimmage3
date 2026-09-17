# Instructions

How to use the Spreadsheet Visualizer: point it at a folder of similarly
structured `.xlsx` files, pick one or more values to track, and get a graph
of how those values change across the whole set.

## 1. Requirements

- Python 3.9 or later
- Your data as a folder of `.xlsx` files. Each file should have the same
  general layout (same worksheet, values organized the same way), but the
  tool tolerates some drift - see [What kind of spreadsheets this expects](#4-what-kind-of-spreadsheets-this-expects) below.
- No internet connection is required to run it or to view the generated
  graphs.

## 2. Setup (one time)

From the project folder:

```bash
python3 -m venv .venv
source .venv/bin/activate          # on Windows: .venv\Scripts\activate
python3 -m pip install -r requirements.txt
```

## 3. Running it

### Interactive mode (recommended for first-time use)

```bash
python3 main.py
```

You'll be walked through a short set of prompts:

1. **Enter folder containing Excel files** - type or paste the path to the
   folder holding your `.xlsx` files (for example, `./weather_data`).
2. **Available labels detected** - the tool peeks at the first few files
   and lists names it found (e.g. `TEMP`, `HUMIDITY`), as a hint. You are
   not limited to this list.
3. **Select one or more labels to graph** - type the name(s) you want to
   track, comma-separated (e.g. `TEMP, HUMIDITY`). You can also type an
   exact cell reference instead of a name (e.g. `B2`) if you know exactly
   where the value lives.
4. **How should observations be ordered?** - choose how the files should be
   lined up along the x-axis:
   - `1` Date inside the spreadsheet (looks for a DATE/DATETIME/TIMESTAMP
     label in each file)
   - `2` Date parsed from the filename (e.g. `weather_2026-01-15.xlsx`)
   - `3` File creation date (from the filesystem)
   - `4` File order (alphabetical by filename)
   - `5` Custom order - you'll be shown a numbered list of the discovered
     files and asked to type the order you want as numbers, e.g. `3, 1, 4, 2`
5. **Select graph type** - `1` Line, `2` Scatter, `3` Bar, or `4` Automatic
   (the tool picks Line for trend data, Bar if there's only one point per
   series).
6. When it finishes, it prints a summary and asks whether to open the
   interactive graph in your browser right away.

### Non-interactive mode (flags only, no prompts)

Useful for repeat runs or scripting. Any flag you omit falls back to an
interactive prompt for just that piece of information.

```bash
python3 main.py \
  --input-folder ./weather_data \
  --series TEMP HUMIDITY \
  --order filename \
  --chart auto
```

Available flags:

| Flag | Meaning |
| --- | --- |
| `--input-folder PATH` | Folder of `.xlsx` files to process |
| `--output-folder PATH` | Where to write results (default: `output/`) |
| `--series NAME [NAME ...]` | One or more labels or cell references to extract |
| `--order {filename,file_creation_date,date_in_filename,date_in_sheet,custom}` | Ordering method |
| `--custom-order NAME [NAME ...]` | Filenames in the exact order to use (required with `--order custom` in non-interactive mode) |
| `--chart {line,scatter,bar,auto}` | Chart type |
| `--sheet NAME` | Worksheet name, if it isn't the first/only sheet |
| `--open-graph` | Automatically open the interactive HTML chart when done |

### Try it on the bundled sample data

```bash
python3 main.py --input-folder test_data/valid --series TEMP HUMIDITY \
    --order filename --chart auto
```

## 4. What kind of spreadsheets this expects

- Each file should have a `Property` / `Value` style layout: a cell
  containing the name (e.g. `TEMP`), with the actual value in the cell
  immediately to its **right**, on the same row.
- The label does **not** need to be in the same row or cell across files -
  the tool searches each file for the label you asked for, so the table can
  shift position between files.
- Extra columns, extra rows, and unrelated data elsewhere in the sheet are
  ignored.
- A file missing your chosen label, or with a blank/non-numeric value, does
  **not** stop the run - that file is skipped for that series and reported
  as a warning, and everything else keeps processing.
- Files with a `~$` prefix (Excel's temporary lock files) are automatically
  skipped.

## 5. What you get back

Every run writes to a timestamped set of files under `output/` (or your
`--output-folder`), so repeated runs never overwrite each other:

| File | Contents |
| --- | --- |
| `output/graphs/<run-id>_chart.png` | Static image of the chart |
| `output/graphs/<run-id>_chart.html` | Interactive chart - open it in any browser, no internet needed |
| `output/csv/<run-id>_extracted_data.csv` | Every extracted data point: source file, sheet, series, cell, order, value |
| `output/reports/<run-id>_run_summary.txt` | Per-series found/missing/invalid counts and run settings |
| `output/reports/<run-id>_warnings.txt` | Every warning raised during the run (missing labels, invalid values, ordering fallbacks) |

## 6. Troubleshooting

- **"no .xlsx files found"** - double-check the folder path, and that the
  files end in `.xlsx` (not `.xls` or `.csv`).
- **A series shows 0/N files found** - the label text must match exactly
  (case-insensitive, extra spaces are ignored). Check the "Available labels
  detected" list, or open one file to confirm the exact spelling.
- **Ordering by date looks wrong** - "date in filename" expects an
  `YYYY-MM-DD`-style date somewhere in the filename; "date in sheet" looks
  for a `DATE`, `DATETIME`, or `TIMESTAMP` label. If neither applies, use
  `--order filename` or `--order custom`.
- **Want to rerun tests to confirm everything still works?**
  ```bash
  python3 -m unittest discover -s tests -t . -v
  ```
