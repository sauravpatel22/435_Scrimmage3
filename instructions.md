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
   track, comma-separated (e.g. `TEMP, HUMIDITY`), or the number(s) shown
   next to them in the detected list above (e.g. `1, 2`). You can also type
   an exact cell reference instead of a name (e.g. `B2`) if you know
   exactly where the value lives.
4. **How should observations be ordered?** - choose how the files should be
   lined up along the x-axis:
   - `1` Current order (alphabetical by filename)
   - `2` Date created (from the filesystem)
   - `3` Custom order - you'll be shown a numbered list of the discovered
     files and asked to type the order you want as numbers, e.g. `3, 1, 4, 2`
5. **Select graph type** - `1` Line, `2` Scatter, or `3` Bar.
6. When it finishes, it prints a summary - including any warnings from the
   run - and tells you where the PNG chart was saved.

### Non-interactive mode (flags only, no prompts)

Useful for repeat runs or scripting. Any flag you omit falls back to an
interactive prompt for just that piece of information.

```bash
python3 main.py \
  --input-folder ./weather_data \
  --series TEMP HUMIDITY \
  --order filename \
  --chart line
```

Available flags:

| Flag | Meaning |
| --- | --- |
| `--input-folder PATH` | Folder of `.xlsx` files to process |
| `--output-folder PATH` | Where to write the graph (default: `output/`) |
| `--series NAME [NAME ...]` | One or more labels or cell references to extract |
| `--order {filename,file_creation_date,custom}` | Ordering method |
| `--custom-order NAME [NAME ...]` | Filenames in the exact order to use (required with `--order custom` in non-interactive mode) |
| `--chart {line,scatter,bar}` | Chart type |
| `--sheet NAME` | Worksheet name, if it isn't the first/only sheet |

### Try it on the bundled sample data

```bash
python3 main.py --input-folder test_data/valid --series TEMP HUMIDITY \
    --order filename --chart line
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

The only output file is the graph, written under `output/graphs/` (or your
`--output-folder`), timestamped so repeated runs never overwrite each other:

| File | Contents |
| --- | --- |
| `output/graphs/<run-id>_chart.png` | The chart, as a static image |

Everything else - files processed, data points extracted, and any warnings
(missing labels, invalid values, ordering fallbacks) - is printed to the
console when the run finishes. Nothing else is written to disk.

## 6. Troubleshooting

- **"no .xlsx files found"** - double-check the folder path, and that the
  files end in `.xlsx` (not `.xls` or `.csv`).
- **A series shows 0/N files found** - the label text must match exactly
  (case-insensitive, extra spaces are ignored). Check the "Available labels
  detected" list, or open one file to confirm the exact spelling.
- **Ordering by "date created" looks wrong** - this uses the file's
  filesystem creation time, which changes if a file is copied or
  re-downloaded. If that isn't reliable for your files, use
  `--order filename` or `--order custom` instead.
- **Want to rerun tests to confirm everything still works?**
  ```bash
  python3 -m unittest discover -s tests -t . -v
  ```
