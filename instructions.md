# Instructions

How to use the Spreadsheet Visualizer: point it at a folder of `.xlsx` files,
say what you want to track, and get a graph of how it changes across the set.

## 1. What you need

- Python 3.9 or later.
- A folder of `.xlsx` files that are roughly the same shape. They do not have
  to be identical - see [section 4](#4-what-it-expects-of-your-files).
- No internet connection, at any point.

## 2. Setup (once)

```bash
python3 -m venv .venv
source .venv/bin/activate          # Windows: .venv\Scripts\activate
python3 -m pip install -r requirements.txt
```

## 3. Running it

### Guided mode

```bash
python3 main.py
```

It asks four things, and re-asks anything you get wrong instead of quitting:

1. **Which folder?** Type or paste the path.
2. **What do you want to graph?** It lists the labels and column headers it
   found. Type names, or the numbers next to them, separated by commas. The
   full set of things you can type is in [section 5](#5-ways-to-say-what-you-want).
3. **What order should the files go in?** Filename, a date in the filename, a
   date inside the sheet, or an order you type yourself.
4. **Which chart?** Line, scatter or bar.

### Flags mode

Anything you pass as a flag is not asked about, so a full set of flags runs
with no prompts at all:

```bash
python3 main.py \
  --input-folder ./weather_data \
  --series TEMP HUMIDITY \
  --order date_in_filename \
  --chart line
```

| Flag | What it does |
| --- | --- |
| `--input-folder PATH` | The folder of `.xlsx` files |
| `--output-folder PATH` | Where the graph goes (default: `output/`) |
| `--series ... ` | What to graph; see [section 5](#5-ways-to-say-what-you-want) |
| `--order {filename,date_in_filename,date_in_sheet,custom}` | X-axis order |
| `--custom-order NAME ...` | Filenames in the order you want, with `--order custom` |
| `--chart {line,scatter,bar}` | Chart type |
| `--sheet NAME` | Worksheet name, if it isn't the first one |
| `--recursive` | Include `.xlsx` files in subfolders |
| `--title TEXT` | Chart title |
| `--xlabel TEXT` / `--ylabel TEXT` | Axis labels |
| `--right-axis SERIES ...` | Put these series on a second y-axis |

### Try it before using your own data

Two sets of sample files ship with the tool:

```bash
# Weather readings across 10 files
python3 main.py --input-folder test_data/valid --series TEMP HUMIDITY \
    --order filename --chart line

# One line per student, from grid-shaped sheets
python3 main.py --input-folder test_data/grid --series "*:Exam" \
    --order filename --chart line --title "Exam scores" --ylabel Score
```

## 4. What it expects of your files

- **Two shapes work.** Either a label with its value in the cell to its
  **right** (`TEMP | 72`), or a grid with row labels down the side and column
  headers across the top.
- **Labels can move.** Each file is searched for the label you asked for, so
  the table can sit anywhere and the rows can be in any order.
- **Extra rows, columns and stray notes are ignored.**
- **Nothing stops the run.** A file missing your label, or holding a blank or
  non-numeric value, is reported as a warning and skipped for that series.
- **Blank is never zero.** A missing value leaves a gap in the line; it is
  never quietly plotted as 0.
- **Excel lock files** (`~$...xlsx`) are skipped automatically.

## 5. Ways to say what you want

| You type | Meaning |
| --- | --- |
| `TEMP` | Find the label `TEMP`, take the value to its right |
| `B2` | Read cell B2 |
| `Bob:Exam` | Row `Bob` crossed with column `Exam` |
| `Bob:*` | Every column on Bob's row, one series per column |
| `*:Exam` | Every row's `Exam` value, one series per row |
| `cell:T5` | Force `T5` to mean the cell, not a label |
| `label:T5` | Force `T5` to mean a label, not the cell |
| `TEMP=Outside temp` | Any of the above, renamed in the legend |

If a name could be either a label or a cell address, such as `T5`, `Q1` or
`CO2`, an actual label of that name always wins. The run tells you which
reading it used, and `cell:` / `label:` override it.

## 6. What you get back

One PNG per run, under `output/graphs/`, named with a timestamp so runs never
overwrite each other. Everything else - how many files were read, how many had
data, and every warning - is printed when the run finishes. Repeated warnings
are grouped so a 500-file run stays readable.

If nothing was extracted, no chart is written and the run says so rather than
saving an empty picture.

## 7. If something looks wrong

- **"no .xlsx files found"** - check the path, and that the files really end
  in `.xlsx`. Add `--recursive` if they are in subfolders.
- **A series found nothing** - the label has to match the text in the cell
  (case and surrounding spaces do not matter). Check the list the tool prints,
  or open one file to confirm the spelling.
- **It read a cell when you meant a label** - use `label:NAME`.
- **The files are in the wrong order** - `filename` order is alphabetical, so
  `Jan10` sorts before `Jan5`. If your filenames carry dates, use
  `--order date_in_filename`; if the date is inside the sheet, use
  `--order date_in_sheet`; otherwise set the order yourself with
  `--order custom`.
- **Two series with very different ranges** - put one on its own axis with
  `--right-axis HUMIDITY`.

## 8. Known limits

- A value must sit to the **right** of its label, or at a row/column crossing.
  A sheet with the value to the left of the label, or with labels across the
  top and values underneath, is reported rather than guessed at.
- Values carrying units or symbols (`72F`, `$1,234`) are refused as
  non-numeric. Strip them in Excel first, or point at a plain number cell.
- Dates written as text, or as the numbers Excel stores dates as, are not
  recognised as dates; the run says so and falls back to a sortable value.
- Labels in a script your computer has no font for appear as empty boxes in
  the image. The numbers are still correct.
