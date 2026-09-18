# 435_Scrimmage3

## Authors

Saurav Patel & Vedant Patel

## Purpose

A local, offline Python tool for graphing values across a folder of similarly
structured Excel (`.xlsx`) files. Point it at the folder, say what you want to
track - a label like `TEMP`, a cell like `B2`, or a row and column like
`Bob:Exam` - and it pulls that value out of every file and saves a chart of
how it changes across the set.

It is built for the messy reality of real spreadsheets. Labels are searched
for rather than assumed to sit at a fixed address, so tables can move between
files. A file with a missing label, a blank cell or text where a number should
be is reported and skipped, never fatal. 500 files read in about 3 seconds.

## Quick start

```bash
python3 -m venv .venv && source .venv/bin/activate
pip install -r requirements.txt

python3 main.py                                  # guided: asks what it needs
python3 main.py --input-folder test_data/valid \
    --series TEMP HUMIDITY --order filename --chart line
```

See [instructions.md](instructions.md) for the full walk-through.

## What you can ask for

| You type | You get |
| --- | --- |
| `TEMP` | The value to the right of the label `TEMP` |
| `TEMP, HUMIDITY` | Both, as two series on one chart |
| `B2` | Whatever is in cell B2 of every file |
| `Bob:Exam` | Where row `Bob` meets column `Exam` - for grid-shaped sheets |
| `Bob:*` | Every column on Bob's row, one series each |
| `*:Exam` | Every row's Exam value - one series per student, ticker, etc. |
| `cell:T5` / `label:T5` | Force the reading, for names that look like addresses |
| `TEMP=Outside temp` | Any of the above, renamed in the legend |

Files can be ordered along the x-axis by filename, by a date in the filename,
by a date inside the sheet, or by an order you type yourself.

## How it fits together

```text
435_Scrimmage3/
├── main.py         The command line: flags, the guided prompts, the summary
├── sheets.py       Everything about spreadsheets: finding files, reading a
│                   worksheet once into memory, working out what a selection
│                   refers to, pulling values out, deciding the x-axis order
├── chart.py        Drawing: line, scatter and bar PNGs, one shared x-axis
├── requirements.txt    openpyxl and matplotlib - nothing else, nothing remote
│
├── input/user_files/   Somewhere to drop your own .xlsx files
├── output/graphs/      Generated charts, one timestamped PNG per run
│
├── test_data/          Fixtures used by the tests, and handy for demos
│   ├── valid/              10 clean, standard-layout files
│   ├── missing_values/     A missing label and a blank value
│   ├── shifted_labels/     10 different table positions and row orders
│   ├── invalid_data/       10 ways a value can be unusable
│   ├── mixed_structure/    Extra rows/columns, title rows, odd spacing
│   ├── grid/               Students x sessions, for row-by-column selection
│   ├── edge_cases/         33 files, each awkward in one specific way
│   ├── multi_sheet/        Data hiding behind a cover sheet
│   ├── broken/             Files that are not really workbooks at all
│   └── make_edge_cases.py  Regenerates the three folders above
│
└── tests/              82 tests, run against the fixtures above
    ├── test_sheets.py              Reading, selecting, ordering
    ├── test_extract_and_chart.py   Whole runs, and what lands on the chart
    ├── test_robustness.py          The "try to break it" tests
    └── test_main.py                The command line layer
```

Three rules keep the code honest, and the tests check each of them:

1. **Labels are searched for, never assumed.** A table that moves between
   files still resolves.
2. **Bad data warns, it never stops the run.** One unreadable file does not
   cost you the other 499.
3. **Every series shares one x-axis.** A file missing a value leaves a visible
   gap rather than shifting that series out of step with the others.

`tests/test_robustness.py` is where those rules are attacked on purpose:
labels that look like cell addresses, labels containing `:` or `=`, values of
infinity, sheets with no data, workbooks that are not workbooks, data behind a
cover sheet, and dates written five different ways. Add a new nasty file to
`test_data/make_edge_cases.py` when you think of another way to break it.

## Working on it

```bash
python3 -m unittest discover -s tests -t . -v   # 82 tests
python3 test_data/make_edge_cases.py            # rebuild the awkward fixtures
```
