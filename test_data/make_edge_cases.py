"""
make_edge_cases.py - builds the awkward fixtures in test_data/edge_cases/
and test_data/broken/.

Run it from the project root to regenerate them:

    python3 test_data/make_edge_cases.py

Each file is deliberately unpleasant in exactly one way, so that when a test
fails you know which kind of nastiness caused it. Keeping the generator in
the repo means anyone can see what is inside a fixture without opening Excel,
and can add a new flavour of broken file in one place.
"""

from datetime import date, datetime
from pathlib import Path

import openpyxl
from openpyxl.utils import get_column_letter

EDGE = Path(__file__).resolve().parent / "edge_cases"
BROKEN = Path(__file__).resolve().parent / "broken"
MULTI = Path(__file__).resolve().parent / "multi_sheet"


def _save(folder: Path, name: str, rows, title="Sheet1"):
    workbook = openpyxl.Workbook()
    worksheet = workbook.active
    worksheet.title = title
    for row in rows:
        worksheet.append(list(row))
    workbook.save(folder / name)
    return workbook


def build():
    for folder in (EDGE, BROKEN, MULTI):
        folder.mkdir(parents=True, exist_ok=True)

    # --- values that are technically numbers, but awkward ------------------
    _save(EDGE, "numbers_as_text.xlsx", [("Property", "Value"), ("TEMP", "72"), ("HUMIDITY", " 45.5 ")])
    _save(EDGE, "scientific_notation.xlsx", [("TEMP", 7.2e1), ("HUMIDITY", 4.5e-2)])
    _save(EDGE, "negative_and_zero.xlsx", [("TEMP", -40), ("HUMIDITY", 0)])
    _save(EDGE, "very_large_numbers.xlsx", [("TEMP", 1.7e308), ("HUMIDITY", 12345678901234567890)])
    _save(EDGE, "currency_and_units.xlsx", [("TEMP", "72F"), ("HUMIDITY", "$1,234")])

    # --- labels that fight the selection syntax ---------------------------
    _save(EDGE, "label_with_colon.xlsx", [("Time: AM", 6), ("TEMP", 70)])
    _save(EDGE, "label_with_equals.xlsx", [("A=B", 42), ("TEMP", 71)])
    _save(EDGE, "label_with_asterisk.xlsx", [("TEMP*", 73), ("TEMP", 74)])
    _save(EDGE, "label_looks_like_cell.xlsx", [("Ticker", "Close"), ("T5", 12.5), ("CO2", 400), ("Q1", 88)])

    # --- labels that are the same text, written differently ----------------
    _save(EDGE, "case_variation.xlsx", [("temp", 75), ("Humidity", 46)])
    _save(EDGE, "padded_label.xlsx", [("  TEMP  ", 76), ("HUMIDITY ", 47)])
    _save(EDGE, "nonbreaking_space.xlsx", [("TEMP ", 77), ("HUMIDITY ", 48)])
    _save(EDGE, "unicode_labels.xlsx", [("TEMPÉRATURE", 21), ("温度", 22), ("🌡 TEMP", 23)])
    _save(EDGE, "very_long_label.xlsx", [("TEMP" + "_x" * 120, 79), ("TEMP", 80)])

    # --- shapes that are not the expected label/value pair -----------------
    _save(EDGE, "label_in_last_column.xlsx", [("Value", "TEMP")])          # nothing to the right
    _save(EDGE, "value_to_the_left.xlsx", [(72, "TEMP"), (45, "HUMIDITY")])
    _save(EDGE, "transposed.xlsx", [("TEMP", "HUMIDITY"), (81, 49)])        # labels across the top
    _save(EDGE, "duplicate_labels.xlsx", [("TEMP", 82), ("TEMP", 999), ("HUMIDITY", 50)])
    _save(EDGE, "empty_sheet.xlsx", [])
    _save(EDGE, "headers_only.xlsx", [("Property", "Value")])
    _save(EDGE, "blank_rows_between.xlsx", [("TEMP", 83), (), (), ("HUMIDITY", 51)])

    # --- big, sparse, and hidden ------------------------------------------
    wide = [["col%d" % i for i in range(60)], ["TEMP", 84] + [None] * 58]
    _save(EDGE, "very_wide.xlsx", wide)

    tall = [("Property", "Value")] + [(f"FILLER_{i}", i) for i in range(1500)] + [("TEMP", 85)]
    _save(EDGE, "very_tall.xlsx", tall)

    workbook = openpyxl.Workbook()
    worksheet = workbook.active
    worksheet.append(["TEMP", 86])
    worksheet.append(["HUMIDITY", 52])
    worksheet.row_dimensions[1].hidden = True          # hidden, but still real data
    workbook.save(EDGE / "hidden_row.xlsx")

    workbook = openpyxl.Workbook()
    worksheet = workbook.active
    worksheet["A1"] = "TEMP"
    worksheet["B1"] = 87
    worksheet.merge_cells("A2:B2")                      # merged cell below
    worksheet["A3"] = "HUMIDITY"
    worksheet["B3"] = 53
    workbook.save(EDGE / "merged_cells.xlsx")

    workbook = openpyxl.Workbook()
    worksheet = workbook.active
    worksheet["A1"] = "TEMP"
    worksheet["B1"] = "=1+1"                            # formula, never calculated
    workbook.save(EDGE / "uncached_formula.xlsx")

    # --- dates in every awkward form --------------------------------------
    _save(EDGE, "date_real.xlsx", [("DATE", date(2026, 5, 4)), ("TEMP", 88)])
    _save(EDGE, "date_datetime.xlsx", [("DATE", datetime(2026, 5, 5, 14, 30)), ("TEMP", 89)])
    _save(EDGE, "date_as_text.xlsx", [("DATE", "2026-05-06"), ("TEMP", 90)])
    _save(EDGE, "date_serial_number.xlsx", [("DATE", 46147), ("TEMP", 91)])
    _save(EDGE, "date_missing.xlsx", [("TEMP", 92)])

    # --- a grid whose headers are not where you would hope -----------------
    _save(EDGE, "grid_shifted_header.xlsx", [
        (None, None, None),
        (None, "Quiz", "Exam"),
        ("Alice", 70, 80),
        ("Bob", 60, 90),
    ])
    _save(EDGE, "grid_missing_student.xlsx", [
        ("Student", "Quiz", "Exam"),
        ("Alice", 71, 81),
    ])

    # --- several worksheets ------------------------------------------------
    workbook = openpyxl.Workbook()
    workbook.active.title = "Notes"
    workbook.active["A1"] = "nothing useful here"
    data = workbook.create_sheet("Readings")
    data.append(["TEMP", 93])
    data.append(["HUMIDITY", 54])
    workbook.create_sheet("Empty")
    workbook.save(MULTI / "two_sheets_01.xlsx")

    workbook = openpyxl.Workbook()
    workbook.active.title = "Notes"
    data = workbook.create_sheet("Readings")
    data.append(["TEMP", 94])
    data.append(["HUMIDITY", 55])
    workbook.save(MULTI / "two_sheets_02.xlsx")

    # --- files that are not really workbooks at all ------------------------
    (BROKEN / "zero_byte.xlsx").write_bytes(b"")
    (BROKEN / "actually_csv.xlsx").write_text("TEMP,72\nHUMIDITY,45\n")
    (BROKEN / "truncated.xlsx").write_bytes(b"PK\x03\x04broken-zip-header-only")
    _save(BROKEN, "fine.xlsx", [("TEMP", 95), ("HUMIDITY", 56)])   # one good file among them

    print(f"wrote {len(list(EDGE.glob('*.xlsx')))} edge-case files to {EDGE}")
    print(f"wrote {len(list(MULTI.glob('*.xlsx')))} multi-sheet files to {MULTI}")
    print(f"wrote {len(list(BROKEN.glob('*.xlsx')))} broken files to {BROKEN}")


if __name__ == "__main__":
    build()
