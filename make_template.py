"""
Creates the blank Excel workbook that attendance.py fills in.

Run this once to produce a fresh, empty tracker:

    python make_template.py attendance.xlsx

The file it produces is the one to keep in the repository as a template. Do not
put real student names in that copy.
"""

import sys

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

HEADER_FILL = PatternFill("solid", fgColor="D9D9D9")
INPUT_FILL = PatternFill("solid", fgColor="FFF2CC")   # cells a person may edit
HEADER_FONT = Font(name="Arial", bold=True)
BODY_FONT = Font(name="Arial")
NOTE_FONT = Font(name="Arial", italic=True, color="808080")
TITLE_FONT = Font(name="Arial", bold=True, size=12)


def add_sheet(workbook, title, headers, widths, note=None):
    sheet = workbook.create_sheet(title)
    for column, heading in enumerate(headers, start=1):
        cell = sheet.cell(1, column, heading)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = Alignment(horizontal="center")
        sheet.column_dimensions[get_column_letter(column)].width = widths[column - 1]
    sheet.freeze_panes = "A2"
    if note:
        cell = sheet.cell(1, len(headers) + 2, note)
        cell.font = NOTE_FONT
    return sheet


def build(path):
    workbook = Workbook()
    workbook.remove(workbook.active)

    # ---- Students -------------------------------------------------------
    # Rebuilt by the program on every run. Every figure here is an Excel
    # formula reading from the Log sheet, so it can be audited by clicking it.
    add_sheet(
        workbook,
        "Students",
        ["Student", "Grade", "Code", "Threshold",
         "Total Occurrences", "Detentions Earned", "Count Toward Next"],
        [26, 8, 8, 11, 18, 18, 18],
        note="Rebuilt automatically. Do not type in this sheet.",
    )

    # ---- Log ------------------------------------------------------------
    # The one sheet that holds real data. Everything else is derived from it.
    log = add_sheet(
        workbook,
        "Log",
        ["Date", "Student", "Grade", "Code", "Period", "Raw Entry"],
        [12, 26, 8, 8, 9, 12],
        note="The record of truth. To undo a mistake, delete the row and re-run the program.",
    )
    log.cell(2, 8, "Written by the program. Rows appear here as reports are recorded.").font = NOTE_FONT

    # ---- Detentions -----------------------------------------------------
    add_sheet(
        workbook,
        "Detentions",
        ["Date Earned", "Student", "Grade", "Code", "Detention #", "Assigned By", "Served"],
        [13, 26, 8, 8, 13, 14, 10],
        note="Columns F and G are yours to fill in. Columns A-E are rebuilt each run.",
    )

    # ---- Settings -------------------------------------------------------
    settings = workbook.create_sheet("Settings")
    settings["A1"] = "Which codes earn a detention"
    settings["A1"].font = TITLE_FONT
    settings.merge_cells("A1:D1")

    headers = ["Code", "Description", "Counts? (YES/NO)", "Detention Every"]
    for column, heading in enumerate(headers, start=1):
        cell = settings.cell(2, column, heading)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = Alignment(horizontal="center")

    # Only L is switched on. The others are listed so that turning one on is a
    # matter of changing NO to YES -- no programming required.
    rows = [
        ("L", "Late / tardy", "YES", 3),
        ("T", "Tardy to class", "NO", 3),
        ("X", "Unexcused absence", "NO", 3),
        ("W", "Withdrawn", "NO", ""),
        ("M", "Medical", "NO", ""),
        ("I", "In-school", "NO", ""),
    ]
    for offset, (code, description, counts, every) in enumerate(rows):
        row = 3 + offset
        settings.cell(row, 1, code).font = BODY_FONT
        settings.cell(row, 2, description).font = BODY_FONT
        for column, value in ((3, counts), (4, every)):
            cell = settings.cell(row, column, value)
            cell.font = BODY_FONT
            cell.fill = INPUT_FILL
            cell.alignment = Alignment(horizontal="center")

    for column, width in zip("ABCD", (10, 26, 18, 18)):
        settings.column_dimensions[column].width = width

    note = settings.cell(11, 1,
        "Shaded cells are the only ones to edit. Descriptions above are placeholders "
        "supplied by the developer -- replace them with your own code meanings.")
    note.font = NOTE_FONT

    # ---- Read Me --------------------------------------------------------
    readme = workbook.create_sheet("Read Me")
    lines = [
        ("How this workbook works", TITLE_FONT),
        ("", BODY_FONT),
        ("Log        Every absence code ever recorded. One row = one student, "
         "one code, one period, one date. This is the only sheet holding real data.", BODY_FONT),
        ("Students   Running totals. Every cell is a formula reading the Log sheet. "
         "Rebuilt automatically -- do not type here.", BODY_FONT),
        ("Detentions Every detention earned, with the date it was earned. "
         "Fill in 'Assigned By' and 'Served' yourself.", BODY_FONT),
        ("Settings   Which codes count, and how many it takes. Edit the shaded cells.", BODY_FONT),
        ("", BODY_FONT),
        ("A Log row looks like this:", BODY_FONT),
        ("    Date        Student            Grade   Code   Period   Raw Entry", NOTE_FONT),
        ("    2026-09-14  Ellison, Margot    8       L      0        L-", NOTE_FONT),
        ("", BODY_FONT),
        ("To fix a mistake: delete the offending row from the Log sheet, save, "
         "and run the program again. Everything else recalculates.", BODY_FONT),
        ("", BODY_FONT),
        ("This workbook holds student data. Keep it where student records are "
         "normally kept, and do not email or upload it.", BODY_FONT),
    ]
    for offset, (text, font) in enumerate(lines, start=1):
        cell = readme.cell(offset, 1, text)
        cell.font = font
    readme.column_dimensions["A"].width = 110

    workbook.save(path)
    print(f"Created {path}")


if __name__ == "__main__":
    build(sys.argv[1] if len(sys.argv) > 1 else "attendance.xlsx")
