"""
Build RSI_2026_Schedule.xlsx, the master editable spreadsheet.

Two sheets:
  - "Data": one row per student (name, title, mentor, field, pronunciation,
    email). This is a lookup table, not meant to be reformatted.
  - "Schedule": a timetable grid -- one row per time slot, one column-group
    per room (concurrent rooms side by side). You type/change a student's
    name in a room's Name cell; the Field/Title/Mentor cells next to it are
    formulas (VLOOKUP into Data) and update automatically. Reordering who
    presents when is just cutting/pasting names between cells.

The initial Name cells are pre-filled using the same mentor+field-aware
packing heuristic as before (see lib_scheduling.py), as a draft to hand-edit
from -- not a final answer. Feel free to drag names around; the formulas
will keep up.

After hand-editing the Schedule sheet, run 03_export_site_csv.py to refresh
data/schedule.csv, which is what index.html actually reads.

Rerun this script only if you want to throw away hand-edits and regenerate
a fresh algorithmic draft (e.g. the roster changed substantially).
"""
import csv
from pathlib import Path

from openpyxl import Workbook
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

from lib_scheduling import (
    BLOCK_LABELS, BLOCKS, DAYS, FIELD_COLORS_HEX, LUNCH_LABEL, ROOMS,
    SECTION_CAPACITY, build_mentor_blocks, pack, slot_times,
)

ROOT = Path(__file__).resolve().parent.parent
ROSTER_PATH = ROOT / "data" / "roster.csv"
OUT_PATH = ROOT / "RSI_2026_Schedule.xlsx"

DATA_HEADERS = ["Full Name", "Title", "Mentor", "Field", "First Pron", "Last Pron", "Email", "Kerb"]

# Schedule-sheet column layout: Time, then 4 columns per room (Name/Field/Title/Mentor).
ROOM_COL_START = {room: 2 + 4 * i for i, room in enumerate(ROOMS)}  # 32-124->B, 32-141->F, 32-155->J
SUBHEADERS = ["Name", "Field", "Title", "Mentor"]

HEADER_FILL = PatternFill("solid", fgColor="4F46E5")
DAY_FILL = PatternFill("solid", fgColor="1E293B")
BLOCK_FILL = PatternFill("solid", fgColor="E2E8F0")
LUNCH_FILL = PatternFill("solid", fgColor="FEF3C7")
SUBHEADER_FILL = PatternFill("solid", fgColor="F1F5F9")
THIN = Side(style="thin", color="CBD5E1")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


def load_roster():
    with open(ROSTER_PATH, newline="", encoding="utf-8") as f:
        return list(csv.DictReader(f))


def build_data_sheet(wb, students):
    ws = wb.active
    ws.title = "Data"
    ws.append(DATA_HEADERS)
    for cell in ws[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = HEADER_FILL
        cell.alignment = Alignment(horizontal="center")
    for s in students:
        ws.append([
            s["full_name"], s["title"], s["mentor"], s["field"],
            s["first_pron"], s["last_pron"], s["email"], s["kerb"],
        ])
    widths = [24, 46, 26, 26, 20, 20, 22, 14]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:H{len(students) + 1}"
    return len(students)


def merge_row(ws, row, col_start, col_end, value, font=None, fill=None, align=None):
    ws.merge_cells(start_row=row, start_column=col_start, end_row=row, end_column=col_end)
    cell = ws.cell(row=row, column=col_start, value=value)
    if font:
        cell.font = font
    if fill:
        cell.fill = fill
    cell.alignment = align or Alignment(horizontal="center", vertical="center")
    return cell


def build_schedule_sheet(wb, sections):
    ws = wb.create_sheet("Schedule")
    total_cols = 1 + 4 * len(ROOMS)

    merge_row(ws, 1, 1, total_cols, "RSI 2026 Symposium Schedule",
              font=Font(bold=True, size=16, color="1E293B"))
    merge_row(ws, 2, 1, total_cols,
              "Edit a Name cell to reassign a slot -- Field / Title / Mentor update automatically.",
              font=Font(italic=True, size=10, color="64748B"))

    row = 4
    name_ranges = []  # (column_letter, first_row, last_row) for data validation

    sections_by_key = {(s["day"]["date"], s["block"], s["room"]): s["students"] for s in sections}

    for day in DAYS:
        row += 1
        merge_row(ws, row, 1, total_cols, day["display"],
                  font=Font(bold=True, size=13, color="FFFFFF"), fill=DAY_FILL)
        row += 1

        for block in BLOCKS:
            merge_row(ws, row, 1, total_cols, BLOCK_LABELS[block],
                      font=Font(bold=True, size=11, color="1E293B"), fill=BLOCK_FILL)
            row += 1

            # Room header row (merged over each room's 4 columns)
            ws.cell(row=row, column=1, value="Time").font = Font(bold=True)
            ws.cell(row=row, column=1).fill = SUBHEADER_FILL
            for room in ROOMS:
                start_col = ROOM_COL_START[room]
                merge_row(ws, row, start_col, start_col + 3, f"Room {room}",
                          font=Font(bold=True, color="FFFFFF"), fill=HEADER_FILL)
            row += 1

            # Sub-header row
            for room in ROOMS:
                start_col = ROOM_COL_START[room]
                for j, label in enumerate(SUBHEADERS):
                    c = ws.cell(row=row, column=start_col + j, value=label)
                    c.font = Font(bold=True, size=9)
                    c.fill = SUBHEADER_FILL
                    c.alignment = Alignment(horizontal="center")
            row += 1

            capacity = SECTION_CAPACITY[block]
            times = slot_times(block, capacity)
            first_data_row = row

            for i in range(capacity):
                start_label, end_label = times[i]
                ws.cell(row=row, column=1, value=f"{start_label} – {end_label}").font = Font(size=9)
                for room in ROOMS:
                    start_col = ROOM_COL_START[room]
                    name_col_letter = get_column_letter(start_col)
                    roster_here = sections_by_key.get((day["date"], block, room), [])
                    prefill = roster_here[i]["full_name"] if i < len(roster_here) else ""
                    name_cell = ws.cell(row=row, column=start_col, value=prefill)
                    name_cell.border = BORDER
                    field_cell = ws.cell(
                        row=row, column=start_col + 1,
                        value=f'=IFERROR(VLOOKUP(${name_col_letter}{row},Data!$A:$D,4,FALSE),"")')
                    title_cell = ws.cell(
                        row=row, column=start_col + 2,
                        value=f'=IFERROR(VLOOKUP(${name_col_letter}{row},Data!$A:$D,2,FALSE),"")')
                    mentor_cell = ws.cell(
                        row=row, column=start_col + 3,
                        value=f'=IFERROR(VLOOKUP(${name_col_letter}{row},Data!$A:$D,3,FALSE),"")')
                    for c in (field_cell, title_cell, mentor_cell):
                        c.border = BORDER
                        c.alignment = Alignment(wrap_text=True, vertical="center")
                        c.font = Font(size=9)
                    title_cell.alignment = Alignment(wrap_text=True, vertical="center")
                row += 1

            last_data_row = row - 1
            for room in ROOMS:
                name_col_letter = get_column_letter(ROOM_COL_START[room])
                name_ranges.append((name_col_letter, first_data_row, last_data_row))

            if block == "AM":
                merge_row(ws, row, 1, total_cols, LUNCH_LABEL,
                          font=Font(bold=True, italic=True, color="92400E"), fill=LUNCH_FILL)
                row += 1
        row += 1  # gap between days

    # Column widths
    ws.column_dimensions["A"].width = 20
    for room in ROOMS:
        start_col = ROOM_COL_START[room]
        ws.column_dimensions[get_column_letter(start_col)].width = 22      # Name
        ws.column_dimensions[get_column_letter(start_col + 1)].width = 15  # Field
        ws.column_dimensions[get_column_letter(start_col + 2)].width = 42  # Title
        ws.column_dimensions[get_column_letter(start_col + 3)].width = 22  # Mentor

    # Dropdown of valid names on every Name cell, sourced from Data sheet.
    num_students = ws.parent["Data"].max_row - 1
    dv = DataValidation(type="list", formula1=f"=Data!$A$2:$A${num_students + 1}",
                         allow_blank=True, showDropDown=False)
    ws.add_data_validation(dv)
    for col_letter, first_row, last_row in name_ranges:
        dv.add(f"{col_letter}{first_row}:{col_letter}{last_row}")

    # Conditional formatting: color each Field cell by subject, across the
    # whole sheet, so it stays correct even after edits.
    field_cols = [get_column_letter(ROOM_COL_START[room] + 1) for room in ROOMS]
    max_row = row + 5
    for field_name, (bg, fg) in FIELD_COLORS_HEX.items():
        fill = PatternFill("solid", fgColor=bg)
        font = Font(color=fg, size=9)
        for col_letter in field_cols:
            ws.conditional_formatting.add(
                f"{col_letter}4:{col_letter}{max_row}",
                CellIsRule(operator="equal", formula=[f'"{field_name}"'], fill=fill, font=font),
            )

    ws.sheet_view.showGridLines = False
    return ws


def main():
    students = load_roster()
    field_blocks = build_mentor_blocks(students)
    used_sections, unused_sections = pack(field_blocks)

    wb = Workbook()
    n = build_data_sheet(wb, students)
    build_schedule_sheet(wb, used_sections)

    wb.save(OUT_PATH)
    print(f"Wrote {OUT_PATH} ({n} students, {len(used_sections)} sections)")
    if unused_sections:
        print(f"{len(unused_sections)} sections left empty in the draft:")
        for s in unused_sections:
            print(f"  {s['day']['label']} {s['block']} {s['room']}")


if __name__ == "__main__":
    main()
