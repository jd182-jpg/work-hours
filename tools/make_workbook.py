#!/usr/bin/env python3
"""
Build the live work-hours workbook.

Creates a .xlsx with two sheets:

  Summary    Current pay period at a glance, then a month-by-month grid with
             the 1st-15th subtotal, the 16th-EOM subtotal, and the month total.
             Every number is a live formula driven off the Hours Log table, so
             the totals update themselves the moment a row lands.

  Hours Log  A table named HoursLog holding one row per day worked:
             Date, Day, Time In, Time Out, Hours, Period, Notes.
             The web app appends rows here through the Microsoft Graph API.

Usage:
    python3 make_workbook.py --year 2026 --out "Work Hours.xlsx"
"""

import argparse
import os

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo

# Leggett house palette: navy ground, blue accent.
NAVY = "0F2440"
BLUE = "2E6FB7"
LIGHT = "EAF1F9"
RULE = "C9D6E4"
GREY = "6B7A8C"

MONTHS = ["January", "February", "March", "April", "May", "June",
          "July", "August", "September", "October", "November", "December"]

HEADERS = ["Date", "Day", "Time In", "Time Out", "Hours", "Period", "Notes"]

thin = Side(style="thin", color=RULE)


def style_header(cell):
    cell.font = Font(bold=True, color="FFFFFF", size=11)
    cell.fill = PatternFill("solid", fgColor=NAVY)
    cell.alignment = Alignment(horizontal="center", vertical="center")
    cell.border = Border(bottom=thin)


def build_summary(ws, year):
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 22
    for col in "BCD":
        ws.column_dimensions[col].width = 16
    ws.column_dimensions["E"].width = 3

    ws["A1"] = "Work Hours"
    ws["A1"].font = Font(bold=True, size=18, color=NAVY)
    ws["A2"] = "Logged from the phone app. Totals below are live formulas."
    ws["A2"].font = Font(size=10, color=GREY, italic=True)
    ws.row_dimensions[1].height = 26

    # ---- Current pay period -------------------------------------------------
    ws["A4"] = "CURRENT PAY PERIOD"
    ws["A4"].font = Font(bold=True, size=10, color=BLUE)

    ws["A5"] = "Period start"
    ws["B5"] = ("=IF(DAY(TODAY())<=15,DATE(YEAR(TODAY()),MONTH(TODAY()),1),"
                "DATE(YEAR(TODAY()),MONTH(TODAY()),16))")
    ws["A6"] = "Period end"
    ws["B6"] = ("=IF(DAY(TODAY())<=15,DATE(YEAR(TODAY()),MONTH(TODAY()),15),"
                "EOMONTH(TODAY(),0))")
    ws["A7"] = "Hours this period"
    ws["B7"] = ('=SUMIFS(HoursLog[Hours],HoursLog[Date],">="&$B$5,'
                'HoursLog[Date],"<="&$B$6)')
    ws["A8"] = "Days worked"
    ws["B8"] = ('=COUNTIFS(HoursLog[Date],">="&$B$5,HoursLog[Date],"<="&$B$6)')

    for row in range(5, 9):
        ws.cell(row=row, column=1).font = Font(size=11, color=NAVY)
        ws.cell(row=row, column=2).font = Font(bold=True, size=11, color=NAVY)
    ws["B5"].number_format = "mmm d, yyyy"
    ws["B6"].number_format = "mmm d, yyyy"
    ws["B7"].number_format = "0.00"
    ws["B7"].fill = PatternFill("solid", fgColor=LIGHT)
    ws["B8"].number_format = "0"

    # ---- Month-by-month grid ------------------------------------------------
    top = 11
    ws.cell(row=top - 1, column=1, value="%s BY PAY PERIOD" % year).font = Font(
        bold=True, size=10, color=BLUE)

    for idx, label in enumerate(["Month", "1st - 15th", "16th - End", "Month Total"]):
        style_header(ws.cell(row=top, column=1 + idx, value=label))

    for i, name in enumerate(MONTHS):
        r = top + 1 + i
        m = i + 1
        first = "DATE(%d,%d,1)" % (year, m)
        ws.cell(row=r, column=1, value="%s %d" % (name, year))
        ws.cell(row=r, column=2, value=(
            '=SUMIFS(HoursLog[Hours],HoursLog[Date],">="&%s,'
            'HoursLog[Date],"<="&DATE(%d,%d,15))' % (first, year, m)))
        ws.cell(row=r, column=3, value=(
            '=SUMIFS(HoursLog[Hours],HoursLog[Date],">="&DATE(%d,%d,16),'
            'HoursLog[Date],"<="&EOMONTH(%s,0))' % (year, m, first)))
        ws.cell(row=r, column=4, value="=B%d+C%d" % (r, r))

        for c in range(1, 5):
            cell = ws.cell(row=r, column=c)
            cell.border = Border(bottom=thin)
            if c == 1:
                cell.font = Font(size=11, color=NAVY)
            else:
                cell.number_format = "0.00"
                cell.alignment = Alignment(horizontal="right")
                cell.font = Font(size=11, bold=(c == 4),
                                 color=NAVY if c == 4 else GREY)
        ws.cell(row=r, column=4).fill = PatternFill("solid", fgColor=LIGHT)

    total_row = top + 13
    ws.cell(row=total_row, column=1, value="YEAR TOTAL").font = Font(
        bold=True, size=11, color="FFFFFF")
    for c in range(1, 5):
        cell = ws.cell(row=total_row, column=c)
        cell.fill = PatternFill("solid", fgColor=NAVY)
        cell.font = Font(bold=True, size=11, color="FFFFFF")
    for c in range(2, 5):
        col = get_column_letter(c)
        cell = ws.cell(row=total_row, column=c,
                       value="=SUM(%s%d:%s%d)" % (col, top + 1, col, top + 12))
        cell.number_format = "0.00"
        cell.alignment = Alignment(horizontal="right")
        cell.fill = PatternFill("solid", fgColor=NAVY)
        cell.font = Font(bold=True, size=11, color="FFFFFF")

    ws.freeze_panes = "A%d" % (top + 1)


def build_log(ws):
    ws.sheet_view.showGridLines = False
    widths = [13, 8, 11, 11, 10, 12, 34]
    for i, w in enumerate(widths):
        ws.column_dimensions[get_column_letter(i + 1)].width = w

    ws.append(HEADERS)
    # A table needs at least one data row. The app fills this placeholder on the
    # first sync instead of appending, so no blank row is left behind.
    ws.append([None] * len(HEADERS))

    table = Table(displayName="HoursLog", ref="A1:G2")
    table.tableStyleInfo = TableStyleInfo(
        name="TableStyleMedium2", showRowStripes=True, showColumnStripes=False)
    ws.add_table(table)

    for c in range(1, len(HEADERS) + 1):
        ws.cell(row=1, column=c).font = Font(bold=True, color="FFFFFF")

    ws.cell(row=2, column=1).number_format = "mm/dd/yyyy"
    ws.cell(row=2, column=3).number_format = "h:mm AM/PM"
    ws.cell(row=2, column=4).number_format = "h:mm AM/PM"
    ws.cell(row=2, column=5).number_format = "0.00"

    # Force the Period column to Text. The app already avoids sending a value
    # Excel could mistake for a date, but a cell left on General format applies
    # that guesswork to anything landing in it later, by hand or otherwise.
    ws.cell(row=2, column=6).number_format = "@"
    ws.column_dimensions["F"].number_format = "@"

    ws.freeze_panes = "A2"


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--year", type=int, default=2026)
    ap.add_argument("--out", default="Work Hours.xlsx")
    args = ap.parse_args()

    wb = Workbook()
    build_summary(wb.active, args.year)
    wb.active.title = "Summary"
    build_log(wb.create_sheet("Hours Log"))
    wb.save(args.out)
    print("wrote %s (%d bytes)" % (args.out, os.path.getsize(args.out)))


if __name__ == "__main__":
    main()
