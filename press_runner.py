"""Get press releases and output them to xlsx, along with some metrics"""

# import os
import sys

import xlsxwriter

from fmp import fmp_check_symbols
from press_utilities import get_diff_between_releases, get_frames, write_press_comments
from workbook_utilities import close_workbook, set_global_font

TICKER = input("Ticker: ")

if TICKER not in fmp_check_symbols([TICKER]):
    print(f"[Error] Symbol not found: {TICKER}")
    print("[Error] Exiting")

print(f"\n[Processing] Symbol {TICKER} checked!")

frames = get_frames(TICKER)

if len(frames) == 1:
    print(f"\n[Error] No press releases found for {TICKER}")
    print("[ERROR] Quitting\n")
    sys.exit(1)

press_frame = frames[0]
comment_frame = frames[1]

average_days = get_diff_between_releases(press_frame)
num_rows = press_frame.shape[0]

WORKBOOK_NAME = "Workbooks/Press.xlsx"
WORKSHEET_NAME = "Press Releases"
WORKBOOK = xlsxwriter.Workbook(WORKBOOK_NAME)
set_global_font(WORKBOOK)
worksheet = WORKBOOK.add_worksheet(WORKSHEET_NAME)

cell_format = WORKBOOK.add_format({"font": "Tenorite", "border": 1})
worksheet.write("D2", f"Average days per release: {average_days}", cell_format)

average_format = WORKBOOK.add_format(
    {"font": "Tenorite", "num_format": "0.00%", "border": 1}
)
worksheet.write("I2", f"=AVERAGE(I5:I{num_rows + 5})", average_format)
worksheet.write("J2", f"=AVERAGE(J5:J{num_rows + 5})", average_format)
worksheet.write("K2", f"=AVERAGE(K5:K{num_rows + 5})", average_format)

press_frame.write_excel(
    workbook=WORKBOOK,
    worksheet=WORKSHEET_NAME,
    position="B4",
    table_style="TableStyleMedium17",
    header_format={"bold": True},
    column_formats={
        "Days Ahead: 1": {"num_format": "#,##0.00"},
        "Days Ahead: 3": {"num_format": "#,##0.00"},
        "Days Ahead: 5": {"num_format": "#,##0.00"},
        "%: PC -> 1": {"num_format": "0.00%", "border": 1},  # type: ignore
        "%: PC -> 3": {"num_format": "0.00%", "border": 1},  # type: ignore
        "%: PC -> 5": {"num_format": "0.00%", "border": 1},  # type: ignore
    },
    column_widths={
        "Symbol": 90,
        "Date": 160,
        "Title": 1800,
        "Previous Close": 120,
        "Days Ahead: 1": 120,
        "Days Ahead: 3": 120,
        "Days Ahead: 5": 120,
        "%: PC -> 1": 90,
        "%: PC -> 3": 90,
        "%: PC -> 5": 90,
    },
    conditional_formats={
        "%: PC -> 1": "3_color_scale",
        "%: PC -> 3": "3_color_scale",
        "%: PC -> 5": "3_color_scale",
    },
)

press_frame = press_frame.drop("Symbol")
press_frame = press_frame.drop("Title")
write_press_comments(worksheet, comment_frame, "D", 5)
close_workbook(WORKBOOK, WORKBOOK_NAME)

