# pylint: disable=line-too-long
"""Scrapes tradereconomics for commodities data"""

import polars

from commodities_utilities import get_commodities_frames, write_url
from workbook_utilities import close_workbook, create_workbook

WORKBOOK_NAME = "Workbooks/Commodities.xlsx"
WORKBOOK = create_workbook(WORKBOOK_NAME)
worksheet_name = "Commodities"
worksheet = WORKBOOK.add_worksheet(worksheet_name)


print("[Downloading] Gathering Data")
tables = get_commodities_frames()

column = "B"
row = 4

for table in tables:
    columns = ["Weekly", "Monthly", "YoY"]
    for col in columns:
        table[col] = table[col].str.rstrip("%").astype("float") / (100.0)

    frame = polars.from_pandas(table)
    frame = frame.drop("Price")
    frame = frame.drop("Day")
    frame = frame.drop("%")

    cell = column + str(row)

    frame.write_excel(
        workbook=WORKBOOK,
        worksheet=worksheet_name,
        position=cell,
        header_format={
            "bold": True,
            "font_color": "white",
            "bg_color": "black",
            "font": "Tenorite",
            "border": 2,
            "valign": "center",
        },
        column_widths={
            frame.columns[0]: 210,
            # "%": 100,
            "Weekly": 100,
            "Monthly": 100,
            "YoY": 100,
            "Date": 100,
        },
        column_formats={
            frame.columns[0]: {"font_color": "white", "bg_color": "black", "font": "Tenorite", "num_format": "0.00%", "border": 2},  # type: ignore
            "Weekly": {"font": "Tenorite", "num_format": "0.00%", "border": 2, "bold": True},  # type: ignore
            "Monthly": {"font": "Tenorite", "num_format": "0.00%", "border": 2, "bold": True},  # type: ignore
            "YoY": {"font": "Tenorite", "num_format": "0.00%", "border": 2, "bold": True},  # type: ignore
            "Date": {"font_color": "white", "bg_color": "black", "font": "Tenorite", "num_format": "0.00%", "border": 2, "bold": True},  # type: ignore
        },
        conditional_formats={
            "Weekly": "3_color_scale",
            "Monthly": "3_color_scale",
            "YoY": "3_color_scale",
        },
    )

    row = row + frame.shape[0] + 2

write_url(WORKBOOK, worksheet, "B2:F2")


print("[Writing] Writing Workbook")
close_workbook(WORKBOOK, WORKBOOK_NAME)
print("[Writing] Opening Workbook")

