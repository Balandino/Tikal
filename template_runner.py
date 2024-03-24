"""Build the Trade Idea Template"""


from module_loader import add_module
from workbook_utilities import close_workbook, create_workbook

ticker = "NVDA"

WORKBOOK_NAME = f"Workbooks/{ticker}_template.xlsx"
WORKSHEET_NAME = "Press Releases"
WORKBOOK = create_workbook(WORKBOOK_NAME)
worksheet = WORKBOOK.add_worksheet(ticker)


worksheet.set_column(0, 0, 15)


MAIN_COLOUR = "#1B201E"
DATA_COLOUR_1 = "#E7E6E6"
DATA_COLOUR_2 = "#D0CECE"
GUIDE_COLOUR = "#FFC000"

MODULES = [
    {"name": "profile", "workbook": WORKBOOK, "worksheet": worksheet, "ticker": ticker}
]

FORMATS = {
    "sub_heading_left": WORKBOOK.add_format(
        {
            "bold": 1,
            "fg_color": MAIN_COLOUR,
            "font_color": "white",
            "text_wrap": True,
            "align": "left",
            "valign": "vcenter",
            "border": 2,
            "border_color": "black",
            "font_size": 10,
        }
    ),
    "sub_heading_center": WORKBOOK.add_format(
        {
            "bold": 1,
            "fg_color": MAIN_COLOUR,
            "font_color": "white",
            "text_wrap": True,
            "align": "center",
            "valign": "vcenter",
            "border": 2,
            "border_color": "black",
            "font_size": 10,
        }
    ),
    "main_heading_left": WORKBOOK.add_format(
        {
            "bold": 1,
            "fg_color": MAIN_COLOUR,
            "font_color": "white",
            "text_wrap": True,
            "align": "left",
            "valign": "vcenter",
            "border": 2,
            "border_color": "black",
            "font_size": 12,
        }
    ),
    "main_heading_center": WORKBOOK.add_format(
        {
            "bold": 1,
            "fg_color": MAIN_COLOUR,
            "font_color": "white",
            "text_wrap": True,
            "align": "center",
            "valign": "vcenter",
            "border": 2,
            "border_color": "black",
            "font_size": 12,
        }
    ),
    "data_type_1_left": WORKBOOK.add_format(
        {
            "bold": 1,
            "fg_color": DATA_COLOUR_1,
            "font_color": "black",
            "text_wrap": True,
            "align": "left",
            "valign": "vcenter",
            "border": 2,
            "border_color": "black",
            "font_size": 10,
        }
    ),
    "data_type_1_center": WORKBOOK.add_format(
        {
            "bold": 1,
            "fg_color": DATA_COLOUR_1,
            "font_color": "black",
            "text_wrap": True,
            "align": "center",
            "valign": "vcenter",
            "border": 2,
            "border_color": "black",
            "font_size": 10,
        }
    ),
    "data_type_2_left": WORKBOOK.add_format(
        {
            "bold": 1,
            "fg_color": DATA_COLOUR_2,
            "font_color": "black",
            "text_wrap": True,
            "align": "left",
            "valign": "vcenter",
            "border": 2,
            "border_color": "black",
            "font_size": 10,
        }
    ),
    "data_type_2_center": WORKBOOK.add_format(
        {
            "bold": 1,
            "fg_color": DATA_COLOUR_2,
            "font_color": "black",
            "text_wrap": True,
            "align": "center",
            "valign": "vcenter",
            "border": 2,
            "border_color": "black",
            "font_size": 10,
        }
    ),
}


START_CELL = "A1"
for module in MODULES:
    START_CELL = add_module(START_CELL, module, FORMATS)


close_workbook(WORKBOOK, WORKBOOK_NAME)

