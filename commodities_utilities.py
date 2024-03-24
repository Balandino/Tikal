# pylint: disable=line-too-long
"""Utility functions for commodities_runner"""


import io

import pandas
import requests
from pandas import DataFrame
from xlsxwriter import Workbook
from xlsxwriter.workbook import Worksheet


def get_commodities_frames() -> list[DataFrame]:
    """
    Returns list of Pandas dataframes read from html page

    Returns:
        List of pandas frames for iteration

    """
    url = "https://tradingeconomics.com/commodities"

    header_info = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/58.0.3029.110 Safari/537.3"
    }

    data = requests.get(url, headers=header_info, timeout=5)
    return pandas.read_html(io.StringIO(data.text))


def write_url(WORKBOOK: Workbook, worksheet: Worksheet, cell_range: str):
    """
    Write url to merged range

    Args:
        WORKBOOK: Workbook to create format
        worksheet: Worksheet to write data to
        cell_range: Cell range to write to
    """
    merge_format = WORKBOOK.add_format(
        {
            "bold": 1,
            "border": 2,
            "align": "center",
            "valign": "vcenter",
            # "fg_color": "black",
            "color": "blue",
        }
    )

    worksheet.merge_range(
        cell_range, "https://tradingeconomics.com/commodities", merge_format
    )  # pyright: ignore[reportGeneralTypeIssues]

