"""Refreshes sector data in main"""

from datetime import datetime

import requests
import xlwings


def get_data(sector_dict: dict, key: str, percentage: bool) -> str:
    """
    Returns the data for the key, or - if key not present

    Args:
        sector_dict: Dict of sub sector data
        key: The relevant key

    Returns:
        Data or -
    """
    if key in sector_dict:
        if percentage:
            return str(float(sector_dict[key]) / 100)

        return sector_dict[key]

    return "-"


def run(sheet_name: str):
    """
    Writes scraped sector data into the passed in workbook

    Args:
        sheet_name: The name of the shee to write to
    """
    url = "https://stockanalysis.com/api/aggregation/industries_by_sector/?cols=industry_name,profitMargin,change,ch1m,chYTD,ch3y,stocks"  # pylint: disable=line-too-long

    full_data = requests.get(url, timeout=5).json()["data"]

    sectors = [
        "Communication Services",
        "Consumer Discretionary",
        "Consumer Staples",
        "Energy",
        "Financials",
        "Healthcare",
        "Industrials",
        "Materials",
        "Real Estate",
        "Technology",
        "Utilities",
    ]

    data_dict: dict[str, list] = {
        "Industry": [],
        "1M Change": [],
        "YTD": [],
        "3Y Change": [],
        "Stocks": [],
    }

    for sector in sectors:
        sector_data = full_data[sector]
        for sub_sector in sector_data:
            data_dict["Industry"].append(get_data(sub_sector, "industry_name", False))
            data_dict["1M Change"].append(get_data(sub_sector, "ch1m", True))
            data_dict["YTD"].append(get_data(sub_sector, "chYTD", True))
            data_dict["3Y Change"].append(get_data(sub_sector, "ch3y", True))
            data_dict["Stocks"].append(get_data(sub_sector, "stocks", False))

    workbook = xlwings.Book.caller()
    sheet = workbook.sheets[sheet_name]
    sheet.range("B4").options(transpose=True).value = data_dict["Industry"]
    sheet.range("C4").options(transpose=True).value = data_dict["1M Change"]
    sheet.range("D4").options(transpose=True).value = data_dict["YTD"]
    sheet.range("E4").options(transpose=True).value = data_dict["3Y Change"]
    sheet.range("F4").options(transpose=True).value = data_dict["Stocks"]

    sheet.range("K4").options(transpose=True).value = data_dict["Industry"]
    sheet.range("L4").options(transpose=True).value = data_dict["1M Change"]
    sheet.range("M4").options(transpose=True).value = data_dict["YTD"]
    sheet.range("N4").options(transpose=True).value = data_dict["3Y Change"]
    sheet.range("O4").options(transpose=True).value = data_dict["Stocks"]

    update_time = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
    sheet["Q13"].value = update_time
    sheet.range("Q13").select()

