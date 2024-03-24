"""Utility functions for ratio_runner.py"""


import polars
from xlsxwriter import Workbook
from xlsxwriter.worksheet import Worksheet

from fmp import (fmp_balance_sheet_annual, fmp_check_symbols, fmp_key_metrics,
                 fmp_ratios, get_data_no_title)


def get_tickers() -> list[str]:
    """
    Returns a list of checked tickers.

    Returns:
        List of upper tickers

    """
    tickers = [
        symbol.upper()
        for symbol in input(
            "Enter a single or space separated list of tickers: "
        ).split(" ")
    ]

    checked_list = fmp_check_symbols(tickers)
    output_list = []

    print("")
    for ticker in tickers:
        if ticker in checked_list:
            print(f"[Processing] {ticker} Checked")
            output_list.append(ticker)
        else:
            print(f"[Error] {ticker} Not found")

    print("")
    return output_list


def add_text(WORKBOOK: Workbook, worksheet: Worksheet, cell_range: str, text: str):
    """
    Adds text with black background and white text

    Args:
        WORKBOOK: Workbook object to contain the data
        worksheet: The current worksheet
        cell_Range: The cell range to be merged. e.g K68:L68
        text: The text to write
    """
    heading_format = WORKBOOK.add_format(
        {
            "bold": 1,
            "fg_color": "black",
            "font_color": "white",
            "text_wrap": True,
            "align": "left",
            "valign": "top",
            "border": 2,
            "border_color": "red",
        }
    )

    worksheet.merge_range(
        cell_range, text, heading_format
    )  # pyright: ignore[reportGeneralTypeIssues]


def get_ratios_frame(ticker: str, ratios: list) -> polars.DataFrame:
    """
    Obtain a polars DataFrame with the passed in ratios

    Data should be passed in in the following format:
    [
        [
            "dictionary key",
            "Title",
            "ratios or metrics, whichever one it's in",
            Whether to round to 2dp or not,
        ],
        ...
    ]


    Args:
        ticker: Symbol for fmp
        ratios: A list of lists


    Returns:
        DataFrame
    """
    ratios_json = fmp_ratios(ticker)
    metrics_json = fmp_key_metrics(ticker)

    ratio_dict = {}
    for ratio in ratios:
        if ratio[2] == "ratios":
            ratio_dict[ratio[1]] = get_data_no_title(ratios_json, ratio[0], ratio[3])
        else:
            ratio_dict[ratio[1]] = get_data_no_title(metrics_json, ratio[0], ratio[3])

    working_cap = ratio_dict["Working Capital (M)"]

    working_cap_m = []
    for num in working_cap:
        working_cap_m.append(str(num)[:-6])

    ratio_dict["Working Capital (M)"] = working_cap_m

    ratio_dict["Working Capital to Assets"] = get_working_cap_to_assets(
        metrics_json, fmp_balance_sheet_annual(ticker)
    )

    ratio_frame = polars.DataFrame(ratio_dict).reverse().transpose(include_header=True)
    ratio_frame = ratio_frame.rename(
        {
            "column": ticker,
            "column_0": "1",
            "column_1": "2",
            "column_2": "3",
            "column_3": "4",
            "column_4": "5",
        }
    )

    return ratio_frame


def get_working_cap_to_assets(ratios_json, balance_sheet_json):
    """
    Calculates and returns working capital to assets

    Args:
        ratios_json: The ratio json from fmp for workingCapital

    Returns:
        List of working capital to assets
    """
    working_cap_list = get_data_no_title(ratios_json, "workingCapital", False)
    total_assets = get_data_no_title(balance_sheet_json, "totalAssets", False)

    if len(total_assets) < len(working_cap_list):
        total_assets.extend([0] * (len(working_cap_list) - len(total_assets)))

    if len(working_cap_list) < len(total_assets):
        working_cap_list.extend([0] * (len(total_assets) - len(total_assets)))

    data = []
    for count, entry in enumerate(working_cap_list):
        if entry == 0 or total_assets[count] == 0:
            data.append(0)
        else:
            data.append(round(entry / total_assets[count], 2))

    return data
