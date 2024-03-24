"""Creates a spreadhsheet with various ratios and a company description for each ticker"""
# pylint: disable=C0103

import sys

from fmp import fmp_company_profile
from ratios_utilities import add_text, get_ratios_frame, get_tickers
from workbook_utilities import close_workbook, create_workbook

RATIOS = [
    [
        "date",
        "Date",
        "ratios",
        False,
    ],
    [
        "currentRatio",
        "Current Ratio",
        "ratios",
        True,
    ],
    [
        "quickRatio",
        "Quick Ratio",
        "ratios",
        True,
    ],
    [
        "returnOnAssets",
        "ROA",
        "ratios",
        True,
    ],
    [
        "returnOnEquity",
        "ROE",
        "ratios",
        True,
    ],
    [
        "roic",
        "ROIC",
        "metrics",
        True,
    ],
    [
        "interestCoverage",
        "Interest Coverage",
        "metrics",
        True,
    ],
    [
        "priceToSalesRatio",
        "Price to Sales",
        "ratios",
        True,
    ],
    [
        "bookValuePerShare",
        "BVPS",
        "metrics",
        True,
    ],
    [
        "debtToEquity",
        "Debt to Equity",
        "metrics",
        True,
    ],
    [
        "debtToAssets",
        "Debt to Assets",
        "metrics",
        True,
    ],
    [
        "freeCashFlowYield",
        "Free Cashflow Yield",
        "metrics",
        True,
    ],
    [
        "assetTurnover",
        "Free Cashflow Yield",
        "ratios",
        True,
    ],
    [
        "workingCapital",
        "Working Capital (M)",
        "metrics",
        True,
    ],
]


WORKBOOK_NAME = "Workbooks/Ratios.xlsx"
WORKSHEET_NAME = "Ratios"
WORKBOOK = create_workbook(WORKBOOK_NAME)
worksheet = WORKBOOK.add_worksheet(WORKSHEET_NAME)
TICKERS = get_tickers()

if len(TICKERS) < 1:
    print("[ERROR] No matching tickers found, quitting")
    sys.exit(0)

ratios_column = 1
ratios_row = 1

for ticker in TICKERS:
    ratio_frame = get_ratios_frame(ticker, RATIOS)
    profile = fmp_company_profile(ticker)

    ratio_frame.write_excel(
        workbook=WORKBOOK,
        worksheet=WORKSHEET_NAME,
        position=[ratios_row, ratios_column],  # type: ignore
        table_style="TableStyleDark3",
        column_widths={
            ticker: 160,
            "1": 80,
            "2": 80,
            "3": 80,
            "4": 80,
            "5": 80,
        },
        column_formats={
            ticker: {"bold": True},  # type: ignore
            "1": {"bold": True},  # type: ignore
            "2": {"bold": True},  # type: ignore
            "3": {"bold": True},  # type: ignore
            "4": {"bold": True},  # type: ignore
            "5": {"bold": True},  # type: ignore
        },
        header_format={"bold": True},
    )

    bottom_row = str(ratios_row + ratio_frame.shape[0] - 1)
    cell_range = f"I{str(ratios_row + 1)}:AD{bottom_row}"
    add_text(WORKBOOK, worksheet, cell_range, profile["description"])

    bottom_row = str(ratios_row + ratio_frame.shape[0] + 1)
    cell_range = f"I{bottom_row}:P{bottom_row}"
    add_text(WORKBOOK, worksheet, cell_range, profile["website"])

    cell_range = f"R{bottom_row}:U{bottom_row}"
    add_text(WORKBOOK, worksheet, cell_range, profile["exchangeShortName"])

    cell_range = f"W{bottom_row}:Y{bottom_row}"
    add_text(WORKBOOK, worksheet, cell_range, profile["industry"])

    cell_range = f"AA{bottom_row}:AD{bottom_row}"
    add_text(
        WORKBOOK,
        worksheet,
        cell_range,
        f"No. Full Time Employees: {int(profile['fullTimeEmployees']):,}",
    )

    ratios_row = ratios_row + ratio_frame.shape[0] + 2
    print(f"[Processing] {ticker} Completed!")


print("")
close_workbook(WORKBOOK, WORKBOOK_NAME)
