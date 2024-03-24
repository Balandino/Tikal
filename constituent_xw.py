"""Test workpad for new code """

import io
from datetime import datetime

import pandas
import requests
import xlwings


def run(sheet_name: str):
    """
    Refreshes constituent performance data

    Args:
        sheet_name: Name of sheet to write to
    """
    header = {
        "User-Agent": "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/50.0.2661.75 Safari/537.36",  # pylint: disable=line-too-long
        "X-Requested-With": "XMLHttpRequest",
    }

    # URLs to scrape
    urls = [
        "https://www.slickcharts.com/sp500/performance",
        "https://www.slickcharts.com/nasdaq100/performance",
        "https://www.slickcharts.com/dowjones/performance",
    ]

    # Number of top rows to get (How many constituents from the top to gather)
    top_count = [20, 20, 15]

    # Number of bottom rows to get (How many constituents from the bottom to gather)
    bottom_count = [21, 21, 15]

    # Position to write top data
    top_positions = [
        ["D6", "E6", "F6"],
        ["L6", "M6", "N6"],
        ["D66", "E66", "F66"],
    ]

    # Position to write bottom data
    bottom_positions = [
        ["D28", "E28", "F28"],
        ["L28", "M28", "N28"],
        ["D83", "E83", "F83"],
    ]

    for count, url in enumerate(urls):
        html_data = requests.get(url, timeout=5, headers=header)

        returns_table = pandas.read_html(io.StringIO(html_data.text))[0]
        top = returns_table.head(top_count[count])
        bottom = returns_table.tail(bottom_count[count])

        company = top["Company"].tolist()
        tickers = top["Symbol"].tolist()
        ytd = top["YTD Return"].tolist()

        workbook = xlwings.Book.caller()
        sheet = workbook.sheets[sheet_name]
        sheet.range(top_positions[count][0]).options(transpose=True).value = company
        sheet.range(top_positions[count][1]).options(transpose=True).value = tickers
        sheet.range(top_positions[count][2]).options(transpose=True).value = ytd

        company = bottom["Company"].tolist()
        tickers = bottom["Symbol"].tolist()
        ytd = bottom["YTD Return"].tolist()

        sheet.range(bottom_positions[count][0]).options(transpose=True).value = company
        sheet.range(bottom_positions[count][1]).options(transpose=True).value = tickers
        sheet.range(bottom_positions[count][2]).options(transpose=True).value = ytd

        update_time = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
        sheet["S15"].value = update_time
