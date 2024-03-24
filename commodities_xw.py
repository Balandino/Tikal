# pylint: disable=line-too-long
"""Test workpad for new code """


import io
from datetime import datetime

import pandas
import requests
import xlwings


def run(sheet_name: str):
    """Scrapes commodity data for Excel"""
    url = "https://tradingeconomics.com/commodities"

    header_info = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/58.0.3029.110 Safari/537.3"
    }

    data = requests.get(url, headers=header_info, timeout=5)
    frames = pandas.read_html(io.StringIO(data.text))

    column = "B"
    row = 4

    for frame in frames:
        columns = ["Weekly", "Monthly", "YoY"]
        for col in columns:
            frame[col] = frame[col].str.rstrip("%").astype("float") / (100.0)

        frame = frame.drop(columns=["%", "Price", "Day"])

        cell = column + str(row)
        row = row + frame.shape[0] + 2

        workbook = xlwings.Book.caller()
        sheet = workbook.sheets[sheet_name]
        sheet.range(cell).options(pandas.DataFrame, header=1, index=False).value = frame

        update_time = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
        sheet["I15"].value = update_time
        sheet.range("I15").select()

