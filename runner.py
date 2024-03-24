# pylint: disable=line-too-long
"""Test workpad for new code """

import concurrent.futures
import io

import pandas
import polars
import requests
from polars import DataFrame

from fmp import fmp_check_symbols, fmp_historical_prices


def multithreading_last_100(ticker: str):
    """
    Obtains last 100 closes for ticker.  Designed to be used called from get_last_100_historical

    Args:
        ticker: Ticker to retrieve information from

    Returns:
        dict{ticker: DataFrame}

    """
    data = fmp_historical_prices(ticker)
    historics = data["historical"]
    close_list = []

    for count in range(0, 101):
        close_list.append(historics[count]["close"])

    return {ticker: close_list}


def get_last_100_historical(tickers: list) -> dict[str, DataFrame]:
    """
    Obtain historical data simaultaneously

    Args:
        tickers: List of tickers to obtain data for

    Returns:
        Dictionary with tickers as keys and lists of historical closes as values

    """
    with concurrent.futures.ThreadPoolExecutor() as executor:
        futures = []
        historical = {}

        for ticker in tickers:
            futures.append(executor.submit(multithreading_last_100, ticker))

        for future in concurrent.futures.as_completed(futures):
            historical.update(future.result())

    return historical


def run():
    """Run"""
    url = "https://stockanalysis.com/etf/ura/holdings/"

    header_info = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/58.0.3029.110 Safari/537.3"
    }

    data = requests.get(url, headers=header_info, timeout=5)
    frames = pandas.read_html(io.StringIO(data.text))
    frame = polars.from_pandas(frames[0]).drop(["Shares"])

    num_rows = frame.shape[0]
    print(frame)

    tickers = []
    for count in range(0, num_rows):
        tickers.append(frame.item(count, 1))

    print(tickers)
    good_tickers = fmp_check_symbols(tickers)
    historical_data = get_last_100_historical(good_tickers)
    print(historical_data)


run()

