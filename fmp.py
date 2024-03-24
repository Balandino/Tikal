# pylint: disable=line-too-long
"""Functions to download data from Financial Modelling Prep"""

import requests

# Globals
API_KEY = ""
LIMIT = 5  # Used to limit amount of returned data in some calls


def fmp_check_symbols(input_list: list[str]) -> list[str]:
    """
    Checks tickers are supported by FMP.

    Args:
        input_list: List of tickers to be checked.

    Returns: List of tickers confirmed to be supported by FMP
    """
    output_list = []

    symbol_list = get_data_no_title(fmp_symbol_list(), "symbol", False)
    for ticker in input_list:
        if ticker is None:
            continue

        ticker = ticker.upper()
        if ticker in symbol_list:
            output_list.append(ticker.upper())

    return output_list


def get_data(json_data, key: str, title: str, do_round: bool) -> list:
    """
    Designed to be used on the fmp functions that return lists of JSON objects, this will extract
    the targetted value in each object within that list and return a list of those values.  If an
    object is missing a value, then NA is appended in its plac

    e

    Args:
        key: The key to access the targetted JSON data
        title: A string that can be pre-pended to the returned list
        do_round: Whether the data should be rounded to 2dp

    Returns: A list of the gathered data
    """

    data = [title]

    for json_object in json_data:
        if do_round:
            data.append(round(json_object[key], 2))
        else:
            data.append(json_object[key])
    return data


def get_data_no_title(json_data, key: str, do_round: bool) -> list:
    """
    The same as get_data but will exclude the title

    Args:
        key: The key to access the targetted JSON data
        title: A string that can be pre-pended to the returned list, pass "" if not desired
        do_round: Whether the data should be rounded to 2dp

    Returns: A list of the gathered data

    """

    data = get_data(json_data, key, "NA", do_round)
    data.pop(0)
    return data


def check_for_error(json_data):
    """
    Checks for standard FMP error message on call failure

    Args:
        json_data: JSON to be checked
    """
    return "Error Message" in json_data


def fmp_symbol_list() -> list[dict]:
    """Returns json array of all fmp supported symbols

    Returns:
        List of dictionaries with symbol information

    [
      {
        "symbol": "KZMS.ME",
        "name": "The Open Joint Stock Company Krasnokamsk Metal Mesh Works",
        "price": 226,
        "exchange": "MCX",
        "exchangeShortName": "MCX",
        "type": "stock"
      },
      ...
    """
    url = f"https://financialmodelingprep.com/api/v3/stock/list?apikey={API_KEY}"

    return requests.get(url, timeout=5).json()


def fmp_ratios(ticker: str) -> list[dict]:
    """
    Returns ratios for the ticker

    Args:
        ticker: Symbol for fmp

    Returns:
        List of dictionaries with ratios
    """
    url = (
        "https://financialmodelingprep.com/api/v3/ratios/"
        + ticker
        + "?apikey="
        + API_KEY
        + "&limit="
        + str(LIMIT)
    )

    return requests.get(url, timeout=5).json()


def fmp_key_metrics(ticker: str) -> list[dict]:
    """
    Returns key metrics for the ticker

    Args:
        ticker: Symbol for fmp

    Returns:
        List of dictionaries with key metrics
    """
    url = (
        "https://financialmodelingprep.com/api/v3/key-metrics/"
        + ticker
        + "?apikey="
        + API_KEY
        + "&limit="
        + str(LIMIT)
    )

    return requests.get(url, timeout=5).json()


def fmp_company_profile(ticker: str) -> dict:
    """
    Returns the company profile for a ticker

    Args:
        ticker: Symbol for FMP

    Returns:
        List object with a json dictionary

    """
    url = (
        "https://financialmodelingprep.com/api/v3/profile/"
        + ticker
        + "?apikey="
        + API_KEY
    )

    return requests.get(url, timeout=5).json()[0]


def fmp_press_releases(ticker: str) -> list[dict]:
    """
    Returns json array of press release information

    Args:
        ticker: Symbol to search for on FMP

    Returns:
        List of json dictionaries

    [
      {
        "symbol": "NVDA",
        "date": "2023-05-25 17:00:00",
        "title": "NVIDIA ANNOUNCES UPCOMING EVENTS FOR FINANCIAL COMMUNITY",
        "text": "SANTA CLARA, CALIF., MAY 25, 2023 (GLOBE NEWSWIRE) -- NVIDIA WILL PRE...
      },
      ...
    """
    url = f"https://financialmodelingprep.com/api/v3/press-releases/{ticker}?apikey={API_KEY}"
    return requests.get(url, timeout=5).json()


def fmp_sales_per_segment(ticker: str) -> list[dict]:
    """
    Returns a list of the sales per product for a company

    Args:
        ticker: Ticker for FMP

    Returns:
        List of dictionaries

    [
      {
        "2023-01-29": {
          "Automotive": 903000000,
          "Data Center": 15005000000,
          "Gaming": 9067000000,
          "OEM and Other": 455000000,
          "Professional Visualization": 1544000000
        }
      },
      ...

    """
    url = f"https://financialmodelingprep.com/api/v4/revenue-product-segmentation?symbol={ticker}&structure=flat&period=annual&apikey={API_KEY}"
    return requests.get(url, timeout=5).json()


def fmp_sales_per_region(ticker: str) -> list[dict]:
    """

    Returns a list of the sales per region for a company

    Args:
        ticker: Ticker for FMP

    Returns:
        List of dictionaries

       [
      {
        "2022-09-24": {
          "CHINA": 74200000000,
          "Other Countries": 172269000000,
          "UNITED STATES": 147859000000
        }
      },
      ...
    """
    url = f"https://financialmodelingprep.com/api/v4/revenue-geographic-segmentation?symbol={ticker}&structure=flat&apikey={API_KEY}"
    return requests.get(url, timeout=5).json()


def fmp_historical_prices(ticker: str) -> dict:
    """
    Returns historical price data.  Note: price information stored under 'historical' key

    Args:
        ticker: Symbol to search for on FMP

    Returns:
        List of json dictionaries

    {
      "symbol": "NVDA",
      "historical": [
        {
          "date": "2023-05-26",
          "close": 389.0214
        },
        {
          "date": "2023-05-25",
          "close": 379.8
        },
        ...
    """

    url = f"https://financialmodelingprep.com/api/v3/historical-price-full/{ticker}?serietype=line&apikey={API_KEY}"
    return requests.get(url, timeout=5).json()


def fmp_balance_sheet_annual(ticker) -> list[dict]:
    """
    Returns the annual balance sheet for the symbol

    Args:
        ticker: Symbol for FMP

    Returns:
        List of json dictionaries

    """
    url = (
        "https://financialmodelingprep.com/api/v3/balance-sheet-statement/"
        + ticker
        + "?limit="
        + str(LIMIT)
        + "&apikey="
        + API_KEY
    )

    return requests.get(url, timeout=5).json()


