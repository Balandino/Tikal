"""Scrape historical data from Yahoo Finance"""

import calendar
import time
from enum import Enum
from urllib.error import HTTPError

import polars


class YahooInterval(Enum):
    """Intervals for get_historical_prices"""

    DAY = "1d"
    WEEK = "1wk"
    MONTH = "1mo"


def get_historical_prices(ticker: str, interval: Enum):
    """
    Scrapes historical price data for a symbol

    Args:
        ticker: Symbol to be used for lookup.  For example, NVDA or ^GSPC
        interval: Time interval to be used.  Use YahooInterval Enum.

    Returns:
        Polars DataFrame (DataFrame blank if Symbol not found)


    Scrapes yahoo finance for historical data.  In the event of the symbol not being located an
    empty DataFrame will be returned.  If successful, returns the following descending structure:

               |------------------------------------------------------------------------|
    Col Title: | Date       | Open   | High   | Low    | Close  | Adj Close | Volume    |
    Type:      | str        | f64    | f64    | f64    | f64    | f64       | i64       |
    Row 1:     | 2023-08-28 | 464.82 | 499.27 | 448.88 | 485.09 | 485.09    | 311355600 |
    Row 2:     | 2023-08-21 | 444.94 | 502.66 | 442.22 | 460.18 | 460.18    | 431021100 |
    Row 3:     | 2023-08-14 | 404.86 | 452.68 | 403.11 | 432.99 | 432.99    | 292926600 |

    Note that any rows with null data will be excluded.
    """

    # Setting period1 to -2208988800 should set the start date to roughly
    # 01/01/1900, defaulting the search to the max if this is date is too old
    url = (
        "https://query1.finance.yahoo.com/v7/finance/download/"
        + ticker
        + "?period1=-2208988800"
        + "&period2="
        + str(calendar.timegm(time.gmtime()))  # Current time as timestamp
        + "&interval="
        + interval.value
        + "&events=history&includeAdjustedClose=true"
    )

    try:
        return (
            polars.read_csv(
                url,
                ignore_errors=True,
                schema={
                    "Date": polars.Utf8,
                    "Open": polars.Float64,
                    "High": polars.Float64,
                    "Low": polars.Float64,
                    "Close": polars.Float64,
                    "Adj Close": polars.Float64,
                    "Volume": polars.Int64,
                },
            )
            .sort("Date", descending=True)
            .drop_nulls()
            .with_columns(
                polars.col("Open", "High", "Low", "Close", "Adj Close").round(2)
            )
        )

    except HTTPError:
        return polars.DataFrame()
