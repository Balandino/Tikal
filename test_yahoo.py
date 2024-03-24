"""Unittests for yahoo.py"""

import unittest
from datetime import datetime

import polars

from yahoo import YahooInterval, get_historical_prices


class TestYahoo(unittest.TestCase):

    """Unit tests for yahoo.py"""

    TICKERS = ["AMD", "H", "^STOXX50E"]  # self.TICKERS
    FAULTY_TICKER = "NNVVDDAA"
    GOOD_TICKER = "NVDA"

    def test_get_historical_prices(self):
        """Unit test"""
        # shape[1] == columns, shape[0] == Rows

        # Test shape of frame and no nulls
        for ticker in self.TICKERS:
            day_frame = get_historical_prices(ticker, YahooInterval.DAY)
            self.assertTrue(day_frame.shape[1] == 7)
            self.assertTrue(day_frame.shape[0] > 3)

            null_frame = day_frame.select(polars.all().is_null().sum())
            self.assertTrue(null_frame.shape[1] == 7)
            self.assertTrue(null_frame.shape[0] == 1)

            week_frame = get_historical_prices(ticker, YahooInterval.WEEK)
            self.assertTrue(week_frame.shape[1] == 7)
            self.assertTrue(week_frame.shape[0] > 3)

            null_frame = week_frame.select(polars.all().is_null().sum())
            self.assertTrue(null_frame.shape[1] == 7)
            self.assertTrue(null_frame.shape[0] == 1)

            month_frame = get_historical_prices(ticker, YahooInterval.MONTH)
            self.assertTrue(month_frame.shape[1] == 7)
            self.assertTrue(month_frame.shape[0] > 3)

            null_frame = month_frame.select(polars.all().is_null().sum())
            self.assertTrue(null_frame.shape[1] == 7)
            self.assertTrue(null_frame.shape[0] == 1)

        # Shape should be [0,0] if symbol not found
        month_frame = get_historical_prices(self.FAULTY_TICKER, YahooInterval.MONTH)
        self.assertTrue(month_frame.shape[1] == 0)
        self.assertTrue(month_frame.shape[0] == 0)

        # Check ordered descending
        day_frame = get_historical_prices(self.GOOD_TICKER, YahooInterval.DAY)
        newest_date = day_frame.item(0, 0)
        older_date = day_frame.item(5, 0)

        newest_date_dt = datetime.strptime(newest_date, "%Y-%m-%d")
        older_date_dt = datetime.strptime(older_date, "%Y-%m-%d")
        self.assertTrue(newest_date_dt > older_date_dt)

        # Check rounded to 2dp
        self.assertTrue(len(str(day_frame.item(0, 1)).split(".")[1]) < 3)
        self.assertTrue(len(str(day_frame.item(0, 2)).split(".")[1]) < 3)
        self.assertTrue(len(str(day_frame.item(0, 3)).split(".")[1]) < 3)
        self.assertTrue(len(str(day_frame.item(0, 4)).split(".")[1]) < 3)
        self.assertTrue(len(str(day_frame.item(0, 5)).split(".")[1]) < 3)
