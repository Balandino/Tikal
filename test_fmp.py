"""Unittests for fmp.py"""
import json
import unittest

import requests

from fmp import (
    check_for_error,
    fmp_balance_sheet_annual,
    fmp_check_symbols,
    fmp_company_profile,
    fmp_historical_prices,
    fmp_key_metrics,
    fmp_press_releases,
    fmp_ratios,
    fmp_sales_per_region,
    fmp_sales_per_segment,
    get_data,
    get_data_no_title,
)


class TestFMP(unittest.TestCase):

    """Unit tests for fmp.py"""

    TICKERS = ["NVDA"]  # self.TICKERS
    GOOD_TICKER = "NVDA"

    def test_check_symbols(self):
        """Returns list of tickers confirmed to be working"""
        self.assertTrue(len(fmp_check_symbols(self.TICKERS)) == 1)

    def test_check_ratios(self):
        """Checks ratios call"""
        for ticker in self.TICKERS:
            ratios = fmp_ratios(ticker)

            data = get_data_no_title(ratios, "currentRatio", True)
            self.assertTrue(len(data) > 3)

    def test_check_metrics(self):
        """Checks key metrics call"""
        for ticker in self.TICKERS:
            metrics = fmp_key_metrics(ticker)

            data = get_data_no_title(metrics, "roic", True)
            self.assertTrue(len(data) > 3)

    def test_check_balance_sheet(self):
        """Checks key balance sheet call"""
        for ticker in self.TICKERS:
            balance = fmp_balance_sheet_annual(ticker)

            data = get_data_no_title(balance, "totalAssets", True)
            self.assertTrue(len(data) > 3)

    def test_company_profile(self):
        """Check description call"""
        for ticker in self.TICKERS:
            self.assertTrue(fmp_company_profile(ticker)["symbol"] == ticker)

    def test_fmp_press_releases(self):
        """Some tickers have no entries, so be careful with this one"""
        for ticker in self.TICKERS:
            self.assertTrue(len(fmp_press_releases(ticker)) > 20)

    def test_fmp_historical_prices(self):
        """Should return two entries in a list: ticker and historical price list"""
        for ticker in self.TICKERS:
            self.assertTrue(len(fmp_historical_prices(ticker)) == 2)

    def test_fmp_sales_per_region(self):
        """Should return a list of dictionaries"""
        for ticker in self.TICKERS:
            self.assertTrue(len(fmp_sales_per_region(ticker)) > 3)

    def test_fmp_sales_per_segment(self):
        """Should return a list of dictionaries"""
        for ticker in self.TICKERS:
            self.assertTrue(len(fmp_sales_per_segment(ticker)) > 3)

    def test_get_data(self):
        """Data created in the test function itself"""

        json_str = """
        [
            { "num": 1.134 },
            { "num": 3.245 },
            { "num": 7.789 }
        ]
        """

        data = json.loads(json_str)

        parsed_data = get_data(data, "num", "Ticker", False)
        self.assertTrue(parsed_data[0] == "Ticker")
        self.assertTrue(parsed_data[1] == 1.134)
        self.assertTrue(parsed_data[2] == 3.245)
        self.assertTrue(parsed_data[3] == 7.789)

        parsed_data = get_data(data, "num", "Ticker", True)
        self.assertTrue(parsed_data[0] == "Ticker")
        self.assertTrue(parsed_data[1] == 1.13)
        self.assertTrue(parsed_data[2] == 3.25)
        self.assertTrue(parsed_data[3] == 7.79)

    def test_get_data_no_title(self):
        """Data created in the test function itself"""

        json_str = """
        [
            { "num": 1.134 },
            { "num": 3.245 },
            { "num": 7.789 }
        ]
        """

        data = json.loads(json_str)

        parsed_data = get_data_no_title(data, "num", False)
        self.assertTrue(parsed_data[0] == 1.134)
        self.assertTrue(parsed_data[1] == 3.245)
        self.assertTrue(parsed_data[2] == 7.789)

    def test_check_for_errors(self):
        """Pulls on fmp_historical_prices, so an error there may break this"""
        faulty_url = (
            "https://financialmodelingprep.com/api/v4/score?symbol=AAPL&apikey=ERROR"
        )
        self.assertTrue(check_for_error(requests.get(faulty_url, timeout=5).json()))
        self.assertFalse(check_for_error(fmp_historical_prices(self.GOOD_TICKER)))

