"""Unit testing the calculation functions wihtin press_utilities.py"""

import unittest

import polars

from press_utilities import (add_closing_prices, add_percentage_cols,
                             get_diff_between_releases, get_frames)


class TestPress(unittest.TestCase):
    """Unit tests"""

    TICKERS = ["NVDA"]  # self.TICKERS

    def test_calculations(self):
        """Checks the closing prices and calculations on the completed press frame"""
        ticker = "NVDA"
        press_frame = polars.read_csv("Test Data\\NVDA_Press.csv")
        press_frame = add_closing_prices(ticker, press_frame)
        press_frame = add_percentage_cols(press_frame)

        self.assertTrue(press_frame.item(70, 3) == 56.38)
        self.assertTrue(press_frame.item(70, 4) == 57.9)
        self.assertTrue(press_frame.item(70, 5) == 55.26)
        self.assertTrue(press_frame.item(70, 6) == 55.26)
        self.assertTrue(str(press_frame.item(70, 7))[0:6] == "0.0269")
        self.assertTrue(str(press_frame.item(70, 8))[0:7] == "-0.0198")
        self.assertTrue(str(press_frame.item(70, 9))[0:7] == "-0.0198")

        self.assertTrue(press_frame.item(67, 3) == 31.77)
        self.assertTrue(press_frame.item(67, 4) == 32.79)
        self.assertTrue(press_frame.item(67, 5) == 33.38)
        self.assertTrue(press_frame.item(67, 6) == 33.38)
        self.assertTrue(str(press_frame.item(67, 7))[0:6] == "0.0321")
        self.assertTrue(str(press_frame.item(67, 8))[0:6] == "0.0506")
        self.assertTrue(str(press_frame.item(67, 9))[0:6] == "0.0506")

        self.assertTrue(press_frame.item(0, 3) == 471.16)
        self.assertTrue(press_frame.item(0, 4) == 460.18)
        self.assertTrue(press_frame.item(0, 5) == 468.35)
        self.assertTrue(press_frame.item(0, 6) == 487.84)

        self.assertTrue(str(press_frame.item(0, 7))[0:7] == "-0.0233")
        self.assertTrue(str(press_frame.item(0, 8))[0:7] == "-0.0059")
        self.assertTrue(str(press_frame.item(0, 9))[0:6] == "0.0354")

    def test_get_frames(self):
        """Tests the full get_frames function"""
        for ticker in self.TICKERS:
            frames = get_frames(ticker)
            self.assertTrue(len(frames) == 2)

            press_frame = frames[0]
            self.assertTrue(press_frame.shape[1] == 10)

            comment_frame = frames[1]
            self.assertTrue(comment_frame.shape[1] == 1)

            frames = get_frames("MKS.L")
            self.assertTrue(len(frames) == 1)

    def test_average_days(self):
        """Tests the average days function"""
        press_frame = polars.read_csv("Test Data\\NVDA_Press.csv")
        self.assertTrue(get_diff_between_releases(press_frame) == 28)
