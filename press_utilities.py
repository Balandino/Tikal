"""Utilities to aid press_runner.py"""


from datetime import date, datetime, timedelta

import polars
from xlsxwriter.worksheet import Worksheet

from fmp import fmp_historical_prices, fmp_press_releases


def get_diff_between_releases(press_frame: polars.DataFrame) -> int:
    """
    Returns the number representing the average days between press releases

    Args:
        press_frame: Frame with dates

    Returns:
        Number
    """

    num_rows = press_frame.shape[0]
    count = 0
    total = 0
    row_num = 1

    for _ in range(num_rows - 2):
        date1 = press_frame.item(row_num, 1).split(" ")[0]
        date1_dt: datetime = datetime.strptime(date1, "%Y-%m-%d")

        date2 = press_frame.item(row_num + 1, 1).split(" ")[0]
        date2_dt: datetime = datetime.strptime(date2, "%Y-%m-%d")

        total = total + (date1_dt - date2_dt).days
        count = count + 1
        row_num = row_num + 1

    return int(total / count)


def write_press_comments(
    worksheet: Worksheet,
    comment_frame: polars.DataFrame,
    col_letter: str,
    row_num: int,
):
    """
    Args:
        worksheet: The worksheet to be written to
        comment_frame: The single column frame containing the comments
        col_letter: Column letter to start at
        row_num: Row number to start at
    """

    for count in range(comment_frame.shape[0]):
        cell = col_letter + str(row_num)
        worksheet.write_comment(
            cell,
            comment_frame.item(count, 0),
            {
                "x_scale": 9,
                "y_scale": 3,
                "color": "white",
                "font": "Tenorite",
                "font_size": 11,
            },
        )

        row_num = row_num + 1


def get_frames(ticker: str) -> list[polars.DataFrame]:
    """
    Returns List with 2 elements.
    0 = Completed Press frame, with closing prices and percentage changes
    1 = Press comments frame

    Args:
        ticker: Symbol for FMP

    Returns:
        2 Element list with the frames
    """

    press_frame = polars.DataFrame(fmp_press_releases(ticker))

    if press_frame.shape[0] == 0:
        return [press_frame]

    comment_frame = polars.DataFrame(press_frame.get_column("text"))
    press_frame = press_frame.drop("text")

    # Generate Test Data
    # press_frame.write_csv("Test Data\\NVDA_Press.csv")
    with open("Test Data\\Price_Dict.txt", "w", encoding="utf-8") as dict_file:
        for key, value in get_price_dict(ticker).items():
            dict_file.write(f"{key}: {value}\n")

    press_frame = add_closing_prices(ticker, press_frame)
    press_frame = add_percentage_cols(press_frame)
    press_frame = press_frame.rename(
        {"symbol": "Symbol", "date": "Date", "title": "Title"}
    )

    return [press_frame, comment_frame]


def add_percentage_cols(press_frame: polars.DataFrame) -> polars.DataFrame:
    """
    Adds percentage change columns onto the end of the data frame

    Args:
        press_frame: Press frame with closing prices attached

    Returns:
        Press frame with percentage columns appended
    """

    days_ahead = [1, 3, 5]

    for num in days_ahead:
        press_frame = press_frame.with_columns(
            (
                (polars.col(f"Days Ahead: {num}") - polars.col("Previous Close"))
                / polars.col("Previous Close")
            ).alias(f"%: PC -> {num}")
        )

    return press_frame


def add_closing_prices(ticker: str, press_frame: polars.DataFrame) -> polars.DataFrame:
    """
    Returns a Polars DataFrame with press releases and closing prices

    Args:
        ticker: Symbol compatible with FMP
        press_frame: DataFrame of fmp_press_releases, without text column

    Returns:
        DataFrame of press releases and closing prices
    """

    # To look up prices efficiently
    price_dict = get_price_dict(ticker)

    closing_prices = get_prev_close_series(press_frame, price_dict)
    press_frame = press_frame.with_columns(
        polars.lit(polars.Series(closing_prices)).alias("Previous Close")
    )

    days_ahead_list = [1, 3, 5]
    for days_ahead in days_ahead_list:
        closing_prices = get_close_series(press_frame, price_dict, days_ahead)
        press_frame = press_frame.with_columns(
            polars.lit(polars.Series(closing_prices)).alias(f"Days Ahead: {days_ahead}")
        )

    return press_frame


def get_prev_close_series(press_frame: polars.DataFrame, price_dict: dict) -> list[str]:
    """
    Returns a series of the nearest previous closes based on the dates in press_frame

    Args:
        press_frame: Press frame containing dates
        price_dict: Dict of closing prices with dates as keys

    Returns:

    """

    close_series: list[str] = []
    num_rows = press_frame.shape[0]

    for count in range(num_rows):
        current_date = press_frame.item(count, 1).split()[0]
        close = get_previous_close(current_date, price_dict)
        close_series.append(close)

    return close_series


def get_close_series(
    press_frame: polars.DataFrame, price_dict: dict, days_ahead: int
) -> list[str]:
    """
    Returns a series of closes that are days_ahead of each date in the press_frame

    Args:
        press_frame: Press frame containing dates
        price_dict: Dict of closing prices with dates as keys
        days_ahead: Number of days ahead each close should be from the price

    Returns:
        List of closing prices

    """

    close_series: list[str] = []
    num_rows = press_frame.shape[0]

    for count in range(num_rows):
        current_date = press_frame.item(count, 1).split()[0]
        close = get_close(current_date, price_dict, days_ahead)
        close_series.append(close)

    return close_series


def get_close(start_date: str, price_dict: dict, days_ahead: int) -> str:
    """
    Returns the closing price for days_ahead after date argument.  If a close cannot be found
    for that date, then will check further.  If the date goes beyond today's date then the last
    price found will be returned

    Args:
     start_date: Starting date in format - 2023-12-20
     price_dict: Dict of closes with dates as keys
     days_ahead: Number of days ahead to search for close


    Returns:
        Price close string from price_dict
    """

    today: str = date.today().strftime("%Y-%m-%d")
    target_date: datetime = datetime.strptime(start_date, "%Y-%m-%d")
    target_date = target_date + timedelta(days=days_ahead)
    target_date_str = str(target_date).split(" ", maxsplit=1)[0]

    if target_date_str in price_dict:
        return price_dict[target_date_str]

    while True:
        target_date = target_date + timedelta(days=1)
        target_date_str = str(target_date).split(" ", maxsplit=1)[0]

        if target_date_str >= today:
            if start_date in price_dict:
                return price_dict[start_date]

            return get_previous_close(start_date, price_dict)

        if target_date_str in price_dict:
            return price_dict[target_date_str]


def get_previous_close(start_date: str, price_dict: dict) -> str:
    """
    Returns the nearest previous close.  Called as a last resort from get_close

    Args:
        start_date: Date to start search from
        price_dict: Dictionary of prices

    Returns:
        Price close string from price_dict
    """

    target_date: datetime = datetime.strptime(start_date, "%Y-%m-%d")
    target_date = target_date - timedelta(days=1)
    target_date_str = str(target_date).split(" ", maxsplit=1)[0]

    while True:
        if target_date_str in price_dict:
            return price_dict[target_date_str]

        target_date = target_date - timedelta(days=1)
        target_date_str = str(target_date).split(" ", maxsplit=1)[0]


def get_price_dict(ticker: str) -> dict[str, str]:
    """
    Returns a dictionary with dates as keys and closes as values for the last 30 years

    Args:
        ticker: Ticker symbol for FMP

    Returns:
        Dictionary
    """

    price_dict: dict[str, str] = {}
    prices = fmp_historical_prices(ticker)["historical"]

    for price_data in prices:
        price_dict[price_data["date"]] = price_data["close"]

    return price_dict

