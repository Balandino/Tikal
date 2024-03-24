"""Utility functions for runner_indexes.py"""

import concurrent.futures
import time

from polars import DataFrame
from xlsxwriter import Workbook
from xlsxwriter.worksheet import Worksheet

from yahoo import YahooInterval, get_historical_prices


def multithreading_download(ticker: str) -> dict[str, DataFrame]:
    """
    Obtains historical data for the ticker.  Designed to be used with get_all_historical_data

    Args:
        ticker: Ticker symbol for historical data

    Returns: dict{ticker: DataFrame}

    """
    return {ticker: get_historical_prices(ticker, YahooInterval.DAY)}


def get_all_historical_data(tickers: list) -> dict[str, DataFrame]:
    """
    Dowload all historical data and return as a dictionary

    Args:
        tickers: List of tickers to lookup on Yahoo Finance

    Returns: dict{ticker: DataFrame, ticker: DataFrame...}

    """
    with concurrent.futures.ThreadPoolExecutor() as executor:
        futures = []
        historical = {}

        for ticker in tickers:
            futures.append(executor.submit(multithreading_download, ticker=ticker))

        for future in concurrent.futures.as_completed(futures):
            historical.update(future.result())

    return historical


def add_indexes(
    workbook: Workbook,
    indexes: list[list[str]],
    historical: dict[str, DataFrame],
    colors: dict[str, str],
):
    """
    Adds the index data to the spreadsheet

    Args:
        workbook: Workbook to contain the data
        indexes: The index data
        historical: Dictionary of tickers and ta DataFrame of their historical data
        colors: Colors with keys matching index 4 index data
    """
    tickers: list[str] = []
    longest_name = 0

    names_col = "B"
    current_row = 5
    url_positions = {}

    for count, index_data in enumerate(indexes):
        tickers.append(index_data[0])
        current_type = index_data[4]
        url_positions[index_data[0]] = f"{names_col}{current_row}"

        # Matches the cell positions on the main page that the index names are written
        # to.  This allows the data pagesd to hyperlink back to the correct cell on the
        # heading page
        if (count + 1) < len(indexes) and indexes[count + 1][4] != current_type:
            current_row = current_row + 2
        else:
            current_row = current_row + 1

        # Get the longest name to manage print output neatly
        if len(index_data[1]) > longest_name:
            longest_name = len(index_data[1]) + 1

    # Add index pages
    for index_data in indexes:
        INDEX_START_TIME = time.perf_counter()
        if index_data[4] == "Index":
            index_data[4] = colors[index_data[4]]
            add_index(
                workbook,
                index_data,
                historical[index_data[0]],
                url_positions[index_data[0]],
            )
        else:
            index_data[4] = colors[index_data[4]]
            add_vol_index(
                workbook,
                index_data,
                historical[index_data[0]],
                url_positions[index_data[0]],
            )

        print(
            f"[Processed] {index_data[1]:{longest_name}} ({time.perf_counter()-INDEX_START_TIME:0.2f} seconds)"
        )


def add_title_page(
    WORKBOOK: Workbook,
    indexes: list[list[str]],
    colors: dict[str, str],
    historical: dict[str, DataFrame],
) -> Worksheet:
    """
    Adds main title page to Workbook

    Args:
        WORKBOOK: Workbook object to add page to
        indexes: List of lists contianing index data
        colors: Dictionary of colours to use

    Returns:
        The worksheet
    """
    worksheet = WORKBOOK.add_worksheet("Indexes")
    worksheet.set_column(1, 1, 15.0)
    worksheet.set_column(2, 2, 27.0)
    add_heading(WORKBOOK, worksheet, "B2:D3", "Data")
    add_heading(WORKBOOK, worksheet, "F2:J2", "Coloured by section & column")
    add_heading(WORKBOOK, worksheet, "F3", "Last 5")
    add_heading(WORKBOOK, worksheet, "G3", "Last 20")
    add_heading(WORKBOOK, worksheet, "H3", "Last 40")
    add_heading(WORKBOOK, worksheet, "I3", "Last 60")
    add_heading(WORKBOOK, worksheet, "J3", "Last 90")

    add_index_data(indexes, WORKBOOK, worksheet, historical, colors)

    worksheet.set_tab_color("black")
    return worksheet


def add_index_data(
    indexes: list[list[str]],
    WORKBOOK: Workbook,
    worksheet: Worksheet,
    historical: dict[str, DataFrame],
    colors: dict[str, str],
):
    """
    Writes the index name, colours and percentage move cells for the heading page

    Args:
        indexes: List of index data
        WORKBOOK: Workbook for adding formats to
        worksheet: Worksheet to write to
        historical: The Historical data to use for calculations
        colors: Dictionary of colours to use
    """
    row = 5
    start_row = row
    percentage_cols = ["F", "G", "H", "I", "J"]
    num_closes = [4, 19, 39, 59, 89]  # Offset by 1
    name_cols = ["B", "C", "D"]
    for count, index_data in enumerate(indexes):
        cell_format = WORKBOOK.add_format(
            {
                "font": "Tenorite",
                "bg_color": colors[index_data[4]],
                "border": 2,
                "align": "center",
            }
        )

        worksheet.write((name_cols[0] + str(row)), index_data[0], cell_format)
        worksheet.write((name_cols[1] + str(row)), index_data[1])

        cell_format = WORKBOOK.add_format(
            {
                "font": "Tenorite",
                "bg_color": "black",
                "bold": 1,
                "font_color": "white",
                "border": 2,
                "align": "center",
            }
        )

        last_close = historical[index_data[0]].item(0, 1)
        worksheet.write((name_cols[2] + str(row)), last_close, cell_format)

        url_cell = "K75" if index_data[4] == "Index" else "G75"
        worksheet.write_url(
            (name_cols[1] + str(row)),
            string=index_data[1],
            url=f"internal:'{index_data[1]}'!{url_cell}",
        )  # pyright: ignore[reportGeneralTypeIssues])

        cell_format = WORKBOOK.add_format({"num_format": "0.00%", "border": 1})

        for num, col in enumerate(percentage_cols):
            percent_move = get_percent_move(historical[index_data[0]], num_closes[num])
            worksheet.write((col + str(row)), percent_move, cell_format)

        current_type = index_data[4]

        # If moved to new type of data (e.g Index to vol), add a new line
        if (count + 1) < len(indexes) and indexes[count + 1][4] != current_type:
            for letter in percentage_cols:
                worksheet.conditional_format(
                    f"{letter}{start_row}:{letter}{row}", {"type": "3_color_scale"}
                )  # pyright: ignore[reportGeneralTypeIssues])

            row = row + 2
            start_row = row
        else:
            row = row + 1

        if count == (len(indexes) - 1):
            for letter in percentage_cols:
                worksheet.conditional_format(
                    f"{letter}{start_row}:{letter}{row}", {"type": "3_color_scale"}
                )  # pyright: ignore[reportGeneralTypeIssues])


def get_percent_move(historical_data: DataFrame, num_closes: int) -> float:
    """
    Gets the % difference between closing prices.  Returns 0.0 if num_closes greater
    than the number of available closes.  num_closes should be 1 less than target as
    first entry is inclusive.  For example, put 4 for change over last 5 closes

    Args:
        historical_data: DataFrame of historical data from Yahoo Finance
        num_closes: The number of close moves to capture

    Returns:
        Float representing percentage difference

    """
    close_list = historical_data.head(100).select("Close").to_series().to_list()

    if num_closes > len(close_list):
        return 0.0

    return (close_list[0] - close_list[num_closes]) / close_list[num_closes]


def add_url(
    WORKBOOK: Workbook, worksheet: Worksheet, cell_range: str, text: str, url: str
):
    """
    Adds a url to a specified cell range

    Args:
        WORKBOOK: Workbook for adding format to
        worksheet: Worksheet to write to
        cell_range: Cell range, single or multiple cells
        text: Text to display
        url: Url to use
    """

    merge_format = WORKBOOK.add_format(
        {
            "bold": 1,
            "border": 2,
            "align": "center",
            "valign": "vcenter",
            # "fg_color": "black",
            "color": "blue",
        }
    )

    if ":" in cell_range:
        worksheet.merge_range(
            cell_range, text, merge_format
        )  # pyright: ignore[reportGeneralTypeIssues]

    worksheet.write_url(
        cell_range,
        string=text,
        url=url,
        cell_format=merge_format,
    )  # pyright: ignore[reportGeneralTypeIssues])


def add_vol_index(
    WORKBOOK: Workbook, index_data: list, historical: DataFrame, url_cell: str
):
    """
    Adds a volatility index to the workbook along with 2 x charts

    Args:
        WORKBOOK: The workbook to have the new sheets added to
        index_data: List of lists
        historical: The historical data to process
        url_cell: The cell reference on the main page to hyperlink to


    Each list within index data should be structured like so:
    [
        "yahoo_ticker",
        "Name",
        "Country",
        "Description"
        "Tab Colour"
    ]

    For example:
    [
        "^GVIX",
        "VIX",
        "U.S",
        "The VIX is...
        "red"
    ],

    """

    worksheet_name = index_data[1]
    worksheet = WORKBOOK.add_worksheet(worksheet_name)
    worksheet.set_tab_color(index_data[4])
    add_url(
        WORKBOOK,
        worksheet,
        "G75:H75",
        "Back to main page",
        f"internal:'Indexes'!{url_cell}",
    )

    add_historical_vol_data(WORKBOOK, worksheet_name, historical)

    chart_data = {
        "workbook": WORKBOOK,
        "worksheet": worksheet,
        "worksheet_name": worksheet_name,
        "chart_name": f"{worksheet_name} (Last 60)",
        "series_name": "VIX",
        "series_color": "black",
        "x_axis_range": "E2:E61",
        "y_axis_range": "A2:A61",
        "x_axis_name": "Date",
        "x_axis_major_unit": 10,
        "y_axis_name": "Close",
        "position_cell": "G2",
        "width_px": 1728,
        "height_px": 640,
    }

    add_index_line_chart(chart_data)

    chart_data = {
        "workbook": WORKBOOK,
        "worksheet": worksheet,
        "worksheet_name": worksheet_name,
        "chart_name": worksheet_name,
        "series_name": "VIX",
        "series_color": "black",
        "x_axis_range": "E2:E8000",
        "y_axis_range": "A2:A8000",
        "x_axis_name": "Date",
        "x_axis_major_unit": 400,
        "y_axis_name": "Close",
        "position_cell": "G35",
        "width_px": 1728,
        "height_px": 640,
    }

    add_index_line_chart(chart_data)
    add_heading(WORKBOOK, worksheet, "G68:K68", index_data[2])
    add_heading(WORKBOOK, worksheet, "G70:AG70", index_data[3])


def add_index(
    WORKBOOK: Workbook, index_data: list, historical: DataFrame, url_cell: str
):
    """
    Adds an index to the workbook along with 2 x Bull/Bear charts

    Args:
        WORKBOOK: The workbook to have the new sheets added to
        index_data: List of lists
        historical: The historical data to process
        url_cell: The cell reference on the main page to hyperlink to

    Each list within index data should be structured like so:
    [
        "yahoo_ticker",
        "Name",
        "Country",
        "Description"
        "Tab Colour"
    ]

    For example:
    [
        "^GSPC",
        "S&P 500",
        "U.S",
        "The S&P 500 Index, or Standard & Poor....
        "green"
    ],

    """

    worksheet_name = index_data[1]
    worksheet = WORKBOOK.add_worksheet(worksheet_name)
    worksheet.set_tab_color(index_data[4])
    add_url(
        WORKBOOK,
        worksheet,
        "K75:L75",
        "Back to main page",
        f"internal:'Indexes'!{url_cell}",
    )

    add_historical_index_data(WORKBOOK, worksheet_name, historical)

    chart_data = {
        "workbook": WORKBOOK,
        "worksheet": worksheet,
        "worksheet_name": worksheet_name,
        "chart_name": f"{worksheet_name} (Last 60)",
        "series_name": "Bull",
        "series_name2": "Bear",
        "series_color": "green",
        "series_color2": "red",
        "x_axis_range": "E2:E61",
        "x_axis_range2": "I2:I61",
        "y_axis_range": "A2:A61",
        "x_axis_name": "Date",
        "x_axis_major_unit": 10,
        "y_axis_name": "Close",
        "position_cell": "K2",
        "width_px": 1728,
        "height_px": 640,
        "bull_bear": True,
    }

    add_index_line_chart(chart_data)

    chart_data = {
        "workbook": WORKBOOK,
        "worksheet": worksheet,
        "worksheet_name": worksheet_name,
        "chart_name": worksheet_name,
        "series_name": "Bull",
        "series_name2": "Bear",
        "series_color": "green",
        "series_color2": "red",
        "x_axis_range": "E2:E8000",
        "x_axis_range2": "I2:I8000",
        "y_axis_range": "A2:A8000",
        "x_axis_name": "Date",
        "x_axis_major_unit": 400,
        "y_axis_name": "Close",
        "position_cell": "K35",
        "width_px": 1728,
        "height_px": 640,
        "bull_bear": True,
    }

    add_index_line_chart(chart_data)

    add_heading(WORKBOOK, worksheet, "K68:O68", index_data[2])
    add_heading(WORKBOOK, worksheet, "K70:AK70", index_data[3])


def add_heading(WORKBOOK: Workbook, worksheet: Worksheet, cell_range: str, text: str):
    """
    Adds text with black background and white text

    Args:
        WORKBOOK: Workbook object to contain the data
        worksheet: The current worksheet
        cell_Range: The cell range to be merged. e.g K68:L68
        text: The text to write
    """
    heading_format = WORKBOOK.add_format(
        {
            "bold": 1,
            "align": "center",
            "valign": "vcenter",
            "fg_color": "black",
            "font_color": "white",
            "font": "Tenorite",
        }
    )

    if ":" in cell_range:
        worksheet.merge_range(
            cell_range, text, heading_format
        )  # pyright: ignore[reportGeneralTypeIssues]
    else:
        worksheet.write(
            cell_range, text, heading_format
        )  # pyright: ignore[reportGeneralTypeIssues]


def add_historical_vol_data(
    WORKBOOK: Workbook, worksheet_name: str, historical: DataFrame
):
    """
    Downloads and adds historical data to the current sheet

    Args:
        WORKBOOK: Workbook object to contain the data
        worksheet_name: The name of the current sheet
        historical: Polars DataFrame of the historical data
    """

    historical_prices_frame = historical.drop(["Adj Close", "Volume"])

    historical_prices_frame.write_excel(
        workbook=WORKBOOK,
        worksheet=worksheet_name,
        position="A1",
        header_format={"font": "Tenorite", "bold": True},
        column_formats={
            "Date": {"font": "Tenorite", "num_format": "dd/mm/yyyy"},
            "Open": {"font": "Tenorite"},
            "High": {"font": "Tenorite"},
            "Low": {"font": "Tenorite"},
            "Close": {"font": "Tenorite"},
        },
        column_widths={
            "Date": 90,
            "Open": 90,
            "High": 90,
            "Low": 90,
            "Close": 90,
        },
    )


def add_historical_index_data(
    WORKBOOK: Workbook, worksheet_name: str, historical: DataFrame
):
    """
    Downloads and adds historical data to the current sheet along with Bull & Bear index columns

    Args:
        WORKBOOK: Workbook object to contain the data
        worksheet_name: The name of the current sheet
        historical: Polars DataFrame of the historical data
    """

    historical_prices_frame = historical.drop(["Adj Close", "Volume"])

    historical_prices_frame.write_excel(
        workbook=WORKBOOK,
        worksheet=worksheet_name,
        position="A1",
        header_format={"font": "Tenorite", "bold": True},
        formulas={
            "Rolling Index High": f"=MAX([@High]:$C${historical_prices_frame.shape[0]})",
            "Rolling Bear Market Level": "=[@Rolling Index High]*0.8",
            "Bull Market Index": "=IF([@Close]>[@Rolling Bear Market Level],[@Close],#N/A)",
            "Bear Market Index": "=IF([@Close]<=[@Rolling Bear Market Level],[@Close],#N/A)",
        },
        column_formats={
            "Date": {"font": "Tenorite", "num_format": "dd/mm/yyyy"},
            "Open": {"font": "Tenorite"},
            "High": {"font": "Tenorite"},
            "Low": {"font": "Tenorite"},
            "Close": {"font": "Tenorite"},
            "Bull Market Index": {"align": "right"},
            "Bear Market Index": {"align": "right"},
        },
        column_widths={
            "Date": 90,
            "Open": 90,
            "High": 90,
            "Low": 90,
            "Close": 90,
            "Rolling Index High": 150,
            "Rolling Bear Market Level": 150,
            "Bull Market Index": 150,
            "Bear Market Index": 150,
        },
    )


def add_index_line_chart(chart: dict):
    """
    Builds and inserts a line chart with bull and bear colourings based on the dictionary entries:

    Args:
        chart: Dictionary containing chart options

    chart = {
        "workbook": WORKBOOK,  -- The xlsxwriter Workbook object
        "worksheet": worksheet, -- The xlwsxwriter worksheet within the workbook
        "worksheet_name": worksheet_name, -- worksheet name for chart cell references
        "chart_name": worksheet_name, -- Chart Title
        "series_name": "Bull", -- Name of first series
        "series_name2": "Bear", -- Name of second series
        "series_color": "green" -- Color of series 1 line
        "series_color2": "red" -- Color of series 2 line
        "x_axis_range": "E1:E60", -- Range of cells on worksheet_name to get x values from
        "x_axis_range2": "I1:I60", -- Range of cells on worksheet_name for series 2 x values
        "y_axis_range": "A2:A60", -- Range of cells on worksheet_name to get y values from
        "x_axis_name": "Date", -- Title of x axis
        "y_axis_name": "Close", -- Title of y axis
        "position_cell": "H2", -- Cell to position top left corner of the table
        "width_px": 1280, -- Chart width in pixels, use multiple of default column width (64 px)
        "height_px": 800, -- Chart height in pixels, use multiple of default column width (64 px)
        "bull_bear": True -- Adds bear series if present.  Requires x_axis_range2 & series_name2
     }
    """
    new_chart = chart["workbook"].add_chart({"type": "line"})

    new_chart.add_series(  # pyright: ignore [reportOptionalMemberAccess]
        {
            "name": chart["series_name"],
            "categories": f'=\'{chart["worksheet_name"]}\'!{chart["y_axis_range"]}',
            "values": f'=\'{chart["worksheet_name"]}\'!{chart["x_axis_range"]}',
            "line": {"color": chart["series_color"], "width": 0.25},
        }
    )

    if chart.get("bull_bear"):
        new_chart.add_series(  # pyright: ignore [reportOptionalMemberAccess]
            {
                "name": chart["series_name2"],
                "categories": f'=\'{chart["worksheet_name"]}\'!{chart["y_axis_range"]}',
                "values": f'=\'{chart["worksheet_name"]}\'!{chart["x_axis_range2"]}',
                "line": {"color": chart["series_color2"], "width": 0.25},
            }
        )

    new_chart.set_title(  # pyright: ignore [reportOptionalMemberAccess]
        {"name": chart["chart_name"]}
    )
    new_chart.set_x_axis(  # pyright: ignore [reportOptionalMemberAccess]
        {
            "date_axis": True,
            "reverse": True,
            "name": chart["x_axis_name"],
            "interval_unit": chart["x_axis_major_unit"],
            "interval_tick": chart["x_axis_major_unit"],
        }
    )
    new_chart.set_y_axis(  # pyright: ignore [reportOptionalMemberAccess]
        {
            "name": chart["y_axis_name"],
            "major_gridlines": {
                "visible": True,
                "line": {"width": 0.5, "dash_type": "dash", "color": "gray"},
            },
        }
    )

    new_chart.set_size(
        {"width": chart["width_px"], "height": chart["height_px"]}
    )  # pyright: ignore

    new_chart.set_legend({"none": True})

    new_chart.show_na_as_empty_cell()

    chart["worksheet"].insert_chart(
        chart["position_cell"], new_chart, {"x_offset": 0, "y_offset": 0}
    )

