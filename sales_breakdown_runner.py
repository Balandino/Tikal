# pylint: disable=line-too-long
"""Gets Sales Per Region & Sales Per Segment Data"""

from xlsxwriter.worksheet import Worksheet

from fmp import fmp_check_symbols, fmp_sales_per_region, fmp_sales_per_segment
from workbook_utilities import close_workbook, create_workbook


def process_sales(
    worksheet: Worksheet, row: int, col: int, data: list[dict], title: str
) -> int:
    """
    Writes the Sales data to the worksheet

    Args:
        worksheet: Worksheet to write to
        row: Row to start writing to
        col: Column to start at
        data: List of dictionaries from FMP call
        title: Title text to write

    Returns:
        An int representing the next row to start writing data to
    """
    worksheet.write(row, col, title)
    row = row + 2

    for count in range(0, 5):
        block = data[count]
        for key in block:
            block_date = key
            sales_data = block[block_date]
            worksheet.write(row, col, block_date)
            # print(block_date)
            row = row + 1
            for sales in sales_data:
                worksheet.write(row, col, sales)
                worksheet.write(row, col + 1, sales_data[sales])
                row = row + 1
                # print(f"Key: {sales} -> {sales_data[sales]}")

            row = row + 1

    row = row + 3
    return row


# def process_sales_2(data: list[dict]):
#     data_to_return = {}
#
#     date_data = []
#     processed_count = 0
#
#     for count in range(0, 5):
#         block = data[count]
#         for block_date in block:
#             sales_data = block[block_date]
#             date_data.append(block_date)
#             for sales_heading in sales_data:
#                 if sales_heading not in data_to_return:
#                     data_to_return[sales_heading] = [0] * processed_count
#
#                 data_to_return[sales_heading].append(sales_data[sales_heading])
#
#         processed_count = processed_count + 1
#
#     # print(data_to_return)
#     data_to_return["Dates"] = date_data
#     return data_to_return
#
#
# def write_sales(worksheet: Worksheet, row: int, col: int, data: dict) -> int:
#     """
#     Writes processed sales data to the worksheet
#
#     Args:
#         worksheet: Worksheet to write to
#         row: Row to start writing to
#         col: Column to start writing to
#         data: Processed sales data
#
#     Returns:
#         Number of the next row to start writing to
#     """
#     dates = data.pop("Dates")
#     num = 0
#     col_num = col
#     while num < len(dates):
#         worksheet.write(row, col_num, dates[num])
#         col_num += 1
#         num += 1


def run():
    """Runner"""

    WORKBOOK_NAME = "Workbooks/Sales_Per.xlsx"
    WORKBOOK = create_workbook(WORKBOOK_NAME)
    worksheet_name = "Sales"
    worksheet = WORKBOOK.add_worksheet(worksheet_name)

    ticker = input("Ticker: ")

    if ticker not in fmp_check_symbols([ticker]):
        print(f"[Error] Symbol not found: {ticker}")
        print("[Error] Exiting")

    region_sales = fmp_sales_per_region(ticker)
    segment_sales = fmp_sales_per_segment(ticker)

    row = 1
    col = 1

    row = process_sales(worksheet, row, col, region_sales, "Sales Per Region")
    print()
    process_sales(worksheet, row, col, segment_sales, "Sales Per Segment")

    print("Writing Workbook")
    close_workbook(WORKBOOK, WORKBOOK_NAME)
    print("Opening Workbook")


run()

