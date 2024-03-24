"""Runs sector screener for improved visualisation"""

import xlwings


def run(sheet_name: str):
    """
    Runs the screener

    Args:
        sheet_name: Name of active sheet in workbook
    """

    workbook = xlwings.Book.caller()
    sheet = workbook.sheets[sheet_name]
    last_row = sheet.range("L1").end("down").row
    error_list: list[bool] = []

    print(f"Last Row: {last_row}")
    error_list.append(checkLayout(sheet))

    if checkErrors(error_list):
        return

    print("Layout checks complete")

    error_list.append(check_growth_calc(sheet, "M", "N", "R", last_row))  # Revenue FY0
    error_list.append(check_growth_calc(sheet, "N", "O", "S", last_row))  # Revenue FY1
    error_list.append(check_growth_calc(sheet, "O", "P", "T", last_row))  # Revenue FY2
    print("Revenue Checked")

    error_list.append(check_growth_calc(sheet, "U", "V", "Z", last_row))  # EG 1
    error_list.append(check_growth_calc(sheet, "V", "W", "AA", last_row))  # EG 2
    error_list.append(check_growth_calc(sheet, "W", "X", "AB", last_row))  # EG 3
    print("EG Checked")

    if checkErrors(error_list):
        return

    error_list.append(check_peg(sheet, "AC", "Z", "AF", last_row))  # PEG F1
    error_list.append(check_peg(sheet, "AD", "AA", "AG", last_row))  # PEG F2
    error_list.append(check_peg(sheet, "AE", "AB", "AH", last_row))  # PEG F3
    print("PEG Checked")

    if checkErrors(error_list):
        return

    colour_columns = [
        "L",
        "R",
        "S",
        "T",
        "Z",
        "AA",
        "AB",
        "AC",
        "AD",
        "AE",
        "AF",
        "AG",
        "AH",
        "AL",
        "AM",
        "AN",
        "AO",
        "AP",
        "AQ",
    ]
    reverse_colour_columns = ["AI"]

    colour_scale(workbook, colour_columns, last_row, False)
    colour_scale(workbook, reverse_colour_columns, last_row, True)
    print("Colour Scale Added")

    macro = workbook.macro("Filterer")
    macro("")

    space_columns = [
        "R:R",
        "V:V",
        "AB:AB",
        "AF:AF",
        "AJ:AJ",
        "AO:AO",
        "AV:AV",
    ]

    add_new_cols(sheet, space_columns)
    print("Columns Added")

    average_columns = [
        "S",
        "T",
        "U",
        "AC",
        "AD",
        "AE",
        "AG",
        "AH",
        "AI",
        "AK",
        "AL",
        "AM",
    ]
    add_averages(sheet, average_columns, last_row)
    print("Averages Added")

    macro = workbook.macro("StockDataPrice")
    macro("")
    print("Tickers converted to stocks and price added")

    # ranking_columns = [
    #     "S",
    #     "T",
    #     "U",
    #     "AC",
    #     "AD",
    #     "AE",
    #     "AG",
    #     "AH",
    #     "AI",
    #     "AK",
    #     "AL",
    #     "AM",
    # ]

    sheet.autofit(axis="columns")


def add_new_cols(sheet, columns: list[str]):
    """
    Inserts new columns

    Args:
        sheet: Sheet with data
        columns: List of column letters in format A:A, B:B
    """
    for col in columns:
        sheet.range(col).insert("right")


def add_averages(sheet: xlwings.Sheet, columns: list[str], last_row: int):
    """
    Adds average formula to cells below main columns

    Args:
        sheet: Sheet to write to
        columns: List of column letters
        last_row: Last used row in columns
    """
    for col in columns:
        average_cell = f"{col}{last_row + 2}"
        average_range = f"{col}2:{col}{last_row}"
        sheet[average_cell].formula = f"=Average({average_range})"
        sheet[average_cell].api.Borders.Weight = 2


def colour_scale(
    workbook: xlwings.Book, columns: list[str], last_row: int, reverse: bool
):
    """
    Uses the workbook macros to colour scale the columns

    Args:
        workbook: Workbookwith macros
        cells: List of column letters
        last_row: Last row num
        reverse: Use reverse colours where red it high and green is low
    """
    if reverse:
        macro = workbook.macro("ReverseColourer")
    else:
        macro = workbook.macro("Colourer")

    for col in columns:
        cell_range = f"{col}2:{col}{last_row}"
        macro(cell_range)


def checkErrors(error_list: list[bool]) -> bool:
    """
    Returns True if error in list, otherwise False.  Holds up program
    via input command to allow user to see output

    Args:
        error_list:

    Returns:
        True if errors, False otherwise

    """
    for error in error_list:
        if error is True:
            print("Errors found, can't proceed.  Press Enter to exit")
            input()
            return True

    return False


def check_peg(
    sheet: xlwings.Sheet, pe_col: str, eg_col: str, peg_col: str, last_row: int
) -> bool:
    """
    Checks PEG.  If the EG is 100% and the PEG is wrong, the PEG will be re-calculated on the principle
    that check_growth_calc has adjusted the EG due to a shift from negative to positive

    Args:
        sheet: Sheet containing data
        pe_col: Column letter of P/E to use
        eg_col: Column letter of EG to use
        peg_col: Column letter of PEG figure
        last_row: Last row of data

    Returns:
        True if errors found, otherwise False
    """
    errors = False

    for cur_row in range(2, last_row + 1):
        pe_value = sheet[f"{pe_col}{str(cur_row)}"].value
        eg_value = sheet[f"{eg_col}{str(cur_row)}"].value
        peg_value = sheet[f"{peg_col}{str(cur_row)}"].value

        if pe_value in ("NULL", "NaN"):
            continue

        if eg_value in ("NULL", "NaN"):
            continue

        pe_value_f = float(pe_value)  # pyright: ignore[reportGeneralTypeIssues]
        eg_value_f = float(eg_value)  # pyright: ignore[reportGeneralTypeIssues]
        peg_value_f = float(peg_value)  # pyright: ignore[reportGeneralTypeIssues]
        peg_value_f = round(peg_value, 4)  # pyright: ignore[reportGeneralTypeIssues]

        calc_peg = round((pe_value_f / eg_value_f) / 100, 4)

        if peg_value_f != calc_peg:
            if eg_value_f in (1, -1):
                # If 100% or -100% then code likely adjusted EG due to shift from negative to positive
                # or vice versa, so re-calculate PEG
                sheet[f"{peg_col}{str(cur_row)}"].value = float(pe_value_f / eg_value_f)
            else:
                sheet.range(f"{peg_col}{str(cur_row)}").color = "#FF0000"
                errors = True

    return errors


def check_growth_calc(
    sheet: xlwings.Sheet, old_col: str, new_col: str, calc_col: str, last_row: int
) -> bool:
    """
    Checks each growth calculation on the spreadsheet for errors

    Args:
        sheet: Sheet with data
        old_col: Column letter of older data for calculation
        new_col: Column letter of newer data for calculation
        calc_col: Column letter of data that sheet has calculated
        last_row: Last row of data on sheet

    Returns:
        True if errors were found, otherwise false
    """
    errors = False

    for cur_row in range(2, last_row + 1):
        old_value = sheet[f"{old_col}{str(cur_row)}"].value
        new_value = sheet[f"{new_col}{str(cur_row)}"].value

        if old_value in ("NULL", "NaN"):
            continue

        if new_value in ("NULL", "NaN"):
            continue

        new_value_f = float(new_value)  # pyright: ignore[reportGeneralTypeIssues]
        old_value_f = float(old_value)  # pyright: ignore[reportGeneralTypeIssues]

        # If moving from negative to positive then should be 100%, no greater
        if old_value_f < 0 < new_value_f:
            growth_val = sheet.range(f"{calc_col}{str(cur_row)}").value = 1
            continue

        # If moving from positive to negative than loss can be no more than -100%
        if old_value_f > 0 > new_value_f:
            growth_val = sheet.range(f"{calc_col}{str(cur_row)}").value = -1
            continue

        if new_value_f < 0 and old_value_f < 0:
            my_growth_val_f = (abs(new_value_f) - abs(old_value_f)) / old_value_f
        else:
            my_growth_val_f = (new_value_f - old_value_f) / old_value_f

        my_growth_val_s = str(my_growth_val_f)[:6]
        my_growth_val_f = round(my_growth_val_f, 4)

        growth_val = sheet.range(f"{calc_col}{str(cur_row)}").value
        growth_val_f = float(growth_val)  # pyright: ignore[reportGeneralTypeIssues]
        growth_val_s = str(growth_val)[:6]
        growth_val_f = round(growth_val_f, 4)

        if growth_val_f != my_growth_val_f:
            # Sometimes there is a 0.01 rounding error in the numbers.  A straight
            # string comparison of the unrounded numbers is a safe fallback
            if growth_val_s != my_growth_val_s:
                sheet.range(f"{calc_col}{str(cur_row)}").color = "#FF0000"
                errors = True

    return errors


def checkLayout(sheet) -> bool:
    """
    Confirm columns are in correct positions in case screener has changed

    Args:
        sheet: Sheet to check

    Returns:
        True if errors were found
    """

    cells = [
        "L1",
        "M1",
        "N1",
        "O1",
        "P1",
        "Q1",
        "R1",
        "S1",
        "T1",
        "U1",
        "V1",
        "W1",
        "X1",
        "Y1",
        "Z1",
        "AA1",
        "AB1",
        "AC1",
        "AD1",
        "AE1",
        "AF1",
        "AG1",
        "AH1",
        "AI1",
        "AJ1",
        "AK1",
        "AL1",
        "AM1",
        "AN1",
        "AO1",
        "AP1",
        "AQ1",
        "AR1",
    ]

    column_names = [
        "Market Cap",
        "Revenue FY0 - Previous Financial Year",
        "Revenue FY1 - Current Financial Year",
        "Revenue FY2 - Next Financial Year",
        "Revenue FY3",
        "Revenue NTM",
        "Revenue Growth FY1",
        "Revenue Growth FY2",
        "Revenue Growth FY3",
        "EPS FY0 - Previous Financial Year",
        "EPS FY1 - Current Financial Year",
        "EPS FY2 - Next Financial Year",
        "EPS FY3",
        "EPS NTM",
        "EG F1",
        "EG F2",
        "EG F3",
        "PE FY1",
        "PE FY2",
        "PE FY3",
        "PEG F1",
        "PEG F2",
        "PEG F3",
        "Debt/Equity %",
        "Net Profit Margin (LTM)",
        "Dividend Yield",
        "Earnings Surprise % FQ-3",
        "Earnings Surprise % FQ-2",
        "Earnings Surprise % FQ-1",
        "Earnings Surprise % - FQ0",
        "Percent Change in FY1  EPS Estimates (Prev 60 Days)",
        "Percent Change in FY2 EPS Estimates (Prev 60 Days)",
        "Next EPS Report Date",
    ]

    errors = False

    for count, cell in enumerate(cells):
        cell_value = sheet.range(cell).value

        if cell_value != column_names[count]:
            print(
                f"[ERROR] Cell {cell} expected to have {column_names[count]}, found {cell_value}"
            )
            errors = True

    return errors

