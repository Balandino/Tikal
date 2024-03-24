"""The profile module """

from xlsxwriter.format import Format

from module_utilities import (
    add_col_data,
    add_merge_col_data,
    add_merged_data,
    add_textbox,
    append_offset_cell,
    get_offset_cell,
)


def load_profile(start_cell: str, module: dict, formats: dict[str, Format]) -> str:
    """
    Profile module

    Args:
        start_cell: The cell to start writing to
        module: Data dictionary
        formats: Dictionary of formats to use

    Returns:
        Cell position for next module to start at, e,g A20
    """

    write_heading(start_cell, module, formats)
    write_ticker_info(start_cell, module, formats)
    write_description(start_cell, module)

    return ""


def write_description(start_cell: str, module: dict):
    """
    Inserts textbox for description

    Args:
        start_cell: Cell for top left corner
        module: Dictionary containing worksheet
    """
    cell = get_offset_cell(start_cell, 16, 0)
    add_textbox(cell, module["worksheet"], "Profile")


def write_ticker_info(start_cell: str, module: dict, formats: dict[str, Format]):
    """
    Write ticker data, such as exchange and sector

    Args:
        start_cell: Cell to base offsets from
        module: Data dictionary
        formats: Formats dictionary
    """
    data = [
        "Ticker",
        "Current Price",
        "Direction",
        "Exchange",
        "Sector",
        "Industry",
        "Date",
    ]

    cell = get_offset_cell(start_cell, 7, 0)
    add_col_data(cell, data, module["worksheet"], formats["sub_heading_left"])

    cell = get_offset_cell(start_cell, 7, 1)
    add_merge_col_data(cell, data, 2, module["worksheet"], formats["data_type_1_left"])


def write_heading(start_cell: str, module: dict, formats: dict[str, Format]):
    """
    Adds the main heading text

    Args:
        start_cell: Cell to base offsets from
        module: Data dictionary
    """
    cell = get_offset_cell(start_cell, 1, 0)
    cell = append_offset_cell(cell, 3, 8)

    add_merged_data(
        cell, module["ticker"], module["worksheet"], formats["main_heading_center"]
    )

