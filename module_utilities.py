"""Utility functions for template runner"""

from xlsxwriter.format import Format
from xlsxwriter.utility import xl_cell_to_rowcol, xl_rowcol_to_cell
from xlsxwriter.worksheet import Worksheet


def add_merged_data(cell_range: str, text: str, worksheet: Worksheet, fmt: Format):
    """
    Adds text with black background and white text

    Args:
        cell_Range: The cell range to be merged. e.g K68:L68
        text: The text to write
        worksheet: The current worksheet
        fmt: The format to use
    """
    worksheet.merge_range(
        cell_range, text, fmt
    )  # pyright: ignore[reportGeneralTypeIssues]


def get_offset_cell(start_cell: str, row_offset: int, col_offset: int) -> str:
    """
    Gets the cell range at the offset designated by the two offset values

    Args:
        start_cell: The first cell and one to base the offsets relative to
        row_offset: The row offset from start_cell
        col_offset: The column offset from start_cell

    Returns:
        The cell at the offset position, e.g C5

    """
    cell_positions = xl_cell_to_rowcol(start_cell)
    new_row = cell_positions[0] + row_offset
    new_col = cell_positions[1] + col_offset
    return xl_rowcol_to_cell(new_row, new_col)  # type: ignore


def append_offset_cell(start_cell: str, row_offset: int, col_offset: int) -> str:
    """
    Returns a two cell range with start_cell as the first and the second cell based on the offset.
    For example: A1:C5

    Args:
        start_cell: The first cell and one to base the offsets relative to
        row_offset: The row offset from start_cell
        col_offset: The column offset from start_cell

    Returns:
        Cell range, e.g A1:C5

    """
    new_cell = get_offset_cell(start_cell, row_offset, col_offset)
    return f"{start_cell}:{new_cell}"  # type: ignore


def add_col_data(start_cell: str, data: list, worksheet: Worksheet, fmt: Format):
    """
    Writes a list of data in column format

    Args:
        start_cell: Cell to begin writing in
        data: List of data to write
        worksheet: Worksheet to write to
        fmt: Format to use
    """
    coords = xl_cell_to_rowcol(start_cell)
    row = coords[0]
    col = coords[1]
    count = 0

    while count < len(data):
        worksheet.write(row, col, data[count], fmt)
        row = row + 1
        count = count + 1


def add_textbox(cell: str, worksheet: Worksheet, text: str):
    """
    Adds textbox to the worksheet

    Args:
        cell: Cell fopr top left corner
        worksheet: Worksheet to write to
        text: The text to write
    """
    worksheet.insert_textbox(
        cell,
        text,
        {
            "line": {"color": "black", "width": 2},
            "width": 622,
            "height": 260,
            "align": {"vertical": "middle", "horizontal": "top"},
            "font": {"size": 12, "name": "Tenorite"},
        },
    )


def add_merge_col_data(
    start_cell: str, data: list, num_cols: int, worksheet: Worksheet, fmt: Format
):
    """
    Writes a list of data in column format, merging num_cols number of columns

    Args:
        start_cell: Cell to begin writing in
        data: List of data to write
        worksheet: Worksheet to write to
        fmt: Format to use
    """
    cell: str = start_cell
    count = 0

    while count < len(data):
        coords = xl_cell_to_rowcol(cell)
        row = coords[0]
        col = coords[1]

        cell = append_offset_cell(cell, 0, num_cols)
        add_merged_data(cell, data[count], worksheet, fmt)

        cell = xl_rowcol_to_cell(row + 1, col)  # type: ignore
        count = count + 1

