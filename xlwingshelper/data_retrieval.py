import xlwings as xw
from typing import List, Tuple

from xlwingshelper.utility_functions import *


def get_max_row(sheet):
    """
    旧名 : all
    Finds the maximum row number with data in any column of the given Excel sheet.

    Args:
        sheet: The Excel sheet to be analyzed.

    Returns:
        The maximum row number with data.

    Raises:
        ValueError: If the sheet is empty or other unexpected conditions occur.

    """
    max_col = sheet.range(1, sheet.cells.last_cell.column).end("left").column

    last_row = 1
    for i in range(1, max_col + 1):
        max_row = sheet.range(sheet.cells.last_cell.row, i).end("up").row
        if max_row > last_row:
            last_row = max_row

    return last_row


def get_last_cell(sheet):
    """
    旧名 : last
    Finds the coordinates of the last cell with data in the specified Excel sheet.

    This function locates the rightmost column with data in the first row and then
    iterates through each column to find the bottommost row with data.

    Example:
        last_col, last_row = get_last_cell_coordinates(sheet)

    Args:
        sheet (xlwings.Sheet): The Excel sheet to be analyzed.

    Returns:
        tuple: A tuple containing the index of the last column with data and the index of the last row with data.

    """
    last_col = sheet.range(1, sheet.cells.last_cell.column).end("left").column
    last_row = 1

    for i in range(1, last_col + 1):
        current_row = sheet.range(sheet.cells.last_cell.row, i).end("up").row
        if current_row > last_row:
            last_row = current_row

    return last_col, last_row


def get_last_row_in_col(sheet, col=1):
    """
    Finds the last row with data in a specified column of the given Excel sheet.

    Example:
        last_row = get_last_row_in_col(sheet, col=1)

    Args:
        sheet (xlwings.Sheet): The Excel sheet to be analyzed.
        col (int): The column index (1-based) to find the last row with data.

    Returns:
        int: The row number of the last cell with data in the specified column.

    Raises:
        ValueError: If the specified column index is invalid or out of range.

    """
    if not isinstance(col, int) or col < 1:
        raise ValueError("Invalid column index. It must be an integer greater than 0.")

    # Convert column index to letter representation for use with xlwings
    col_letter = num_to_alpha(col)
    last_row = sheet.range(sheet.cells.last_cell.row, col_letter).end("up").row
    return last_row


def get_col_values(
    sheet, colstart: int = 0, colend: int = 0
) -> Tuple[int, int, List[List]]:
    """
    旧名 : all_col,col_dict
    Extracts the values of columns from the specified Excel sheet.

    Args:
        sheet: The Excel sheet from which to extract values.
        colstart (int): The starting column index (0-indexed).
        colend (int): The ending column index (0-indexed). If 0, reads till the last column.

    Returns:
        Tuple containing the last column index, last row index, and a list of column values.

    Raises:
        ValueError: If `colstart` or `colend` are out of range.

    """
    last_col, last_row = get_last_cell(sheet)

    if colend == 0 or colend > last_col:
        colend = last_col
    if colstart < 0 or colend < colstart:
        raise ValueError("Invalid column start or end index")

    col_values = []
    for col in range(colstart, colend):
        strcol = num_to_alpha(col + 1)
        col_values.append(sheet.range(strcol + "1:" + strcol + str(last_row)).value)

    return last_col, last_row, col_values


def get_row_values(sheet, rowstart: int = 0, rowend: int = 0):
    """
    旧名 : all_row
    Extracts values of rows from the specified Excel sheet and returns them as a list of lists.

    Args:
        sheet: The Excel sheet from which to extract values.
        rowstart (int): The starting row index (1-indexed).
        rowend (int): The ending row index (1-indexed). If 0, reads till the last row.

    Returns:
        Tuple containing the last column index, last row index, and a list of row values.

    Raises:
        ValueError: If `rowstart` or `rowend` are out of range.

    """
    last_col, last_row = get_last_cell(sheet)

    if rowend == 0 or rowend > last_row:
        rowend = last_row
    if rowstart < 1 or rowend < rowstart:
        raise ValueError("Invalid row start or end index")

    row_values = []
    strlastcol = num_to_alpha(last_col)
    for row in range(rowstart, rowend):
        row_values.append(
            sheet.range("A" + str(row + 1) + ":" + strlastcol + str(row + 1)).value
        )

    return last_col, last_row, row_values
