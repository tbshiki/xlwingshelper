import logging
from typing import Any, List, Tuple, cast

import xlwings as xw

# loggingを設定
logging.basicConfig(level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s")

RangeValues = List[List[Any]] | List[Any]


def get_last_cell(sheet: xw.Sheet) -> Tuple[int, int]:
    """
    指定されたExcelシートでデータを持つ最後のセルの座標を見つけます。
    これはused_rangeに基づいています。

    Args:
        sheet (xw.Sheet): 分析対象のExcelシート。

    Returns:
        tuple: 最後の列のインデックスと最後の行のインデックスを含むタプル。(列, 行)
               シートが空の場合やエラーが発生した場合は(0, 0)を返します。
    """
    try:
        # 完全に空のシートのused_rangeはA1です。A1も空かどうかを確認します。
        if sheet.used_range.address == "$A$1" and sheet.range("A1").value is None:
            return 0, 0
        last_cell = sheet.used_range.last_cell
        return last_cell.column, last_cell.row
    except Exception as e:
        logging.error(f"get_last_cellで例外発生: {e}")
        return 0, 0


def get_max_row(sheet: xw.Sheet) -> int:
    """
    指定されたExcelシートのいずれかの列にあるデータを持つ最大行番号を見つけます。

    Args:
        sheet (xw.Sheet): 分析対象のExcelシート。

    Returns:
        int: データを持つ最大行番号。シートが空またはエラーの場合は0を返します。
    """
    try:
        _, last_row = get_last_cell(sheet)
        return last_row
    except Exception as e:
        logging.error(f"get_max_rowで例外発生: {e}")
        return 0


def get_last_row_in_col(sheet: xw.Sheet, col: int | str = 1) -> int:
    """
    指定されたExcelシートの指定された列でデータを持つ最後の行を見つけます。

    Args:
        sheet (xw.Sheet): 分析対象のExcelシート。
        col (int or str): 列インデックス（1から始まる）または文字（例: 'A'）。

    Returns:
        int: 指定された列のデータを持つ最後のセルの行番号。

    Raises:
        ValueError: 指定された列インデックスが無効な場合。
    """
    try:
        if isinstance(col, str):
            col_num = sheet.range(f"{col}1").column
        elif isinstance(col, int):
            if col < 1:
                raise ValueError("列インデックスは1以上でなければなりません。")
            col_num = col
        else:
            raise ValueError("列は整数または文字列でなければなりません。")

        return sheet.range(sheet.cells.last_cell.row, col_num).end("up").row
    except Exception as e:
        logging.error(f"get_last_row_in_colで例外発生: {e}")
        if "ValueError" in str(e):
            raise ValueError(f"無効な列識別子: {col}") from e
        return 0


def get_col_values(sheet: xw.Sheet, col_start: int = 1, col_end: int = 0) -> Tuple[int, int, RangeValues]:
    """
    指定されたExcelシートから列の値を抽出します。

    Args:
        sheet (xw.Sheet): 値を抽出するExcelシート。
        col_start (int): 開始列インデックス（1から始まる）。
        col_end (int): 終了列インデックス（1から始まる）。0の場合、使用されている最後の列まで読み取ります。

    Returns:
        Tuple[int, int, List[List]]: 以下を含むタプル:
            - 使用された最後の列インデックス
            - 使用された最後の行インデックス
            - 各内部リストが列の値を表すリストのリスト。
    """
    try:
        last_col, last_row = get_last_cell(sheet)
        if last_col == 0:
            return 0, 0, []

        if col_end == 0 or col_end > last_col:
            col_end = last_col
        if col_start < 1 or col_end < col_start:
            raise ValueError("無効な開始または終了列インデックスです。")

        raw_data = sheet.range((1, col_start), (last_row, col_end)).options(transpose=True).value
        if raw_data is None:
            return last_col, last_row, []

        if col_start == col_end:
            return last_col, last_row, [raw_data]
        return last_col, last_row, cast(RangeValues, raw_data)
    except Exception as e:
        logging.error(f"get_col_valuesで例外発生: {e}")
        return 0, 0, []


def get_row_values(sheet: xw.Sheet, row_start: int = 1, row_end: int = 0) -> Tuple[int, int, RangeValues]:
    """
    指定されたExcelシートから行の値を抽出します。

    Args:
        sheet (xw.Sheet): 値を抽出するExcelシート。
        row_start (int): 開始行インデックス（1から始まる）。
        row_end (int): 終了行インデックス（1から始まる）。0の場合、使用されている最後の行まで読み取ります。

    Returns:
        Tuple[int, int, List[List]]: 以下を含むタプル:
            - 使用された最後の列インデックス
            - 使用された最後の行インデックス
            - 各内部リストが行の値を表すリストのリスト。
    """
    try:
        last_col, last_row = get_last_cell(sheet)
        if last_row == 0:
            return 0, 0, []

        if row_end == 0 or row_end > last_row:
            row_end = last_row
        if row_start < 1 or row_end < row_start:
            raise ValueError("無効な開始または終了行インデックスです。")

        raw_data = sheet.range((row_start, 1), (row_end, last_col)).options(ndim=2).value
        if raw_data is None:
            return last_col, last_row, []
        return last_col, last_row, cast(RangeValues, raw_data)
    except Exception as e:
        logging.error(f"get_row_valuesで例外発生: {e}")
        return 0, 0, []
