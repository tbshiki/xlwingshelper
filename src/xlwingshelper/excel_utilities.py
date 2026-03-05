import logging
import os

import xlwings as xw

# loggingを設定
logging.basicConfig(level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s")


def FreezePanes(ws=None, row=1, col=0):
    """
    ウインドウ枠の固定

    Args:
        ws (xw.Sheet, optional): 固定したいワークシート
        row (int, optional): 固定したい行番号
        col (int, optional): 固定したい列番号
    """
    try:
        if ws is None:
            ws = xw.books.active.sheets.active

        ws.activate()
        wb = ws.book

        aw = wb.app.api.ActiveWindow
        if aw is None:
            logging.warning("ActiveWindowが取得できません。")
            return False

        aw.FreezePanes = False
        aw.SplitColumn = col
        aw.SplitRow = row
        aw.FreezePanes = True
        return True
    except Exception as e:
        logging.error(f"FreezePanesで例外発生: {e}")
        return False


def FreezePanes0(ws=None):
    """
    ウインドウ枠固定の解除

    Args:
        ws (xw.Sheet, optional): 操作対象シート
    """
    try:
        if ws is None:
            ws = xw.books.active.sheets.active

        ws.activate()
        wb = ws.book
        aw = wb.app.api.ActiveWindow

        if aw is None:
            logging.warning("ActiveWindowが取得できません")
            return False

        aw.FreezePanes = False
        aw.SplitColumn = 0
        aw.SplitRow = 0
        return True
    except Exception as e:
        logging.error(f"FreezePanes0で例外発生: {e}")
        return False


def apply_autofilter(sheet, data_row_count, last_col, header_row=1, first_col=1):
    """
    ヘッダー付き範囲に AutoFilter を安全に適用する。
    - data_row_count <= 0 の場合は適用せず、既存フィルタ状態のみ解除する
    - COM 例外時は AutoFilter(1) をフォールバックで試行する

    Args:
        sheet (xw.Sheet): 対象シート
        data_row_count (int): ヘッダーを除くデータ行数
        last_col (int): 最終列番号（1始まり）
        header_row (int, optional): ヘッダー行番号. Defaults to 1.
        first_col (int, optional): 開始列番号. Defaults to 1.

    Returns:
        bool: AutoFilter を適用できた場合 True。未適用/失敗時は False。
    """
    if sheet is None:
        logging.warning("apply_autofilter: sheet が None のため処理をスキップします。")
        return False

    try:
        row_count = int(data_row_count)
    except Exception:
        row_count = 0

    try:
        if bool(sheet.api.FilterMode):
            sheet.api.ShowAllData()
    except Exception:
        pass

    try:
        sheet.api.AutoFilterMode = False
    except Exception:
        pass

    if row_count <= 0:
        return False

    if last_col < first_col:
        logging.warning("apply_autofilter: 列指定が不正です。first_col=%s, last_col=%s", first_col, last_col)
        return False

    last_row = header_row + row_count  # +1 はヘッダー行
    target = sheet.range((header_row, first_col), (last_row, last_col))

    try:
        target.api.AutoFilter()
        return True
    except Exception as exc:
        logging.warning("apply_autofilter: AutoFilter() 失敗。フォールバックを試行します: %s", exc)

    try:
        target.api.AutoFilter(1)
        return True
    except Exception as exc:
        logging.warning("apply_autofilter: AutoFilter(1) も失敗しました: %s", exc)
        return False


def check_wb_create(save_wb_path):
    """
    同名のブックが存在するかチェックしてブックを作成します。
    存在する場合、既存のブックをリネームしてバックアップします。

    Args:
        save_wb_path (str): 保存するブックのパス

    Returns:
        xw.Book or None: 作成されたワークブック、または失敗した場合はNone
    """
    try:
        if os.path.exists(save_wb_path):
            base, ext = os.path.splitext(save_wb_path)
            counter = 2
            while True:
                new_path = f"{base} ({counter}){ext}"
                try:
                    os.rename(save_wb_path, new_path)
                    logging.info(f"既存のファイル '{save_wb_path}' を '{new_path}' にリネームしました。")
                    break
                except OSError:
                    counter += 1
                    if counter > 50:
                        logging.error("50回リネームを試みましたが、ユニークなファイル名を作成できませんでした。")
                        return None
        wb = xw.Book()
        wb.save(save_wb_path)
        return wb
    except Exception as e:
        logging.error(f"check_wb_createで例外発生: {e}")
        return None


def check_sheet_add(sheet_name, wb=None, position=0):
    """
    同名シートが存在するかチェックしてシートを追加します。
    存在する場合、既存のシートをリネームします。

    Args:
        sheet_name (str): シート名
        wb (xw.Book, optional): 対象のワークブック. Defaults to active book.
        position (int, optional): 追加する位置. Defaults to 0.

    Returns:
        xw.Sheet or None: 追加されたシート、または失敗した場合はNone
    """
    try:
        if wb is None:
            wb = xw.books.active
    except Exception as e:
        logging.error(f"アクティブなワークブックの取得に失敗: {e}")
        return None

    try:
        return wb.sheets.add(sheet_name, before=wb.sheets[position])
    except ValueError:
        logging.warning(f"シート '{sheet_name}' は既に存在します。既存のシートをリネームして新しいシートを追加します。")
        try:
            existing_sheet = wb.sheets[sheet_name]
            all_sheet_names = [sh.name for sh in wb.sheets]
            counter = 2
            while True:
                new_sheet_name = f"{sheet_name} ({counter})"
                if new_sheet_name not in all_sheet_names:
                    break
                counter += 1
                if counter > 50:
                    logging.error("50回試みましたが、ユニークなシート名を作成できませんでした。")
                    return None
            existing_sheet.name = new_sheet_name
            return wb.sheets.add(sheet_name, before=wb.sheets[position])
        except Exception as e:
            logging.error(f"シートの追加/リネーム中にエラーが発生: {e}")
            return None


def check_sheet_copy(sheet_source_name, wb_source=None, wb_destination=None, position=0):
    """
    同名シートが存在するかチェックしてシートをコピーします。
    コピー先に同名シートが存在する場合、既存のシートをリネームします。

    Args:
        sheet_source_name (str): コピー元のシート名
        wb_source (xw.Book, optional): コピー元のワークブック. Defaults to active book.
        wb_destination (xw.Book, optional): コピー先のワークブック. Defaults to source book.
        position (int, optional): コピー先の位置. Defaults to 0.

    Returns:
        xw.Sheet or None: コピーされたシート、または失敗した場合はNone
    """
    try:
        if wb_source is None:
            wb_source = xw.books.active
        if wb_destination is None:
            wb_destination = wb_source
    except Exception as e:
        logging.error(f"アクティブなワークブックの取得に失敗: {e}")
        return None

    try:
        all_dest_sheet_names = [sh.name for sh in wb_destination.sheets]
        if sheet_source_name in all_dest_sheet_names:
            logging.warning(f"コピー先のブックにシート '{sheet_source_name}' は既に存在します。既存のシートをリネームします。")
            existing_sheet = wb_destination.sheets[sheet_source_name]
            counter = 2
            while True:
                new_sheet_name = f"{sheet_source_name} ({counter})"
                if new_sheet_name not in all_dest_sheet_names:
                    break
                counter += 1
                if counter > 50:
                    logging.error("50回試みましたが、ユニークなシート名を作成できませんでした。")
                    return None
            existing_sheet.name = new_sheet_name

        source_sheet = wb_source.sheets[sheet_source_name]
        return source_sheet.copy(before=wb_destination.sheets[position])
    except Exception as e:
        logging.error(f"シートのコピー中にエラーが発生: {e}")
        return None
