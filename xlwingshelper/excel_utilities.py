import xlwings as xw
import os


def FreezePanes(ws=None, row=1, col=0):
    """
    ウインドウ枠の固定

    Args:
        ws (xw.Sheet, optional): 固定したいワークシート
        row (int, optional): 固定したい行番号
        col (int, optional): 固定したい列番号
    """
    try:
        # wsがなければアクティブなものを取得
        if ws is None:
            ws = xw.books.active.sheets.active

        ws.activate()
        wb = ws.book  # wsからbookを確実に取得

        aw = wb.app.api.ActiveWindow
        if aw is None:
            print("ActiveWindowが取得できません。")
            return False

        aw.FreezePanes = False
        aw.SplitColumn = col
        aw.SplitRow = row
        aw.FreezePanes = True
        return True
    except Exception as e:
        print(f"FreezePanesで例外発生: {e}")
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
            print("ActiveWindowが取得できません")
            return False

        aw.FreezePanes = False
        aw.SplitColumn = 0
        aw.SplitRow = 0
        return True
    except Exception as e:
        print(f"FreezePanes0で例外発生: {e}")
        return False


def check_wb_create(save_wb_path, extension=".xlsx"):
    """同名のブックが存在するかチェックしてブック作成
    Args:
        book_name (str): ブック名
        wb (xw.Book, optional): xw.Book. Defaults to None.
        position (int, optional): 位置. Defaults to 0.

    Returns:
        bool: True:存在する, False:存在しない
        wb: ワークブック
    """

    if os.path.exists(save_wb_path):
        counter = 2
        while True:
            try:
                os.rename(
                    save_wb_path,
                    str(os.path.splitext(save_wb_path)[0]) + " (" + str(counter) + ")" + extension,
                )
            except:
                pass
            else:
                break
            if counter > 50:  # 50も作成してたらおかしいので終了
                return False
            counter += 1
    wb = xw.Book()
    wb.save(save_wb_path)

    return wb


def check_sheet_add(sheet_name, wb=None, position=0):
    """同名シートが存在するかチェックしてシート追加

    Args:
        sheet_name (str): シート名
        wb (xw.Book, optional): xw.Book. Defaults to None.
        position (int, optional): 位置. Defaults to 0.

    Returns:
        bool: True:存在する, False:存在しない
        sh: シート
    """

    if wb == None:  # Excelが起動していない場合はFalseを返す
        try:
            wb = xw.books.active
        except:
            return False

    try:
        sh = wb.sheets.add(sheet_name, before=wb.sheets[position])
    except:
        sh = wb.sheets[sheet_name]
        all_sh_name = [sh.name for sh in wb.sheets]
        counter = 2

        while True:
            if f"{sheet_name} ({counter})" in all_sh_name:
                counter += 1
                if counter > 50:
                    return False  # 50も作成してたらおかしいのでその場合はFalseを返す
            else:
                break

        sh.name = f"{sheet_name} ({counter})"
        sh = wb.sheets.add(sheet_name, before=wb.sheets[position])

    return sh


def check_sheet_copy(sheet_source_name, wb_source=None, wb_destination=None, position=0):
    """同名シートが存在するかチェックしてシートコピー

    Args:
        sheet_name (str): シート名
        wb (xw.Book, optional): xw.Book. Defaults to None.
        position (int, optional): 位置. Defaults to 0.

    Returns:
        bool: True:存在する, False:存在しない
        sh: シート
    """

    if wb_destination == None:  # Excelが起動していない場合はFalseを返す
        try:
            wb_destination = xw.books.active
        except:
            return False

    all_sh_name = [sh.name for sh in wb_destination.sheets]

    if sheet_source_name in all_sh_name:
        # 同名シートが存在するので(*)を付ける
        counter = 2

        while True:
            if f"{sheet_source_name} ({counter})" in all_sh_name:
                counter += 1
                if counter > 50:
                    return False  # 50も作成してたらおかしいのでその場合はFalseを返す
            else:
                break
        sheet = wb_destination.sheets[sheet_source_name]
        sheet.name = f"{sheet_source_name} ({counter})"

    add_sh = wb_source.sheets[sheet_source_name].copy(before=wb_destination.sheets[position])

    return add_sh
