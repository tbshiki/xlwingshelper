import logging
import sys
import ctypes
from ctypes import wintypes

# loggingを設定
logging.basicConfig(level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s")

# Windows環境でのみwin32関連のモジュールをインポート
win32api = None
win32con = None
if sys.platform == "win32":
    try:
        import win32api
        import win32con
    except ImportError:
        logging.error("pywin32がインストールされていません。'pip install pywin32' を実行してください。")
        pass

FLASHW_ALL = 0x00000003
FLASHW_TIMERNOFG = 0x0000000C


class FLASHWINFO(ctypes.Structure):
    _fields_ = [
        ("cbSize", wintypes.UINT),
        ("hwnd", wintypes.HWND),
        ("dwFlags", wintypes.DWORD),
        ("uCount", wintypes.UINT),
        ("dwTimeout", wintypes.DWORD),
    ]


def _get_hwnd(application):
    try:
        return application.app.hwnd
    except Exception:
        return None


def highlight_window(application, flash=True, flash_count=5, activate=True):
    """
    Excelウィンドウを目立たせる（前面化 + タスクバー点滅）。

    Args:
        application: xlwingsのBook/Appオブジェクト
        flash (bool): 点滅を行うか
        flash_count (int): 点滅回数
        activate (bool): 点滅前にExcelを前面化するか

    Returns:
        bool: 処理が実行できた場合はTrue
    """
    if sys.platform != "win32":
        return False

    hwnd = _get_hwnd(application)
    if hwnd is None:
        return False

    if activate:
        try:
            application.app.activate(steal_focus=True)
        except Exception:
            pass

    if not flash:
        return True

    try:
        flash_info = FLASHWINFO(
            ctypes.sizeof(FLASHWINFO),
            hwnd,
            FLASHW_ALL | FLASHW_TIMERNOFG,
            max(1, int(flash_count)),
            0,
        )
        ctypes.windll.user32.FlashWindowEx(ctypes.byref(flash_info))
        return True
    except Exception as exc:
        logging.debug(f"ウィンドウ点滅の実行に失敗しました: {exc}")
        return False


def MessageBox(
    application,
    alert="エラーが発生しました",
    title="エラー",
    button="MB_OK",
    icon="MB_ICONERROR",
    flash=True,
    flash_count=5,
):
    """
    Windows環境でのみメッセージボックスを表示します。

    Args:
        application: xlwingsのAppオブジェクト
        alert (str): メッセージボックスに表示するテキスト
        title (str): メッセージボックスのタイトル
        button (str): 表示するボタンの種類 (例: "MB_OK", "MB_OKCANCEL")
        icon (str): 表示するアイコンの種類 (例: "MB_ICONERROR", "MB_ICONINFORMATION")
        flash (bool): 表示前にExcelウィンドウを点滅させるか
        flash_count (int): 点滅回数

    Returns:
        int: ユーザーがクリックしたボタンを示す整数値
    """
    if sys.platform != "win32":
        logging.warning("MessageBox はWindows 環境でのみ利用可能です。")
        return None

    if win32api is None or win32con is None:
        logging.error("pywin32が利用できないため MessageBox を表示できません。")
        return None

    try:
        highlight_window(application, flash=flash, flash_count=flash_count, activate=True)
        button_flag = win32con.__dict__.get(button, win32con.MB_OK)
        icon_flag = win32con.__dict__.get(icon, win32con.MB_ICONERROR)
        flags = button_flag | icon_flag

        return win32api.MessageBox(application.app.hwnd, alert, title, flags)
    except Exception as e:
        logging.error(f"MessageBoxの表示中にエラーが発生しました: {e}")
        return None
