import logging
import sys

# loggingを設定
logging.basicConfig(level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s")

# Windows環境でのみwin32関連のモジュールをインポート
if sys.platform == "win32":
    try:
        import win32api
        import win32con
    except ImportError:
        logging.error("pywin32がインストールされていません。'pip install pywin32' を実行してください。")

        # pywin32がなくてもプログラムが停止しないように、ダミーの関数を定義
        def MessageBox(*args, **kwargs):
            logging.error("MessageBox はWindows 環境でのみ利用可能です。")
            return None


def MessageBox(application, alert="エラーが発生しました", title="エラー", button="MB_OK", icon="MB_ICONERROR"):
    """
    Windows環境でのみメッセージボックスを表示します。

    Args:
        application: xlwingsのAppオブジェクト
        alert (str): メッセージボックスに表示するテキスト
        title (str): メッセージボックスのタイトル
        button (str): 表示するボタンの種類 (例: "MB_OK", "MB_OKCANCEL")
        icon (str): 表示するアイコンの種類 (例: "MB_ICONERROR", "MB_ICONINFORMATION")

    Returns:
        int: ユーザーがクリックしたボタンを示す整数値
    """
    if sys.platform != "win32":
        logging.warning("MessageBox はWindows 環境でのみ利用可能です。")
        return None

    try:
        button_flag = win32con.__dict__.get(button, win32con.MB_OK)
        icon_flag = win32con.__dict__.get(icon, win32con.MB_ICONERROR)
        flags = button_flag | icon_flag

        return win32api.MessageBox(application.app.hwnd, alert, title, flags)
    except Exception as e:
        logging.error(f"MessageBoxの表示中にエラーが発生しました: {e}")
        return None
