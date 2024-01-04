import win32api  # 参考:http://housoubu.mizusasi.net/data/prog/p003.html
import win32con


# Windows環境のみ メッセージボックスを表示する
def MessageBox(application, alert="エラーが発生しました", title="エラー", button="MB_OK", icon="MB_ICONERROR"):
    flg = win32api.MessageBox(
        application.app.hwnd, alert, title, win32con.__dict__[button] | win32con.__dict__[icon],
    )
    return flg
