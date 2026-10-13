# -*- coding: utf-8 -*-
"""桌面通用：平台判断、窗口标题、提示框、单实例。"""
import sys

IS_WIN = sys.platform == "win32"

#: 控制台窗口标题。找窗口、单实例都靠它，不要改。
APP_TITLE = "PPT 控制台"

_MUTEX = None   # 保持引用，进程结束前互斥体一直有效


def message_box(text, error=True):
    if IS_WIN:
        import ctypes
        ctypes.windll.user32.MessageBoxW(0, str(text), APP_TITLE, 0x10 if error else 0x40)
    else:
        print(text)


def ensure_single_instance():
    """已经开着一个控制台时，把那个窗口叫到前面，返回 False（调用方应退出）。"""
    global _MUTEX
    if not IS_WIN:
        return True
    import ctypes
    k32, u32 = ctypes.windll.kernel32, ctypes.windll.user32
    _MUTEX = k32.CreateMutexW(None, False, "Local\\PPT-Presenter-ControlPanel")
    if k32.GetLastError() == 183:  # ERROR_ALREADY_EXISTS
        hwnd = u32.FindWindowW(None, APP_TITLE)   # 另一个进程的窗口，可以用 FindWindow
        if hwnd:
            u32.ShowWindow(hwnd, 9)  # SW_RESTORE
            u32.SetForegroundWindow(hwnd)
        else:
            message_box("PPT 控制台已经在运行了。", error=False)
        return False
    return True
