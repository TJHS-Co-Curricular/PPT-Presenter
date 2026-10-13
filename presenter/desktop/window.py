# -*- coding: utf-8 -*-
"""
程序窗口（pywebview / Edge WebView2）和 Win32 窗口操作。

重要：pywebview 的窗口对象只能在它自己的界面线程里操作。从其他线程（快捷键、网页请求）去改窗口，
会互相等待而卡死（v1.5 时“置顶”一按就无响应就是这个原因）。所以置顶、叫到前面、记大小
一律直接用 Win32 API 对窗口句柄操作，而且只用不会等待界面线程的调用：

    找窗口   InternalGetWindowText（GetWindowText 会给自己进程的窗口发消息并等回应）
    置顶     SetWindowPos + SWP_ASYNCWINDOWPOS
    恢复     ShowWindowAsync
"""
import os
import sys
import threading
import time
import traceback

from presenter.desktop.common import APP_TITLE, IS_WIN

#: 程序窗口是否开着（浏览器模式 / 演示模式下是 False）
WINDOW_OPEN = False
_HWND = None


def main_hwnd():
    """本程序控制台窗口的句柄（只找自己进程里、标题完全相同的可见窗口）。"""
    global _HWND
    if not IS_WIN:
        return None
    import ctypes
    from ctypes import wintypes
    u32 = ctypes.windll.user32
    if _HWND and u32.IsWindow(_HWND):
        return _HWND
    found = []
    me = os.getpid()

    @ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HWND, wintypes.LPARAM)
    def callback(hwnd, _):
        pid = wintypes.DWORD()
        u32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
        if pid.value == me and u32.IsWindowVisible(hwnd):
            buf = ctypes.create_unicode_buffer(256)
            u32.InternalGetWindowText(hwnd, buf, 256)
            if buf.value == APP_TITLE:
                found.append(hwnd)
                return False
        return True

    u32.EnumWindows(callback, 0)
    _HWND = found[0] if found else None
    return _HWND


def set_on_top(on):
    """窗口置顶 / 取消置顶。任何线程都可以调用，不会等待。"""
    hwnd = main_hwnd()
    if not hwnd:
        return False
    import ctypes
    from ctypes import wintypes
    u32 = ctypes.windll.user32
    u32.SetWindowPos.argtypes = [wintypes.HWND, wintypes.HWND, ctypes.c_int, ctypes.c_int,
                                 ctypes.c_int, ctypes.c_int, ctypes.c_uint]
    hwnd_topmost, hwnd_notopmost = -1, -2
    flags = 0x0001 | 0x0002 | 0x0010 | 0x4000  # NOSIZE | NOMOVE | NOACTIVATE | ASYNCWINDOWPOS
    return bool(u32.SetWindowPos(hwnd, hwnd_topmost if on else hwnd_notopmost, 0, 0, 0, 0, flags))


def is_on_top():
    hwnd = main_hwnd()
    if not hwnd:
        return False
    import ctypes
    return bool(ctypes.windll.user32.GetWindowLongW(hwnd, -20) & 0x00000008)  # GWL_EXSTYLE & WS_EX_TOPMOST


def bring_to_front():
    hwnd = main_hwnd()
    if not hwnd:
        return
    import ctypes
    u32 = ctypes.windll.user32
    if u32.IsIconic(hwnd):
        u32.ShowWindowAsync(hwnd, 9)     # SW_RESTORE
    # 前台窗口切换需要一点小技巧，否则 Windows 只会让任务栏闪烁
    u32.keybd_event(0x12, 0, 0, 0)       # Alt 按下
    u32.SetForegroundWindow(hwnd)
    u32.keybd_event(0x12, 0, 2, 0)       # Alt 放开


def window_size():
    """窗口大小（按 96 DPI 换算，和 pywebview 的单位一致）；最小化 / 最大化时返回 None。"""
    hwnd = main_hwnd()
    if not hwnd:
        return None
    import ctypes
    from ctypes import wintypes
    u32 = ctypes.windll.user32
    if u32.IsIconic(hwnd) or u32.IsZoomed(hwnd):
        return None
    r = wintypes.RECT()
    if not u32.GetWindowRect(hwnd, ctypes.byref(r)):
        return None
    try:
        dpi = u32.GetDpiForWindow(hwnd) or 96
    except Exception:
        dpi = 96
    k = 96.0 / dpi
    return [int((r.right - r.left) * k), int((r.bottom - r.top) * k)]


def start_ui_watchdog():
    """
    每 2 秒问一下控制台窗口“还活着吗”（WM_NULL，最多等 1.5 秒）。
    窗口卡住时，把所有线程当时在做什么写进日志，方便查出卡在哪里。
    """
    if not IS_WIN:
        return

    def run():
        import ctypes
        from ctypes import wintypes
        u32 = ctypes.windll.user32
        u32.SendMessageTimeoutW.argtypes = [wintypes.HWND, ctypes.c_uint, wintypes.WPARAM, wintypes.LPARAM,
                                            ctypes.c_uint, ctypes.c_uint, ctypes.POINTER(ctypes.c_size_t)]
        hung_since = None
        while True:
            time.sleep(2)
            hwnd = main_hwnd()
            if not hwnd:
                continue
            res = ctypes.c_size_t()
            t0 = time.time()
            ok = u32.SendMessageTimeoutW(hwnd, 0, 0, 0, 0x0002, 1500, ctypes.byref(res))  # SMTO_ABORTIFHUNG
            if ok:
                if hung_since:
                    print("[窗口] 界面恢复响应（卡了 %.1f 秒）" % (time.time() - hung_since))
                    hung_since = None
                continue
            if hung_since is None:
                hung_since = t0
                print("[窗口] ⚠ 控制台界面没有响应！以下是各线程当时正在做的事：")
                names = {t.ident: t.name for t in threading.enumerate()}
                for ident, frame in sys._current_frames().items():
                    print("---- 线程 %s (%s)" % (names.get(ident, "?"), ident))
                    for line in traceback.format_stack(frame)[-8:]:
                        print(line.rstrip())

    threading.Thread(target=run, daemon=True, name="ui-watchdog").start()


def webview_available():
    try:
        import webview  # noqa: F401
        return True
    except Exception:
        return False


def run_window(url, cfg, save_config, storage_dir):
    """在主线程打开程序窗口；窗口关闭后才返回。"""
    global WINDOW_OPEN
    import webview
    size = cfg.get("window_size") or [1440, 900]
    win = webview.create_window(
        APP_TITLE, url,
        width=max(960, int(size[0])), height=max(620, int(size[1])), min_size=(960, 620),
        background_color="#0e1013",
    )

    def on_shown():
        # 窗口出现后再用 Win32 设置置顶（不用 pywebview 的 on_top）
        for _ in range(40):
            if main_hwnd():
                break
            time.sleep(0.05)
        if cfg.get("window_on_top"):
            set_on_top(True)

    def on_closing():
        size_now = window_size()
        if size_now:
            cfg["window_size"] = size_now
            try:
                save_config(cfg)
            except Exception:
                pass

    try:
        win.events.shown += on_shown
        win.events.closing += on_closing
    except Exception:
        pass
    os.makedirs(storage_dir, exist_ok=True)
    WINDOW_OPEN = True
    # private_mode=False + storage_path：界面里的小设置在重开程序后还在
    webview.start(private_mode=False, storage_path=storage_dir, debug=bool(cfg.get("debug", False)))
    WINDOW_OPEN = False
