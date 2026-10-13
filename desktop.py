# -*- coding: utf-8 -*-
"""
桌面端相关：程序窗口（pywebview / Edge WebView2）、全局快捷键（RegisterHotKey）、单实例、提示框。
只在 Windows 上生效；其他系统下这些功能自动跳过（方便开发测试）。
"""
import os
import sys
import json
import threading
import traceback
import time

IS_WIN = sys.platform == "win32"
APP_TITLE = "PPT 控制台"

# 可以设快捷键的动作
ACTIONS = {
    "go": "GO（开始待命节目）",
    "next": "下一页",
    "prev": "上一页",
    "black": "黑屏 / 恢复",
    "white": "白屏 / 恢复",
    "end": "结束放映",
    "show": "把控制台窗口叫到最前面",
}
DEFAULT_HOTKEYS = {
    "go": "F8",
    "next": "F10",
    "prev": "F9",
    "black": "F11",
    "white": "",
    "end": "",
    "show": "Ctrl+Alt+P",
}

# ---------------------------------------------------------------- 提示框 / 单实例
def message_box(text, error=True):
    if IS_WIN:
        import ctypes
        ctypes.windll.user32.MessageBoxW(0, str(text), APP_TITLE, 0x10 if error else 0x40)
    else:
        print(text)


_MUTEX = None


def ensure_single_instance():
    """已经开着一个控制台时，把那个窗口叫到前面，自己退出。"""
    global _MUTEX
    if not IS_WIN:
        return True
    import ctypes
    k32, u32 = ctypes.windll.kernel32, ctypes.windll.user32
    _MUTEX = k32.CreateMutexW(None, False, "Local\\PPT-Presenter-ControlPanel")
    if k32.GetLastError() == 183:  # ERROR_ALREADY_EXISTS
        hwnd = u32.FindWindowW(None, APP_TITLE)
        if hwnd:
            u32.ShowWindow(hwnd, 9)  # SW_RESTORE
            u32.SetForegroundWindow(hwnd)
        else:
            message_box("PPT 控制台已经在运行了。", error=False)
        return False
    return True


# ---------------------------------------------------------------- 全局快捷键
VK_NAMES = {
    "SPACE": 0x20, "ENTER": 0x0D, "RETURN": 0x0D, "ESC": 0x1B, "ESCAPE": 0x1B, "TAB": 0x09,
    "BACKSPACE": 0x08, "PAGEUP": 0x21, "PGUP": 0x21, "PAGEDOWN": 0x22, "PGDN": 0x22,
    "END": 0x23, "HOME": 0x24, "LEFT": 0x25, "UP": 0x26, "RIGHT": 0x27, "DOWN": 0x28,
    "INSERT": 0x2D, "INS": 0x2D, "DELETE": 0x2E, "DEL": 0x2E, "PAUSE": 0x13, "SCROLLLOCK": 0x91,
    "NUMPADMULTIPLY": 0x6A, "NUMPADADD": 0x6B, "NUMPADSUBTRACT": 0x6D, "NUMPADDECIMAL": 0x6E,
    "NUMPADDIVIDE": 0x6F, "NUMPADENTER": 0x0D,
}
for _i in range(1, 25):
    VK_NAMES["F%d" % _i] = 0x6F + _i
for _i in range(10):
    VK_NAMES["NUMPAD%d" % _i] = 0x60 + _i
MOD_NAMES = {"ALT": 0x1, "CTRL": 0x2, "CONTROL": 0x2, "SHIFT": 0x4, "WIN": 0x8}
MOD_NOREPEAT = 0x4000


def parse_hotkey(text):
    """'Ctrl+Alt+F9' -> (mods, vk)；空字符串返回 None；写错抛 ValueError。"""
    text = (text or "").strip()
    if not text:
        return None
    mods, vk = 0, None
    for part in [p.strip().upper() for p in text.split("+") if p.strip()]:
        if part in MOD_NAMES:
            mods |= MOD_NAMES[part]
        elif part in VK_NAMES:
            vk = VK_NAMES[part]
        elif len(part) == 1 and (part.isalpha() or part.isdigit()):
            vk = ord(part)
        else:
            raise ValueError("不认识的按键：%s" % part)
    if vk is None:
        raise ValueError("缺少主按键：%s" % text)
    if vk == 0x7B:  # F12
        raise ValueError("F12 被 Windows 保留给调试器，不能设成全局快捷键")
    return mods, vk


class HotkeyManager(threading.Thread):
    """
    用 Windows 的 RegisterHotKey 注册全局快捷键：不管哪个程序在前台（包括 PowerPoint 放映中），按了都有效。
    注意：注册了的按键会被控制台“独占”，其他程序收不到，所以默认用 F8–F11 这类放映时用不到的键（F12 被 Windows 保留，不能用）。
    """
    WM_HOTKEY = 0x0312
    WM_APP_RELOAD = 0x8001
    WM_APP_SUSPEND = 0x8002

    def __init__(self, get_bindings, on_fire):
        super().__init__(daemon=True)
        self.get_bindings = get_bindings
        self.on_fire = on_fire
        self.thread_id = None
        self.ids = {}            # id -> action
        self.registered = {}     # action -> 组合键
        self.failed = {}         # action -> 原因
        self.suspended = False
        self.ready = threading.Event()
        self.reloaded = threading.Event()

    # 其他线程调用
    def reload(self, wait=2.0):
        if not IS_WIN or self.thread_id is None:
            return
        import ctypes
        self.reloaded.clear()
        ctypes.windll.user32.PostThreadMessageW(self.thread_id, self.WM_APP_RELOAD, 0, 0)
        self.reloaded.wait(wait)

    def suspend(self, on):
        if not IS_WIN or self.thread_id is None:
            self.suspended = bool(on)
            return
        import ctypes
        self.reloaded.clear()
        ctypes.windll.user32.PostThreadMessageW(self.thread_id, self.WM_APP_SUSPEND, 1 if on else 0, 0)
        self.reloaded.wait(2.0)

    def status(self):
        return {"registered": dict(self.registered), "failed": dict(self.failed),
                "suspended": self.suspended, "available": IS_WIN}

    # 快捷键线程内部
    def _unregister_all(self, user32):
        for hid in list(self.ids):
            user32.UnregisterHotKey(None, hid)
        self.ids.clear()
        self.registered.clear()

    def _register_all(self, user32):
        self._unregister_all(user32)
        self.failed.clear()
        bindings = self.get_bindings() or {}
        used = {}
        for n, action in enumerate(ACTIONS, start=1):
            combo = (bindings.get(action) or "").strip()
            if not combo:
                continue
            try:
                parsed = parse_hotkey(combo)
            except ValueError as e:
                self.failed[action] = str(e)
                continue
            if parsed in used:
                self.failed[action] = "和「%s」重复" % ACTIONS[used[parsed]]
                continue
            mods, vk = parsed
            if user32.RegisterHotKey(None, n, mods | MOD_NOREPEAT, vk):
                self.ids[n] = action
                self.registered[action] = combo
                used[parsed] = action
            else:
                self.failed[action] = "被其他程序占用了"
        if self.failed:
            print("快捷键注册失败：", self.failed)

    def run(self):
        if not IS_WIN:
            self.ready.set()
            return
        import ctypes
        from ctypes import wintypes
        user32, k32 = ctypes.windll.user32, ctypes.windll.kernel32
        self.thread_id = k32.GetCurrentThreadId()
        msg = wintypes.MSG()
        user32.PeekMessageW(ctypes.byref(msg), None, 0, 0, 0)  # 先建立消息队列
        self._register_all(user32)
        self.ready.set()
        while user32.GetMessageW(ctypes.byref(msg), None, 0, 0) > 0:
            if msg.message == self.WM_HOTKEY:
                action = self.ids.get(msg.wParam)
                if action:
                    threading.Thread(target=self._fire, args=(action,), daemon=True).start()
            elif msg.message == self.WM_APP_RELOAD:
                if not self.suspended:
                    self._register_all(user32)
                self.reloaded.set()
            elif msg.message == self.WM_APP_SUSPEND:
                self.suspended = bool(msg.wParam)
                if self.suspended:
                    self._unregister_all(user32)
                else:
                    self._register_all(user32)
                self.reloaded.set()

    def _fire(self, action):
        try:
            self.on_fire(action)
        except Exception:
            traceback.print_exc()


# ---------------------------------------------------------------- 程序窗口
# 重要：pywebview 的窗口对象只能在它自己的界面线程里操作。从其他线程（快捷键、网页请求）去改窗口，
# 会互相等待而卡死（之前“置顶”一按就失去响应就是这个原因）。所以窗口的置顶、叫到前面、记大小，
# 一律直接用 Windows API 对窗口句柄操作，不经过 pywebview。
WINDOW_OPEN = False
_HWND = None


def main_hwnd():
    """找到本程序的控制台窗口句柄（只找自己进程里、标题完全相同的窗口）。"""
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
    def cb(hwnd, _):
        pid = wintypes.DWORD()
        u32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
        if pid.value == me and u32.IsWindowVisible(hwnd):
            # 用 InternalGetWindowText：直接读标题，不向窗口发消息（GetWindowText 对自己进程的窗口
            # 会发 WM_GETTEXT 并等界面线程回应，界面线程忙时就会卡住）
            buf = ctypes.create_unicode_buffer(256)
            u32.InternalGetWindowText(hwnd, buf, 256)
            if buf.value == APP_TITLE:
                found.append(hwnd)
                return False
        return True

    u32.EnumWindows(cb, 0)
    _HWND = found[0] if found else None
    return _HWND


def set_on_top(on):
    """窗口置顶 / 取消置顶。可以在任何线程调用。"""
    hwnd = main_hwnd()
    if not hwnd:
        return False
    import ctypes
    from ctypes import wintypes
    u32 = ctypes.windll.user32
    u32.SetWindowPos.argtypes = [wintypes.HWND, wintypes.HWND, ctypes.c_int, ctypes.c_int,
                                 ctypes.c_int, ctypes.c_int, ctypes.c_uint]
    HWND_TOPMOST, HWND_NOTOPMOST = -1, -2
    # SWP_ASYNCWINDOWPOS：只把请求丢给窗口自己的线程去做，调用方马上返回，绝不互相等待
    flags = 0x0001 | 0x0002 | 0x0010 | 0x4000  # NOSIZE | NOMOVE | NOACTIVATE | ASYNCWINDOWPOS
    return bool(u32.SetWindowPos(hwnd, HWND_TOPMOST if on else HWND_NOTOPMOST, 0, 0, 0, 0, flags))


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
        u32.ShowWindowAsync(hwnd, 9)     # SW_RESTORE（Async：不等界面线程）
    # 前台窗口切换需要一点小技巧，否则 Windows 只会让任务栏闪烁
    u32.keybd_event(0x12, 0, 0, 0)       # Alt 按下
    u32.SetForegroundWindow(hwnd)
    u32.keybd_event(0x12, 0, 2, 0)       # Alt 放开


def window_size():
    hwnd = main_hwnd()
    if not hwnd:
        return None
    import ctypes
    from ctypes import wintypes
    u32 = ctypes.windll.user32
    if u32.IsIconic(hwnd) or u32.IsZoomed(hwnd):
        return None                      # 最小化 / 最大化时不记
    r = wintypes.RECT()
    if not u32.GetWindowRect(hwnd, ctypes.byref(r)):
        return None
    # GetWindowRect 是实际像素；pywebview 的 width/height 是按 DPI 缩放前的大小
    try:
        dpi = u32.GetDpiForWindow(hwnd) or 96
    except Exception:
        dpi = 96
    k = 96.0 / dpi
    return [int((r.right - r.left) * k), int((r.bottom - r.top) * k)]


def start_ui_watchdog():
    """
    每 2 秒问一下控制台窗口“还活着吗”（WM_NULL，最多等 1.5 秒）。
    如果窗口卡住了，就把所有线程当时在做什么写进日志，方便查出卡在哪里。
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
    """在主线程打开程序窗口，窗口关闭后才返回。"""
    global WINDOW_OPEN
    import webview
    size = cfg.get("window_size") or [1440, 900]
    win = webview.create_window(
        APP_TITLE, url,
        width=max(960, int(size[0])), height=max(620, int(size[1])), min_size=(960, 620),
        background_color="#0e1013",
    )

    def on_shown():
        # 窗口出现后再用 Windows API 设置置顶（不用 pywebview 的 on_top）
        for _ in range(40):
            if main_hwnd():
                break
            time.sleep(0.05)
        if cfg.get("window_on_top"):
            set_on_top(True)

    def on_closing():
        sz = window_size()
        if sz:
            cfg["window_size"] = sz
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
    # private_mode=False + storage_path：让“待命”等设置在重开程序后还在
    webview.start(private_mode=False, storage_path=storage_dir, debug=bool(cfg.get("debug", False)))
    WINDOW_OPEN = False


# ---------------------------------------------------------------- 日志 / 调试窗口
class Tee:
    """把 print / 报错同时写进日志文件（和开发时的命令行），每行加上时间。"""

    def __init__(self, logfile=None, base=None):
        self.lock = threading.Lock()
        self.logfile = logfile
        self.base = base
        self._bol = True

    def write(self, s):
        if not s:
            return 0
        with self.lock:
            out = []
            for part in s.splitlines(keepends=True):
                if self._bol:
                    t = time.time()
                    out.append(time.strftime("%H:%M:%S", time.localtime(t)) + ".%03d  " % (int(t * 1000) % 1000))
                out.append(part)
                self._bol = part.endswith("\n")
            text = "".join(out)
            for stream in (self.logfile, self.base):
                if stream is None:
                    continue
                try:
                    stream.write(text)
                    stream.flush()
                except Exception:
                    pass
        return len(s)

    def flush(self):
        pass

    def isatty(self):
        return False


def setup_logging(log_path, frozen):
    """日志超过 5 MB 时改名为 .old 再重新开始。"""
    try:
        if os.path.exists(log_path) and os.path.getsize(log_path) > 5 * 1024 * 1024:
            os.replace(log_path, log_path + ".old")
    except OSError:
        pass
    try:
        f = open(log_path, "a", encoding="utf-8", buffering=1)
    except OSError:
        f = None
    base = None if (frozen or sys.stdout is None) else sys.stdout
    tee = Tee(f, base)
    sys.stdout = sys.stderr = tee
    print("===== %s 启动 =====" % time.strftime("%Y-%m-%d"))
    return tee


class DebugConsole:
    """
    调试窗口 = 另开一个独立的命令行窗口，实时显示日志。
    它是单独的进程：随时关掉都不会影响控制台和放映。
    """

    def __init__(self, log_path, frozen, script_path):
        self.log_path = log_path
        self.frozen = frozen
        self.script_path = script_path
        self.proc = None

    def is_open(self):
        return self.proc is not None and self.proc.poll() is None

    def open(self):
        if self.is_open() or not IS_WIN:
            return self.is_open()
        import subprocess
        if self.frozen:
            cmd = [sys.executable, "--debug-console", self.log_path, str(os.getpid())]
        else:
            cmd = [sys.executable, self.script_path, "--debug-console", self.log_path, str(os.getpid())]
        self.proc = subprocess.Popen(cmd, creationflags=0x00000010)  # CREATE_NEW_CONSOLE
        return True

    def close(self):
        if self.is_open():
            try:
                self.proc.terminate()
            except Exception:
                pass
        self.proc = None
        return False


def _parent_alive(pid):
    if not IS_WIN:
        try:
            os.kill(pid, 0)
            return True
        except OSError:
            return False
    import ctypes
    k32 = ctypes.windll.kernel32
    h = k32.OpenProcess(0x1000, False, pid)  # PROCESS_QUERY_LIMITED_INFORMATION
    if not h:
        return False
    code = ctypes.c_ulong()
    ok = k32.GetExitCodeProcess(h, ctypes.byref(code))
    k32.CloseHandle(h)
    return bool(ok) and code.value == 259  # STILL_ACTIVE


def run_debug_console(log_path, parent_pid):
    """--debug-console 模式：在自己的命令行窗口里实时跟着日志（类似 tail -f），控制台关了就自动结束。"""
    out = sys.stdout
    if IS_WIN:
        import ctypes
        k32 = ctypes.windll.kernel32
        if not k32.GetConsoleWindow():
            k32.AllocConsole()
        k32.SetConsoleTitleW(APP_TITLE + " - Debug")
        k32.SetConsoleOutputCP(65001)
        h = k32.GetStdHandle(-11)
        mode = ctypes.c_ulong()
        if k32.GetConsoleMode(h, ctypes.byref(mode)):
            k32.SetConsoleMode(h, mode.value | 0x0004)  # 打开 ANSI 颜色
        out = open("CONOUT$", "w", encoding="utf-8", buffering=1)
    RED, YEL, GRN, DIM, RST = "\x1b[91m", "\x1b[93m", "\x1b[92m", "\x1b[90m", "\x1b[0m"

    def color(line):
        if any(k in line for k in ("Traceback", "Error", "错误", "失败", "出错", "  File ")):
            return RED + line + RST
        if any(k in line for k in ("忙", "重试", "重新连接", "被占用")):
            return YEL + line + RST
        if "[放映]" in line or "[快捷键]" in line:
            return GRN + line + RST
        if "[缩略图]" in line:
            return DIM + line + RST
        return line

    out.write(DIM + "PPT 控制台调试窗口 —— 实时显示 %s\n这个窗口可以随时关掉，不会影响控制台和放映。\n\n" % log_path + RST)
    pos = 0
    try:
        size = os.path.getsize(log_path)
        pos = max(0, size - 20000)   # 先显示最近的一段
    except OSError:
        pass
    start = pos
    first = True
    buf = b""
    while _parent_alive(parent_pid):
        try:
            size = os.path.getsize(log_path)
            if size < pos:            # 日志被轮换了
                pos = 0
            if size > pos:
                with open(log_path, "rb") as f:
                    f.seek(pos)
                    buf += f.read()
                    pos = f.tell()
                if first and start > 0 and b"\n" in buf:
                    buf = buf.split(b"\n", 1)[1]   # 从中间开始读时，丢掉不完整的第一行
                first = False
                *lines, buf = buf.split(b"\n")
                for ln in lines:
                    out.write(color(ln.decode("utf-8", "replace").rstrip("\r")) + "\n")
                out.flush()
        except OSError:
            pass
        time.sleep(0.05)
