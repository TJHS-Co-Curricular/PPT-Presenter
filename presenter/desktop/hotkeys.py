# -*- coding: utf-8 -*-
"""
全局快捷键：用 Windows 的 RegisterHotKey 注册，不管哪个程序在前台（包括 PowerPoint 放映中）按了都有效。

注意：注册了的按键会被控制台“独占”，其他程序收不到，所以默认用 F8–F11 这类放映时用不到的键。
F12 被 Windows 保留给调试器，任何程序都不能注册。
"""
import threading
import traceback

from presenter.desktop.common import IS_WIN

#: 可以设快捷键的动作（顺序 = 设置界面里的顺序）
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

VK_NAMES = {
    "SPACE": 0x20, "ENTER": 0x0D, "RETURN": 0x0D, "ESC": 0x1B, "ESCAPE": 0x1B, "TAB": 0x09,
    "BACKSPACE": 0x08, "PAGEUP": 0x21, "PGUP": 0x21, "PAGEDOWN": 0x22, "PGDN": 0x22,
    "END": 0x23, "HOME": 0x24, "LEFT": 0x25, "UP": 0x26, "RIGHT": 0x27, "DOWN": 0x28,
    "INSERT": 0x2D, "INS": 0x2D, "DELETE": 0x2E, "DEL": 0x2E, "PAUSE": 0x13, "SCROLLLOCK": 0x91,
    "NUMPADMULTIPLY": 0x6A, "NUMPADADD": 0x6B, "NUMPADSUBTRACT": 0x6D, "NUMPADDECIMAL": 0x6E,
    "NUMPADDIVIDE": 0x6F, "NUMPADENTER": 0x0D,
}
VK_NAMES.update({"F%d" % i: 0x6F + i for i in range(1, 25)})
VK_NAMES.update({"NUMPAD%d" % i: 0x60 + i for i in range(10)})

MOD_NAMES = {"ALT": 0x1, "CTRL": 0x2, "CONTROL": 0x2, "SHIFT": 0x4, "WIN": 0x8}
MOD_NOREPEAT = 0x4000
VK_F12 = 0x7B


def parse_hotkey(text):
    """'Ctrl+Alt+F9' → (mods, vk)；空字符串返回 None；写错抛 ValueError。"""
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
    if vk == VK_F12:
        raise ValueError("F12 被 Windows 保留给调试器，不能设成全局快捷键")
    return mods, vk


class HotkeyManager(threading.Thread):
    """
    快捷键线程：注册 / 注销都必须在同一个线程里做，所以其他线程通过 PostThreadMessage 通知它。

    get_bindings()  返回 {动作: 组合键}
    on_fire(action) 按下时调用（在新线程里，不会挡住消息循环）
    """
    WM_HOTKEY = 0x0312
    WM_APP_RELOAD = 0x8001
    WM_APP_SUSPEND = 0x8002

    def __init__(self, get_bindings, on_fire):
        super().__init__(daemon=True, name="hotkeys")
        self.get_bindings = get_bindings
        self.on_fire = on_fire
        self.thread_id = None
        self.ids = {}            # 注册编号 → 动作
        self.registered = {}     # 动作 → 组合键
        self.failed = {}         # 动作 → 原因
        self.suspended = False
        self.ready = threading.Event()
        self.reloaded = threading.Event()

    # ---- 其他线程调用
    def reload(self, wait=2.0):
        """设置改了之后重新注册。"""
        self._post(self.WM_APP_RELOAD, 0, wait)

    def suspend(self, on):
        """录制快捷键时暂停（否则按下去直接被触发，录不进去）。"""
        if not IS_WIN or self.thread_id is None:
            self.suspended = bool(on)
            return
        self._post(self.WM_APP_SUSPEND, 1 if on else 0, 2.0)

    def status(self):
        return {"registered": dict(self.registered), "failed": dict(self.failed),
                "suspended": self.suspended, "available": IS_WIN}

    def _post(self, msg, wparam, wait):
        if not IS_WIN or self.thread_id is None:
            return
        import ctypes
        self.reloaded.clear()
        ctypes.windll.user32.PostThreadMessageW(self.thread_id, msg, wparam, 0)
        self.reloaded.wait(wait)

    # ---- 快捷键线程内部
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
            print("[快捷键] 注册失败：%s" % self.failed)

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
