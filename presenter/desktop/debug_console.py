# -*- coding: utf-8 -*-
"""
调试窗口：另开一个独立的命令行窗口，实时显示日志（类似 tail -f），按标签上色。

它是单独的进程（``<exe> --debug-console <日志> <父进程 pid>``），所以随时关掉都不会影响控制台和放映；
控制台结束时它会自己退出。如果把命令行窗口直接挂在控制台进程上，关掉它会连控制台一起结束。
"""
import os
import subprocess
import sys
import time

from presenter.desktop.common import APP_TITLE, IS_WIN

CREATE_NEW_CONSOLE = 0x00000010
STILL_ACTIVE = 259


class DebugConsole:
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
        if self.frozen:
            cmd = [sys.executable, "--debug-console", self.log_path, str(os.getpid())]
        else:
            cmd = [sys.executable, self.script_path, "--debug-console", self.log_path, str(os.getpid())]
        self.proc = subprocess.Popen(cmd, creationflags=CREATE_NEW_CONSOLE)
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
    return bool(ok) and code.value == STILL_ACTIVE


_RED, _YEL, _GRN, _DIM, _RST = "\x1b[91m", "\x1b[93m", "\x1b[92m", "\x1b[90m", "\x1b[0m"


def _color(line):
    if any(k in line for k in ("Traceback", "Error", "[错误]", "错误", "失败", "出错", "  File ")):
        return _RED + line + _RST
    if any(k in line for k in ("忙", "重试", "重新连接", "被占用")):
        return _YEL + line + _RST
    if "[放映]" in line or "[快捷键]" in line:
        return _GRN + line + _RST
    if "[缩略图]" in line:
        return _DIM + line + _RST
    return line


def _open_console_output():
    """在 Windows 上给这个进程开一个命令行窗口，并打开 UTF-8 和 ANSI 颜色。"""
    if not IS_WIN:
        return sys.stdout
    import ctypes
    k32 = ctypes.windll.kernel32
    if not k32.GetConsoleWindow():
        k32.AllocConsole()
    k32.SetConsoleTitleW(APP_TITLE + " - Debug")
    k32.SetConsoleOutputCP(65001)
    handle = k32.GetStdHandle(-11)
    mode = ctypes.c_ulong()
    if k32.GetConsoleMode(handle, ctypes.byref(mode)):
        k32.SetConsoleMode(handle, mode.value | 0x0004)  # ENABLE_VIRTUAL_TERMINAL_PROCESSING
    return open("CONOUT$", "w", encoding="utf-8", buffering=1)


def run_debug_console(log_path, parent_pid, tail_bytes=20000):
    """--debug-console 模式的主循环：先显示最近一段日志，再每 50ms 跟一次新内容。"""
    out = _open_console_output()
    out.write(_DIM + "PPT 控制台调试窗口 —— 实时显示 %s\n这个窗口可以随时关掉，不会影响控制台和放映。\n\n"
              % log_path + _RST)
    try:
        pos = max(0, os.path.getsize(log_path) - tail_bytes)
    except OSError:
        pos = 0
    start, first, buf = pos, True, b""
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
                    out.write(_color(ln.decode("utf-8", "replace").rstrip("\r")) + "\n")
                out.flush()
        except OSError:
            pass
        time.sleep(0.05)
