# -*- coding: utf-8 -*-
"""
日志：接管 print / 报错，每行加上毫秒时间，写进 ppt-presenter.log（开发时也同时输出到命令行）。

打包成无控制台的 exe 时 sys.stdout 是 None，不接管的话输出会直接丢失甚至出错，
所以入口脚本一开始就要调用 setup_logging()。
"""
import os
import sys
import threading
import time

MAX_LOG_BYTES = 5 * 1024 * 1024


class Tee:
    """同时写进日志文件和（开发时的）命令行。"""

    def __init__(self, logfile=None, base=None):
        self.lock = threading.Lock()
        self.logfile = logfile
        self.base = base
        self._bol = True     # 下一个字符是不是一行的开头

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
    """日志超过 5 MB 时先改名为 .old 再重新开始。"""
    try:
        if os.path.exists(log_path) and os.path.getsize(log_path) > MAX_LOG_BYTES:
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
