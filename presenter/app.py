# -*- coding: utf-8 -*-
"""
启动流程。

    setup()  扫描节目单、挂上后台任务和状态钩子、启动 PowerPoint 线程（测试也用它）
    run()    完整启动：单实例 → setup → 网页服务 → 快捷键 / 调试窗口 / 卡住检测 → 程序窗口
"""
import logging
import os
import socket
import sys
import threading
import time
import webbrowser

from presenter import background, paths, show
from presenter.config import CFG, save_config
from presenter.core import CTRL, DEBUG, MOCK, WORKER
from presenter.desktop import window
from presenter.desktop.common import ensure_single_instance, message_box
from presenter.desktop.hotkeys import ACTIONS
from presenter.library import refresh_items
from presenter.preload import start_preload
from presenter.shortcuts import HOTKEYS
from presenter.version import VERSION
from presenter.web import create_app

app = create_app()
_setup_done = False


def setup():
    """可以重复调用，只会执行一次。"""
    global _setup_done
    if _setup_done:
        return
    _setup_done = True
    CTRL.with_window = bool(CFG.get("open_with_window", True))
    refresh_items()
    background.register()
    show.register()
    WORKER.start()


def pick_port(want):
    """想要的端口能用就用，否则让系统挑一个空闲端口（只在 127.0.0.1）。"""
    want = int(want or 0)
    if want:
        try:
            with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
                s.bind(("127.0.0.1", want))
            return want
        except OSError:
            print("端口 %d 被占用，改用自动端口" % want)
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        s.bind(("127.0.0.1", 0))
        return s.getsockname()[1]


def serve_forever(port):
    try:
        from waitress import serve
    except ImportError:
        serve = None
    if serve:
        logging.getLogger("waitress").setLevel(logging.ERROR)
        serve(app, host="127.0.0.1", port=port, threads=8, ident=None, channel_timeout=60)
    else:
        app.run(host="127.0.0.1", port=port, threaded=True, debug=False, use_reloader=False)


def wait_until_up(port, timeout=10):
    end = time.time() + timeout
    while time.time() < end:
        try:
            with socket.create_connection(("127.0.0.1", port), timeout=0.3):
                return True
        except OSError:
            time.sleep(0.05)
    return False


def _start_desktop_services():
    HOTKEYS.start()
    if CFG.get("debug_console") or "--debug" in sys.argv:
        DEBUG.open()
    window.start_ui_watchdog()
    HOTKEYS.ready.wait(2)
    print("全局快捷键：%s" % ("、".join("%s=%s" % (ACTIONS[a], k) for a, k in HOTKEYS.registered.items()) or "无"))


def _shutdown():
    """窗口关了 = 程序结束（PowerPoint 不会被关掉；调试窗口发现控制台结束后会自己关掉）。"""
    print("控制台窗口已关闭，程序结束")
    try:
        WORKER.call(CTRL.mark_saved_all, priority=0, timeout=5)   # 之后关 PowerPoint 也不会问要不要保存
        CTRL.save_notes()
    except Exception:
        pass
    os._exit(0)


def run():
    if not ensure_single_instance():
        return
    setup()
    port = pick_port(CFG.get("port", 0))
    url = "http://127.0.0.1:%d/" % port
    print("PPT 控制台 v%s%s | 文件夹：%s | %s" % (VERSION, "（演示模式）" if MOCK else "", CFG["folder"], url))

    threading.Thread(target=serve_forever, args=(port,), daemon=True, name="web").start()
    if not wait_until_up(port):
        message_box("控制台启动失败（网页服务没有起来），详情见 ppt-presenter.log")
        return
    _start_desktop_services()
    if CFG.get("preload_on_start"):
        start_preload()

    if window.webview_available() and not MOCK:
        window.run_window(url, CFG, save_config, paths.WEBVIEW_DIR)
        _shutdown()
    else:
        # 开发 / 测试环境：没有 pywebview 时退回浏览器
        print("未安装 pywebview，改用浏览器打开：", url)
        if not MOCK:
            webbrowser.open(url)
        try:
            while True:
                time.sleep(3600)
        except KeyboardInterrupt:
            pass
