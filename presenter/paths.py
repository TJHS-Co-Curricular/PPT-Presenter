# -*- coding: utf-8 -*-
"""程序用到的所有路径。打包成 exe 和直接用 Python 运行时，位置不同。"""
import os
import sys

#: 是否是 PyInstaller 打包出来的 exe
FROZEN = bool(getattr(sys, "frozen", False))

#: 程序所在文件夹：放 config.json、.cache、日志（exe 或 server.py 旁边）
BASE = os.path.dirname(os.path.abspath(sys.executable)) if FROZEN else \
    os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

#: 只读资源（界面、CHANGELOG）：打包后在 PyInstaller 的临时解压目录里
RES = getattr(sys, "_MEIPASS", BASE)

#: 直接用 Python 运行时的入口脚本（调试窗口要用它再启动一个进程）
ENTRY_SCRIPT = os.path.join(BASE, "server.py")

CONFIG_PATH = os.path.join(BASE, "config.json")
LOG_PATH = os.path.join(BASE, "ppt-presenter.log")

CACHE_ROOT = os.path.join(BASE, ".cache")
THUMBS_DIR = os.path.join(CACHE_ROOT, "thumbs")     # 缩略图 <id>_<mtime>/<页>.png
NOTES_DIR = os.path.join(CACHE_ROOT, "notes")       # 演讲者备注 <id>_<mtime>.json
WEBVIEW_DIR = os.path.join(CACHE_ROOT, "webview")   # 界面窗口（WebView2）的数据

STATIC_DIR = os.path.join(RES, "static")
