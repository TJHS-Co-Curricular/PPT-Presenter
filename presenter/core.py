# -*- coding: utf-8 -*-
"""
全程序共用的实例：

    CTRL    PowerPoint 控制器（Windows 上是真的，其他系统是演示模式）
    WORKER  PowerPoint 线程（app.setup() 里才启动）
    DEBUG   调试窗口
"""
import sys

from presenter import paths
from presenter.desktop.debug_console import DebugConsole
from presenter.powerpoint.controller import PPTController
from presenter.powerpoint.mock import MockController
from presenter.powerpoint.worker import ComWorker

MOCK = sys.platform != "win32"
CTRL = MockController() if MOCK else PPTController()
WORKER = ComWorker(CTRL)
DEBUG = DebugConsole(paths.LOG_PATH, paths.FROZEN, paths.ENTRY_SCRIPT)
