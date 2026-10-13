# -*- coding: utf-8 -*-
"""
Windows 桌面相关功能。只在 Windows 上生效；其他系统下自动跳过（方便开发测试）。

    common         IS_WIN、窗口标题、提示框、单实例
    hotkeys        全局快捷键（RegisterHotKey）
    window         程序窗口（pywebview）与 Win32 窗口操作、界面卡住检测
    logs           日志：给每行加时间并写入文件
    debug_console  调试窗口（独立进程，实时显示日志）
"""
