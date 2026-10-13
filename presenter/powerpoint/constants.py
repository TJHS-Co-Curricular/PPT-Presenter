# -*- coding: utf-8 -*-
"""PowerPoint COM 常量，以及日志 / 界面用的中文名称。"""

# PpSlideShowState
SHOW_RUNNING = 1
SHOW_PAUSED = 2
SHOW_BLACK = 3
SHOW_WHITE = 4
SHOW_DONE = 5          # 放完最后一页后的“放映结束，单击退出”黑屏

PLACEHOLDER_BODY = 2   # ppPlaceholderBody：备注页里的正文占位符
SHOW_ALL = 1           # ppShowAll
MSO_TRUE = -1

STATE_NAMES = {SHOW_RUNNING: "放映中", SHOW_PAUSED: "暂停", SHOW_BLACK: "黑屏", SHOW_WHITE: "白屏",
               SHOW_DONE: "已到结尾"}

NAV_ACTIONS = ("next", "prev", "first", "last", "goto", "black", "white", "end", "focus")
NAV_NAMES = {"next": "下一页", "prev": "上一页", "first": "第一页", "last": "最后一页", "goto": "跳到第 %s 页",
             "black": "黑屏切换", "white": "白屏切换", "end": "结束放映", "focus": "切到放映窗口"}

WINDOW_MINIMIZED = 2   # ppWindowMinimized：带窗口打开后最小化编辑窗口
