# -*- coding: utf-8 -*-
"""全局快捷键按下后做什么（注册由 desktop.hotkeys.HotkeyManager 负责）。"""
from presenter import show
from presenter.config import CFG
from presenter.desktop import window
from presenter.desktop.hotkeys import ACTIONS, HotkeyManager
from presenter.log import log
from presenter.preload import PRELOAD


def on_hotkey(action):
    log("快捷键", ACTIONS.get(action, action))
    if action == "show":
        window.bring_to_front()
        return
    if PRELOAD["running"]:
        log("快捷键", "正在预加载，已忽略")
        return
    try:
        if action == "go":
            show.go("快捷键：")
        elif action in show.BASE_SLIDES:
            show.show_base(action)
        else:
            show.nav(action)      # 翻页、黑屏直接在后台执行，最快，也不依赖界面
    except Exception as e:
        log("快捷键", "%s：%s" % (ACTIONS.get(action, action), e))


HOTKEYS = HotkeyManager(lambda: CFG.get("hotkeys"), on_hotkey)
