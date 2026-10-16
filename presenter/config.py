# -*- coding: utf-8 -*-
"""
config.json：读取、补全默认值、迁移旧版设置、保存。

全程序共用一个字典 ``CFG``（导入时就读好）；改了之后调用 ``save_config(CFG)``。
"""
import json
import os

from presenter import paths
from presenter.desktop.hotkeys import ACTIONS, DEFAULT_HOTKEYS
from presenter.formats import PPT_EXT, VIDEO_EXT, is_temp_file
from presenter.log import log

DEFAULT_CONFIG = {
    "folder": "",               # PPT 所在文件夹（留空 = 程序所在文件夹，或其中的 Example 子文件夹）
    "port": 0,                  # 0 = 自动找空闲端口（只在本机 127.0.0.1，外部连不进来）
    "end_previous_show": True,  # 切换到另一份 PPT 时，自动结束上一份的放映
    "auto_advance": True,       # PPT 放到结尾（“放映结束”黑屏）时，自动开始待命的下一个节目
    "standby": "",              # 当前待命的节目（程序自动记住）
    "preload_on_start": False,  # 启动时自动预加载
    "thumb_width": 480,         # 缩略图宽度（像素）
    "prefetch_thumbs": True,    # PowerPoint 空闲且没在放映时，自动预先生成缩略图
    "work_during_show": False,  # 放映中也生成缩略图 / 读备注（会打断动画，默认关；请改用预加载）
    "open_with_window": True,   # 带窗口打开 PPT（打开后最小化）：内嵌字体、平滑切换才正常
    "debug_console": False,     # 启动时自动打开调试窗口（也可以用 --debug 参数或界面上的 🐞 按钮）
    "window_on_top": False,     # 控制台窗口置顶
    "window_size": [1440, 900],
    "hotkeys": dict(DEFAULT_HOTKEYS),  # 全局快捷键（PowerPoint 在前台时也有效）
}

#: 旧版本用过、现在已经不需要的设置
_OBSOLETE_KEYS = ("host", "pin", "open_browser")


def _has_media(folder):
    try:
        return any(os.path.splitext(n)[1].lower() in PPT_EXT | VIDEO_EXT
                   for n in os.listdir(folder) if not is_temp_file(n))
    except OSError:
        return False


def default_folder():
    example = os.path.join(paths.BASE, "Example")
    if not _has_media(paths.BASE) and os.path.isdir(example):
        return example
    return paths.BASE


#: v2.1 起“黑屏 / 白屏”按钮换成“礼成 / 大会背景”，快捷键跟着搬过去
_RENAMED_ACTIONS = {"black": "finale", "white": "background"}


def _migrate_hotkeys(saved):
    saved = dict(saved or {})
    for old, new in _RENAMED_ACTIONS.items():
        if old in saved and new not in saved and saved[old]:
            saved[new] = saved[old]
    hotkeys = dict(DEFAULT_HOTKEYS)
    hotkeys.update({k: v for k, v in saved.items() if k in ACTIONS})
    # 旧版默认把 GO 设在 F12，但 F12 被 Windows 保留，不能当全局快捷键 → 改成未设定
    for action, combo in hotkeys.items():
        if str(combo).strip().upper() == "F12":
            hotkeys[action] = DEFAULT_HOTKEYS.get(action, "")
    return hotkeys


def save_config(cfg):
    with open(paths.CONFIG_PATH, "w", encoding="utf-8") as f:
        json.dump(cfg, f, ensure_ascii=False, indent=2)


def load_config():
    cfg = dict(DEFAULT_CONFIG)
    exists = os.path.exists(paths.CONFIG_PATH)
    if exists:
        try:
            with open(paths.CONFIG_PATH, "r", encoding="utf-8") as f:
                cfg.update(json.load(f))
        except Exception as e:
            log("提示", "读取 config.json 失败，改用默认设置：%s" % e)
    if not cfg.get("folder"):
        cfg["folder"] = default_folder()
    for k in _OBSOLETE_KEYS:
        cfg.pop(k, None)
    cfg["hotkeys"] = _migrate_hotkeys(cfg.get("hotkeys"))
    if not exists:
        save_config(cfg)
    return cfg


CFG = load_config()
