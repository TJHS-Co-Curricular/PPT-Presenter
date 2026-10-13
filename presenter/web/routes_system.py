# -*- coding: utf-8 -*-
"""状态（长轮询）、预加载、设置、调试窗口、窗口置顶、快捷键。"""
import os

from flask import Blueprint, jsonify, request

from presenter import paths, preload
from presenter.config import CFG, save_config
from presenter.core import DEBUG, WORKER
from presenter.desktop import window
from presenter.desktop.common import IS_WIN
from presenter.desktop.hotkeys import ACTIONS, parse_hotkey
from presenter.library import refresh_items
from presenter.log import log
from presenter.media import MEDIA
from presenter.shortcuts import HOTKEYS
from presenter.status import wait_for_change
from presenter.web.common import body, err, ok

bp = Blueprint("system", __name__)


@bp.route("/api/status")
def api_status():
    """长轮询：带上次的 v，状态没变就等（最多 20 秒），一变化马上返回。"""
    resp = jsonify(wait_for_change(request.args.get("v")))
    resp.headers["Cache-Control"] = "no-store"
    return resp


@bp.route("/api/preload", methods=["POST"])
def api_preload():
    if body().get("action") == "cancel":
        preload.cancel()
        return ok()
    if (WORKER.snapshot.get("live") or {}).get("key"):
        return err("正在放映中，请先结束放映再预加载（预加载会占用 PowerPoint，放映会变卡）", 409)
    if MEDIA.status()["playing"]:
        return err("正在播放视频，请等播完再预加载", 409)
    preload.start_preload()
    return ok()


def _apply_config(data):
    """把设置界面提交的内容写进 CFG；有问题返回错误文字。"""
    if "folder" in data:
        folder = os.path.expandvars(os.path.expanduser(str(data["folder"]).strip().strip('"')))
        if not os.path.isdir(folder):
            return "文件夹不存在：%s" % folder
        CFG["folder"] = os.path.abspath(folder)
    for key in ("end_previous_show", "auto_advance"):
        if key in data:
            CFG[key] = bool(data[key])
    if isinstance(data.get("hotkeys"), dict):
        hotkeys = dict(CFG["hotkeys"])
        for action, combo in data["hotkeys"].items():
            if action not in ACTIONS:
                continue
            combo = str(combo or "").strip()
            try:
                parse_hotkey(combo)
            except ValueError as e:
                return "%s：%s" % (ACTIONS[action], e)
            hotkeys[action] = combo
        CFG["hotkeys"] = hotkeys
    return None


@bp.route("/api/config", methods=["GET", "POST"])
def api_config():
    if request.method == "POST":
        data = body()
        problem = _apply_config(data)
        if problem:
            return err(problem)
        save_config(CFG)
        refresh_items()
        if "hotkeys" in data:
            HOTKEYS.reload()
    return ok(config=CFG, actions=list(ACTIONS.items()), hotkeys=HOTKEYS.status())


@bp.route("/api/hotkeys/suspend", methods=["POST"])
def api_hotkeys_suspend():
    """设置快捷键时先暂停全局快捷键，否则按下去会直接被触发、录不进去。"""
    HOTKEYS.suspend(bool(body().get("on")))
    return ok(hotkeys=HOTKEYS.status())


@bp.route("/api/debug", methods=["GET", "POST"])
def api_debug():
    """打开 / 关闭调试窗口。"""
    if request.method == "POST":
        on = bool(body().get("on"))
        DEBUG.open() if on else DEBUG.close()
        CFG["debug_console"] = on
        save_config(CFG)
        log("调试", "调试窗口%s" % ("打开" if on else "关闭"))
    return ok(open=DEBUG.is_open(), available=IS_WIN, log=paths.LOG_PATH)


@bp.route("/api/window", methods=["GET", "POST"])
def api_window():
    """控制台窗口置顶（直接用 Win32，不经过 pywebview，避免卡死）。"""
    if request.method == "POST":
        on = bool(body().get("on_top"))
        done = window.set_on_top(on)
        CFG["window_on_top"] = on
        save_config(CFG)
        log("窗口", "置顶%s%s" % ("开" if on else "关", "" if done else "（找不到窗口）"))
    return ok(desktop=window.WINDOW_OPEN,
              on_top=window.is_on_top() if window.WINDOW_OPEN else bool(CFG.get("window_on_top")))
