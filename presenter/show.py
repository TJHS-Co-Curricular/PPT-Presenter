# -*- coding: utf-8 -*-
"""
放映流程：开始节目、GO、待命、翻页、关闭，以及“播完自动接下一个”。

编号为 0 的 PPT 是“大会背景”（library.base_item）：
- 大会背景 = 它的第 1 页，礼成 = 它的第 2 页（show_base），切过去不会改变待命；
- 它永远不能被关闭（close_item / close_current 都会拒绝）。

“完全关闭”（close_current）：一个环节结束后，把刚播完的节目从 PowerPoint 里关掉、释放内存；
如果它正在投影，会先切到大会背景再关，投影上不会露出桌面。

待命（CFG["standby"]）由后台管理并记在 config.json：
- 开始任何一个节目（GO、▶、双击），它后面的那个节目自动变成待命；
- 放映到结尾黑屏时自动开始待命的节目（控制台按钮、快捷键、翻页笔直接按在 PowerPoint 上都有效）。
"""
import os
import threading
import time

from presenter.config import CFG, save_config
from presenter.core import CTRL, WORKER
from presenter.library import ITEMS, ITEMS_LOCK, base_item, find_by_key, get_item, is_base, next_item_id, norm
from presenter.log import log
from presenter.media import MEDIA
from presenter.powerpoint.constants import NAV_ACTIONS, SHOW_DONE, STATE_NAMES
from presenter.preload import ensure_not_preloading

_START_LOCK = threading.Lock()
_ADVANCED = {"key": None, "at": 0.0}
_LAST = {"id": None}          # 最近开始的节目（不含大会背景），给“完全关闭”用

#: 动作 → (0 号 PPT 的第几页, 名称)
BASE_SLIDES = {"background": (1, "大会背景"), "finale": (2, "礼成")}
AUTO_ADVANCE_DEBOUNCE = 10    # 秒：同一份 PPT 不会重复触发自动接续


class Busy(RuntimeError):
    """正在切换中（连按好几次只执行一次）。"""


# ------------------------------------------------------------ 待命 / 开始 / GO
def set_standby(iid, why=""):
    iid = iid or ""
    with ITEMS_LOCK:
        if iid and iid not in ITEMS:
            iid = ""
        name = ITEMS[iid]["name"] if iid else "（节目单结束）"
    if CFG.get("standby") != iid:
        CFG["standby"] = iid
        save_config(CFG)
        log("待命", name + ("（%s）" % why if why else ""))
    WORKER.notify()


def start_item(iid, slide=1, why=""):
    """开始放映 / 播放一个节目，并把它后面的节目设为待命。返回 "ppt" / "vlc" / "default"。"""
    ensure_not_preloading()
    item = get_item(iid)
    if item["type"] == "ppt":
        if not _START_LOCK.acquire(blocking=False):
            raise Busy("正在切换中，请稍等")
        try:
            WORKER.call(CTRL.start, item["path"], slide, bool(CFG.get("end_previous_show", True)),
                        priority=0, timeout=600, label=("%s开始放映 " % why) + item["name"])
        finally:
            _START_LOCK.release()
        how = "ppt"
    else:
        log("视频", "%s播放 %s" % (why, item["name"]))
        how = MEDIA.play(item["path"])
    if not is_base(item):
        _LAST["id"] = iid
    set_standby(next_item_id(iid))
    WORKER.notify()
    return how


def go(why=""):
    standby = CFG.get("standby")
    if not standby:
        raise RuntimeError("没有待命的节目（节目单已经播完）")
    return start_item(standby, 1, why)


# ------------------------------------------------------------ 翻页 / 关闭
def nav(action, slide=None):
    if action not in NAV_ACTIONS:
        raise ValueError("未知操作：%s" % action)
    ensure_not_preloading()
    return WORKER.call(CTRL.nav, action, slide, priority=0, timeout=45)


def _start_quiet(item, slide, label):
    """开始放映但不改变待命（大会背景 / 礼成用）。"""
    if not _START_LOCK.acquire(blocking=False):
        raise Busy("正在切换中，请稍等")
    try:
        WORKER.call(CTRL.start, item["path"], slide, bool(CFG.get("end_previous_show", True)),
                    priority=0, timeout=600, label=label)
    finally:
        _START_LOCK.release()
    WORKER.notify()


def show_base(action):
    """大会背景 / 礼成：放映 0 号 PPT 的第 1 / 2 页（待命不变）。"""
    ensure_not_preloading()
    slide, name = BASE_SLIDES[action]
    base = base_item()
    if not base:
        raise RuntimeError("节目单里没有编号为 0 的 PPT（大会背景）")
    _start_quiet(base, slide, "%s（%s 第 %d 页）" % (name, base["name"], slide))


def close_item(iid):
    ensure_not_preloading()
    item = get_item(iid)
    if is_base(item):
        raise RuntimeError("编号 0 是大会背景，不能关闭")
    if item["type"] != "ppt":
        if MEDIA.name == item["name"]:
            MEDIA.stop()
        WORKER.notify()
        return True
    if (WORKER.snapshot.get("live") or {}).get("key") == norm(item["path"]) and base_item():
        show_base("background")       # 正在投影：先切到大会背景，投影上不会露出桌面
    result = WORKER.call(CTRL.close, item["path"], priority=0, timeout=60,
                         label="完全关闭 " + item["name"])
    if _LAST["id"] == item["id"]:
        _LAST["id"] = None
    WORKER.notify()
    return result


def closable_item(snapshot=None):
    """
    “完全关闭”要关哪一个：正在放映的节目；如果正在放的是大会背景，就是刚才播过、还开着的那个节目。
    只读快照，不碰 PowerPoint。没有可关的返回 None。
    """
    snap = snapshot if snapshot is not None else WORKER.snapshot
    live = find_by_key((snap.get("live") or {}).get("key"))
    if live and not is_base(live):
        return live
    if MEDIA.status()["playing"]:
        with ITEMS_LOCK:
            playing = next((it for it in ITEMS.values() if it["name"] == MEDIA.name), None)
        if playing:
            return playing
    last = _LAST["id"]
    with ITEMS_LOCK:
        item = ITEMS.get(last) if last else None
    if item and item["type"] == "ppt" and norm(item["path"]) in set(snap.get("loaded") or []):
        return item
    return None


def close_current():
    item = closable_item()
    if not item:
        raise RuntimeError("没有可以关闭的节目")
    return close_item(item["id"])


# ------------------------------------------------------------ 状态变化（在 PowerPoint 线程里被调用）
def log_live_change(old, new):
    o, n = (old or {}).get("live") or {}, (new or {}).get("live") or {}
    if (o.get("key"), o.get("slide"), o.get("state")) == (n.get("key"), n.get("slide"), n.get("state")):
        return
    if not n:
        log("放映", "已结束")
        return
    log("放映", "%s  第 %s/%s 页  %s" % (os.path.basename(n.get("key") or "?"), n.get("slide") or "-",
                                       n.get("count"), STATE_NAMES.get(n.get("state"), n.get("state"))))


def auto_advance(old, new):
    """放映到了结尾黑屏 → 在新线程里开始待命的节目（不能在 PowerPoint 线程里等自己）。"""
    live = (new or {}).get("live") or {}
    if live.get("state") != SHOW_DONE or not CFG.get("auto_advance", True):
        return
    if not CFG.get("standby"):
        return                       # 节目单播完了，就停在结尾黑屏
    key = live.get("key")
    if _ADVANCED["key"] == key and time.time() - _ADVANCED["at"] < AUTO_ADVANCE_DEBOUNCE:
        return
    _ADVANCED.update(key=key, at=time.time())

    def run():
        try:
            go("自动接续：")
        except Exception as e:
            log("自动接续", "没有成功：%s" % e)

    threading.Thread(target=run, daemon=True, name="auto-advance").start()


def register():
    WORKER.change_hooks.extend([log_live_change, auto_advance])
