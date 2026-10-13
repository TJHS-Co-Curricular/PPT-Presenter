# -*- coding: utf-8 -*-
"""
PowerPoint 线程的空闲任务（队列空的时候才跑，每次只做一小件事）。

    fetch_current_notes  当前页还没读过的备注
    mark_saved           每 3 秒把打开的 PPT 标记为“已保存”（避免保存提示）
    save_notes           每 5 秒把新读的备注写进 .cache
    prefetch_live        放映中：当前页附近的缩略图（默认关闭，见 work_during_show）
    prefetch_notes       没在放映：已加载 PPT 的全部备注
    prefetch_thumbs      没在放映：已加载 PPT 的全部缩略图

**放映中不做重活**：导出缩略图、读备注都会占用 PowerPoint，打断平滑切换和动画（v2.0）。
放映要用的缩略图和备注请事先预加载。预加载期间这些任务全部暂停（预加载自己会做）。
"""
import os
import time

from presenter.cache import thumb_path
from presenter.config import CFG
from presenter.core import CTRL, MOCK, WORKER
from presenter.library import find_by_key, norm, ppt_items
from presenter.log import log
from presenter.powerpoint.constants import SHOW_DONE
from presenter.preload import PRELOAD

_last = {"mark": 0.0, "notes_save": 0.0}
_thumbs_done = set()          # (id, mtime)：已经全部生成过缩略图的 PPT


def _export_if_missing(item, n):
    out = thumb_path(item, n)
    if os.path.exists(out):
        return False
    os.makedirs(os.path.dirname(out), exist_ok=True)
    t0 = time.time()
    CTRL.export_thumb(item["path"], n, out, int(CFG.get("thumb_width", 480)))
    log("缩略图", "预先生成 %s 第 %d 页  %dms" % (item["name"], n, (time.time() - t0) * 1000))
    WORKER.notify()
    return True


def _quiet_during_show():
    """放映中不打扰 PowerPoint（除非设置里打开了 work_during_show）。"""
    return bool(CTRL.live) and not CFG.get("work_during_show", False)


def fetch_current_notes():
    want = CTRL.notes_wanted
    if MOCK or not want or _quiet_during_show():
        return False
    CTRL.notes_wanted = None
    if want in CTRL.notes_cache:
        return False
    CTRL.fetch_notes(*want)
    WORKER.request_refresh()      # 马上刷新状态，把备注带给界面
    return True


def mark_saved():
    if MOCK or not CTRL.pres or time.time() - _last["mark"] < 3:
        return False
    _last["mark"] = time.time()
    CTRL.mark_saved_all()
    return False


def save_notes():
    if CTRL.notes_dirty and time.time() - _last["notes_save"] > 5:
        _last["notes_save"] = time.time()
        CTRL.save_notes()
    return False


def prefetch_live():
    if MOCK or not CFG.get("prefetch_thumbs", True) or not CTRL.live or _quiet_during_show():
        return False
    live = WORKER.snapshot.get("live")
    if not live or not live.get("slide") or live.get("state") == SHOW_DONE:
        return False
    item = find_by_key(live.get("key"))
    if not item:
        return False
    cur, count = live["slide"], live["count"]
    for n in (cur, cur + 1, cur + 2, cur - 1, cur + 3):
        if 1 <= n <= count and _export_if_missing(item, n):
            return True
    return False


def prefetch_notes():
    if MOCK or PRELOAD["running"] or CTRL.live or not CTRL.pres:
        return False
    for key in list(CTRL.pres.keys()):
        deck = CTRL.deck_cache.get(key)
        if not deck or key in CTRL.notes_slow:
            continue
        for n in range(1, deck["count"] + 1):
            if (key, n) not in CTRL.notes_cache:
                CTRL.fetch_notes(key, n)
                return True
    return False


def prefetch_thumbs():
    if MOCK or PRELOAD["running"] or not CFG.get("prefetch_thumbs", True) or CTRL.live or not CTRL.pres:
        return False
    for item in ppt_items():
        tag = (item["id"], item["mtime"])
        key = norm(item["path"])
        if tag in _thumbs_done or key not in CTRL.pres or key not in CTRL.deck_cache:
            continue
        for n in range(1, CTRL.deck_cache[key]["count"] + 1):
            if _export_if_missing(item, n):
                return True
        _thumbs_done.add(tag)
    return False


def register():
    WORKER.idle_jobs.extend([fetch_current_notes, mark_saved, save_notes,
                             prefetch_live, prefetch_notes, prefetch_thumbs])
