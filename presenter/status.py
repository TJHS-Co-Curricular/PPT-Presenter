# -*- coding: utf-8 -*-
"""
发给界面的完整状态。

PowerPoint 线程定时生成一份快照（收到翻页事件马上刷新；放映中 1 秒兜底，没放映 0.5 秒）；
这里只是读快照、再拼上预加载 / 视频 / 待命等信息，
**不排队等 PowerPoint**，所以即使 PowerPoint 正在打开大文件，界面也不会卡。
"v" 是整份状态的指纹，长轮询用它判断“有没有变化”。
"""
import copy
import hashlib
import json
import os
import time

from presenter import show
from presenter.cache import thumb_path
from presenter.config import CFG
from presenter.core import MOCK, WORKER
from presenter.library import ITEMS, ITEMS_LOCK, base_item, norm
from presenter.media import MEDIA
from presenter.preload import PRELOAD


def build_status():
    st = copy.deepcopy(WORKER.snapshot)
    st["busy"] = WORKER.busy
    st["busy_for"] = int(time.time() - WORKER.busy_since) if WORKER.busy_since else 0
    with ITEMS_LOCK:
        by_key = {norm(it["path"]): it["id"] for it in ITEMS.values()}
        live = st.get("live")
        if live:
            live["id"] = by_key.get(live.pop("key", None))
            # 告诉界面哪些页的缩略图已经做好，可以直接显示
            item = ITEMS.get(live["id"]) if live["id"] else None
            if item and live.get("slide"):
                live["ready"] = [n for n in range(live["slide"] - 1, live["slide"] + 4)
                                 if 1 <= n <= live["count"] and os.path.exists(thumb_path(item, n))]
    st["loaded"] = [by_key[k] for k in st.get("loaded", []) if k in by_key]
    st["preload"] = dict(PRELOAD, errors=list(PRELOAD["errors"]),
                         font_warnings=[dict(w) for w in PRELOAD["font_warnings"]])
    st["media"] = MEDIA.status()
    st["mock"] = MOCK
    st["standby"] = CFG.get("standby") or ""
    base = base_item()
    st["base"] = base["id"] if base else ""
    closable = show.closable_item(WORKER.snapshot)
    st["closable"] = {k: closable[k] for k in ("id", "num", "title", "type")} if closable else None
    st["auto_advance"] = bool(CFG.get("auto_advance", True))
    if not st.get("error"):
        st.pop("error", None)
    st["v"] = hashlib.md5(json.dumps(st, sort_keys=True, ensure_ascii=False).encode("utf-8")).hexdigest()[:10]
    return st


def wait_for_change(since, timeout=20):
    """长轮询：状态和 since 一样就等，一变化（或超时）就返回新状态。"""
    st = build_status()
    if not since or st["v"] != since:
        return st
    deadline = time.time() + timeout
    while time.time() < deadline:
        with WORKER.changed:
            WORKER.changed.wait(0.5)   # 0.5 秒兜底：视频播完、预加载进度这类变化不一定会 notify
        st = build_status()
        if st["v"] != since:
            break
    return st
