# -*- coding: utf-8 -*-
"""MockController：非 Windows 环境下的演示模式（没有 PowerPoint），用来测试界面和流程。"""
from presenter.library import norm
from presenter.powerpoint.constants import SHOW_BLACK, SHOW_DONE, SHOW_RUNNING, SHOW_WHITE
from presenter.powerpoint.controller import PPTController


class MockController(PPTController):
    def __init__(self):
        super().__init__()
        self.state = {}

    def _alive(self):
        return True

    def poll_interval(self):
        return 0.1

    def mark_saved_all(self):
        return True

    def load(self, path):
        key = norm(path)
        self.pres[key] = True
        n = 8 + (sum(map(ord, key)) % 20)
        return self.deck_cache.setdefault(key, {"count": n, "titles": ["第 %d 页" % i for i in range(1, n + 1)],
                                                "hidden": [False] * n, "ratio": 16 / 9})

    def start(self, path, slide=None, end_previous=True):
        self.load(path)
        self.live = norm(path)
        self.state = {"slide": int(slide or 1), "state": SHOW_RUNNING}
        return True

    def nav(self, action, slide=None):
        if not self.live:
            raise RuntimeError("目前没有正在放映的 PPT")
        n = self.deck_cache[self.live]["count"]
        s = self.state
        if action == "next":
            if s["state"] in (SHOW_BLACK, SHOW_WHITE):
                s["state"] = SHOW_RUNNING
            elif s["slide"] >= n:
                s["state"] = SHOW_DONE
            else:
                s["slide"] += 1
        elif action == "prev":
            if s["state"] in (SHOW_BLACK, SHOW_WHITE, SHOW_DONE):
                s["state"] = SHOW_RUNNING
            else:
                s["slide"] = max(1, s["slide"] - 1)
        elif action == "first":
            s.update(slide=1, state=SHOW_RUNNING)
        elif action == "last":
            s.update(slide=n, state=SHOW_RUNNING)
        elif action == "goto":
            s.update(slide=int(slide), state=SHOW_RUNNING)
        elif action == "black":
            s["state"] = SHOW_RUNNING if s["state"] == SHOW_BLACK else SHOW_BLACK
        elif action == "white":
            s["state"] = SHOW_RUNNING if s["state"] == SHOW_WHITE else SHOW_WHITE
        elif action == "end":
            self.live = None
        return True

    def close(self, path):
        key = norm(path)
        self.pres.pop(key, None)
        if self.live == key:
            self.live = None
        return True

    def fonts(self, path):
        return [("演示字体", False), ("黑体", False)]

    def export_thumb(self, path, n, out, width):
        raise RuntimeError("演示模式没有缩略图")

    def status(self):
        st = {"powerpoint": True, "live": None, "loaded": list(self.pres.keys()), "mock": True}
        if self.live:
            s = self.state
            st["live"] = {"key": self.live, "slide": None if s["state"] == SHOW_DONE else s["slide"],
                          "count": self.deck_cache[self.live]["count"], "state": s["state"],
                          "notes": "演示模式的备注：第 %d 页" % s["slide"]}
        return st
