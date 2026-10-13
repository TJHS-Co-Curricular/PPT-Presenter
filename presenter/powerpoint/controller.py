# -*- coding: utf-8 -*-
"""
PPTController：通过 COM 控制 PowerPoint。

所有方法都只能在 ComWorker 线程里调用（见 worker.py），不要在别的线程直接调用。

规则（都是踩过的坑）：
- PPT 一律只读打开，并随时标记为“已保存”，否则 PowerPoint 会弹出模态的
  “Save your changes to this file?”，让所有 COM 调用卡住。
- **要带窗口打开**（打开后自动最小化）。无窗口打开（WithWindow=False）时 PowerPoint 不载入内嵌字体，
  平滑切换（Morph）等效果也可能不正常（v2.0 修复）。
- 用 PowerPoint 的事件得知翻页（events.py），放映中尽量少打扰 PowerPoint：
  频繁的 COM 调用会插队打断动画渲染。status() 只读轻量属性，并缓存放映窗口。
"""
import json
import os
import time

from presenter import cache, paths
from presenter.library import norm
from presenter.log import log
from presenter.powerpoint import events
from presenter.powerpoint.com_errors import is_busy_error
from presenter.powerpoint.constants import (
    MSO_TRUE,
    PLACEHOLDER_BODY,
    SHOW_ALL,
    SHOW_BLACK,
    SHOW_DONE,
    SHOW_RUNNING,
    SHOW_WHITE,
    WINDOW_MINIMIZED,
)

NOTES_SLOW_MS = 800          # 读一页备注超过这个时间，就停用这份 PPT 的备注
STATUS_SLOW_MS = 300         # 读状态超过这个时间，输出诊断日志
REENUM_SECONDS = 3           # 放映窗口缓存多久重新确认一次（发现别人手动开的放映）

# 状态刷新间隔（秒）
POLL_IDLE = 0.5              # 没在放映
POLL_SHOW_WITH_EVENTS = 1.0  # 放映中、有翻页事件：只是兜底（例如在 PowerPoint 里按 B 黑屏没有事件）
POLL_SHOW_NO_EVENTS = 0.25   # 放映中、没有事件：只能靠查询


class PPTController:
    def __init__(self):
        self.app = None
        self.pres = {}              # norm(path) → Presentation
        self.live = None            # norm(path)：当前放映中的文件
        self.deck_cache = {}        # norm(path) → {"count", "titles", "hidden", "ratio"}
        self.notes_cache = {}       # (norm(path), 页) → 备注文字
        self.notes_slow = set()     # 读备注太慢、已停用备注的 PPT
        self.notes_wanted = None    # 等空闲时去读的 (key, 页)
        self.notes_dirty = set()    # 备注有更新、还没写进 .cache 的 PPT
        self._loaded_checked = 0.0
        self._slow_logged = 0.0
        self.with_window = True     # 带窗口打开（app.setup() 按设置改）
        self.dirty = False          # PowerPoint 事件通知“有变化”，需要马上刷新状态
        self.events_seen = 0
        self._events = None         # 事件对象（保留引用，否则会被回收）
        self._live_win = None       # 缓存的放映窗口
        self._live_win_checked = 0.0

    def poll_interval(self):
        """PowerPoint 线程多久刷新一次状态。放映中尽量少打扰 PowerPoint。"""
        if self.dirty:
            return 0.0
        if not self.live:
            return POLL_IDLE
        return POLL_SHOW_WITH_EVENTS if self._events is not None else POLL_SHOW_NO_EVENTS

    # ------------------------------------------------------------ 连接
    def _alive(self):
        if self.app is None:
            return False
        try:
            _ = self.app.Version          # 探测：PowerPoint 还在不在
            return True
        except Exception as e:
            if is_busy_error(e):
                return True   # 只是忙，不是关掉了
            self.reset()
            return False

    def reset(self):
        """PowerPoint 被关掉 / 崩溃后，丢掉所有旧的连接。"""
        self.app = None
        self._events = None
        self.pres.clear()
        self.live = None
        self._live_win = None

    def _app(self):
        if self._alive():
            return self.app
        import win32com.client
        last = None
        for progid in ("PowerPoint.Application", "KWPP.Application"):   # KWPP = WPS 演示
            try:
                self.app = win32com.client.Dispatch(progid)
                break
            except Exception as e:
                last = e
        if self.app is None:
            raise RuntimeError("无法启动 PowerPoint：%s" % last)
        try:
            self.app.Visible = True
        except Exception:
            pass
        self._events = events.attach(self.app, self)
        return self.app

    # ------------------------------------------------------------ 文件
    @staticmethod
    def _mark_saved(p):
        """放映、跳页、导出缩略图都会让 PowerPoint 觉得文件改过了；标记为已保存就不会弹保存提示。"""
        try:
            if not p.Saved:
                p.Saved = MSO_TRUE
        except Exception:
            pass

    def mark_saved_all(self):
        for p in list(self.pres.values()):
            self._mark_saved(p)
        return True

    def _find_open(self, key):
        app = self._app()
        for i in range(1, app.Presentations.Count + 1):
            p = app.Presentations(i)
            try:
                if norm(p.FullName) == key:
                    return p
            except Exception:
                pass
        return None

    def _get_pres(self, path, open_if_needed=True):
        key = norm(path)
        p = self.pres.get(key)
        if p is not None:
            try:
                _ = p.FullName            # 探测：文件还开着吗
                return p
            except Exception:
                self.pres.pop(key, None)
        p = self._find_open(key)
        if p is not None and self.with_window and not self._has_window(p) and self._show_window(key) is None:
            # 以前（v1.9 及更早）无窗口打开、还留在 PowerPoint 里的：关掉重新开，才能载入内嵌字体
            log("PowerPoint", "重新打开（之前是无窗口打开，内嵌字体可能没载入）：%s" % os.path.basename(path))
            self._mark_saved(p)
            p.Close()
            p = None
        if p is None:
            if not open_if_needed:
                return None
            p = self._open(path)
        self.pres[key] = p
        return p

    @staticmethod
    def _has_window(p):
        try:
            return p.Windows.Count > 0
        except Exception:
            return True

    def _open(self, path):
        """
        只读打开（不锁文件、不会误存）。带窗口打开，内嵌字体和平滑切换才正常；打开后把编辑窗口最小化。
        参数：FileName, ReadOnly=True, Untitled=False, WithWindow
        """
        p = self._app().Presentations.Open(path, True, False, bool(self.with_window))
        self._mark_saved(p)
        if self.with_window:
            try:
                p.Windows(1).WindowState = WINDOW_MINIMIZED
            except Exception:
                pass
        return p

    def load(self, path):
        """打开（如果还没开）并读取页数、标题、隐藏页；同时载入备注的硬盘缓存。"""
        key = norm(path)
        p = self._get_pres(path)
        if key not in self.deck_cache:
            titles, hidden = [], []
            count = p.Slides.Count
            for i in range(1, count + 1):
                s = p.Slides(i)
                title = ""
                try:
                    if s.Shapes.HasTitle:
                        title = s.Shapes.Title.TextFrame.TextRange.Text.strip()
                except Exception:
                    pass
                titles.append(title.replace("\r", " ").replace("\v", " ")[:60])
                try:
                    hidden.append(bool(s.SlideShowTransition.Hidden))
                except Exception:
                    hidden.append(False)
            try:
                ratio = float(p.PageSetup.SlideWidth) / float(p.PageSetup.SlideHeight)
            except Exception:
                ratio = 16 / 9
            self.deck_cache[key] = {"count": count, "titles": titles, "hidden": hidden, "ratio": ratio}
            self._load_notes_file(path)
        return self.deck_cache[key]

    def fonts(self, path):
        """PPT 用到的字体：[(名字, 是否内嵌)]。"""
        p = self._get_pres(path)
        out = []
        for i in range(1, p.Fonts.Count + 1):
            try:
                f = p.Fonts(i)
                out.append((str(f.Name), bool(f.Embedded)))
            except Exception:
                pass
        return out

    def close(self, path):
        key = norm(path)
        w = self._show_window(key)
        if w is not None:
            w.View.Exit()
        p = self._get_pres(path, open_if_needed=False)
        if p is not None:
            self._mark_saved(p)
            p.Close()
        self.pres.pop(key, None)
        if self.live == key:
            self.live = None
            self._live_win = None
        return True

    # ------------------------------------------------------------ 放映
    def _show_window(self, key):
        if not self._alive():
            return None
        app = self.app
        for i in range(1, app.SlideShowWindows.Count + 1):
            w = app.SlideShowWindows(i)
            try:
                if norm(w.Presentation.FullName) == key:
                    return w
            except Exception:
                pass
        return None

    def _live_window(self):
        """
        当前放映窗口。缓存起来，平时只用一个轻量调用确认它还在；
        每隔几秒再看一下放映窗口数量，有别人手动开的放映才重新列举。
        """
        if not self._alive():
            return None
        w = self._live_win
        if w is not None:
            try:
                _ = w.View.State          # 探测：放映窗口还在吗
                valid = True
            except Exception as e:
                valid = is_busy_error(e)
            if valid:
                if time.time() - self._live_win_checked < REENUM_SECONDS:
                    return w
                self._live_win_checked = time.time()
                try:
                    if self.app.SlideShowWindows.Count <= 1:
                        return w
                except Exception:
                    return w
            else:
                self._live_win = None
        self._live_win_checked = time.time()
        if self.live:
            w = self._show_window(self.live)
            if w is not None:
                self._live_win = w
                return w
        # 有人在电脑上手动开了放映，也接管过来
        if self.app.SlideShowWindows.Count > 0:
            w = self.app.SlideShowWindows(self.app.SlideShowWindows.Count)
            try:
                self.live = norm(w.Presentation.FullName)
            except Exception:
                pass
            self._live_win = w
            return w
        self.live = None
        self._live_win = None
        return None

    def start(self, path, slide=None, end_previous=True):
        """
        开始放映（已经在放映就切过去）。先开新的再结束旧的，投影上不会闪出桌面。
        end_previous：是否结束上一份 PPT 的放映。
        """
        key = norm(path)
        p = self._get_pres(path)
        info = self.load(path)
        old = self.live
        w = self._show_window(key)
        if w is None:
            settings = p.SlideShowSettings
            try:
                settings.RangeType = SHOW_ALL
            except Exception:
                pass
            w = settings.Run()
        if slide and 1 <= int(slide) <= info["count"]:
            if int(slide) != 1 or w.View.State == SHOW_DONE:
                w.View.GotoSlide(int(slide))
        if w.View.State in (SHOW_BLACK, SHOW_WHITE):
            w.View.State = SHOW_RUNNING
        try:
            w.Activate()
        except Exception:
            pass
        self.live = key
        self._live_win, self._live_win_checked = w, time.time()
        if end_previous and old and old != key:
            ow = self._show_window(old)
            if ow is not None:
                try:
                    ow.View.Exit()
                except Exception:
                    pass
        return True

    def nav(self, action, slide=None):
        w = self._live_window()
        if w is None:
            raise RuntimeError("目前没有正在放映的 PPT")
        v = w.View
        if action in ("next", "prev"):
            if v.State in (SHOW_BLACK, SHOW_WHITE):
                v.State = SHOW_RUNNING        # 黑 / 白屏时按翻页 = 先恢复画面
            elif action == "next":
                v.Next()
            else:
                v.Previous()
        elif action == "first":
            v.First()
        elif action == "last":
            v.Last()
        elif action == "goto":
            v.GotoSlide(int(slide))
            if v.State != SHOW_RUNNING:
                v.State = SHOW_RUNNING
        elif action == "black":
            v.State = SHOW_RUNNING if v.State == SHOW_BLACK else SHOW_BLACK
        elif action == "white":
            v.State = SHOW_RUNNING if v.State == SHOW_WHITE else SHOW_WHITE
        elif action == "end":
            v.Exit()
            self.live = None
            self._live_win = None
        elif action == "focus":
            w.Activate()
        else:
            raise ValueError("未知操作：%s" % action)
        return True

    # ------------------------------------------------------------ 缩略图
    def export_thumb(self, path, n, out, width):
        p = self._get_pres(path)
        info = self.load(path)
        height = int(round(width / info["ratio"]))
        p.Slides(int(n)).Export(out, "PNG", int(width), height)
        self._mark_saved(p)
        return out

    # ------------------------------------------------------------ 备注
    @staticmethod
    def _notes(slide):
        # 先用最快的方式（备注正文固定是第 2 个占位符），不行再逐个找
        try:
            return slide.NotesPage.Shapes.Placeholders(2).TextFrame.TextRange.Text.replace("\r", "\n").strip()
        except Exception:
            pass
        try:
            shapes = slide.NotesPage.Shapes
            for i in range(1, shapes.Count + 1):
                sh = shapes(i)
                try:
                    if sh.PlaceholderFormat.Type == PLACEHOLDER_BODY and sh.HasTextFrame:
                        return sh.TextFrame.TextRange.Text.replace("\r", "\n").strip()
                except Exception:
                    pass
        except Exception:
            pass
        return ""

    def fetch_notes(self, key, idx):
        """
        读一页的演讲者备注（只在空闲 / 预加载时做，不挡翻页）。
        读备注会打开备注页，含大图 / 视频的页面可能要好几秒；超过 NOTES_SLOW_MS 就停用这份 PPT 的备注。
        """
        p = self.pres.get(key)
        if p is None or key in self.notes_slow:
            return False
        t0 = time.time()
        try:
            text = self._notes(p.Slides(idx))
        except Exception:
            text = ""
        ms = (time.time() - t0) * 1000
        self.notes_cache[(key, idx)] = text
        self.notes_dirty.add(key)
        name = os.path.basename(key)
        if ms > NOTES_SLOW_MS:
            self.notes_slow.add(key)
            log("备注", "%s 第 %d 页读取备注花了 %dms，太慢，这份 PPT 停用备注显示（不影响放映）" % (name, idx, ms))
        elif ms > 200:
            log("备注", "%s 第 %d 页 %dms" % (name, idx, ms))
        return True

    def _load_notes_file(self, path):
        f = cache.notes_path(path)
        if not f or not os.path.exists(f):
            return
        try:
            with open(f, "r", encoding="utf-8") as fh:
                data = json.load(fh)
            key = norm(path)
            for k, v in (data.get("notes") or {}).items():
                self.notes_cache[(key, int(k))] = v
            if data.get("slow"):
                self.notes_slow.add(key)
        except Exception as e:
            log("备注", "读取缓存失败 %s：%s" % (os.path.basename(path), e))

    def save_notes(self):
        """把有更新的备注写进 .cache（不碰 PowerPoint，任何线程都可以调用）。"""
        for key in list(self.notes_dirty):
            self.notes_dirty.discard(key)
            f = cache.notes_path(key)
            if not f:
                continue
            notes = {str(i): t for (k, i), t in list(self.notes_cache.items()) if k == key}
            try:
                os.makedirs(paths.NOTES_DIR, exist_ok=True)
                with open(f, "w", encoding="utf-8") as fh:
                    json.dump({"notes": notes, "slow": key in self.notes_slow}, fh, ensure_ascii=False)
            except OSError as e:
                log("备注", "写入缓存失败：%s" % e)

    # ------------------------------------------------------------ 状态
    def status(self):
        """当前状态（每 0.1 秒一次）。只读轻量属性；超过 STATUS_SLOW_MS 会输出每一步的耗时。"""
        t = [time.time()]
        marks = []

        def mark(name):
            t.append(time.time())
            marks.append("%s %dms" % (name, (t[-1] - t[-2]) * 1000))

        self.dirty = False
        st = {"powerpoint": self._alive(), "live": None, "loaded": [], "events": self._events is not None}
        mark("连接")
        if not st["powerpoint"]:
            return st
        # “哪些文件还开着”比较费时，每 2 秒检查一次就够
        now = time.time()
        if now - self._loaded_checked > 2:
            self._loaded_checked = now
            for k, p in list(self.pres.items()):
                try:
                    _ = p.FullName        # 探测：文件还开着吗
                except Exception as e:
                    if not is_busy_error(e):
                        self.pres.pop(k, None)
            mark("文件列表")
        st["loaded"] = list(self.pres.keys())
        w = self._live_window()
        mark("放映窗口")
        if w is not None:
            v = w.View
            state = v.State
            idx = None
            if state != SHOW_DONE:
                try:
                    idx = v.Slide.SlideIndex
                except Exception:
                    idx = None
            mark("页码")
            key = self.live
            # 备注只取已经读好的；没读过的交给空闲任务（idle_fetch_notes）
            notes = self.notes_cache.get((key, idx), "") if idx else ""
            if idx and (key, idx) not in self.notes_cache and key not in self.notes_slow:
                self.notes_wanted = (key, idx)
            dc = self.deck_cache.get(key)
            count = dc["count"] if dc else w.Presentation.Slides.Count
            st["live"] = {"key": key, "slide": idx, "count": count, "state": state, "notes": notes,
                          "notes_off": key in self.notes_slow,
                          "notes_missing": bool(idx) and (key, idx) not in self.notes_cache}
        total = (time.time() - t[0]) * 1000
        if total > STATUS_SLOW_MS and time.time() - self._slow_logged > 5:
            self._slow_logged = time.time()
            log("诊断", "读取放映状态花了 %dms：%s" % (total, "，".join(marks)))
        return st
