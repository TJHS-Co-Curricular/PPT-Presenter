# -*- coding: utf-8 -*-
"""
PPT 控制台 —— 多份 PowerPoint / 视频的集中控制（本机桌面程序）
运行：python server.py   (Windows + Microsoft PowerPoint)
"""
import os
import sys
import json
import re
import time
import socket
import queue
import shutil
import hashlib
import itertools
import threading
import subprocess
import traceback
import copy
import webbrowser
import logging

from flask import Flask, jsonify, request, send_file, send_from_directory, abort

FROZEN = getattr(sys, "frozen", False)
# BASE：放 config.json / 缓存的地方（exe 所在文件夹）；RES：界面文件（打包后在临时解压目录里）
BASE = os.path.dirname(os.path.abspath(sys.executable if FROZEN else __file__))
RES = getattr(sys, "_MEIPASS", BASE)
CONFIG_PATH = os.path.join(BASE, "config.json")
LOG_PATH = os.path.join(BASE, "ppt-presenter.log")

import desktop  # noqa: E402
import version  # noqa: E402

# 版本号来自 CHANGELOG.md（打包时一起放进 exe）
VERSION = version.get_version(RES, BASE)

# 调试窗口模式：这个进程只负责在自己的命令行窗口里实时显示日志
if len(sys.argv) >= 4 and sys.argv[1] == "--debug-console":
    desktop.run_debug_console(sys.argv[2], int(sys.argv[3]))
    sys.exit(0)

# print 和报错都写进日志文件（打包成无控制台的 exe 时，不这样做会直接丢失）
desktop.setup_logging(LOG_PATH, FROZEN)
DEBUG = desktop.DebugConsole(LOG_PATH, FROZEN, os.path.abspath(__file__))
CACHE_DIR = os.path.join(BASE, ".cache", "thumbs")

DEFAULT_CONFIG = {
    "folder": "",               # PPT 所在文件夹（留空 = exe 所在文件夹；若那里没有 PPT 而有 Example 子文件夹，就用 Example）
    "port": 0,                  # 0 = 自动找一个空闲端口（只在本机 127.0.0.1 上，外部连不进来）
    "end_previous_show": True,  # 切换到另一份 PPT 时，自动结束上一份的放映
    "auto_advance": True,       # PPT 放到结尾（“放映结束”黑屏）时，自动开始待命的下一个节目
    "standby": "",              # 当前待命的节目（程序自动记住）
    "preload_on_start": False,  # 启动时自动把所有 PPT 打开到后台
    "thumb_width": 480,
    "prefetch_thumbs": True,    # PowerPoint 空闲且没在放映时，自动预先生成缩略图
    "debug_console": False,     # 启动时自动打开调试窗口（也可以用 --debug 参数或界面上的 🐞 按钮）
    "window_on_top": False,     # 控制台窗口置顶
    "window_size": [1440, 900],
    "hotkeys": dict(desktop.DEFAULT_HOTKEYS),  # 全局快捷键（PowerPoint 在前台时也有效）
}


PPT_EXT = {".ppt", ".pptx", ".pptm", ".pps", ".ppsx", ".ppsm"}
VIDEO_EXT = {".mp4", ".mov", ".avi", ".mkv", ".wmv", ".m4v", ".mpg", ".mpeg"}
AUDIO_EXT = {".mp3", ".wav", ".m4a", ".aac", ".flac", ".wma"}

# PowerPoint 常量
ppSlideShowRunning, ppSlideShowPaused, ppSlideShowBlackScreen, ppSlideShowWhiteScreen, ppSlideShowDone = 1, 2, 3, 4, 5
ppPlaceholderBody = 2


# ---------------------------------------------------------------- 配置
def load_config():
    cfg = dict(DEFAULT_CONFIG)
    if os.path.exists(CONFIG_PATH):
        try:
            with open(CONFIG_PATH, "r", encoding="utf-8") as f:
                cfg.update(json.load(f))
        except Exception as e:
            print("读取 config.json 失败：", e)
    if not cfg.get("folder"):
        cfg["folder"] = default_folder()
    # 旧版的远程设置不再需要
    for k in ("host", "pin", "open_browser"):
        cfg.pop(k, None)
    hk = dict(desktop.DEFAULT_HOTKEYS)
    hk.update({k: v for k, v in (cfg.get("hotkeys") or {}).items() if k in desktop.ACTIONS})
    # 旧版默认把 GO 设在 F12，但 F12 被 Windows 保留，不能当全局快捷键
    for k, v in hk.items():
        if str(v).strip().upper() == "F12":
            hk[k] = desktop.DEFAULT_HOTKEYS.get(k, "") if k != "go" else "F8"
    cfg["hotkeys"] = hk
    if not os.path.exists(CONFIG_PATH):
        save_config(cfg)
    return cfg


def default_folder():
    def has_media(d):
        try:
            return any(os.path.splitext(n)[1].lower() in PPT_EXT | VIDEO_EXT
                       for n in os.listdir(d) if not n.startswith(("~$", "._")))
        except OSError:
            return False
    ex = os.path.join(BASE, "Example")
    if not has_media(BASE) and os.path.isdir(ex):
        return ex
    return BASE


def save_config(cfg):
    with open(CONFIG_PATH, "w", encoding="utf-8") as f:
        json.dump(cfg, f, ensure_ascii=False, indent=2)


CFG = load_config()


# ---------------------------------------------------------------- 文件扫描
def norm(p):
    return os.path.normcase(os.path.abspath(p))


def item_id(path):
    return hashlib.md5(norm(path).encode("utf-8")).hexdigest()[:12]


def natural_key(s):
    return [int(t) if t.isdigit() else t.lower() for t in re.split(r"(\d+)", s)]


def scan_folder(folder):
    items = []
    if not folder or not os.path.isdir(folder):
        return items
    for name in os.listdir(folder):
        if name.startswith("~$") or name.startswith("._") or name.startswith("."):
            continue  # PowerPoint 锁文件 / macOS 残留
        full = os.path.join(folder, name)
        if not os.path.isfile(full):
            continue
        ext = os.path.splitext(name)[1].lower()
        if ext in PPT_EXT:
            kind = "ppt"
        elif ext in VIDEO_EXT:
            kind = "video"
        elif ext in AUDIO_EXT:
            kind = "audio"
        else:
            continue
        stem = os.path.splitext(name)[0]
        # 只把“短编号”当作序号（例如 0、5.4、10、5.1-5.3），2025 这类年份不算
        m = re.match(r"^\s*(\d{1,2}(?:\.\d{1,2})?(?:\s*-\s*\d{1,2}(?:\.\d{1,2})?)?)(?!\d)\s*[.、_\s-]?\s*(.+)$", stem)
        num, title = "", stem
        if m:
            num, title = m.group(1).replace(" ", ""), m.group(2)
        st = os.stat(full)
        items.append({
            "id": item_id(full),
            "path": full,
            "name": name,
            "title": title,
            "num": num,
            "type": kind,
            "size": st.st_size,
            "mtime": int(st.st_mtime),
        })
    items.sort(key=lambda it: (0 if it["num"] else 1, natural_key(it["name"])))
    return items


ITEMS = {}
ITEMS_LOCK = threading.Lock()


def refresh_items():
    lst = scan_folder(CFG["folder"])
    with ITEMS_LOCK:
        ITEMS.clear()
        for it in lst:
            ITEMS[it["id"]] = it
    # 待命的节目不在列表里了（换了文件夹等），就从第一个开始
    if CFG.get("standby") not in ITEMS:
        CFG["standby"] = lst[0]["id"] if lst else ""
    return lst


def get_item(iid):
    with ITEMS_LOCK:
        it = ITEMS.get(iid)
    if not it:
        abort(404, "找不到这个项目，请刷新列表")
    return it


# ---------------------------------------------------------------- COM 工作线程
# PowerPoint 忙碌（正在打开大文件、正在开始放映、播放内嵌视频）时会拒绝调用，这些错误码要等一下再重试
HR_BUSY = {-2147418111, -2147417846}          # RPC_E_CALL_REJECTED, RPC_E_SERVERCALL_RETRYLATER
HR_DEAD = {-2147417848, -2147023174, -2147023170, -2147220995}  # 断开 / RPC 服务器不可用 / 对象已不存在


def com_hresult(e):
    try:
        return e.args[0] if e.args and isinstance(e.args[0], int) else getattr(e, "hresult", None)
    except Exception:
        return None


def is_busy_error(e):
    if com_hresult(e) in HR_BUSY:
        return True
    # 动态调度在 PowerPoint 忙时查不到成员名，会变成 AttributeError: PowerPoint.Application.Presentations
    return isinstance(e, AttributeError) and str(e).startswith(("PowerPoint.", "KWPP."))


def is_dead_error(e):
    return com_hresult(e) in HR_DEAD


NAV_NAMES = {"next": "下一页", "prev": "上一页", "first": "第一页", "last": "最后一页", "goto": "跳到第 %s 页",
             "black": "黑屏切换", "white": "白屏切换", "end": "结束放映", "focus": "切到放映窗口"}
STATE_NAMES = {1: "放映中", 2: "暂停", 3: "黑屏", 4: "白屏", 5: "已到结尾"}


def describe(fn, args, label=None):
    """把一个 PowerPoint 操作写成看得懂的一句话（给日志用）。"""
    if label:
        return label
    name = getattr(fn, "__name__", str(fn))
    base = lambda p: os.path.basename(str(p))
    if name == "nav" and args:
        t = NAV_NAMES.get(args[0], str(args[0]))
        return t % (args[1] if len(args) > 1 else "") if "%s" in t else t
    if name == "export_thumb" and len(args) >= 2:
        return "[缩略图] %s 第 %s 页" % (base(args[0]), args[1])
    if name in ("load", "start", "close") and args:
        return {"load": "打开", "start": "开始放映", "close": "关闭"}[name] + " " + base(args[0])
    return name


def log_live_change(old, new):
    o, n = (old or {}).get("live") or {}, (new or {}).get("live") or {}
    if (o.get("key"), o.get("slide"), o.get("state")) == (n.get("key"), n.get("slide"), n.get("state")):
        return
    if not n:
        print("[放映] 已结束")
        return
    print("[放映] %s  第 %s/%s 页  %s" % (os.path.basename(n.get("key") or "?"), n.get("slide") or "-",
                                        n.get("count"), STATE_NAMES.get(n.get("state"), n.get("state"))))


class ComWorker(threading.Thread):
    """
    所有 PowerPoint 调用都在这一个线程里执行。
    PowerPoint 自己一次只能处理一个 COM 调用，多开线程去调用它只会得到更多“Call was rejected”，
    所以这里保持单线程，但：控制命令优先、自动重试忙碌错误、空闲时更新状态快照和预生成缩略图。
    网页请求（多线程）只读快照，不会因为 PowerPoint 忙而卡住。
    """

    def __init__(self, ctrl):
        super().__init__(daemon=True)
        self.ctrl = ctrl
        self.q = queue.PriorityQueue()
        self.counter = itertools.count()
        self.busy = None            # 正在做的事（显示在界面上）
        self.busy_since = None
        self.snapshot = {"powerpoint": False, "live": None, "loaded": []}
        self.snapshot_at = 0.0
        self.idle_jobs = []         # 空闲时执行的低优先级函数，返回 True 表示做了事
        self.changed = threading.Condition()   # 状态一变就叫醒等待中的网页请求

    def notify(self):
        with self.changed:
            self.changed.notify_all()

    def call(self, fn, *args, priority=0, timeout=180, label=None, **kw):
        ev = threading.Event()
        box = {}
        self.q.put((priority, next(self.counter), fn, args, kw, ev, box, label))
        if not ev.wait(timeout):
            raise TimeoutError("PowerPoint 没有响应（可能有对话框挡住了，请到电脑上看看）")
        if "err" in box:
            raise box["err"]
        return box.get("res")

    def _run_with_retry(self, fn, args, kw):
        deadline = time.time() + 30
        revived = False
        self.retries = 0
        while True:
            try:
                return fn(*args, **kw)
            except Exception as e:
                if is_busy_error(e) and time.time() < deadline:
                    if self.retries == 0 and fn is not self.ctrl.status:
                        print("[PowerPoint] 忙碌中，等待重试…（%s）" % describe(fn, args))
                    self.retries += 1
                    self._pump(0.25)
                    continue
                if is_dead_error(e) and not revived:
                    # PowerPoint 被关掉或崩溃了：丢掉旧连接，重新连一次
                    print("[PowerPoint] 连接断开，重新连接…")
                    revived = True
                    self.ctrl.reset()
                    continue
                raise

    @staticmethod
    def _pump(sec):
        if sys.platform == "win32":
            import pythoncom
            end = time.time() + sec
            while time.time() < end:
                pythoncom.PumpWaitingMessages()
                time.sleep(0.02)
        else:
            time.sleep(sec)

    def _refresh_snapshot(self):
        try:
            snap = self._run_with_retry(self.ctrl.status, (), {})
        except Exception as e:
            snap = dict(self.snapshot)
            snap["error"] = str(e)
        self.snapshot_at = time.time()
        if snap != self.snapshot:
            log_live_change(self.snapshot, snap)
            try:
                on_live_change(self.snapshot, snap)
            except Exception:
                traceback.print_exc()
            self.snapshot = snap
            self.notify()

    def run(self):
        if sys.platform == "win32":
            import pythoncom
            pythoncom.CoInitialize()
        while True:
            try:
                job = self.q.get(timeout=0.08)
            except queue.Empty:
                job = None
            if job is None:
                # 空闲：先刷新状态，再做一点后台工作
                # 每 0.1 秒读一次页码：在 PowerPoint 里直接按键或用翻页笔，网页也会马上跟上
                if time.time() - self.snapshot_at > 0.1:
                    self._refresh_snapshot()
                for idle in self.idle_jobs:
                    try:
                        if idle():
                            break
                    except Exception:
                        pass
                continue
            _, _, fn, args, kw, ev, box, label = job
            self.busy, self.busy_since = label, time.time()
            if label:
                self.notify()
            t0 = time.time()
            try:
                box["res"] = self._run_with_retry(fn, args, kw)
                if fn is not self.ctrl.status:
                    print("[PowerPoint] %s  %dms%s" % (describe(fn, args, label), (time.time() - t0) * 1000,
                                                      "（重试 %d 次）" % self.retries if self.retries else ""))
            except Exception as e:
                ms = (time.time() - t0) * 1000
                if isinstance(e, (RuntimeError, ValueError, TimeoutError)):
                    # 我们自己的提示（例如“目前没有正在放映的 PPT”），不用印整段 Traceback
                    print("[提示] %s：%s" % (describe(fn, args, label), e))
                else:
                    print("[错误] %s 失败（%dms）：" % (describe(fn, args, label), ms))
                    traceback.print_exc()
                box["err"] = e
            finally:
                had_label = self.busy is not None
                self.busy = self.busy_since = None
                if had_label:
                    self.notify()
            # 先更新状态快照再通知网页，这样网页马上拿到的就是新的页码
            if fn is not self.ctrl.status:
                self._refresh_snapshot()
            ev.set()


# ---------------------------------------------------------------- PowerPoint 控制
class PPTController:
    def __init__(self):
        self.app = None
        self.pres = {}          # norm(path) -> Presentation
        self.live = None        # norm(path) 当前放映中的文件
        self.deck_cache = {}    # norm(path) -> deck info
        self.notes_cache = {}
        self.notes_slow = set()     # 读备注太慢、已停用备注的 PPT
        self.notes_wanted = None    # 等空闲时去读的 (key, 页码)

    # --- 基础
    def _alive(self):
        if self.app is None:
            return False
        try:
            self.app.Version
            return True
        except Exception as e:
            if is_busy_error(e):
                return True   # 只是忙，不是关掉了
            self.reset()
            return False

    @staticmethod
    def _mark_saved(p):
        """
        放映、跳页、导出缩略图都会让 PowerPoint 觉得文件“改过了”，关闭时就会弹出
        “Save your changes to this file?”（只读打开也一样）。标记为已保存就不会再问。
        """
        try:
            if not p.Saved:
                p.Saved = -1  # msoTrue
        except Exception:
            pass

    def mark_saved_all(self):
        for p in list(self.pres.values()):
            self._mark_saved(p)
        return True

    def reset(self):
        self.app = None
        self.pres.clear()
        self.live = None

    def _app(self):
        if self._alive():
            return self.app
        import win32com.client
        last = None
        for progid in ("PowerPoint.Application", "KWPP.Application"):
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
        return self.app

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
                p.FullName
                return p
            except Exception:
                self.pres.pop(key, None)
        p = self._find_open(key)
        if p is None:
            if not open_if_needed:
                return None
            # ReadOnly=True, Untitled=False, WithWindow=False：后台打开，不弹编辑窗口，不锁文件
            p = self._app().Presentations.Open(path, True, False, False)
            self._mark_saved(p)
        self.pres[key] = p
        return p

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
        if not self._alive():
            return None
        if self.live:
            w = self._show_window(self.live)
            if w is not None:
                return w
        # 有人在电脑上手动开了放映，也接管过来
        if self.app.SlideShowWindows.Count > 0:
            w = self.app.SlideShowWindows(self.app.SlideShowWindows.Count)
            try:
                self.live = norm(w.Presentation.FullName)
            except Exception:
                pass
            return w
        self.live = None
        return None

    # --- 公开操作
    def load(self, path):
        key = norm(path)
        p = self._get_pres(path)
        if key not in self.deck_cache:
            titles, hidden = [], []
            n = p.Slides.Count
            for i in range(1, n + 1):
                s = p.Slides(i)
                t = ""
                try:
                    if s.Shapes.HasTitle:
                        t = s.Shapes.Title.TextFrame.TextRange.Text.strip()
                except Exception:
                    pass
                titles.append(t.replace("\r", " ").replace("\v", " ")[:60])
                try:
                    hidden.append(bool(s.SlideShowTransition.Hidden))
                except Exception:
                    hidden.append(False)
            try:
                ratio = float(p.PageSetup.SlideWidth) / float(p.PageSetup.SlideHeight)
            except Exception:
                ratio = 16 / 9
            self.deck_cache[key] = {"count": n, "titles": titles, "hidden": hidden, "ratio": ratio}
        return self.deck_cache[key]

    def start(self, path, slide=None):
        key = norm(path)
        p = self._get_pres(path)
        info = self.load(path)
        old = self.live
        w = self._show_window(key)
        if w is None:
            ss = p.SlideShowSettings
            try:
                ss.RangeType = 1  # ppShowAll
            except Exception:
                pass
            w = ss.Run()
        if slide and 1 <= int(slide) <= info["count"]:
            if int(slide) != 1 or w.View.State == ppSlideShowDone:
                w.View.GotoSlide(int(slide))
        if w.View.State in (ppSlideShowBlackScreen, ppSlideShowWhiteScreen):
            w.View.State = ppSlideShowRunning
        try:
            w.Activate()
        except Exception:
            pass
        self.live = key
        if CFG.get("end_previous_show", True) and old and old != key:
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
        if action == "next":
            if v.State in (ppSlideShowBlackScreen, ppSlideShowWhiteScreen):
                v.State = ppSlideShowRunning
            else:
                v.Next()
        elif action == "prev":
            if v.State in (ppSlideShowBlackScreen, ppSlideShowWhiteScreen):
                v.State = ppSlideShowRunning
            else:
                v.Previous()
        elif action == "first":
            v.First()
        elif action == "last":
            v.Last()
        elif action == "goto":
            v.GotoSlide(int(slide))
            if v.State != ppSlideShowRunning:
                v.State = ppSlideShowRunning
        elif action == "black":
            v.State = ppSlideShowRunning if v.State == ppSlideShowBlackScreen else ppSlideShowBlackScreen
        elif action == "white":
            v.State = ppSlideShowRunning if v.State == ppSlideShowWhiteScreen else ppSlideShowWhiteScreen
        elif action == "end":
            v.Exit()
            self.live = None
        elif action == "focus":
            w.Activate()
        else:
            raise ValueError("未知操作：%s" % action)
        return True

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
        return True

    def export_thumb(self, path, n, out, width):
        p = self._get_pres(path)
        info = self.load(path)
        h = int(round(width / info["ratio"]))
        p.Slides(int(n)).Export(out, "PNG", int(width), h)
        self._mark_saved(p)
        return out

    def _notes(self, slide):
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
                    if sh.PlaceholderFormat.Type == ppPlaceholderBody and sh.HasTextFrame:
                        return sh.TextFrame.TextRange.Text.replace("\r", "\n").strip()
                except Exception:
                    pass
        except Exception:
            pass
        return ""

    def fetch_notes(self, key, idx):
        """
        读一页的演讲者备注（在空闲时做，不挡翻页）。
        有些 PPT 读备注特别慢（页面里有大图 / 视频），超过 0.8 秒就停用这份 PPT 的备注，保证翻页速度。
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
        if ms > 800:
            self.notes_slow.add(key)
            print("[备注] %s 第 %d 页读取备注花了 %dms，太慢，这份 PPT 停用备注显示（不影响放映）"
                  % (os.path.basename(key), idx, ms))
        elif ms > 200:
            print("[备注] %s 第 %d 页 %dms" % (os.path.basename(key), idx, ms))
        return True

    def status(self):
        T = [time.time()]
        marks = []

        def mark(name):
            T.append(time.time())
            marks.append("%s %dms" % (name, (T[-1] - T[-2]) * 1000))

        st = {"powerpoint": self._alive(), "live": None, "loaded": []}
        mark("连接")
        if not st["powerpoint"]:
            return st
        # 检查“哪些文件还开着”比较费时，每 2 秒做一次就够；页码每次都读
        now = time.time()
        if now - getattr(self, "_loaded_checked", 0) > 2:
            self._loaded_checked = now
            for k, p in list(self.pres.items()):
                try:
                    p.FullName
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
            if state != ppSlideShowDone:
                try:
                    idx = v.Slide.SlideIndex
                except Exception:
                    idx = None
            mark("页码")
            key = self.live
            # 备注不在这里读（太慢会拖住翻页），只取已经读好的；没读过的交给空闲时去读
            notes = self.notes_cache.get((key, idx), "") if idx else ""
            if idx and (key, idx) not in self.notes_cache and key not in self.notes_slow:
                self.notes_wanted = (key, idx)
            dc = self.deck_cache.get(key)
            count = dc["count"] if dc else w.Presentation.Slides.Count
            st["live"] = {"key": key, "slide": idx, "count": count, "state": state, "notes": notes,
                          "notes_off": key in self.notes_slow}
        total = (time.time() - T[0]) * 1000
        if total > 300 and time.time() - getattr(self, "_slow_logged", 0) > 5:
            self._slow_logged = time.time()
            print("[诊断] 读取放映状态花了 %dms：%s" % (total, "，".join(marks)))
        return st


class MockController(PPTController):
    """非 Windows 环境下的演示模式，方便测试界面。"""

    def __init__(self):
        super().__init__()
        self.state = {}
        self.cur = None

    def _alive(self):
        return True

    def load(self, path):
        key = norm(path)
        self.pres[key] = True
        n = 8 + (sum(map(ord, key)) % 20)
        return self.deck_cache.setdefault(key, {"count": n, "titles": ["第 %d 页" % i for i in range(1, n + 1)],
                                                "hidden": [False] * n, "ratio": 16 / 9})

    def start(self, path, slide=None):
        key = norm(path)
        self.load(path)
        self.live = key
        self.state = {"slide": int(slide or 1), "state": 1}
        return True

    def nav(self, action, slide=None):
        if not self.live:
            raise RuntimeError("目前没有正在放映的 PPT")
        n = self.deck_cache[self.live]["count"]
        s = self.state
        if action == "next":
            if s["state"] in (3, 4): s["state"] = 1
            elif s["slide"] >= n: s["state"] = 5
            else: s["slide"] += 1
        elif action == "prev":
            if s["state"] == 5: s["state"] = 1
            elif s["state"] in (3, 4): s["state"] = 1
            else: s["slide"] = max(1, s["slide"] - 1)
        elif action == "first": s.update(slide=1, state=1)
        elif action == "last": s.update(slide=n, state=1)
        elif action == "goto": s.update(slide=int(slide), state=1)
        elif action == "black": s["state"] = 1 if s["state"] == 3 else 3
        elif action == "white": s["state"] = 1 if s["state"] == 4 else 4
        elif action == "end": self.live = None
        return True

    def close(self, path):
        key = norm(path)
        self.pres.pop(key, None)
        if self.live == key:
            self.live = None
        return True

    def export_thumb(self, path, n, out, width):
        raise RuntimeError("演示模式没有缩略图")

    def status(self):
        st = {"powerpoint": True, "live": None, "loaded": list(self.pres.keys()), "mock": True}
        if self.live:
            s = self.state
            st["live"] = {"key": self.live, "slide": None if s["state"] == 5 else s["slide"],
                          "count": self.deck_cache[self.live]["count"], "state": s["state"],
                          "notes": "演示模式的备注：第 %d 页" % s["slide"]}
        return st


MOCK = sys.platform != "win32"
CTRL = MockController() if MOCK else PPTController()
WORKER = ComWorker(CTRL)


# ---------------------------------------------------------------- 视频 / 音频
class MediaPlayer:
    def __init__(self):
        self.proc = None
        self.name = None
        self._vlc = None
        self._vlc_at = 0

    def vlc(self):
        if time.time() - self._vlc_at > 30:
            self._vlc, self._vlc_at = self.find_vlc(), time.time()
        return self._vlc

    @staticmethod
    def find_vlc():
        cands = [shutil.which("vlc")]
        for env in ("ProgramFiles", "ProgramFiles(x86)"):
            base = os.environ.get(env)
            if base:
                cands.append(os.path.join(base, "VideoLAN", "VLC", "vlc.exe"))
        for c in cands:
            if c and os.path.exists(c):
                return c
        return None

    def play(self, path):
        self.stop()
        self.name = os.path.basename(path)
        vlc = self.vlc()
        if vlc:
            self.proc = subprocess.Popen([vlc, "--fullscreen", "--play-and-exit", "--no-video-title-show",
                                          "--no-qt-privacy-ask", path])
            return "vlc"
        if MOCK:
            self.proc = None
            return "mock"
        os.startfile(path)  # 用系统默认播放器
        return "default"

    def stop(self):
        if self.proc and self.proc.poll() is None:
            try:
                self.proc.terminate()
            except Exception:
                pass
        self.proc = None
        self.name = None

    def status(self):
        playing = bool(self.proc and self.proc.poll() is None)
        if not playing and self.proc:
            self.proc = None
            self.name = None
        return {"playing": playing, "name": self.name, "vlc": bool(self.vlc())}


MEDIA = MediaPlayer()


# ---------------------------------------------------------------- 预加载
PRELOAD = {"running": False, "done": 0, "total": 0, "current": None, "errors": []}


def preload_all():
    if PRELOAD["running"]:
        return
    with ITEMS_LOCK:
        targets = [it for it in ITEMS.values() if it["type"] == "ppt"]
    targets.sort(key=lambda it: natural_key(it["name"]))
    PRELOAD.update(running=True, done=0, total=len(targets), errors=[])
    WORKER.notify()

    def run():
        for it in targets:
            PRELOAD["current"] = it["name"]
            try:
                WORKER.call(CTRL.load, it["path"], priority=8, timeout=600, label="预加载 " + it["name"])
            except Exception as e:
                PRELOAD["errors"].append("%s：%s" % (it["name"], e))
            PRELOAD["done"] += 1
            WORKER.notify()
        PRELOAD.update(running=False, current=None)
        WORKER.notify()

    threading.Thread(target=run, daemon=True).start()


# ---------------------------------------------------------------- 缩略图
def thumb_path(it, n):
    return os.path.join(CACHE_DIR, "%s_%d" % (it["id"], it["mtime"]), "%d.png" % n)


PREFETCH_DONE = set()


def _export_if_missing(it, n):
    out = thumb_path(it, n)
    if os.path.exists(out):
        return False
    os.makedirs(os.path.dirname(out), exist_ok=True)
    t0 = time.time()
    CTRL.export_thumb(it["path"], n, out, int(CFG.get("thumb_width", 480)))
    print("[缩略图] 预先生成 %s 第 %d 页  %dms" % (it["name"], n, (time.time() - t0) * 1000))
    WORKER.notify()
    return True


def idle_prefetch_live():
    """放映中：只预先生成当前页和后面两页的缩略图，翻页时网页上的画面就能立刻换。"""
    if MOCK or not CFG.get("prefetch_thumbs", True) or not CTRL.live:
        return False
    live = WORKER.snapshot.get("live")
    if not live or not live.get("slide") or live.get("state") == ppSlideShowDone:
        return False
    with ITEMS_LOCK:
        it = next((x for x in ITEMS.values() if norm(x["path"]) == live.get("key")), None)
    if not it:
        return False
    cur, count = live["slide"], live["count"]
    for n in (cur, cur + 1, cur + 2, cur - 1, cur + 3):
        if 1 <= n <= count and _export_if_missing(it, n):
            return True
    return False


def idle_prefetch_thumbs():
    """PowerPoint 空闲、而且没有在放映时，把已加载 PPT 的缩略图一张一张预先生成。"""
    if MOCK or not CFG.get("prefetch_thumbs", True) or CTRL.live or not CTRL.pres:
        return False
    with ITEMS_LOCK:
        items = [it for it in ITEMS.values() if it["type"] == "ppt"]
    for it in items:
        tag = (it["id"], it["mtime"])
        key = norm(it["path"])
        if tag in PREFETCH_DONE or key not in CTRL.pres or key not in CTRL.deck_cache:
            continue
        for n in range(1, CTRL.deck_cache[key]["count"] + 1):
            if _export_if_missing(it, n):
                return True   # 每次只做一张，马上回去看有没有控制命令
        PREFETCH_DONE.add(tag)
    return False


_LAST_MARK = [0.0]


def idle_mark_saved():
    """每 3 秒把控制台打开的 PPT 标记为“已保存”，这样任何时候关闭 PowerPoint 都不会弹出保存提示。"""
    if MOCK or not CTRL.pres or time.time() - _LAST_MARK[0] < 3:
        return False
    _LAST_MARK[0] = time.time()
    CTRL.mark_saved_all()
    return False   # 很快，不占用这一轮空闲


def idle_fetch_notes():
    """空闲时读当前页的备注；读好了马上推给界面。"""
    want = getattr(CTRL, "notes_wanted", None)
    if MOCK or not want:
        return False
    CTRL.notes_wanted = None
    if want in CTRL.notes_cache:
        return False
    CTRL.fetch_notes(*want)
    WORKER.snapshot_at = 0      # 让下一轮立刻刷新状态，把备注带出去
    return True


def idle_prefetch_notes():
    """没在放映时，把已加载 PPT 的备注一页一页先读好（读得慢的 PPT 会自动停用备注）。"""
    if MOCK or CTRL.live or not CTRL.pres:
        return False
    for key in list(CTRL.pres.keys()):
        dc = CTRL.deck_cache.get(key)
        if not dc or key in CTRL.notes_slow:
            continue
        for n in range(1, dc["count"] + 1):
            if (key, n) not in CTRL.notes_cache:
                CTRL.fetch_notes(key, n)
                return True
    return False


WORKER.idle_jobs.append(idle_fetch_notes)
WORKER.idle_jobs.append(idle_mark_saved)
WORKER.idle_jobs.append(idle_prefetch_live)
WORKER.idle_jobs.append(idle_prefetch_notes)
WORKER.idle_jobs.append(idle_prefetch_thumbs)
WORKER.start()


# ---------------------------------------------------------------- Web
app = Flask(__name__, static_folder=None)


def err(msg, code=400):
    if code >= 500 and not isinstance(msg, (RuntimeError, ValueError, TimeoutError)):
        print("[错误] %s %s → %s" % (request.method, request.path, msg))
    return jsonify({"ok": False, "error": str(msg)}), code


@app.route("/")
def index():
    return send_from_directory(os.path.join(RES, "static"), "index.html")


@app.route("/static/<path:p>")
def static_files(p):
    return send_from_directory(os.path.join(RES, "static"), p)


def public_item(it, loaded_keys, live_key):
    key = norm(it["path"])
    d = {k: it[k] for k in ("id", "name", "title", "num", "type", "size")}
    d["loaded"] = key in loaded_keys
    d["live"] = key == live_key
    return d


@app.route("/api/items")
def api_items():
    if request.args.get("refresh"):
        refresh_items()
    with ITEMS_LOCK:
        lst = sorted(ITEMS.values(), key=lambda it: (0 if it["num"] else 1, natural_key(it["name"])))
    return jsonify({"ok": True, "version": VERSION, "folder": CFG["folder"], "folder_exists": os.path.isdir(CFG["folder"]),
                    "items": [public_item(it, set(), None) for it in lst]})


def build_status():
    # 只读快照，不排队等 PowerPoint —— 即使 PowerPoint 正在打开大文件，界面也不会卡
    st = copy.deepcopy(WORKER.snapshot)
    st["busy"] = WORKER.busy
    st["busy_for"] = int(time.time() - WORKER.busy_since) if WORKER.busy_since else 0
    with ITEMS_LOCK:
        by_key = {norm(it["path"]): it["id"] for it in ITEMS.values()}
    live = st.get("live")
    if live:
        live["id"] = by_key.get(live.pop("key", None))
        # 告诉网页哪些页的缩略图已经做好，可以直接显示
        it = ITEMS.get(live["id"]) if live["id"] else None
        if it and live.get("slide"):
            live["ready"] = [n for n in range(live["slide"] - 1, live["slide"] + 4)
                             if 1 <= n <= live["count"] and os.path.exists(thumb_path(it, n))]
    st["loaded"] = [by_key[k] for k in st.get("loaded", []) if k in by_key]
    st["preload"] = dict(PRELOAD, errors=list(PRELOAD["errors"]))
    st["media"] = MEDIA.status()
    st["mock"] = MOCK
    st["standby"] = CFG.get("standby") or ""
    st["auto_advance"] = bool(CFG.get("auto_advance", True))
    st.pop("error", None) if not st.get("error") else None
    st["v"] = hashlib.md5(json.dumps(st, sort_keys=True, ensure_ascii=False).encode("utf-8")).hexdigest()[:10]
    return st


@app.route("/api/status")
def api_status():
    """
    长轮询：网页带上上次拿到的版本号 v，状态没变就在这里等（最多 20 秒），
    一变化（翻页、黑屏、开始放映……）马上返回。比定时查询快，也几乎不占资源。
    """
    since = request.args.get("v")
    st = build_status()
    if since and st["v"] == since:
        deadline = time.time() + 20
        while time.time() < deadline:
            with WORKER.changed:
                WORKER.changed.wait(0.5)   # 0.5 秒兜底，顺便覆盖视频结束、预加载进度这类变化
            st = build_status()
            if st["v"] != since:
                break
    resp = jsonify(st)
    resp.headers["Cache-Control"] = "no-store"
    return resp


@app.route("/api/deck/<iid>")
def api_deck(iid):
    it = get_item(iid)
    if it["type"] != "ppt":
        return err("不是 PPT")
    try:
        info = WORKER.call(CTRL.load, it["path"], priority=2, timeout=600, label="打开 " + it["name"])
    except Exception as e:
        return err(e, 500)
    return jsonify({"ok": True, "id": iid, **info})


START_LOCK = threading.Lock()


def ordered_items():
    with ITEMS_LOCK:
        return sorted(ITEMS.values(), key=lambda it: (0 if it["num"] else 1, natural_key(it["name"])))


def next_item_id(iid):
    lst = ordered_items()
    for k, it in enumerate(lst):
        if it["id"] == iid:
            return lst[k + 1]["id"] if k + 1 < len(lst) else ""
    return ""


def set_standby(iid, why=""):
    iid = iid or ""
    with ITEMS_LOCK:
        if iid and iid not in ITEMS:
            iid = ""
        name = ITEMS[iid]["name"] if iid else "（节目单结束）"
    if CFG.get("standby") != iid:
        CFG["standby"] = iid
        save_config(CFG)
        print("[待命] %s%s" % (name, "（%s）" % why if why else ""))
    WORKER.notify()


def start_item(iid, slide=1, why=""):
    """开始放映 / 播放一个节目，并把它后面的节目设为待命。"""
    it = get_item(iid)
    if it["type"] == "ppt":
        # 连按好几次只执行一次
        if not START_LOCK.acquire(blocking=False):
            raise RuntimeError("正在切换中，请稍等")
        try:
            WORKER.call(CTRL.start, it["path"], slide, priority=0, timeout=600,
                        label=("%s开始放映 " % why) + it["name"])
        finally:
            START_LOCK.release()
        how = "ppt"
    else:
        print("[视频] %s播放 %s" % (why, it["name"]))
        how = MEDIA.play(it["path"])
    set_standby(next_item_id(iid))
    WORKER.notify()
    return how


def go(why=""):
    sb = CFG.get("standby")
    if not sb:
        raise RuntimeError("没有待命的节目（节目单已经播完）")
    return start_item(sb, 1, why)


@app.route("/api/start", methods=["POST"])
def api_start():
    data = request.get_json(force=True, silent=True) or {}
    try:
        how = start_item(data.get("id"), data.get("slide") or 1)
        return jsonify({"ok": True, "player": how})
    except RuntimeError as e:
        return err(e, 409)
    except Exception as e:
        return err(e, 500)


@app.route("/api/go", methods=["POST"])
def api_go():
    try:
        return jsonify({"ok": True, "player": go()})
    except RuntimeError as e:
        return err(e, 409)
    except Exception as e:
        return err(e, 500)


@app.route("/api/standby", methods=["POST"])
def api_standby():
    data = request.get_json(force=True, silent=True) or {}
    set_standby(data.get("id"), "手动设定")
    return jsonify({"ok": True})


# ---------------------------------------------------------------- 自动接下一个节目
_ADVANCED = {"key": None, "at": 0.0}


def on_live_change(old, new):
    """
    放映到了结尾（PowerPoint 的“放映结束”黑屏）→ 自动开始待命的节目。
    不管是控制台按钮、快捷键，还是翻页笔直接在 PowerPoint 上按的“下一页”，都会走到这里。
    """
    n = (new or {}).get("live") or {}
    if n.get("state") != ppSlideShowDone or not CFG.get("auto_advance", True):
        return
    if not CFG.get("standby"):
        return                       # 节目单播完了，就停在结尾黑屏
    key = n.get("key")
    if _ADVANCED["key"] == key and time.time() - _ADVANCED["at"] < 10:
        return                       # 已经在切换了，不要重复触发
    _ADVANCED.update(key=key, at=time.time())

    def run():
        try:
            go("自动接续：")
        except Exception as e:
            print("[自动接续] 没有成功：%s" % e)

    threading.Thread(target=run, daemon=True).start()   # 不能在 PowerPoint 线程里等自己


@app.route("/api/nav", methods=["POST"])
def api_nav():
    data = request.get_json(force=True, silent=True) or {}
    try:
        WORKER.call(CTRL.nav, data.get("action"), data.get("slide"), priority=0, timeout=45)
        return jsonify({"ok": True})
    except Exception as e:
        return err(e, 500)


@app.route("/api/close", methods=["POST"])
def api_close():
    data = request.get_json(force=True, silent=True) or {}
    it = get_item(data.get("id"))
    try:
        WORKER.call(CTRL.close, it["path"], priority=0, timeout=60)
        return jsonify({"ok": True})
    except Exception as e:
        return err(e, 500)


@app.route("/api/media/stop", methods=["POST"])
def api_media_stop():
    MEDIA.stop()
    WORKER.notify()
    return jsonify({"ok": True})


@app.route("/api/preload", methods=["POST"])
def api_preload():
    preload_all()
    return jsonify({"ok": True})


@app.route("/api/thumb/<iid>/<int:n>")
def api_thumb(iid, n):
    it = get_item(iid)
    if it["type"] != "ppt":
        abort(404)
    out = thumb_path(it, n)
    if not os.path.exists(out):
        os.makedirs(os.path.dirname(out), exist_ok=True)
        try:
            WORKER.call(CTRL.export_thumb, it["path"], n, out, int(CFG.get("thumb_width", 480)),
                        priority=5, timeout=120)
        except Exception:
            abort(404)
        if not os.path.exists(out):
            abort(404)
    resp = send_file(out, mimetype="image/png")
    resp.headers["Cache-Control"] = "max-age=86400"
    return resp


@app.route("/api/config", methods=["GET", "POST"])
def api_config():
    if request.method == "POST":
        data = request.get_json(force=True, silent=True) or {}
        if "folder" in data:
            f = os.path.expandvars(os.path.expanduser(str(data["folder"]).strip().strip('"')))
            if not os.path.isdir(f):
                return err("文件夹不存在：%s" % f)
            CFG["folder"] = os.path.abspath(f)
        if "end_previous_show" in data:
            CFG["end_previous_show"] = bool(data["end_previous_show"])
        if "auto_advance" in data:
            CFG["auto_advance"] = bool(data["auto_advance"])
        if isinstance(data.get("hotkeys"), dict):
            hk = dict(CFG["hotkeys"])
            for k, v in data["hotkeys"].items():
                if k in desktop.ACTIONS:
                    v = str(v or "").strip()
                    try:
                        desktop.parse_hotkey(v)
                    except ValueError as e:
                        return err("%s：%s" % (desktop.ACTIONS[k], e))
                    hk[k] = v
            CFG["hotkeys"] = hk
        save_config(CFG)
        refresh_items()
        if "hotkeys" in data:
            HOTKEYS.reload()
    return jsonify({"ok": True, "config": CFG, "actions": list(desktop.ACTIONS.items()), "hotkeys": HOTKEYS.status()})


@app.route("/api/debug", methods=["GET", "POST"])
def api_debug():
    """打开 / 关闭调试窗口。"""
    if request.method == "POST":
        data = request.get_json(force=True, silent=True) or {}
        on = bool(data.get("on"))
        DEBUG.open() if on else DEBUG.close()
        CFG["debug_console"] = on
        save_config(CFG)
        print("[调试] 调试窗口%s" % ("打开" if on else "关闭"))
    return jsonify({"ok": True, "open": DEBUG.is_open(), "available": desktop.IS_WIN, "log": LOG_PATH})


@app.route("/api/window", methods=["GET", "POST"])
def api_window():
    """控制台窗口置顶。直接用 Windows API，不经过 pywebview（避免卡死）。"""
    if request.method == "POST":
        data = request.get_json(force=True, silent=True) or {}
        on = bool(data.get("on_top"))
        ok = desktop.set_on_top(on)
        CFG["window_on_top"] = on
        save_config(CFG)
        print("[窗口] 置顶%s%s" % ("开" if on else "关", "" if ok else "（找不到窗口）"))
    return jsonify({"ok": True, "desktop": desktop.WINDOW_OPEN,
                    "on_top": desktop.is_on_top() if desktop.WINDOW_OPEN else bool(CFG.get("window_on_top"))})


@app.route("/api/hotkeys/suspend", methods=["POST"])
def api_hotkeys_suspend():
    """设置快捷键时先暂停全局快捷键，否则按下去会直接被触发、录不进去。"""
    data = request.get_json(force=True, silent=True) or {}
    HOTKEYS.suspend(bool(data.get("on")))
    return jsonify({"ok": True, "hotkeys": HOTKEYS.status()})


# ---------------------------------------------------------------- 全局快捷键
def on_hotkey(action):
    print("[快捷键] %s" % desktop.ACTIONS.get(action, action))
    if action == "show":
        desktop.bring_to_front()
        return
    if action == "go":
        try:
            go("快捷键：")
        except Exception as e:
            print("[快捷键] GO：%s" % e)
        return
    # 翻页、黑屏直接在后台执行，最快，也不依赖网页
    try:
        WORKER.call(CTRL.nav, action, priority=0, timeout=45)
    except Exception as e:
        print("快捷键 %s：%s" % (action, e))


HOTKEYS = desktop.HotkeyManager(lambda: CFG.get("hotkeys"), on_hotkey)


# ---------------------------------------------------------------- 启动
def pick_port(want):
    want = int(want or 0)
    if want:
        try:
            with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as t:
                t.bind(("127.0.0.1", want))
            return want
        except OSError:
            print("端口 %d 被占用，改用自动端口" % want)
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as t:
        t.bind(("127.0.0.1", 0))
        return t.getsockname()[1]


def serve_forever(port):
    try:
        from waitress import serve
    except ImportError:
        serve = None
    if serve:
        logging.getLogger("waitress").setLevel(logging.ERROR)
        serve(app, host="127.0.0.1", port=port, threads=8, ident=None, channel_timeout=60)
    else:
        app.run(host="127.0.0.1", port=port, threaded=True, debug=False, use_reloader=False)


def wait_until_up(port, timeout=10):
    end = time.time() + timeout
    while time.time() < end:
        try:
            with socket.create_connection(("127.0.0.1", port), timeout=0.3):
                return True
        except OSError:
            time.sleep(0.05)
    return False


def main():
    if not desktop.ensure_single_instance():
        return
    refresh_items()
    port = pick_port(CFG.get("port", 0))
    url = "http://127.0.0.1:%d/" % port
    print("PPT 控制台 v%s%s | 文件夹：%s | %s" % (VERSION, "（演示模式）" if MOCK else "", CFG["folder"], url))

    threading.Thread(target=serve_forever, args=(port,), daemon=True).start()
    if not wait_until_up(port):
        desktop.message_box("控制台启动失败（网页服务没有起来），详情见 ppt-presenter.log")
        return
    HOTKEYS.start()
    if CFG.get("debug_console") or "--debug" in sys.argv:
        DEBUG.open()
    desktop.start_ui_watchdog()
    HOTKEYS.ready.wait(2)
    print("全局快捷键：%s" % ("、".join("%s=%s" % (desktop.ACTIONS[a], k) for a, k in HOTKEYS.registered.items()) or "无"))
    if CFG.get("preload_on_start"):
        preload_all()

    if desktop.webview_available() and not MOCK:
        desktop.run_window(url, CFG, save_config, os.path.join(BASE, ".cache", "webview"))
        # 窗口关了 = 程序结束（PowerPoint 不会被关掉；调试窗口发现控制台结束后也会自己关掉）
        print("控制台窗口已关闭，程序结束")
        try:
            WORKER.call(CTRL.mark_saved_all, priority=0, timeout=5)   # 之后关 PowerPoint 也不会问要不要保存
        except Exception:
            pass
        os._exit(0)
    else:
        # 开发 / 测试环境：没有 pywebview 时退回浏览器
        print("未安装 pywebview，改用浏览器打开：", url)
        if not MOCK:
            webbrowser.open(url)
        try:
            while True:
                time.sleep(3600)
        except KeyboardInterrupt:
            pass


if __name__ == "__main__":
    try:
        main()
    except Exception:
        traceback.print_exc()
        desktop.message_box("PPT 控制台出错了：\n\n%s\n详情见 ppt-presenter.log" % traceback.format_exc(limit=3))
