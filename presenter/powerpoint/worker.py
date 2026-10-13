# -*- coding: utf-8 -*-
"""
ComWorker：唯一能调用 PowerPoint 的线程。

PowerPoint 一次只能处理一个外部（COM）调用，多线程同时调用只会得到更多 “Call was rejected”，
所以所有操作都排进这个线程的优先级队列：

    0 放映 / 翻页 / 关闭    1 状态    2 打开文件    5 缩略图    8 预加载

另外它还负责：
- PowerPoint 忙碌时自动等待重试（最多 30 秒），断开时重新连接一次；
- 空闲时刷新状态快照（网页只读快照，不排队），状态变化时调用 change_hooks 并唤醒长轮询。
  刷新间隔由 controller.poll_interval() 决定：PowerPoint 事件通知有变化时立刻刷新，
  放映中只低频兜底（频繁查询会打断动画）；
- 空闲时运行 idle_jobs（每次只做一小件事，马上回来看有没有新命令）。

⚠ 绝不能在这个线程里调用 WORKER.call（自己等自己 = 死锁）。需要从状态变化触发操作时，另开线程。
"""
import itertools
import os
import queue
import sys
import threading
import time
import traceback

from presenter.log import log, log_exception
from presenter.powerpoint.com_errors import is_busy_error, is_dead_error
from presenter.powerpoint.constants import NAV_NAMES

BUSY_RETRY_SECONDS = 30
IDLE_POLL_SECONDS = 0.03     # 队列空时多久醒来一次（顺便处理 PowerPoint 事件消息）


def describe(fn, args, label=None):
    """把一个 PowerPoint 操作写成看得懂的一句话（给日志用）。"""
    if label:
        return label
    name = getattr(fn, "__name__", str(fn))

    def base(p):
        return os.path.basename(str(p))

    if name == "nav" and args:
        text = NAV_NAMES.get(args[0], str(args[0]))
        return text % (args[1] if len(args) > 1 else "") if "%s" in text else text
    if name == "export_thumb" and len(args) >= 2:
        return "[缩略图] %s 第 %s 页" % (base(args[0]), args[1])
    if name == "prep_slide" and len(args) >= 3:
        return "[预加载] %s 第 %s 页" % (args[0]["name"], args[2])
    if name in ("load", "start", "close") and args:
        return {"load": "打开", "start": "开始放映", "close": "关闭"}[name] + " " + base(args[0])
    return name


class ComWorker(threading.Thread):
    def __init__(self, ctrl):
        super().__init__(daemon=True, name="powerpoint")
        self.ctrl = ctrl
        self.q = queue.PriorityQueue()
        self.counter = itertools.count()
        self.busy = None             # 正在做的事（显示在界面上）
        self.busy_since = None
        self.retries = 0
        self.snapshot = {"powerpoint": False, "live": None, "loaded": []}
        self.snapshot_at = 0.0
        self.idle_jobs = []          # 空闲任务：返回 True 表示这一轮做了事
        self.change_hooks = []       # 状态变化时调用 hook(old, new)
        self.changed = threading.Condition()

    # ------------------------------------------------------------ 其他线程调用
    def notify(self):
        """唤醒所有在等状态变化的网页请求（长轮询）。"""
        with self.changed:
            self.changed.notify_all()

    def call(self, fn, *args, priority=0, timeout=180, label=None, **kw):
        """
        把 fn(*args) 排进 PowerPoint 线程执行，等它做完并返回结果（或抛出它的异常）。
        label：显示在界面上的“PowerPoint 忙碌：xxx”，也用于日志。
        """
        ev = threading.Event()
        box = {}
        self.q.put((priority, next(self.counter), fn, args, kw, ev, box, label))
        if not ev.wait(timeout):
            raise TimeoutError("PowerPoint 没有响应（可能有对话框挡住了，请到电脑上看看）")
        if "err" in box:
            raise box["err"]
        return box.get("res")

    def request_refresh(self):
        """让下一轮空闲时立刻刷新状态。"""
        self.snapshot_at = 0

    # ------------------------------------------------------------ 线程内部
    def _run_with_retry(self, fn, args, kw):
        deadline = time.time() + BUSY_RETRY_SECONDS
        revived = False
        self.retries = 0
        while True:
            try:
                return fn(*args, **kw)
            except Exception as e:
                if is_busy_error(e) and time.time() < deadline:
                    if self.retries == 0 and fn is not self.ctrl.status:
                        log("PowerPoint", "忙碌中，等待重试…（%s）" % describe(fn, args))
                    self.retries += 1
                    self._pump(0.25)
                    continue
                if is_dead_error(e) and not revived:
                    log("PowerPoint", "连接断开，重新连接…")
                    revived = True
                    self.ctrl.reset()
                    continue
                raise

    @staticmethod
    def _pump(seconds):
        if sys.platform == "win32":
            import pythoncom
            end = time.time() + seconds
            while time.time() < end:
                pythoncom.PumpWaitingMessages()
                time.sleep(0.02)
        else:
            time.sleep(seconds)

    @staticmethod
    def _pump_events():
        """处理排队的 Windows 消息，PowerPoint 的翻页事件才会送到 events.AppEvents。"""
        if sys.platform == "win32":
            import pythoncom
            pythoncom.PumpWaitingMessages()

    def _refresh_snapshot(self):
        try:
            snap = self._run_with_retry(self.ctrl.status, (), {})
        except Exception as e:
            snap = dict(self.snapshot)
            snap["error"] = str(e)
        self.snapshot_at = time.time()
        if snap != self.snapshot:
            old = self.snapshot
            for hook in self.change_hooks:
                try:
                    hook(old, snap)
                except Exception:
                    traceback.print_exc()
            self.snapshot = snap
            self.notify()

    def _run_idle(self):
        self._pump_events()
        if time.time() - self.snapshot_at >= self.ctrl.poll_interval():
            self._refresh_snapshot()
        for job in self.idle_jobs:
            try:
                if job():
                    break
            except Exception:
                pass

    def _run_job(self, job):
        _, _, fn, args, kw, ev, box, label = job
        self.busy, self.busy_since = label, time.time()
        if label:
            self.notify()
        t0 = time.time()
        try:
            box["res"] = self._run_with_retry(fn, args, kw)
            ms = (time.time() - t0) * 1000
            quiet = getattr(fn, "_quiet_under_ms", None)   # 批量小任务（例如预加载每一页）只记慢的
            if fn is not self.ctrl.status and (quiet is None or ms >= quiet or self.retries):
                log("PowerPoint", "%s  %dms%s" % (describe(fn, args, label), ms,
                                                 "（重试 %d 次）" % self.retries if self.retries else ""))
        except Exception as e:
            ms = (time.time() - t0) * 1000
            if isinstance(e, (RuntimeError, ValueError, TimeoutError)):
                log("提示", "%s：%s" % (describe(fn, args, label), e))   # 预期内的情况
            else:
                log_exception("错误", "%s 失败（%dms）：" % (describe(fn, args, label), ms))
            box["err"] = e
        finally:
            had_label = self.busy is not None
            self.busy = self.busy_since = None
            if had_label:
                self.notify()
        # 先更新状态快照再通知调用方，这样网页马上拿到的就是新的页码
        if fn is not self.ctrl.status:
            self._refresh_snapshot()
        ev.set()

    def run(self):
        if sys.platform == "win32":
            import pythoncom
            pythoncom.CoInitialize()
        while True:
            try:
                job = self.q.get(timeout=IDLE_POLL_SECONDS)
            except queue.Empty:
                self._run_idle()
                continue
            self._run_job(job)
