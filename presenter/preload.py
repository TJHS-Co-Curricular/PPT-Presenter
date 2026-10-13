# -*- coding: utf-8 -*-
"""
完整预加载（建议典礼开始前做一次）：

    第 1 步  打开全部 PPT（只读、带窗口并最小化），检查字体
    第 2 步  逐页读取演讲者备注、生成缩略图，全部存进 .cache

做完以后放映时就不用再临时读取，翻页延迟最低；.cache 里已有的会直接跳过。
期间禁止一切放映操作（界面被遮住，后台也拒绝命令，全局快捷键暂停），可以随时取消。
"""
import os
import threading
import time
import traceback

from presenter.cache import thumb_path
from presenter.config import CFG
from presenter.core import CTRL, MOCK, WORKER
from presenter.fonts import missing_fonts
from presenter.library import norm, ppt_items
from presenter.log import log

#: 预加载状态（会原样发给界面）
PRELOAD = {"running": False, "phase": 0, "phase_name": "", "done": 0, "total": 0, "current": None, "detail": "",
           "errors": [], "cancel": False, "started": 0, "finished_at": 0, "summary": "", "stats": {},
           "font_warnings": []}    # [{"name": PPT 文件名, "fonts": [没安装也没内嵌的字体]}]


class PreloadRunning(RuntimeError):
    pass


def ensure_not_preloading():
    """所有会操作 PowerPoint 的入口都要先调用它（预加载时 PowerPoint 同时做两件事，两边都会很慢）。"""
    if PRELOAD["running"]:
        raise PreloadRunning("正在预加载，请等完成（或先取消预加载）再操作")


def cancel():
    if PRELOAD["running"]:
        PRELOAD["cancel"] = True
        log("预加载", "用户取消，做完当前这一页就停")
        WORKER.notify()


def _new_stats(decks):
    return {"decks": decks, "slides": 0, "thumbs_new": 0, "thumbs_cached": 0,
            "notes_new": 0, "notes_cached": 0, "notes_off": 0}


def _make_prep_slide(stats):
    def prep_slide(item, key, n):
        """一页的准备工作：读备注 + 生成缩略图（在 PowerPoint 线程里执行）。"""
        if MOCK:
            time.sleep(0.01)
            return
        if key not in CTRL.notes_slow:
            if (key, n) in CTRL.notes_cache:
                stats["notes_cached"] += 1
            else:
                CTRL.fetch_notes(key, n)
                stats["notes_new"] += 1
        out = thumb_path(item, n)
        if os.path.exists(out):
            stats["thumbs_cached"] += 1
        else:
            os.makedirs(os.path.dirname(out), exist_ok=True)
            CTRL.export_thumb(item["path"], n, out, int(CFG.get("thumb_width", 480)))
            stats["thumbs_new"] += 1

    prep_slide._quiet_under_ms = 1000   # 日志只记慢的页
    return prep_slide


def _summary(stats):
    took = time.time() - PRELOAD["started"]
    text = ("已取消。" if PRELOAD["cancel"] else "完成！") + \
        "%d 份 PPT、%d 页；新生成缩略图 %d 张（已有 %d 张），新读备注 %d 页（已有 %d 页）%s；用时 %d 分 %02d 秒。" % (
            stats["decks"], stats["slides"], stats["thumbs_new"], stats["thumbs_cached"],
            stats["notes_new"], stats["notes_cached"],
            "，%d 份 PPT 备注太慢已停用" % stats["notes_off"] if stats["notes_off"] else "",
            took // 60, took % 60)
    if PRELOAD["font_warnings"]:
        text += " ⚠ %d 份 PPT 用了没安装的字体（见下方）。" % len(PRELOAD["font_warnings"])
    if PRELOAD["errors"]:
        text += " 有 %d 个问题，详情见日志。" % len(PRELOAD["errors"])
    return text


def _open_all(targets):
    for item in targets:
        if PRELOAD["cancel"]:
            return
        PRELOAD.update(current=item["name"], detail="打开中")
        WORKER.notify()
        try:
            WORKER.call(CTRL.load, item["path"], priority=8, timeout=600, label="预加载 打开 " + item["name"])
            _check_fonts(item)
        except Exception as e:
            PRELOAD["errors"].append("%s：打开失败 %s" % (item["name"], e))
        PRELOAD["done"] += 1


def _check_fonts(item):
    """没安装也没内嵌的字体，放映时会被替换成别的字体。"""
    try:
        used = WORKER.call(CTRL.fonts, item["path"], priority=8, timeout=60)
    except Exception as e:
        log("提示", "%s：无法读取字体列表 %s" % (item["name"], e))
        return
    missing = missing_fonts(used)
    embedded = [n for n, emb in used if emb]
    if embedded:
        log("预加载", "%s：内嵌字体 %s" % (item["name"], "、".join(embedded)))
    if missing:
        PRELOAD["font_warnings"].append({"name": item["name"], "fonts": missing})
        log("预加载", "⚠ %s 用了这台电脑没安装、也没内嵌的字体：%s（放映时会被替换）"
            % (item["name"], "、".join(missing)))


def _prepare_slides(targets, stats):
    jobs = []
    for item in targets:
        deck = CTRL.deck_cache.get(norm(item["path"]))
        if deck:
            stats["slides"] += deck["count"]
            jobs += [(item, n, deck["count"]) for n in range(1, deck["count"] + 1)]
    PRELOAD.update(phase=2, phase_name="生成缩略图和备注", done=0, total=len(jobs))
    WORKER.notify()
    prep_slide = _make_prep_slide(stats)
    last_notify = deck_t0 = 0.0
    for item, n, count in jobs:
        if PRELOAD["cancel"]:
            break
        PRELOAD.update(current=item["name"], detail="第 %d / %d 页" % (n, count))
        key = norm(item["path"])
        if n == 1:
            deck_t0 = time.time()
        try:
            WORKER.call(prep_slide, item, key, n, priority=8, timeout=120)
        except Exception as e:
            PRELOAD["errors"].append("%s 第 %d 页：%s" % (item["name"], n, e))
        PRELOAD["done"] += 1
        if n == count:
            CTRL.save_notes()          # 每做完一份就把备注写进 .cache
            log("预加载", "%s：%d 页完成，%.1f 秒%s" % (item["name"], count, time.time() - deck_t0,
                                                   "（备注太慢已停用）" if key in CTRL.notes_slow else ""))
        if time.time() - last_notify > 0.2:
            last_notify = time.time()
            WORKER.notify()
    CTRL.save_notes()


def start_preload():
    """在后台线程开始预加载；已经在跑就什么都不做。"""
    if PRELOAD["running"]:
        return
    targets = ppt_items()
    stats = _new_stats(len(targets))
    PRELOAD.update(running=True, phase=1, phase_name="打开 PPT", done=0, total=len(targets), current=None,
                   detail="", errors=[], cancel=False, started=time.time(), finished_at=0, summary="", stats=stats,
                   font_warnings=[])
    log("预加载", "开始：%d 份 PPT" % len(targets))
    WORKER.notify()

    def run():
        try:
            _open_all(targets)
            if not PRELOAD["cancel"]:
                _prepare_slides(targets, stats)
        except Exception as e:
            traceback.print_exc()
            PRELOAD["errors"].append(str(e))
        finally:
            stats["notes_off"] = len(CTRL.notes_slow)
            summary = _summary(stats)
            log("预加载", summary)
            for e in PRELOAD["errors"][:20]:
                log("预加载", "问题：" + e)
            PRELOAD.update(running=False, current=None, detail="", finished_at=time.time(), summary=summary)
            WORKER.notify()

    threading.Thread(target=run, daemon=True, name="preload").start()
