# -*- coding: utf-8 -*-
"""节目单、页面信息、缩略图、界面文件。"""
import os

from flask import Blueprint, abort, request, send_file, send_from_directory

from presenter import paths
from presenter.cache import thumb_path
from presenter.config import CFG
from presenter.core import CTRL, MOCK, WORKER
from presenter.library import get_item, ordered_items, refresh_items
from presenter.version import VERSION
from presenter.web.common import err, ok

bp = Blueprint("library", __name__)

_PUBLIC_FIELDS = ("id", "name", "title", "num", "type", "size")


@bp.route("/")
def index():
    return send_from_directory(paths.STATIC_DIR, "index.html")


@bp.route("/static/<path:p>")
def static_files(p):
    return send_from_directory(paths.STATIC_DIR, p)


@bp.route("/api/items")
def api_items():
    if request.args.get("refresh"):
        refresh_items()
    return ok(version=VERSION, folder=CFG["folder"], folder_exists=os.path.isdir(CFG["folder"]),
              items=[{k: it[k] for k in _PUBLIC_FIELDS} for it in ordered_items()])


@bp.route("/api/deck/<iid>")
def api_deck(iid):
    item = get_item(iid)
    if item["type"] != "ppt":
        return err("不是 PPT")
    try:
        info = WORKER.call(CTRL.load, item["path"], priority=2, timeout=600, label="打开 " + item["name"])
    except Exception as e:
        return err(e, 500)
    return ok(id=iid, **info)


@bp.route("/api/thumb/<iid>/<int:n>")
def api_thumb(iid, n):
    item = get_item(iid)
    if item["type"] != "ppt" or MOCK:      # 演示模式没有缩略图，界面会显示页码占位
        abort(404)
    out = thumb_path(item, n)
    if not os.path.exists(out):
        if CTRL.live and not CFG.get("work_during_show", False):
            abort(404)                     # 放映中不临时导出（会打断动画），界面显示页码占位
        os.makedirs(os.path.dirname(out), exist_ok=True)
        try:
            WORKER.call(CTRL.export_thumb, item["path"], n, out, int(CFG.get("thumb_width", 480)),
                        priority=5, timeout=120)
        except Exception:
            abort(404)
        if not os.path.exists(out):
            abort(404)
    resp = send_file(out, mimetype="image/png")
    resp.headers["Cache-Control"] = "max-age=86400"
    return resp
