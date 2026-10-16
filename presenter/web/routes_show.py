# -*- coding: utf-8 -*-
"""放映、GO、待命、翻页、关闭、视频。"""
from flask import Blueprint

from presenter import show
from presenter.core import WORKER
from presenter.media import MEDIA
from presenter.web.common import body, guarded, ok

bp = Blueprint("show", __name__)


@bp.route("/api/start", methods=["POST"])
def api_start():
    data = body()
    return guarded(lambda: ok(player=show.start_item(data.get("id"), data.get("slide") or 1)))


@bp.route("/api/go", methods=["POST"])
def api_go():
    return guarded(lambda: ok(player=show.go()))


@bp.route("/api/standby", methods=["POST"])
def api_standby():
    show.set_standby(body().get("id"), "手动设定")
    return ok()


@bp.route("/api/nav", methods=["POST"])
def api_nav():
    data = body()

    def run():
        show.nav(data.get("action"), data.get("slide"))
        return ok()

    return guarded(run)


@bp.route("/api/close", methods=["POST"])
def api_close():
    data = body()

    def run():
        show.close_item(data.get("id"))
        return ok()

    return guarded(run)


@bp.route("/api/base", methods=["POST"])
def api_base():
    """大会背景 / 礼成：{"action": "background" | "finale"}。"""
    action = body().get("action")

    def run():
        if action not in show.BASE_SLIDES:
            raise ValueError("未知操作：%s" % action)
        show.show_base(action)
        return ok()

    return guarded(run)


@bp.route("/api/close_current", methods=["POST"])
def api_close_current():
    """完全关闭刚播完 / 正在播的节目（编号 0 除外）。"""

    def run():
        show.close_current()
        return ok()

    return guarded(run)


@bp.route("/api/media/stop", methods=["POST"])
def api_media_stop():
    MEDIA.stop()
    WORKER.notify()
    return ok()
