# -*- coding: utf-8 -*-
"""接口通用：读请求、统一的错误回复。"""
from flask import jsonify, request

from presenter.library import ItemNotFound
from presenter.log import log


def body():
    """请求的 JSON 内容（没有或格式不对时返回 {}）。"""
    return request.get_json(force=True, silent=True) or {}


def ok(**data):
    return jsonify(dict(ok=True, **data))


def err(msg, code=400):
    """
    统一的错误回复 {"ok": false, "error": "..."}，界面会把 error 显示成提示。
    预期内的情况（RuntimeError / ValueError / TimeoutError，例如“正在预加载”）用 4xx，不记日志；
    其他意外错误用 5xx 并记进日志。
    """
    if code >= 500 and not isinstance(msg, (RuntimeError, ValueError, TimeoutError)):
        log("错误", "%s %s → %s" % (request.method, request.path, msg))
    return jsonify({"ok": False, "error": str(msg)}), code


def guarded(fn):
    """执行 fn()：预期内的错误回 409，意外错误回 500。"""
    try:
        return fn()
    except ItemNotFound:
        raise                              # 交给 create_app 里的 404 处理
    except (RuntimeError, ValueError, TimeoutError) as e:
        return err(e, 409)
    except Exception as e:
        return err(e, 500)
