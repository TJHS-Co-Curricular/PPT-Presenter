# -*- coding: utf-8 -*-
"""
本机网页接口（只监听 127.0.0.1）。界面是 static/index.html，所有数据都走 /api。

    routes_library  节目单、页面信息、缩略图、界面文件
    routes_show     放映、GO、待命、翻页、关闭、视频
    routes_system   状态（长轮询）、预加载、设置、调试窗口、窗口置顶、快捷键
"""
from flask import Flask

from presenter.library import ItemNotFound
from presenter.web.common import err
from presenter.web.routes_library import bp as library_bp
from presenter.web.routes_show import bp as show_bp
from presenter.web.routes_system import bp as system_bp


def create_app():
    app = Flask(__name__, static_folder=None)
    app.register_blueprint(library_bp)
    app.register_blueprint(show_bp)
    app.register_blueprint(system_bp)

    @app.errorhandler(ItemNotFound)
    def _not_found(e):
        return err(e, 404)

    return app
