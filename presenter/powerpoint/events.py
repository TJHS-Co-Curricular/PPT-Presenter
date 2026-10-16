# -*- coding: utf-8 -*-
"""
PowerPoint 的应用程序事件：翻页、动画点击、放映开始 / 结束时，PowerPoint 主动通知我们。

有了事件就不用每 0.1 秒去问 PowerPoint “现在第几页”——那样的查询会插队打断动画渲染
（v1.9 日志里一次状态查询要 1.6 秒，平滑切换和进入动画因此不正常）。

事件处理函数只做一件事：把 controller.dirty 设为 True，PowerPoint 线程看到后马上刷新状态。
处理函数里不要再调用 COM（PowerPoint 正在等我们返回）。

事件靠 Windows 消息送达，所以 PowerPoint 线程必须经常 PumpWaitingMessages（worker.py 已经在做）。
如果注册失败（例如 WPS 没有类型库），会退回较快的定时查询，功能不受影响。
"""
from presenter.log import log


class AppEvents:
    """
    win32com.client.WithEvents 会用无参数的方式创建它，所以用类属性指向控制器。
    方法名和参数名（OnSlideShowNextSlide、Wn …）由 PowerPoint 规定，不能改。
    """
    controller = None

    def _touch(self):
        c = AppEvents.controller
        if c is not None:
            c.dirty = True
            c.events_seen += 1

    def OnSlideShowBegin(self, Wn):
        self._touch()

    def OnSlideShowNextSlide(self, Wn):
        self._touch()

    def OnSlideShowNextClick(self, Wn, nEffect):
        self._touch()

    def OnSlideShowEnd(self, Pres):
        self._touch()


def attach(app, controller):
    """给 PowerPoint 应用程序挂上事件。成功返回事件对象（要一直保留引用），失败返回 None。"""
    try:
        import win32com.client
        AppEvents.controller = controller
        handler = win32com.client.WithEvents(app, AppEvents)
        log("PowerPoint", "已连接翻页事件（放映中不再频繁查询，动画更流畅）")
        return handler
    except Exception as e:
        log("提示", "无法连接 PowerPoint 事件，改用定时查询：%s" % e)
        return None
