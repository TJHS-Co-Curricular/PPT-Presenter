# -*- coding: utf-8 -*-
"""
统一的日志输出。

所有日志都是一行 ``[标签] 内容``；时间戳由 desktop.logs.Tee 自动加上，并同时写进 ppt-presenter.log。
标签约定（调试窗口会按标签上色）：

    PowerPoint 操作及耗时    放映   页码 / 状态变化     待命      视频
    快捷键      缩略图       备注   窗口      调试      诊断      预加载    自动接续
    提示   预期内的情况（不打印 Traceback）
    错误   意外错误（调用 log_exception 附上 Traceback）
"""
import traceback


def log(tag, msg):
    print("[%s] %s" % (tag, msg))


def log_exception(tag, msg):
    """记录意外错误，并附上完整的 Traceback。"""
    print("[%s] %s" % (tag, msg))
    traceback.print_exc()
