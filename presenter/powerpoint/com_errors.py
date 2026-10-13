# -*- coding: utf-8 -*-
"""
COM 错误分类。

忙碌：PowerPoint 正在打开大文件、开始放映、播放内嵌视频时会拒绝调用。等一下重试就好，
     **不是**关掉了（v1.2 之前把忙碌当成关闭，结果丢掉了所有已打开的文件）。
断开：PowerPoint 被关掉或崩溃了，需要重新连接。

注意：当前 pywin32 没有 pythoncom.IID_IMessageFilter，不能注册 COM 消息过滤器，只能手动重试。
"""

HR_BUSY = {
    -2147418111,   # RPC_E_CALL_REJECTED
    -2147417846,   # RPC_E_SERVERCALL_RETRYLATER
}
HR_DEAD = {
    -2147417848,   # RPC_E_DISCONNECTED
    -2147023174,   # RPC_S_SERVER_UNAVAILABLE
    -2147023170,   # RPC_S_CALL_FAILED
    -2147220995,   # CO_E_OBJNOTCONNECTED
}


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
