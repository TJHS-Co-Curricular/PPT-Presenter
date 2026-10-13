# -*- coding: utf-8 -*-
"""
PPT 控制台 —— 入口。

    python server.py            正常启动（Windows + PowerPoint；其他系统是演示模式）
    python server.py --debug    同时打开调试窗口

代码都在 presenter/ 包里，见 CLAUDE.md。
"""
import sys
import traceback


def main():
    # 调试窗口模式：这个进程只负责在自己的命令行窗口里实时显示日志
    if len(sys.argv) >= 4 and sys.argv[1] == "--debug-console":
        from presenter.desktop.debug_console import run_debug_console
        run_debug_console(sys.argv[2], int(sys.argv[3]))
        return

    # 必须最先做：之后所有 print 和报错都写进日志（打包成无控制台的 exe 时，不这样做会直接丢失）
    from presenter import paths
    from presenter.desktop.logs import setup_logging
    setup_logging(paths.LOG_PATH, paths.FROZEN)

    from presenter.desktop.common import message_box
    try:
        from presenter.app import run
        run()
    except Exception:
        traceback.print_exc()
        message_box("PPT 控制台出错了：\n\n%s\n详情见 ppt-presenter.log" % traceback.format_exc(limit=3))


if __name__ == "__main__":
    main()
