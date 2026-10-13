# -*- coding: utf-8 -*-
"""
PowerPoint 控制（COM）。

    constants   PowerPoint 的常量和中文名称
    com_errors  COM 错误分类：忙碌（等一下重试）/ 断开（重新连接）
    controller  PPTController：打开、放映、翻页、缩略图、备注、读取状态
    mock        MockController：非 Windows 下的演示模式
    worker      ComWorker：唯一能调用 PowerPoint 的线程
"""
