# -*- coding: utf-8 -*-
"""
PPT 控制台（PPT-Presenter）后端。

模块按功能划分（依赖只从下往上，没有循环引用）：

    基础      paths · version · log · formats · config
    节目单    library · cache
    PowerPoint powerpoint/（constants · com_errors · controller · mock · worker）
    共享实例  core（CTRL / WORKER / DEBUG）
    业务      media · preload · show · background · status · shortcuts
    网页接口  web/（library · show · system 三组路由）
    桌面      desktop/（common · hotkeys · window · logs · debug_console）
    启动      app

这个文件故意保持为空：导入 presenter 不应产生任何副作用。
"""
