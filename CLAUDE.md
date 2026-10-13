# CLAUDE.md — PPT 控制台（PPT-Presenter）

给在这个仓库里工作的 AI 助手 / 开发者看的说明。改代码前请先读完“必须遵守的规则”。

## 这是什么
Windows 桌面程序：在一个窗口里集中控制多份 PowerPoint 和视频，用于典礼、晚会按节目单逐个播放。
纯本机运行（只监听 127.0.0.1），用 PyInstaller 打包成单个 exe。界面文字、代码注释、日志一律用**简体中文**。

## 目录结构
```
server.py                  入口：调试窗口模式 / 接管日志 / 调用 presenter.app.run()
presenter/                 后端（按功能分模块，依赖只从下往上，没有循环引用）
  paths.py                 所有路径（exe 旁边的 config / .cache / 日志，打包后的资源目录）
  version.py               从 CHANGELOG.md 读版本号（python -m presenter.version）
  log.py                   log(标签, 内容) / log_exception
  formats.py               支持的文件类型
  config.py                config.json 读写、默认值、旧版迁移；全局字典 CFG
  library.py               扫描文件夹 → 节目单（编号排序）、ITEMS、get_item / next_item_id
  cache.py                 .cache 里缩略图 / 备注的文件位置
  powerpoint/
    constants.py           PowerPoint 常量（SHOW_DONE 等）和中文名称
    com_errors.py          COM 错误分类：忙碌（重试） / 断开（重连）
    events.py              PowerPoint 应用程序事件（翻页 / 动画点击 / 放映开始结束）→ controller.dirty
    controller.py          PPTController：打开、放映、翻页、缩略图、备注、读状态
    mock.py                MockController：非 Windows 的演示模式
    worker.py              ComWorker：唯一能调 PowerPoint 的线程（优先级队列、重试、快照、空闲任务）
  core.py                  共享实例 CTRL / WORKER / DEBUG
  fonts.py                 字体检查：没安装也没内嵌的字体（注册表 + GDI）
  media.py                 视频播放（VLC / 系统播放器）
  preload.py               完整预加载、ensure_not_preloading()
  show.py                  待命 / 开始 / GO / 翻页 / 关闭、播完自动接续（状态钩子）
  background.py            空闲任务：预取缩略图和备注、标记已保存、写备注缓存
  status.py                拼给界面的完整状态、长轮询 wait_for_change
  shortcuts.py             全局快捷键按下后做什么、HOTKEYS 实例
  web/
    __init__.py            create_app()，注册三组路由、404 处理
    common.py              body() / ok() / err() / guarded()
    routes_library.py      /  /static  /api/items  /api/deck  /api/thumb
    routes_show.py         /api/start  /api/go  /api/standby  /api/nav  /api/close  /api/media/stop
    routes_system.py       /api/status  /api/preload  /api/config  /api/hotkeys/suspend  /api/debug  /api/window
  app.py                   setup()（测试也用）/ run()（完整启动）
  desktop/                 Windows 桌面功能（非 Windows 自动跳过）
    common.py              IS_WIN、APP_TITLE、提示框、单实例
    hotkeys.py             ACTIONS / DEFAULT_HOTKEYS / parse_hotkey / HotkeyManager
    window.py              pywebview 窗口、Win32 置顶 / 叫到前面 / 记大小、界面卡住检测
    logs.py                Tee：每行加时间写进日志
    debug_console.py       调试窗口（独立进程，tail 日志并上色）
static/index.html          整个界面（HTML + CSS + JS 一个文件，无构建步骤）
build.bat / start.bat      打包 / 开发运行
CHANGELOG.md               更新记录，最上面的版本号就是程序版本
ruff.toml                  代码风格检查设置
```
运行时生成在 exe（或 server.py）旁边：`config.json`、`.cache/`（thumbs、notes、webview）、`ppt-presenter.log`。

## 想改什么，去哪里
| 想做的事 | 改哪里 |
|---|---|
| 新的放映操作（例如“跳到某节”） | `powerpoint/controller.py` 加方法 → `show.py` 包一层（记得 `ensure_not_preloading()`）→ `web/routes_show.py` 加接口 → 界面 |
| 新的设置项 | `config.py` 的 `DEFAULT_CONFIG` → `web/routes_system.py` 的 `_apply_config` → 界面设置对话框 |
| 新的快捷键动作 | `desktop/hotkeys.py` 的 `ACTIONS` / `DEFAULT_HOTKEYS` → `shortcuts.on_hotkey` |
| 新的后台预取 | `background.py` 写一个空闲任务并加进 `register()` |
| 状态里多一个字段 | 轻量的放 `controller.status()`，其他的放 `status.build_status()` |
| 新的文件类型 | `formats.py` |

## 线程
```
主线程        pywebview 窗口（WinForms + WebView2 界面线程），窗口关闭 → os._exit(0)
web（waitress 8 线程）  Flask 接口；/api/status 是长轮询（状态变化立即返回，最多等 20 秒）
powerpoint（ComWorker） 所有 COM 调用；优先级 0 放映翻页 / 1 状态 / 2 打开 / 5 缩略图 / 8 预加载；
                        空闲时处理 Windows 消息（PowerPoint 事件靠它送达）、按 poll_interval() 刷新快照
                        （事件通知立即刷新；放映中 1 秒兜底；没放映 0.5 秒），变化时调 change_hooks 并唤醒长轮询，再跑 idle_jobs
hotkeys       RegisterHotKey + GetMessage 循环
ui-watchdog   每 2 秒 SendMessageTimeout(WM_NULL) 检查窗口，卡住时把所有线程调用栈写进日志
preload       预加载流程（每一页是一个独立的 WORKER.call，方便取消）
auto-advance  播完自动接续（由状态钩子另开线程）
调试窗口      独立子进程（exe --debug-console <日志> <pid>）
```

## 必须遵守的规则（都是踩过的坑）
1. **所有 PowerPoint COM 调用只能在 ComWorker 线程里做**：通过 `WORKER.call(fn, ..., priority=, label=)`。
   **绝不能在 ComWorker 线程里调用 `WORKER.call`**（自己等自己 = 死锁）。`change_hooks` 和 `idle_jobs` 都在这个线程里运行，需要触发操作时另开线程（参考 `show.auto_advance`）。
2. **PowerPoint 忙碌 ≠ 关闭**（`powerpoint/com_errors.py`）：`RPC_E_CALL_REJECTED` / `RETRYLATER`，以及动态调度下的 `AttributeError: PowerPoint.Application.xxx` 都是“忙”，由 worker 等待重试（最多 30 秒）。只有 `HR_DEAD` 才 `reset()` 重连。当前 pywin32 **没有** `pythoncom.IID_IMessageFilter`，不要再尝试注册 COM 消息过滤器。
3. **绝不要从非界面线程调用 pywebview 的窗口方法**（`on_top`、`restore`、`evaluate_js`、读 `width` 等），会互相等待导致窗口“无响应”。窗口操作一律用 `desktop/window.py`：找窗口用 `InternalGetWindowText`（不能用 `GetWindowText`），`SetWindowPos` 必须带 `SWP_ASYNCWINDOWPOS`，恢复窗口用 `ShowWindowAsync`。需要界面执行的动作走状态通道或接口，不要用 `evaluate_js` / `js_api`。
4. **放映中尽量不打扰 PowerPoint**：每一次 COM 调用都会插队打断动画 / 平滑切换的渲染（v1.9 时放映中一次状态查询要 1.6 秒，动画因此不正常）。
   - 页码变化靠 `powerpoint/events.py` 的事件得知，不要改回高频轮询；`controller.status()` 只读页码、状态这类轻量属性，放映窗口要缓存（`_live_win`）。
   - 读演讲者备注会打开备注页，含大图 / 视频的页面要好几秒：备注只在空闲且没放映时读，超过 0.8 秒的 PPT 自动停用（`notes_slow`）。
   - **放映中不导出缩略图、不读备注**（`background._quiet_during_show()`、`/api/thumb` 放映中直接 404），都要靠预加载事先做好。`status()` 超过 300ms 会输出 `[诊断]`。
5. **PPT 一律只读、带窗口打开再最小化**（`controller._open()`：`Presentations.Open(path, True, False, True)` + `WindowState = 最小化`）。**不要改回无窗口**（WithWindow=False）：那样 PowerPoint 不载入内嵌字体，Morph 等效果也可能不正常。并且**随时标记为已保存**（`_mark_saved`：打开后、导出缩略图后、每 3 秒、关闭前、退出前），否则 PowerPoint 会弹出模态的 “Save your changes to this file?”，让所有 COM 调用卡住。
6. **缩略图渲染方式变了要提高 `cache.THUMB_RENDER`**，让旧缓存自动作废。
7. **预加载期间禁止操作**：`show.py` 里所有会操作 PowerPoint 的函数都先调用 `ensure_not_preloading()`；全局快捷键（除 `show`）直接忽略；空闲预取暂停。新增入口时也要加。
8. **全局快捷键**：F12 被 Windows 保留给调试器，无法注册（`parse_hotkey` 会拒绝）。默认 F8=GO、F9 上一页、F10 下一页、F11 黑屏、Ctrl+Alt+P 叫出窗口。翻页笔常用的 PageUp / PageDown / B / Esc 不要设成全局快捷键。
9. **入口必须最先 `setup_logging()`**：打包是无控制台模式，`sys.stdout` 为 None。`presenter.config` 在导入时就会读配置（可能打印），所以 `server.py` 先接管日志、再导入 `presenter.app`。
10. 窗口标题必须保持 `PPT 控制台`（`desktop.common.APP_TITLE`），找窗口和单实例都靠它。
11. `import presenter` 不能有副作用；共享实例集中在 `core.py`；`WORKER` 只在 `app.setup()` 里启动。

## 代码风格
- `ruff check .`（设置见 `ruff.toml`：120 字符一行，E / F / W / B / I 规则）必须通过。
- 每个模块开头写中文文档字符串，说明它负责什么；常量用大写放在模块顶部。
- 日志用 `presenter.log.log("标签", "内容")`，标签见 `log.py`；预期内的情况用 `[提示]`，意外错误用 `log_exception`。
- 接口统一用 `web/common.py`：`ok(...)` / `err(msg, code)` / `guarded(fn)`（预期内错误 409，意外错误 500，找不到项目 404）。
- 有意的 COM“探测”写成 `_ = obj.Property  # 探测：…`。

## 开发与测试
- Windows：`start.bat`（Python 直接运行，日志同时输出到命令行）。
- 非 Windows：`python server.py` 进入**演示模式**（MockController，没有 PowerPoint、窗口、快捷键），用浏览器打开日志里的地址。
- 脚本测试：`from presenter import app; app.setup()`，再用 `app.serve_forever(port)` 起服务，可以调用 `shortcuts.on_hotkey("go")`、`show.go()` 等（不要调用 `app.run()`，它会开窗口）。
- 改完至少检查：节目单排序、GO / 待命 / 自动接续、跳页、黑屏、预加载（锁定、取消、完成）、设置对话框（快捷键录制、F12 拒绝）、窄屏布局。

## 版本与发布
1. 每次修改都在 `CHANGELOG.md` **最上面**加新版本：`## [X.Y] - YYYY-MM-DD`，分“新增 / 改进 / 修复 / 移除”写清楚原因和做法。
2. 运行 `build.bat`，生成 `dist\PPT-Presenter-VX.Y.exe`（版本号来自 CHANGELOG，界面右上角也会显示）。
3. `presenter` 包会整个打进 exe（`--collect-submodules presenter`）；新增数据文件要在 `build.bat` 加 `--add-data`。
