# CLAUDE.md — PPT 控制台（PPT-Presenter）

给在这个仓库里工作的 AI 助手 / 开发者看的说明。改代码前请先读完“必须遵守的规则”。

## 这是什么
Windows 桌面程序：在一个窗口里集中控制多份 PowerPoint 和视频，用于典礼、晚会按节目单逐个播放。
纯本机运行（只监听 127.0.0.1），用 PyInstaller 打包成单个 exe。界面文字、代码注释、日志一律用**简体中文**。

## 文件
| 文件 | 作用 |
|---|---|
| `server.py` | 入口。配置、文件扫描、COM 工作线程、PowerPoint 控制、Flask API、待命 / GO / 自动接续、缩略图与备注预取、启动流程 |
| `desktop.py` | 桌面相关：pywebview 窗口、全局快捷键（RegisterHotKey）、Win32 窗口操作、单实例、日志 Tee、调试窗口（独立进程）、界面卡住检测 |
| `version.py` | 从 `CHANGELOG.md` 第一个 `## [X.Y]` 标题读版本号 |
| `static/index.html` | 整个界面（HTML + CSS + JS 一个文件，无构建步骤，无外部依赖） |
| `build.bat` | 打包：读版本 → 装依赖 → PyInstaller → `dist\PPT-Presenter-VX.Y.exe` |
| `start.bat` | 开发时用 Python 直接运行（有命令行窗口） |
| `CHANGELOG.md` | 更新记录，**最上面的版本号就是程序版本** |

运行时生成在 exe（或 server.py）旁边：`config.json`、`.cache/`（缩略图、WebView2 数据）、`ppt-presenter.log`。

## 架构（线程）
```
主线程 ── pywebview 窗口（WinForms + WebView2 界面线程），窗口关闭 → os._exit(0)
waitress（8 线程） ── Flask API；/api/status 是长轮询（状态变化立即返回，最多等 20 秒）
ComWorker（1 线程，STA） ── 所有 PowerPoint COM 调用；优先级队列；空闲时刷新状态快照 + 跑 idle_jobs
HotkeyManager（1 线程） ── RegisterHotKey + GetMessage 循环
ui-watchdog ── 每 2 秒 SendMessageTimeout(WM_NULL) 检查窗口，卡住时把所有线程调用栈写进日志
调试窗口 ── 独立子进程（exe --debug-console <log> <pid>），像 tail -f 一样跟日志
```
- 状态：ComWorker 每 0.1 秒调用 `CTRL.status()` 生成快照；变化时 `WORKER.notify()` 唤醒长轮询；`build_status()` 再拼上预加载、视频、待命等信息并算出版本号 `v`。
- 优先级：0 = 放映 / 翻页 / 关闭，1 = 状态，2 = 打开文件，5 = 缩略图，8 = 预加载；`idle_jobs` 只在队列空时运行，每次只做一小件事。
- 待命（`CFG["standby"]`）由后台管理：`start_item()` 开始任何节目后，把下一个节目设为待命；`go()` 开始待命节目。
- 自动接续：`on_live_change()` 发现放映状态变成 `ppSlideShowDone`（结尾黑屏）且有待命 → 在**新线程**里调用 `go()`。

## 必须遵守的规则（都是踩过的坑）
1. **所有 PowerPoint COM 调用只能在 ComWorker 线程里做**：通过 `WORKER.call(fn, ..., priority=, label=)`。
   **绝不能在 ComWorker 线程里调用 `WORKER.call`**（自己等自己 = 死锁）。需要从状态变化触发操作时，另开线程（参考 `on_live_change`）。
2. **PowerPoint 忙碌 ≠ 关闭**：`RPC_E_CALL_REJECTED` / `RETRYLATER`，以及动态调度下的 `AttributeError: PowerPoint.Application.xxx` 都是“忙”，由 `_run_with_retry` 等待重试（最多 30 秒）。只有 `HR_DEAD` 里的错误码才 `reset()` 重连。
   当前 pywin32 **没有** `pythoncom.IID_IMessageFilter`，不要再尝试注册 COM 消息过滤器。
3. **绝不要从非界面线程调用 pywebview 的窗口方法**（`on_top`、`restore`、`evaluate_js`、读 `width` 等），会互相等待导致窗口“无响应”。窗口操作一律用 `desktop.py` 里的 Win32 函数：
   找窗口用 `InternalGetWindowText`（不能用 `GetWindowText`，它会给自己进程的窗口发消息并等待），`SetWindowPos` 必须带 `SWP_ASYNCWINDOWPOS`，恢复窗口用 `ShowWindowAsync`。
   需要网页执行的动作，走状态通道或 API，不要用 `evaluate_js`。也不要用 pywebview 的 `js_api`。
4. **状态读取路径（`status()`）必须很快**：只读页码、状态这类轻量属性。读演讲者备注会打开备注页，含大图 / 视频的页面要好几秒，所以备注由 `idle_fetch_notes` / `idle_prefetch_notes` 在空闲时读，超过 0.8 秒的 PPT 自动停用备注（`notes_slow`）。`status()` 超过 300ms 会输出 `[诊断]` 日志。
5. **PPT 一律只读、无窗口打开**（`Presentations.Open(path, True, False, False)`），并且**随时标记为已保存**（`_mark_saved`：打开后、导出缩略图后、每 3 秒、关闭前、退出前），否则 PowerPoint 会弹出 “Save your changes to this file?”。这个对话框是模态的，会让所有 COM 调用卡住。
6. **放映中不要做重活**：放映时只预生成当前页附近的缩略图（`idle_prefetch_live`），整份 PPT 的缩略图和备注预取只在没有放映时做。
7. **全局快捷键**：F12 被 Windows 保留给调试器，无法注册（`parse_hotkey` 会拒绝）。默认 F8=GO、F9 上一页、F10 下一页、F11 黑屏、Ctrl+Alt+P 叫出窗口。翻页笔常用的 PageUp / PageDown / B / Esc 不要设成全局快捷键。
8. **打包是无控制台模式**（`--windowed`）：`sys.stdout` 为 None，所以启动时就用 `desktop.setup_logging()` 把输出接到日志文件。任何启动期错误都要写日志并用 `desktop.message_box()` 提示。
9. 窗口标题必须保持 `PPT 控制台`（`desktop.APP_TITLE`），找窗口和单实例都靠它。

## 日志约定
`print()` 即可，`Tee` 会自动加毫秒时间戳并写入 `ppt-presenter.log`。行首用标签：
`[PowerPoint]` 操作及耗时、`[放映]` 页码 / 状态变化、`[待命]`、`[快捷键]`、`[缩略图]`、`[备注]`、`[窗口]`、`[调试]`、`[诊断]`、`[提示]`（预期内的情况，不打印 Traceback）、`[错误]`（意外错误，附 Traceback）。
调试窗口按关键字上色：错误红、忙碌 / 重试黄、`[放映]` `[快捷键]` 绿、`[缩略图]` 灰。

## 开发与测试
- 在 Windows 上：`start.bat`（Python 直接运行，日志同时输出到命令行）。
- 在非 Windows 上：`python server.py` 进入**演示模式**（`MockController`，没有 PowerPoint、没有窗口、没有快捷键），用浏览器打开日志里的地址即可测试界面和流程。
- 也可以在脚本里 `import server`，用 `server.serve_forever(port)` 起服务，再调用 `server.on_hotkey("go")` 等函数测试（不要调用 `main()`，它会打开窗口）。
- 改完界面后至少检查：节目单排序、GO / 待命流转、跳页、黑屏、设置对话框、窄屏布局。

## 版本与发布
1. 每次修改都在 `CHANGELOG.md` **最上面**加新版本：`## [X.Y] - YYYY-MM-DD`，分“新增 / 改进 / 修复 / 移除”写清楚原因和做法。
2. 运行 `build.bat`，生成 `dist\PPT-Presenter-VX.Y.exe`（版本号来自 CHANGELOG，界面右上角也会显示）。
3. 新增 Python 模块要确认 PyInstaller 能找到（直接 import 的会自动打包；动态导入的加到 `build.bat` 的 `--hidden-import`）。新增数据文件要加 `--add-data`。
