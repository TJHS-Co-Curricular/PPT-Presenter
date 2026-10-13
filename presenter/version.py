# -*- coding: utf-8 -*-
"""
版本号：取 CHANGELOG.md 里第一个 “## [X.Y]” 或 “## vX.Y” 标题。

- 程序运行时显示在界面和日志里
- build.bat 用 ``python -m presenter.version`` 打印版本号，给 exe 命名（PPT-Presenter-VX.Y.exe）
"""
import os
import re
import sys

_PATTERN = re.compile(r"^##\s*\[?\s*[vV]?(\d+(?:\.\d+)+)\s*\]?", re.M)


def get_version(*dirs):
    """依次在给定的文件夹里找 CHANGELOG.md；都找不到就返回 "dev"。"""
    for d in dirs:
        if not d:
            continue
        try:
            with open(os.path.join(d, "CHANGELOG.md"), "r", encoding="utf-8") as f:
                m = _PATTERN.search(f.read())
            if m:
                return m.group(1)
        except OSError:
            continue
    return "dev"


def _current():
    from presenter import paths
    return get_version(paths.RES, paths.BASE)


VERSION = _current()


if __name__ == "__main__":
    if VERSION == "dev":
        sys.stderr.write("CHANGELOG.md 里找不到版本号（需要像 “## [1.9] - 2026-10-14” 这样的标题）\n")
        sys.exit(1)
    print(VERSION)
