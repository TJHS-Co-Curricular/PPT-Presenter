# -*- coding: utf-8 -*-
"""
从 CHANGELOG.md 读出最新版本号（第一个 “## [X.Y]” 或 “## vX.Y” 标题）。
- 程序运行时：显示在界面和日志里
- build.bat：`python version.py` 打印版本号，用来给 exe 命名（PPT-Presenter-VX.Y.exe）
"""
import os
import re
import sys

_PATTERN = re.compile(r"^##\s*\[?\s*[vV]?(\d+(?:\.\d+)+)\s*\]?", re.M)


def get_version(*dirs):
    """依次在给定的文件夹里找 CHANGELOG.md；找不到就返回 "dev"。"""
    for d in dirs:
        if not d:
            continue
        path = os.path.join(d, "CHANGELOG.md")
        try:
            with open(path, "r", encoding="utf-8") as f:
                m = _PATTERN.search(f.read())
            if m:
                return m.group(1)
        except OSError:
            continue
    return "dev"


if __name__ == "__main__":
    here = os.path.dirname(os.path.abspath(__file__))
    v = get_version(sys.argv[1] if len(sys.argv) > 1 else here)
    if v == "dev":
        sys.stderr.write("CHANGELOG.md 里找不到版本号（需要像 “## [1.7] - 2026-10-13” 这样的标题）\n")
        sys.exit(1)
    print(v)
