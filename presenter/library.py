# -*- coding: utf-8 -*-
"""
节目单：扫描文件夹里的 PPT / 视频 / 音频，按文件名开头的编号排序。

排序规则：有短编号的在前，按编号大小排（0、1、1.1、1.2、2 … 5.1-5.3、5.4 … 10、11、12；
范围按开头的编号算）；编号相同时按文件名；没有编号的放最后。
2025 这类年份、20b9eb73 这类编号后面紧跟英文字母的，都不算编号。
"""
import hashlib
import os
import re
import threading

from presenter.config import CFG
from presenter.formats import is_temp_file, kind_of

_NUM_RE = re.compile(
    r"^\s*(\d{1,2}(?:\.\d{1,2})?(?:\s*-\s*\d{1,2}(?:\.\d{1,2})?)?)(?![\dA-Za-z])\s*[.、_\s-]?\s*(.+)$")


class ItemNotFound(LookupError):
    """节目单里找不到这个项目（通常是文件夹变了，需要刷新）。"""


def norm(path):
    """统一路径写法（Windows 不分大小写），用作字典的键。"""
    return os.path.normcase(os.path.abspath(path))


def item_id(path):
    return hashlib.md5(norm(path).encode("utf-8")).hexdigest()[:12]


def natural_key(s):
    return [int(t) if t.isdigit() else t.lower() for t in re.split(r"(\d+)", s)]


def num_key(num):
    """'5.1-5.3' → (5, 1)：按开头的编号比大小，所以 1 < 1.1 < 1.2 < 2。"""
    return tuple(int(t) for t in re.findall(r"\d+", num.split("-")[0]))


def sort_key(item):
    if item["num"]:
        return (0, num_key(item["num"]), natural_key(item["name"]))
    return (1, (), natural_key(item["name"]))


def _make_item(full, name, kind):
    stem = os.path.splitext(name)[0]
    m = _NUM_RE.match(stem)
    num, title = (m.group(1).replace(" ", ""), m.group(2)) if m else ("", stem)
    st = os.stat(full)
    return {
        "id": item_id(full),
        "path": full,
        "name": name,
        "title": title,
        "num": num,
        "type": kind,
        "size": st.st_size,
        "mtime": int(st.st_mtime),
    }


def scan_folder(folder):
    items = []
    if not folder or not os.path.isdir(folder):
        return items
    for name in os.listdir(folder):
        if is_temp_file(name):
            continue
        full = os.path.join(folder, name)
        if not os.path.isfile(full):
            continue
        kind = kind_of(os.path.splitext(name)[1])
        if kind:
            items.append(_make_item(full, name, kind))
    items.sort(key=sort_key)
    return items


ITEMS = {}                  # id → 项目
ITEMS_LOCK = threading.Lock()


def refresh_items():
    """重新扫描文件夹。待命的节目如果不在列表里了（例如换了文件夹），改成第一个。"""
    lst = scan_folder(CFG["folder"])
    with ITEMS_LOCK:
        ITEMS.clear()
        for it in lst:
            ITEMS[it["id"]] = it
    if CFG.get("standby") not in ITEMS:
        CFG["standby"] = lst[0]["id"] if lst else ""
    return lst


def is_base(item):
    """编号为 0 的 PPT = 大会背景：第 1 页是大会背景、第 2 页是礼成；不能关闭。"""
    return bool(item) and item["type"] == "ppt" and item["num"] == "0"


def base_item():
    """节目单里编号为 0 的 PPT；没有就返回 None。"""
    return next((it for it in ordered_items() if is_base(it)), None)


def get_item(iid):
    with ITEMS_LOCK:
        it = ITEMS.get(iid)
    if not it:
        raise ItemNotFound("找不到这个项目，请刷新列表")
    return it


def find_by_key(key):
    """按 norm(path) 找项目；找不到返回 None。"""
    with ITEMS_LOCK:
        return next((it for it in ITEMS.values() if norm(it["path"]) == key), None)


def ordered_items():
    with ITEMS_LOCK:
        return sorted(ITEMS.values(), key=sort_key)


def ppt_items():
    return [it for it in ordered_items() if it["type"] == "ppt"]


def next_item_id(iid):
    """节目单里排在 iid 后面的那个；已经是最后一个时返回 ""。"""
    lst = ordered_items()
    for k, it in enumerate(lst):
        if it["id"] == iid:
            return lst[k + 1]["id"] if k + 1 < len(lst) else ""
    return ""
