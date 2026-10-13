# -*- coding: utf-8 -*-
"""
.cache 里的文件位置。文件名带 PPT 的修改时间，PPT 改过之后旧缓存自动失效。

    thumbs/<id>_<mtime>_r<版本>/<页>.png   缩略图（渲染方式改变时提高 THUMB_RENDER，旧缩略图自动作废）
    notes/<id>_<mtime>.json        演讲者备注 {"notes": {"页": 文字}, "slow": 是否停用}
"""
import os

from presenter import paths
from presenter.library import item_id

#: v2.0 起带窗口打开（内嵌字体正确），之前无窗口生成的缩略图字体可能不对，所以换一个版本号
THUMB_RENDER = 2


def thumb_path(item, n):
    return os.path.join(paths.THUMBS_DIR, "%s_%d_r%d" % (item["id"], item["mtime"], THUMB_RENDER), "%d.png" % n)


def notes_path(ppt_path):
    try:
        return os.path.join(paths.NOTES_DIR, "%s_%d.json" % (item_id(ppt_path), int(os.stat(ppt_path).st_mtime)))
    except OSError:
        return None
