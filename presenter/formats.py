# -*- coding: utf-8 -*-
"""支持的文件类型。"""

PPT_EXT = frozenset({".ppt", ".pptx", ".pptm", ".pps", ".ppsx", ".ppsm"})
VIDEO_EXT = frozenset({".mp4", ".mov", ".avi", ".mkv", ".wmv", ".m4v", ".mpg", ".mpeg"})
AUDIO_EXT = frozenset({".mp3", ".wav", ".m4a", ".aac", ".flac", ".wma"})

#: 节目单里显示的类型名
TYPE_PPT, TYPE_VIDEO, TYPE_AUDIO = "ppt", "video", "audio"


def kind_of(ext):
    """按扩展名判断类型；不支持的返回 None。"""
    ext = ext.lower()
    if ext in PPT_EXT:
        return TYPE_PPT
    if ext in VIDEO_EXT:
        return TYPE_VIDEO
    if ext in AUDIO_EXT:
        return TYPE_AUDIO
    return None


def is_temp_file(name):
    """PowerPoint 锁文件（~$）、macOS 残留（._）、隐藏文件。"""
    return name.startswith(("~$", "._", "."))
