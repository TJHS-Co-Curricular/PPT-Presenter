# -*- coding: utf-8 -*-
"""
字体检查：找出 PPT 用到、但这台电脑没安装、也没内嵌在 PPT 里的字体（放映时会被替换成别的字体）。

判断方法（Windows）：
1. 系统字体名单：注册表的字体列表 + GDI 列举的字体家族名；
2. 名单里找不到时（例如中文名“黑体”在英文系统里登记为 SimHei），再让 Windows 实际选一次字体：
   如果选出来的字体和“一个肯定不存在的名字”选出来的不一样，说明这个名字对应得到真实字体。
非 Windows 环境无法判断，返回 None。
"""
import sys

_installed = None
_baseline = None
_PROBE_MISSING = "PPTPresenter-NoSuchFont-9Q"


def _registry_names():
    import winreg
    names = set()
    key_path = r"SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts"
    for root in (winreg.HKEY_LOCAL_MACHINE, winreg.HKEY_CURRENT_USER):
        try:
            key = winreg.OpenKey(root, key_path)
        except OSError:
            continue
        i = 0
        while True:
            try:
                name = winreg.EnumValue(key, i)[0]
            except OSError:
                break
            i += 1
            name = name.split(" (")[0]                 # 去掉 “(TrueType)” 之类
            for part in name.split(" & "):            # “SimSun & NSimSun”
                names.add(part.strip().lower())
    return names


def _logfont_type():
    import ctypes
    from ctypes import wintypes

    class LOGFONTW(ctypes.Structure):
        _fields_ = [("lfHeight", wintypes.LONG), ("lfWidth", wintypes.LONG), ("lfEscapement", wintypes.LONG),
                    ("lfOrientation", wintypes.LONG), ("lfWeight", wintypes.LONG), ("lfItalic", wintypes.BYTE),
                    ("lfUnderline", wintypes.BYTE), ("lfStrikeOut", wintypes.BYTE), ("lfCharSet", wintypes.BYTE),
                    ("lfOutPrecision", wintypes.BYTE), ("lfClipPrecision", wintypes.BYTE),
                    ("lfQuality", wintypes.BYTE), ("lfPitchAndFamily", wintypes.BYTE),
                    ("lfFaceName", wintypes.WCHAR * 32)]
    return LOGFONTW


def _gdi_family_names():
    import ctypes
    from ctypes import wintypes
    LOGFONTW = _logfont_type()
    names = set()
    proc_type = ctypes.WINFUNCTYPE(ctypes.c_int, ctypes.POINTER(LOGFONTW), ctypes.c_void_p,
                                   wintypes.DWORD, wintypes.LPARAM)

    def callback(lf, _tm, _type, _param):
        names.add(lf.contents.lfFaceName.lstrip("@").lower())
        return 1

    hdc = ctypes.windll.user32.GetDC(0)
    try:
        lf = LOGFONTW()
        lf.lfCharSet = 1   # DEFAULT_CHARSET：列出所有字符集
        ctypes.windll.gdi32.EnumFontFamiliesExW(hdc, ctypes.byref(lf), proc_type(callback), 0, 0)
    finally:
        ctypes.windll.user32.ReleaseDC(0, hdc)
    return names


def _selected_face(name):
    """让 Windows 按这个名字选字体，返回实际选到的字体名。"""
    import ctypes
    LOGFONTW = _logfont_type()
    gdi32, user32 = ctypes.windll.gdi32, ctypes.windll.user32
    hdc = user32.GetDC(0)
    try:
        lf = LOGFONTW()
        lf.lfHeight = -16
        lf.lfCharSet = 1
        lf.lfFaceName = name[:31]
        hfont = gdi32.CreateFontIndirectW(ctypes.byref(lf))
        old = gdi32.SelectObject(hdc, hfont)
        buf = ctypes.create_unicode_buffer(64)
        gdi32.GetTextFaceW(hdc, 64, buf)
        gdi32.SelectObject(hdc, old)
        gdi32.DeleteObject(hfont)
        return buf.value
    finally:
        user32.ReleaseDC(0, hdc)


def is_installed(name):
    """True / False；无法判断（非 Windows）时返回 None。"""
    global _installed, _baseline
    if sys.platform != "win32":
        return None
    name = (name or "").strip()
    if not name or name.startswith("+"):          # “+mj-lt” 之类是主题字体占位，不是真名字
        return True
    try:
        if _installed is None:
            _installed = _registry_names() | _gdi_family_names()
            _baseline = _selected_face(_PROBE_MISSING).lower()
        if name.lower() in _installed:
            return True
        face = _selected_face(name).lower()
        return face == name.lower() or face != _baseline
    except Exception:
        return None


def missing_fonts(fonts):
    """fonts：[(名字, 是否内嵌)]。返回没安装也没内嵌的字体名；无法判断时返回 []。"""
    out = []
    for name, embedded in fonts:
        if embedded:
            continue
        if is_installed(name) is False and name not in out:
            out.append(name)
    return out
