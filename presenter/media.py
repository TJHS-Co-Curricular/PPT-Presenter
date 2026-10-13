# -*- coding: utf-8 -*-
"""视频 / 音频播放：有 VLC 就全屏播放（可以停止），没有就用系统默认播放器。"""
import os
import shutil
import subprocess
import sys
import time

VLC_ARGS = ["--fullscreen", "--play-and-exit", "--no-video-title-show", "--no-qt-privacy-ask"]


def find_vlc():
    candidates = [shutil.which("vlc")]
    for env in ("ProgramFiles", "ProgramFiles(x86)"):
        base = os.environ.get(env)
        if base:
            candidates.append(os.path.join(base, "VideoLAN", "VLC", "vlc.exe"))
    return next((c for c in candidates if c and os.path.exists(c)), None)


class MediaPlayer:
    def __init__(self):
        self.proc = None
        self.name = None
        self._vlc = None
        self._vlc_checked = 0.0

    def vlc(self):
        """VLC 路径（每 30 秒重新找一次，不用每次查询都扫硬盘）。"""
        if time.time() - self._vlc_checked > 30:
            self._vlc, self._vlc_checked = find_vlc(), time.time()
        return self._vlc

    def play(self, path):
        self.stop()
        self.name = os.path.basename(path)
        vlc = self.vlc()
        if vlc:
            self.proc = subprocess.Popen([vlc] + VLC_ARGS + [path])
            return "vlc"
        if sys.platform != "win32":
            return "mock"
        os.startfile(path)  # 系统默认播放器（无法从控制台停止）
        return "default"

    def stop(self):
        if self.proc and self.proc.poll() is None:
            try:
                self.proc.terminate()
            except Exception:
                pass
        self.proc = None
        self.name = None

    def status(self):
        playing = bool(self.proc and self.proc.poll() is None)
        if not playing and self.proc:
            self.proc = None
            self.name = None
        return {"playing": playing, "name": self.name, "vlc": bool(self.vlc())}


MEDIA = MediaPlayer()
