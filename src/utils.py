import contextlib
import glob
import os
import re
import socket
import subprocess

import requests

from .config import Config


def get_lyric(lyrics, pos):
    if not lyrics:
        return "", ""
    left, r, idx = 0, len(lyrics) - 1, -1
    while left <= r:
        m = (left + r) // 2
        if lyrics[m][0] <= pos:
            idx, left = m, m + 1
        else:
            r = m - 1
    if idx < 0:
        return lyrics[0][1], lyrics[1][1] if len(lyrics) > 1 else ""
    return lyrics[idx][1], lyrics[idx + 1][1] if idx + 1 < len(lyrics) else ""


def format_output(cfg: Config, state, lyrics, song_key):
    c, d, w = state.cur, state.dur, cfg.bar_width
    pos = int(w * c / d) if d else 0
    thumb = cfg.bar_thumb
    if thumb:
        bar = cfg.bar_filled * pos + thumb + cfg.bar_empty * (w - pos)
    else:
        bar = cfg.bar_filled * pos + cfg.bar_empty * (w - pos)
    l1, l2 = state.lyric1, state.lyric2
    if not l1 and song_key == f"{state.song}-{state.artist}":
        l1, l2 = get_lyric(lyrics, c)
    l1, l2 = l1 or "纯音乐，请欣赏", l2 or ""
    try:
        return cfg.template.format(
            song=state.song,
            artist=state.artist,
            bar=bar,
            time=f"{c // 60}:{c % 60:02d}/{d // 60}:{d % 60:02d}",
            lyric1=l1,
            lyric2=l2,
        )
    except Exception:
        return f"🎵 {state.song} - {state.artist}\n{bar}\n{l1}"


HEADERS = {"User-Agent": "Mozilla/5.0", "Referer": "https://music.163.com/"}


def fetch_lyrics(song, artist):
    with contextlib.suppress(Exception):
        r = requests.post(
            "https://music.163.com/api/search/get",
            data={"s": f"{song} {artist}", "type": 1, "limit": 1},
            headers=HEADERS,
            timeout=3,
        ).json()
        if r.get("result", {}).get("songs"):
            lrc = (
                requests.get(
                    f"https://music.163.com/api/song/lyric?id={r['result']['songs'][0]['id']}&lv=1",
                    headers=HEADERS,
                    timeout=3,
                )
                .json()
                .get("lrc", {})
                .get("lyric", "")
            )
            return sorted(
                [
                    (
                        int(m[1]) * 60
                        + int(m[2])
                        + float(m[3]) * (0.01 if len(m[3]) == 2 else 0.001),
                        m[4].strip(),
                    )
                    for m in re.finditer(r"\[(\d{2}):(\d{2})\.(\d{2,3})\](.*)", lrc)
                    if m[4].strip()
                ],
                key=lambda x: x[0],
            )
    return []


def find_netease() -> str | None:
    patterns = [
        r"%APPDATA%\Microsoft\Windows\Start Menu\Programs\**\*网易云*.lnk",
        r"%PROGRAMDATA%\Microsoft\Windows\Start Menu\Programs\**\*网易云*.lnk",
    ]
    for pat_idx, pattern in enumerate(patterns):
        for lnk in glob.glob(os.path.expandvars(pattern), recursive=True):
            try:
                with open(lnk, "rb") as f:
                    m = re.search(
                        rb"([A-Za-z]:\\[^\x00]+?cloudmusic\.exe)",
                        f.read(),
                        re.IGNORECASE,
                    )
                    if m:
                        p = m.group(1).decode("utf-8", errors="ignore")
                        if os.path.exists(p):
                            return p
            except Exception:
                continue
    return None


# 找呀找呀找端口，找到一个好端口
def find_free_port():
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        s.bind(("127.0.0.1", 0))
        return s.getsockname()[1]


def launch_netease(port=None, path=None):
    import atexit

    exe = path if path and os.path.exists(path) else find_netease()
    if not exe:
        return False, "未找到网易云", None
    if port is None:
        port = find_free_port()
    proc = subprocess.Popen([exe, f"--remote-debugging-port={port}"])

    def _kill_proc():
        with contextlib.suppress(Exception):
            proc.terminate()
            proc.wait(timeout=3)

    atexit.register(_kill_proc)
    return True, exe, port
