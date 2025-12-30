import asyncio
import json
import threading
import time
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from typing import Protocol

import requests
import websockets
from pythonosc import udp_client
from websockets import protocol

from src.config import Config
from src.state import SharedState
from src.utils import fetch_lyrics, format_output, launch_netease

CONFIG_FILE = "ncm_vrchat_config.json"


# 兼容 VIP 界面
JS_GET_STATE = r"""(() => {
    try {
        let r = { song: '', artist: '', cur: 0, dur: 0, play: false, lyric1: '', lyric2: '' };

        // 获取歌曲名
        let songEl = document.querySelector('.cmd-space.title span') 
            || document.querySelector('.main-title')
            || document.querySelector('.two-line .title')
            || document.querySelector('[class*="title"] span');
        r.song = songEl?.innerText?.trim() || songEl?.textContent?.trim() || '';

        // .author, .info.artist
        let artist = document.querySelector('.author');
        r.artist = artist?.innerText?.trim() || '';
        if (!r.artist) { artist = document.querySelector('.info.artist'); r.artist = (artist?.innerText || '').replace(/^歌手[：:]/, '').trim(); }

        // 进度遍历
        let walker = document.createTreeWalker(document.body, NodeFilter.SHOW_TEXT);
        while (walker.nextNode()) {
            let text = walker.currentNode.textContent.trim();
            let m = text.match(/^(\d+):(\d+)\s*\/\s*(\d+):(\d+)$/);
            if (m) {
                r.cur = +m[1] * 60 + +m[2];
                r.dur = +m[3] * 60 + +m[4];
                break;
            }
        }
        
        // 备用 .curtime-thumb
        if (!r.dur) {
            let timeEl = document.querySelector('.curtime-thumb');
            if (timeEl?.innerText) { 
                let m = timeEl.innerText.match(/(\d+):(\d+)\s*\/\s*(\d+):(\d+)/); 
                if (m) { r.cur = +m[1] * 60 + +m[2]; r.dur = +m[3] * 60 + +m[4]; } 
            }
        }

        // cmd-icon-pause, title
        r.play = !!document.querySelector('[class*="cmd-icon-pause"]') || !!document.querySelector('[title*="暂停（Ctrl"]');

        // .line.current
        let curLine = document.querySelector('.line.current');
        if (curLine) {
            r.lyric1 = curLine.innerText?.trim() || '';
            let next = curLine.nextElementSibling;
            if (next && next.classList?.contains('line')) {
                r.lyric2 = next.innerText?.trim() || '';
            }
        }

        return r;
    } catch (e) { return null; }
})()"""


class CallbackProtocol(Protocol):
    def cb_status(self, t: str) -> None: ...
    def cb_song(self, t: str) -> None: ...
    def cb_output(self, t: str) -> None: ...


def netease_thread(
    cfg: Config, shared: SharedState, stop_event: threading.Event, cb: CallbackProtocol
):
    class NeteaseSync:
        def __init__(self):
            self.ws = None
            self.msg_id = 0

        async def connect(self):
            pages = requests.get(
                f"http://127.0.0.1:{cfg.ncm_port}/json", timeout=2
            ).json()
            self.ws = await websockets.connect(
                pages[0]["webSocketDebuggerUrl"], ping_interval=30, ping_timeout=15
            )

        async def eval_js(self, code, timeout=1):
            if not self.ws:
                return None
            
            # 激活页面
            self.msg_id += 1
            bring_id = self.msg_id
            await self.ws.send(json.dumps({"id": bring_id, "method": "Page.bringToFront"}))
            
            # 等待 bringToFront 响应
            try:
                while True:
                    msg = await asyncio.wait_for(self.ws.recv(), timeout=0.5)
                    d = json.loads(msg)
                    if d.get("id") == bring_id:
                        break
            except asyncio.TimeoutError:
                pass
            
            # 执行主代码
            self.msg_id += 1
            msg_id = self.msg_id
            await self.ws.send(json.dumps({
                "id": msg_id, "method": "Runtime.evaluate",
                "params": {"expression": code, "returnByValue": True},
            }))
            
            result = None
            try:
                end_time = asyncio.get_event_loop().time() + timeout
                while True:
                    remaining = end_time - asyncio.get_event_loop().time()
                    if remaining <= 0:
                        break
                    msg = await asyncio.wait_for(self.ws.recv(), timeout=remaining)
                    d = json.loads(msg)
                    if d.get("id") == msg_id:
                        result = d.get("result", {}).get("result", {}).get("value")
                        break
            except asyncio.TimeoutError:
                pass
            except Exception:
                pass
            
            return result

    async def run():
        cb.cb_status("连接中...")
        sync = NeteaseSync()
        retry_count = 0
        max_retries = 3
        while not sync.ws and retry_count < max_retries:
            try:
                await sync.connect()
            except Exception as e:
                retry_count += 1
                cb.cb_status(f"连接失败，重试中... {retry_count}/{max_retries} \n {e}")
                await asyncio.sleep(2)
        if not sync.ws:
            cb.cb_status("连接失败，已达最大重试次数")
            return
        cb.cb_status("已连接")
        
        timeout_count = 0
        
        while not stop_event.is_set():
            try:
                s = await sync.eval_js(JS_GET_STATE)
                
                if s is None:
                    timeout_count += 1
                    if timeout_count >= 3:
                        cb.cb_status("响应超时，重连中...")
                        try:
                            await sync.connect()
                            cb.cb_status("已重连")
                            timeout_count = 0
                        except Exception:
                            pass
                    await asyncio.sleep(0.5)
                    continue
                
                timeout_count = 0
                
                if s.get("song"):
                    need_fetch = False
                    song_to_fetch = ""
                    artist_to_fetch = ""
                    key = ""
                    
                    with shared.lock:
                        # 切歌过渡期
                        if s.get("cur", 0) == 0 and s.get("dur", 0) == 0:
                            if shared.data.dur > 0:
                                if s.get("song") != shared.data.song:
                                    s["cur"] = 0
                                    s["dur"] = shared.data.dur
                                else:
                                    s["cur"] = shared.data.cur
                                    s["dur"] = shared.data.dur
                        
                        shared.data.update(s)
                        shared.last_update = time.time()
                        
                        if shared.data.song:
                            key = f"{shared.data.song}-{shared.data.artist}"
                            if key != shared.song_key:
                                shared.song_key = key
                                shared.lyrics = []
                                need_fetch = True
                                song_to_fetch = shared.data.song
                                artist_to_fetch = shared.data.artist
                    
                    # 启动歌词获取线程
                    if need_fetch:
                        def fetch_task(k, song, artist):
                            try:
                                new_lyrics = fetch_lyrics(song, artist)
                                with shared.lock:
                                    if shared.song_key == k:
                                        shared.lyrics = new_lyrics
                            except Exception:
                                pass
                        threading.Thread(
                            target=fetch_task,
                            args=(key, song_to_fetch, artist_to_fetch),
                            daemon=True
                        ).start()
                
                await asyncio.sleep(0.3)
            except websockets.exceptions.ConnectionClosed:
                await asyncio.sleep(1)
                try:
                    await sync.connect()
                    cb.cb_status("已重连")
                except Exception:
                    cb.cb_status("重连失败，继续尝试...")
            except Exception:
                pass
        
        if sync.ws and sync.ws.state == protocol.State.OPEN:
            await sync.ws.close()

    asyncio.run(run())


def osc_thread(
    cfg: Config, shared: SharedState, stop_event: threading.Event, cb: CallbackProtocol
):
    osc = None
    last_osc = 0
    while not stop_event.is_set():
        with shared.lock:
            state = shared.data.copy()
            lyrics = list(shared.lyrics)
            song_key = shared.song_key
            last_update = shared.last_update
        
        # 推算进度
        if state.play and state.song and last_update > 0:
            elapsed = time.time() - last_update
            state.cur = min(int(state.cur + elapsed), state.dur) if state.dur else state.cur
        
        if state.play and state.song:
            out = format_output(cfg, state, lyrics, song_key)
            now = time.time()
            if now - last_osc >= cfg.refresh_interval:
                if osc is None:
                    osc = udp_client.SimpleUDPClient(cfg.osc_ip, cfg.osc_port)
                osc.send_message("/chatbox/input", [out, True, False])
                last_osc = now
                cb.cb_output(out)
                cb.cb_song(f"播放：{state.song} - {state.artist}")
        else:
            if state.song:
                cb.cb_song(f"暂停：{state.song}")
        time.sleep(0.3)


class App:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("网易云 VRC 助手 - 1.0.2")
        self.root.geometry("460x480")
        self.root.resizable(False, False)
        self.cfg: Config = self.load_cfg()
        self.sync_event = threading.Event()
        self.launch_event = threading.Event()
        self.shared_state = SharedState()
        self.netease_thread = None
        self.osc_thread = None
        self.build_ui()

    def load_cfg(self) -> Config:
        try:
            with open(CONFIG_FILE, encoding="utf-8") as f:
                data = {
                    **Config().model_dump(),
                    **json.load(f),
                }
            data = Config(**data)
        except Exception:  # 默认配置覆盖
            data = Config()
        self.save_cfg(data)
        return data

    @staticmethod
    def save_cfg(config: Config):
        data = config.model_dump(
            mode="json",
            exclude_none=True,
        )
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)

    def build_ui(self):
        m = ttk.Frame(self.root, padding=12)
        m.pack(fill="both", expand=True)

        # 状态栏
        f = ttk.Frame(m)
        f.pack(fill="x", pady=(0, 8))
        self.status = tk.StringVar(value="未连接")
        self.song = tk.StringVar()
        ttk.Label(f, textvariable=self.status, font=("", 10, "bold")).pack(side="left")
        ttk.Label(f, textvariable=self.song, font=("", 9)).pack(side="right")

        # 按钮
        f = ttk.Frame(m)
        f.pack(fill="x", pady=(0, 8))
        self.btn_launch = ttk.Button(
            f, text="启动网易云", command=self.do_launch, width=11
        )
        self.btn_launch.pack(side="left")
        ttk.Button(f, text="选择路径", command=self.do_browse, width=9).pack(
            side="left", padx=4
        )
        self.btn_start = ttk.Button(f, text="开始同步", command=self.do_start, width=9)
        self.btn_start.pack(side="left", padx=4)
        self.btn_stop = ttk.Button(
            f, text="停止", command=self.do_stop, width=5, state="disabled"
        )
        self.btn_stop.pack(side="left")

        # 路径
        f = ttk.Frame(m)
        f.pack(fill="x", pady=(0, 8))
        ttk.Label(f, text="网易云路径：").pack(side="left")
        self.path = tk.StringVar(value=self.cfg.ncm_path or "(自动检测)")
        ttk.Label(f, textvariable=self.path, foreground="gray").pack(
            side="left", padx=4
        )

        # 基础设置
        f = ttk.LabelFrame(m, text="基础设置", padding=6)
        f.pack(fill="x", pady=(0, 8))
        g = ttk.Frame(f)
        g.pack(fill="x")
        ttk.Label(g, text="OSC").grid(row=0, column=0)
        self.e_ip = ttk.Entry(g, width=11)
        self.e_ip.insert(0, self.cfg.osc_ip)
        self.e_ip.grid(row=0, column=1)
        ttk.Label(g, text=":").grid(row=0, column=2)
        self.e_port = ttk.Entry(g, width=5)
        self.e_port.insert(0, str(self.cfg.osc_port))
        self.e_port.grid(row=0, column=3)
        ttk.Label(g, text="刷新").grid(row=0, column=4, padx=(12, 0))
        self.v_interval = tk.DoubleVar(value=self.cfg.refresh_interval)
        ttk.Spinbox(
            g, from_=2, to=10, increment=0.5, width=4, textvariable=self.v_interval
        ).grid(row=0, column=5)
        ttk.Label(g, text="秒").grid(row=0, column=6)

        # 进度条设置
        g2 = ttk.Frame(f)
        g2.pack(fill="x", pady=(6, 0))
        ttk.Label(g2, text="进度条").pack(side="left")
        ttk.Label(g2, text="宽度").pack(side="left", padx=(8, 0))
        self.e_bw = ttk.Entry(g2, width=3)
        self.e_bw.insert(0, str(self.cfg.bar_width))
        self.e_bw.pack(side="left", padx=2)
        ttk.Label(g2, text="已播放").pack(side="left", padx=(8, 0))
        self.e_bf = ttk.Entry(g2, width=3)
        self.e_bf.insert(0, self.cfg.bar_filled)
        self.e_bf.pack(side="left", padx=2)
        ttk.Label(g2, text="滑块").pack(side="left", padx=(8, 0))
        self.e_bt = ttk.Entry(g2, width=3)
        self.e_bt.insert(0, self.cfg.bar_thumb)
        self.e_bt.pack(side="left", padx=2)
        ttk.Label(g2, text="未播放").pack(side="left", padx=(8, 0))
        self.e_be = ttk.Entry(g2, width=3)
        self.e_be.insert(0, self.cfg.bar_empty)
        self.e_be.pack(side="left", padx=2)

        # 输出模板
        f = ttk.LabelFrame(m, text="输出模板", padding=6)
        f.pack(fill="both", expand=True, pady=(0, 8))
        ttk.Label(
            f,
            text="可用变量：{song} {artist} {bar} {time} {lyric1} {lyric2}",
            foreground="gray",
        ).pack(anchor="w")
        self.t_tpl = tk.Text(f, height=3, font=("Consolas", 10))
        self.t_tpl.insert("1.0", self.cfg.template)
        self.t_tpl.pack(fill="both", expand=True, pady=(4, 0))

        # 实时预览
        f = ttk.LabelFrame(m, text="文本预览", padding=6)
        f.pack(fill="both", expand=True)
        self.t_preview = tk.Text(
            f, height=3, font=("Consolas", 10), state="disabled", bg="#f5f5f5"
        )
        self.t_preview.pack(fill="both", expand=True)

        for w in [self.t_tpl, self.e_bw, self.e_bf, self.e_bt, self.e_be]:
            w.bind("<KeyRelease>", lambda e: self.preview())
        self.preview()

    def do_browse(self):
        if p := filedialog.askopenfilename(
            title="选择 cloudmusic.exe",
            filetypes=[("", "cloudmusic.exe"), ("", "*.exe")],
        ):
            self.cfg.ncm_path = p
            self.path.set(p)
            self.save_cfg(self.cfg)

    def do_launch(self):
        ok, r, port = launch_netease(None, self.cfg.ncm_path)
        if ok and port is not None:
            self.cfg.ncm_port = port
            self.status.set(f"已启动 (端口:{port})")
            self.path.set(r)
            self.root.after(3000, lambda: self.status.set("就绪"))
            self.netease_thread = threading.Thread(
                target=netease_thread,
                args=(self.cfg, self.shared_state, self.launch_event, self),
                daemon=True,
            )
            self.netease_thread.start()
            self.btn_launch.config(state="disabled")
        else:
            messagebox.showwarning("提示", f"{r}\n请手动选择路径")

    def stop_launch(self):
        if self.netease_thread and self.netease_thread.is_alive():
            self.launch_event.set()
            self.netease_thread.join(timeout=2)
            self.launch_event = threading.Event()
        self.btn_launch.config(state="normal")
        self.song.set("")

    def preview(self):
        try:
            w = int(self.e_bw.get() or 10)
            thumb = self.e_bt.get() or ""
            pos = w // 2
            bar = (
                (self.e_bf.get() or "▓") * pos
                + thumb
                + (self.e_be.get() or "░") * (w - pos)
            )
            txt = (
                self.t_tpl.get("1.0", "end")
                .strip()
                .format(
                    song="歌曲名称",
                    artist="歌手",
                    bar=bar,
                    time="1:14/5:14",
                    lyric1="当前歌词",
                    lyric2="下句歌词",
                )
            )
            self.t_preview.config(state="normal")
            self.t_preview.delete("1.0", "end")
            self.t_preview.insert("1.0", txt)
            self.t_preview.config(state="disabled")
        except Exception:
            pass

    def update_cfg(self):
        self.cfg.osc_ip = self.e_ip.get()
        self.cfg.osc_port = int(self.e_port.get() or 9000)
        self.cfg.refresh_interval = max(2, min(10, self.v_interval.get()))
        self.cfg.template = self.t_tpl.get("1.0", "end").strip()
        self.cfg.bar_width = int(self.e_bw.get() or 9)
        self.cfg.bar_filled = self.e_bf.get() or "▓"
        self.cfg.bar_thumb = self.e_bt.get() or ""
        self.cfg.bar_empty = self.e_be.get() or "░"

    def cb_status(self, t: str):
        self.root.after(0, lambda: self.status.set(t))

    def cb_song(self, t: str):
        self.root.after(0, lambda: self.song.set(t[:28]))

    def cb_output(self, t: str):
        def f():
            self.t_preview.config(state="normal")
            self.t_preview.delete("1.0", "end")
            self.t_preview.insert("1.0", t)
            self.t_preview.config(state="disabled")

        self.root.after(0, f)

    def do_start(self):
        try:
            self.update_cfg()
            self.save_cfg(self.cfg)
        except Exception as e:
            messagebox.showerror("错误", str(e))
            return

        self.osc_thread = threading.Thread(
            target=osc_thread,
            args=(self.cfg, self.shared_state, self.sync_event, self),
            daemon=True,
        )
        self.osc_thread.start()

        self.btn_start.config(state="disabled")
        self.btn_stop.config(state="normal")
        self.status.set("同步中...")

    def do_stop(self):
        if self.osc_thread and self.osc_thread.is_alive():
            self.sync_event.set()
            self.osc_thread.join(timeout=2)
            self.sync_event = threading.Event()
        self.btn_start.config(state="normal")
        self.btn_stop.config(state="disabled")
        self.status.set("已停止")

    def run(self):
        def on_close():
            self.do_stop()
            self.stop_launch()
            self.root.destroy()
        self.root.protocol("WM_DELETE_WINDOW", on_close)
        self.root.mainloop()


if __name__ == "__main__":
    App().run()
