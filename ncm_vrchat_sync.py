import asyncio, json, time, threading, subprocess, os, winreg, re
import websockets, tkinter as tk
from tkinter import ttk, messagebox, filedialog
from pythonosc import udp_client
import requests

CONFIG_FILE = "ncm_vrchat_config.json"
DEFAULT_CONFIG = {
    "osc_ip": "127.0.0.1", "osc_port": 9000, "ncm_port": 9222, "ncm_path": "",
    "refresh_interval": 3.0, "bar_width": 9, "bar_filled": "▓", "bar_empty": "░", "bar_thumb": "◘",
    "template": "🎵 {song} - {artist}\n{bar} {time}\n{lyric1}\n{lyric2}",
}

# 获取播放状态，检测高亮行（白色 + 22px
JS_GET_STATE = """(()=>{try{let r={song:'',artist:'',cur:0,dur:0,play:false,lyric1:'',lyric2:''};
let t=document.querySelector('.main-title');if(t)r.song=t.innerText||'';
let a=document.querySelector('.author');if(a)r.artist=a.innerText||'';
let m=document.querySelector('.curtime-thumb');if(m){
let x=(m.innerText||'').match(/(\\d+):(\\d+)\\s*\\/\\s*(\\d+):(\\d+)/);
if(x){r.cur=+x[1]*60+ +x[2];r.dur=+x[3]*60+ +x[4];}}
r.play=!!document.querySelector('[class*="cmd-icon-pause"]');
let ul=document.querySelector('ul.lyric');
if(ul&&ul.getBoundingClientRect().height>0){let items=ul.querySelectorAll('li');
let idx=-1;items.forEach((li,i)=>{let s=window.getComputedStyle(li);
if(s.color==='rgb(255, 255, 255)'&&s.fontSize==='22px')idx=i;});
if(idx>=0){r.lyric1=items[idx].innerText?.trim()||'';
for(let j=idx+1;j<items.length;j++){let t=items[j].innerText?.trim();if(t){r.lyric2=t;break;}}}}
return r;}catch(e){return null;}})()"""

HEADERS = {"User-Agent": "Mozilla/5.0", "Referer": "https://music.163.com/"}

# 找呀找呀找端口，找到一个好端口
def find_free_port():
    import socket
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        s.bind(('127.0.0.1', 0))
        return s.getsockname()[1]


def find_netease():
    # 常见路径
    paths = [
        r"C:\Program Files (x86)\Netease\CloudMusic\cloudmusic.exe",
        r"C:\Program Files\Netease\CloudMusic\cloudmusic.exe",
        os.path.expandvars(r"%LOCALAPPDATA%\Netease\CloudMusic\cloudmusic.exe"),
        os.path.expandvars(r"%APPDATA%\Netease\CloudMusic\cloudmusic.exe"),
        os.path.expandvars(r"%LOCALAPPDATA%\Programs\Netease\CloudMusic\cloudmusic.exe"),
    ]
    for p in paths:
        if os.path.exists(p): return p
    
    # 注册表
    reg_paths = [
        (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\网易云音乐"),
        (winreg.HKEY_CURRENT_USER, r"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\网易云音乐"),
        (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\网易云音乐"),
        (winreg.HKEY_CLASSES_ROOT, r"Applications\cloudmusic.exe\shell\open\command"),
    ]
    for root, key in reg_paths:
        try:
            k = winreg.OpenKey(root, key)
            try:
                loc = winreg.QueryValueEx(k, "InstallLocation")[0]
                exe = os.path.join(loc, "cloudmusic.exe")
                if os.path.exists(exe): winreg.CloseKey(k); return exe
            except: pass
            try:
                cmd = winreg.QueryValueEx(k, "")[0]
                m = re.search(r'"([^"]+cloudmusic\.exe)"', cmd, re.IGNORECASE)
                if m and os.path.exists(m.group(1)): winreg.CloseKey(k); return m.group(1)
            except: pass
            winreg.CloseKey(k)
        except: pass
    
    # 开始菜单快捷方式
    try:
        import glob
        for pattern in [r"%APPDATA%\Microsoft\Windows\Start Menu\Programs\**\*网易云*.lnk",
                        r"%PROGRAMDATA%\Microsoft\Windows\Start Menu\Programs\**\*网易云*.lnk"]:
            for lnk in glob.glob(os.path.expandvars(pattern), recursive=True):
                try:
                    with open(lnk, 'rb') as f:
                        m = re.search(rb'([A-Za-z]:\\[^\x00]+?cloudmusic\.exe)', f.read(), re.IGNORECASE)
                        if m:
                            p = m.group(1).decode('utf-8', errors='ignore')
                            if os.path.exists(p): return p
                except: pass
    except: pass
    return None


# 开调试
def launch_netease(port=None, path=None):
    exe = path if path and os.path.exists(path) else find_netease()
    if not exe: return False, "未找到网易云", None
    if port is None:
        port = find_free_port()
    try:
        subprocess.Popen([exe, f"--remote-debugging-port={port}"],
            creationflags=subprocess.DETACHED_PROCESS | subprocess.CREATE_NEW_PROCESS_GROUP)
        return True, exe, port
    except Exception as e: return False, str(e), None


class Sync:
    def __init__(self, cfg, cb):
        self.cfg, self.cb = cfg, cb
        self.ws = self.osc = None
        self.msg_id = self.last_osc = 0
        self.running = False
        self.song_key, self.lyrics = "", []
    
    async def connect(self):
        try:
            if self.ws: await self.ws.close()
        except: pass
        try:
            pages = requests.get(f"http://127.0.0.1:{self.cfg['ncm_port']}/json", timeout=2).json()
            self.ws = await websockets.connect(pages[0]["webSocketDebuggerUrl"], ping_interval=30, ping_timeout=15)
            return True
        except: return False
    
    async def eval_js(self, code):
        if not self.ws: return None
        self.msg_id += 1
        await self.ws.send(json.dumps({"id": self.msg_id, "method": "Runtime.evaluate",
                                       "params": {"expression": code, "returnByValue": True}}))
        async for msg in self.ws:
            d = json.loads(msg)
            if d.get("id") == self.msg_id:
                return d.get("result", {}).get("result", {}).get("value")
    
    def fetch_lyrics(self, song, artist):
        try:
            r = requests.post("https://music.163.com/api/search/get",
                data={"s": f"{song} {artist}", "type": 1, "limit": 1}, headers=HEADERS, timeout=3).json()
            if r.get("result", {}).get("songs"):
                lrc = requests.get(f"https://music.163.com/api/song/lyric?id={r['result']['songs'][0]['id']}&lv=1",
                    headers=HEADERS, timeout=3).json().get("lrc", {}).get("lyric", "")
                # 解析 LRC
                return sorted([(int(m[1])*60+int(m[2])+float(m[3])*(0.01 if len(m[3])==2 else 0.001), m[4].strip())
                    for m in re.finditer(r'\[(\d{2}):(\d{2})\.(\d{2,3})\](.*)', lrc) if m[4].strip()], key=lambda x:x[0])
        except: pass
        return []
    
    def get_lyric(self, pos):
        if not self.lyrics: return "", ""
        l, r, idx = 0, len(self.lyrics)-1, -1
        while l <= r:
            m = (l+r)//2
            if self.lyrics[m][0] <= pos: idx, l = m, m+1
            else: r = m-1
        if idx < 0: return "", ""
        return self.lyrics[idx][1], self.lyrics[idx+1][1] if idx+1 < len(self.lyrics) else ""
    
    def format(self, s):
        c, d, w = s["cur"], s["dur"], self.cfg["bar_width"]
        pos = int(w*c/d) if d else 0
        
        # 进度条
        thumb = self.cfg.get("bar_thumb", "")
        if thumb:
            bar = self.cfg["bar_filled"] * pos + thumb + self.cfg["bar_empty"] * (w - pos)
        else:
            bar = self.cfg["bar_filled"] * pos + self.cfg["bar_empty"] * (w - pos)
        
        # 获取歌词
        l1, l2 = s.get("lyric1"), s.get("lyric2")
        if not l1 and self.song_key == f"{s['song']}-{s['artist']}":
            l1, l2 = self.get_lyric(c)
        l1, l2 = l1 or "纯音乐，请欣赏", l2 or ""
        
        try: return self.cfg["template"].format(song=s["song"], artist=s["artist"], bar=bar,
                                                 time=f"{c//60}:{c%60:02d}/{d//60}:{d%60:02d}", lyric1=l1, lyric2=l2)
        except: return f"🎵 {s['song']} - {s['artist']}\n{bar}\n{l1}"
    
    def send_osc(self, text):
        if time.time() - self.last_osc < self.cfg["refresh_interval"]: return False
        try:
            if not self.osc: self.osc = udp_client.SimpleUDPClient(self.cfg["osc_ip"], self.cfg["osc_port"])
            self.osc.send_message("/chatbox/input", [text, True, False])
            self.last_osc = time.time()
            return True
        except: return False
    
    async def run(self):
        self.running = True
        self.cb["status"]("连接中...")

        for i in range(3):
            if await self.connect(): break
            self.cb["status"](f"重试 {i+1}/3")
            await asyncio.sleep(2)
        if not self.ws:
            self.cb["status"]("连接失败")
            self.running = False
            return
        self.cb["status"]("已连接")
        
        while self.running:
            try:
                s = await self.eval_js(JS_GET_STATE)
                if s and s.get("song"):
                    if s["play"]:
                        self.cb["song"](f"播放: {s['song']} - {s['artist']}")
                        key = f"{s['song']}-{s['artist']}"
                        if key != self.song_key:
                            self.song_key = key
                            self.lyrics = self.fetch_lyrics(s["song"], s["artist"])
                        out = self.format(s)
                        if self.send_osc(out): self.cb["output"](out)
                    else:
                        self.cb["song"](f"暂停: {s['song']}")
                await asyncio.sleep(0.3)
            except websockets.exceptions.ConnectionClosed:
                await asyncio.sleep(1)
                if self.running and await self.connect(): self.cb["status"]("已重连")
            except: await asyncio.sleep(0.5)
        
        if self.ws:
            try: await self.ws.close()
            except: pass
    
    def stop(self): self.running = False


class App:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("网易云VRC助手")
        self.root.geometry("460x480")
        self.root.resizable(False, False)
        self.cfg = self.load_cfg()
        self.sync = None
        self.build_ui()
    
    def load_cfg(self):
        try: return {**DEFAULT_CONFIG, **json.load(open(CONFIG_FILE, encoding="utf-8"))}
        except: return DEFAULT_CONFIG.copy()
    
    def save_cfg(self):
        try: json.dump(self.cfg, open(CONFIG_FILE, "w", encoding="utf-8"), ensure_ascii=False, indent=2)
        except: pass
    
    def build_ui(self):
        m = ttk.Frame(self.root, padding=12)
        m.pack(fill="both", expand=True)
        
        # 状态栏
        f = ttk.Frame(m); f.pack(fill="x", pady=(0,8))
        self.status = tk.StringVar(value="未连接")
        self.song = tk.StringVar()
        ttk.Label(f, textvariable=self.status, font=("",10,"bold")).pack(side="left")
        ttk.Label(f, textvariable=self.song, font=("",9)).pack(side="right")
        
        # 按钮
        f = ttk.Frame(m); f.pack(fill="x", pady=(0,8))
        self.btn_launch = ttk.Button(f, text="启动网易云", command=self.do_launch, width=11)
        self.btn_launch.pack(side="left")
        ttk.Button(f, text="选择路径", command=self.do_browse, width=9).pack(side="left", padx=4)
        self.btn_start = ttk.Button(f, text="开始同步", command=self.do_start, width=9)
        self.btn_start.pack(side="left", padx=4)
        self.btn_stop = ttk.Button(f, text="停止", command=self.do_stop, width=5, state="disabled")
        self.btn_stop.pack(side="left")
        
        # 路径
        f = ttk.Frame(m); f.pack(fill="x", pady=(0,8))
        ttk.Label(f, text="网易云路径:").pack(side="left")
        self.path = tk.StringVar(value=self.cfg.get("ncm_path") or "(自动检测)")
        ttk.Label(f, textvariable=self.path, foreground="gray").pack(side="left", padx=4)
        
        # 基础设置
        f = ttk.LabelFrame(m, text="基础设置", padding=6); f.pack(fill="x", pady=(0,8))
        g = ttk.Frame(f); g.pack(fill="x")
        ttk.Label(g, text="OSC").grid(row=0, column=0)
        self.e_ip = ttk.Entry(g, width=11); self.e_ip.insert(0, self.cfg["osc_ip"]); self.e_ip.grid(row=0, column=1)
        ttk.Label(g, text=":").grid(row=0, column=2)
        self.e_port = ttk.Entry(g, width=5); self.e_port.insert(0, self.cfg["osc_port"]); self.e_port.grid(row=0, column=3)
        ttk.Label(g, text="刷新").grid(row=0, column=4, padx=(12,0))
        self.v_interval = tk.DoubleVar(value=self.cfg["refresh_interval"])
        ttk.Spinbox(g, from_=2, to=10, increment=0.5, width=4, textvariable=self.v_interval).grid(row=0, column=5)
        ttk.Label(g, text="秒").grid(row=0, column=6)
        
        # 进度条设置
        g2 = ttk.Frame(f); g2.pack(fill="x", pady=(6,0))
        ttk.Label(g2, text="进度条").pack(side="left")
        ttk.Label(g2, text="宽度").pack(side="left", padx=(8,0))
        self.e_bw = ttk.Entry(g2, width=3); self.e_bw.insert(0, self.cfg["bar_width"]); self.e_bw.pack(side="left", padx=2)
        ttk.Label(g2, text="已播放").pack(side="left", padx=(8,0))
        self.e_bf = ttk.Entry(g2, width=3); self.e_bf.insert(0, self.cfg["bar_filled"]); self.e_bf.pack(side="left", padx=2)
        ttk.Label(g2, text="滑块").pack(side="left", padx=(8,0))
        self.e_bt = ttk.Entry(g2, width=3); self.e_bt.insert(0, self.cfg.get("bar_thumb", "●")); self.e_bt.pack(side="left", padx=2)
        ttk.Label(g2, text="未播放").pack(side="left", padx=(8,0))
        self.e_be = ttk.Entry(g2, width=3); self.e_be.insert(0, self.cfg["bar_empty"]); self.e_be.pack(side="left", padx=2)
        
        # 输出模板
        f = ttk.LabelFrame(m, text="输出模板", padding=6); f.pack(fill="both", expand=True, pady=(0,8))
        ttk.Label(f, text="可用变量: {song} {artist} {bar} {time} {lyric1} {lyric2}", foreground="gray").pack(anchor="w")
        self.t_tpl = tk.Text(f, height=3, font=("Consolas",10)); self.t_tpl.insert("1.0", self.cfg["template"]); self.t_tpl.pack(fill="both", expand=True, pady=(4,0))
        
        # 实时预览
        f = ttk.LabelFrame(m, text="文本预览", padding=6); f.pack(fill="both", expand=True)
        self.t_preview = tk.Text(f, height=3, font=("Consolas",10), state="disabled", bg="#f5f5f5")
        self.t_preview.pack(fill="both", expand=True)

        for w in [self.t_tpl, self.e_bw, self.e_bf, self.e_bt, self.e_be]:
            w.bind("<KeyRelease>", lambda e: self.preview())
        self.preview()
    
    def do_browse(self):
        p = filedialog.askopenfilename(title="选择cloudmusic.exe", filetypes=[("","cloudmusic.exe"),("","*.exe")])
        if p: self.cfg["ncm_path"] = p; self.path.set(p); self.save_cfg()
    
    def do_launch(self):
        ok, r, port = launch_netease(None, self.cfg.get("ncm_path"))
        if ok:
            self.cfg["ncm_port"] = port
            self.status.set(f"已启动 (端口:{port})")
            self.path.set(r)
            self.root.after(3000, lambda: self.status.set("就绪"))
        else: messagebox.showwarning("提示", f"{r}\n请手动选择路径")
    
    def preview(self):
        try:
            w = int(self.e_bw.get() or 10)
            thumb = self.e_bt.get() or ""
            pos = w // 2
            bar = (self.e_bf.get() or "▓") * pos + thumb + (self.e_be.get() or "░") * (w - pos)
            txt = self.t_tpl.get("1.0","end").strip().format(song="歌曲名称", artist="歌手", bar=bar, time="1:14/5:14", lyric1="当前歌词", lyric2="下句歌词")
            self.t_preview.config(state="normal"); self.t_preview.delete("1.0","end"); self.t_preview.insert("1.0",txt); self.t_preview.config(state="disabled")
        except: pass
    
    def get_cfg(self):
        return {**self.cfg, "osc_ip": self.e_ip.get(), "osc_port": int(self.e_port.get() or 9000),
                "refresh_interval": max(2, min(10, self.v_interval.get())), "template": self.t_tpl.get("1.0","end").strip(),
                "bar_width": int(self.e_bw.get() or 9), "bar_filled": self.e_bf.get() or "▓", 
                "bar_thumb": self.e_bt.get() or "", "bar_empty": self.e_be.get() or "░"}
    
    def cb_status(self, t): self.root.after(0, lambda: self.status.set(t))
    def cb_song(self, t): self.root.after(0, lambda: self.song.set(t[:28]))
    def cb_output(self, t):
        def f(): self.t_preview.config(state="normal"); self.t_preview.delete("1.0","end"); self.t_preview.insert("1.0",t); self.t_preview.config(state="disabled")
        self.root.after(0, f)
    
    def do_start(self):
        try: self.cfg = self.get_cfg(); self.save_cfg()
        except Exception as e: messagebox.showerror("错误", str(e)); return
        self.sync = Sync(self.cfg, {"status": self.cb_status, "song": self.cb_song, "output": self.cb_output})
        threading.Thread(target=lambda: asyncio.run(self.sync.run()), daemon=True).start()
        self.btn_start.config(state="disabled"); self.btn_stop.config(state="normal"); self.btn_launch.config(state="disabled")
    
    def do_stop(self):
        if self.sync: self.sync.stop()
        self.btn_start.config(state="normal"); self.btn_stop.config(state="disabled"); self.btn_launch.config(state="normal")
        self.status.set("已停止"); self.song.set("")
    
    def run(self):
        self.root.protocol("WM_DELETE_WINDOW", lambda: (self.do_stop(), self.root.destroy()))
        self.root.mainloop()


if __name__ == "__main__":
    App().run()
