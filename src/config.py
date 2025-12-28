from pydantic import BaseModel


class Config(BaseModel):
    osc_ip: str = "127.0.0.1"
    osc_port: int = 9000
    ncm_port: int = 9222
    ncm_path: str = ""
    refresh_interval: float = 3.0
    bar_width: int = 9
    bar_filled: str = "▓"
    bar_empty: str = "░"
    bar_thumb: str = "◘"
    template: str = "🎵 {song} - {artist}\n{bar} {time}\n{lyric1}\n{lyric2}"
