import threading


class SongState:
    def __init__(
        self, song="", artist="", cur=0, dur=0, play=False, lyric1="", lyric2=""
    ):
        self.song = song
        self.artist = artist
        self.cur = cur
        self.dur = dur
        self.play = play
        self.lyric1 = lyric1
        self.lyric2 = lyric2

    def update(self, d: dict):
        for k in ["song", "artist", "cur", "dur", "play", "lyric1", "lyric2"]:
            if k in d:
                setattr(self, k, d[k])

    def copy(self):
        return SongState(
            song=self.song,
            artist=self.artist,
            cur=self.cur,
            dur=self.dur,
            play=self.play,
            lyric1=self.lyric1,
            lyric2=self.lyric2,
        )


class SharedState:
    def __init__(self):
        self.lock = threading.Lock()
        self.data = SongState()
        self.song_key = ""
        self.lyrics = []
