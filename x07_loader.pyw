import configparser
from pathlib import Path
import re
import subprocess
import sys
import threading
import time
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

try:
    import serial
    import serial.tools.list_ports
    from serial.serialutil import SerialException
except ImportError:
    subprocess.check_call([sys.executable, "-m", "pip", "install", "pyserial"])
    import serial
    import serial.tools.list_ports
    from serial.serialutil import SerialException

try:
    import termios
    HAS_TERMIOS = True
except ImportError:
    HAS_TERMIOS = False

# ---------- Defaults ----------
DEFAULT_CHAR_DELAY_S = 0.04
DEFAULT_LINE_DELAY_S = 0.20
DEFAULT_LOAD_ADDR = 0x2000
DEFAULT_LOADER_ADDR = 0x1800
POST_LOADER_EXEC_DELAY_S = 3.0
ASM_PRIMER = b"X" * 1024
LOADER_CAS_NAME = "loader.cas"
LOADER_BYTE_COUNT = 162

CAS_FIXED_BASE = 0x0010  # fixed (validated for SEND)
SAVE_IDLE_TIMEOUT_S = 1.25  # end capture if no bytes for this duration (after any data received)

# ---------- X-07 key codes ----------
KEY_ON_BREAK = 0x01
KEY_RETURN   = 0x0D
KEY_SPACE    = 0x20

KEY_HOME     = 0x0B
KEY_CLR      = 0x0C

KEY_INS      = 0x12
KEY_DEL      = 0x16

KEY_RIGHT    = 0x1C
KEY_LEFT     = 0x1D
KEY_UP       = 0x1E
KEY_DOWN     = 0x1F


# ---------- Canon X-07 text encoding ----------
# Imported from basX07.c for compatibility with Canon X-07 national/special characters.
# The mappings are intentionally kept aligned with the original C tables.
X07_UNICODE_MAP = {
    "¥": 0x5C,
    "¿": 0x7F,
    "♠": 0x80,
    "♥": 0x81,
    "♣": 0x82,
    "♦": 0x83,
    "○": 0x84,
    "●": 0x85,
    "Ä": 0x86,
    "Å": 0x87,
    "ä": 0x88,
    "à": 0x89,
    "â": 0x8A,
    "á": 0x8B,
    "å": 0x8C,
    "a̱": 0x8D,
    "Ï": 0x8E,
    "ï": 0x8F,
    "ì": 0x90,
    "î": 0x91,
    "í": 0x92,
    "Ü": 0x93,
    "ü": 0x94,
    "ù": 0x95,
    "û": 0x96,
    "ú": 0x97,
    "É": 0x98,
    "ë": 0x99,
    "è": 0x9A,
    "ê": 0x9B,
    "é": 0x9C,
    "Ö": 0x9D,
    "ö": 0x9E,
    "ò": 0x9F,
    "√": 0xA0,
    "・": 0xA1,
    "「": 0xA2,
    "」": 0xA3,
    "、": 0xA4,
    "。": 0xA5,
    "ヲ": 0xA6,
    "ァ": 0xA7,
    "ィ": 0xA8,
    "ゥ": 0xA9,
    "ェ": 0xAA,
    "ォ": 0xAB,
    "ャ": 0xAC,
    "ュ": 0xAD,
    "ョ": 0xAE,
    "ッ": 0xAF,
    "ー": 0xB0,
    "ア": 0xB1,
    "イ": 0xB2,
    "ウ": 0xB3,
    "エ": 0xB4,
    "オ": 0xB5,
    "カ": 0xB6,
    "キ": 0xB7,
    "ク": 0xB8,
    "ケ": 0xB9,
    "コ": 0xBA,
    "サ": 0xBB,
    "シ": 0xBC,
    "ヌ": 0xBD,
    "ャ": 0xBE,
    "ソ": 0xBF,
    "タ": 0xC0,
    "チ": 0xC1,
    "ツ": 0xC2,
    "テ": 0xC3,
    "ト": 0xC4,
    "ナ": 0xC5,
    "ニ": 0xC6,
    "ヌ": 0xC7,
    "ネ": 0xC8,
    "ノ": 0xC9,
    "ハ": 0xCA,
    "ヒ": 0xCB,
    "フ": 0xCC,
    "ヘ": 0xCD,
    "ホ": 0xCE,
    "マ": 0xCF,
    "ミ": 0xD0,
    "ム": 0xD1,
    "メ": 0xD2,
    "モ": 0xD3,
    "ヤ": 0xD4,
    "ユ": 0xD5,
    "ヨ": 0xD6,
    "ラ": 0xD7,
    "リ": 0xD8,
    "ル": 0xD9,
    "レ": 0xDA,
    "ロ": 0xDB,
    "ワ": 0xDC,
    "ン": 0xDD,
    "゛": 0xDE,
    "゜": 0xDF,
    "ô": 0xE0,
    "ó": 0xE1,
    "o̱": 0xE2,
    "ÿ": 0xE3,
    "Ç": 0xE4,
    "ç": 0xE5,
    "Ñ": 0xE6,
    "ñ": 0xE7,
    "Γ": 0xE8,
    "Σ": 0xE9,
    "Π": 0xEA,
    "Ω": 0xEB,
    "α": 0xEC,
    "β": 0xED,
    "γ": 0xEE,
    "δ": 0xEF,
    "ε": 0xF0,
    "ζ": 0xF1,
    "θ": 0xF2,
    "κ": 0xF3,
    "λ": 0xF4,
    "μ": 0xF5,
    "ρ": 0xF6,
    "π": 0xF7,
    "τ": 0xF8,
    "ϕ": 0xF9,
    "χ": 0xFA,
    "ω": 0xFB,
    "ν": 0xFC,
    "£": 0xFD,
    "¢": 0xFE,
    "÷": 0xFF,
}

X07_ESCAPE_MAP = {
    "\\": 0x5C,
    "YN": 0x5C,
    "?": 0x7F,
    "SP": 0x80,
    "HT": 0x81,
    "DI": 0x82,
    "CL": 0x83,
    "@": 0x84,
    "LD": 0x85,
    ":A": 0x86,
    ".A": 0x87,
    "AN": 0x87,
    ":a": 0x88,
    "`a": 0x89,
    "^a": 0x8A,
    "'a": 0x8B,
    ".a": 0x8C,
    "_a": 0x8D,
    ":I": 0x8E,
    ":i": 0x8F,
    "`i": 0x90,
    "^i": 0x91,
    "'i": 0x92,
    ":U": 0x93,
    ":u": 0x94,
    "`u": 0x95,
    "^u": 0x96,
    "'u": 0x97,
    "'E": 0x98,
    ":e": 0x99,
    "`e": 0x9A,
    "^e": 0x9B,
    "'e": 0x9C,
    ":O": 0x9D,
    ":o": 0x9E,
    "`o": 0x9F,
    "RT": 0xA0,
    "^o": 0xE0,
    "'o": 0xE1,
    "_o": 0xE2,
    ":y": 0xE3,
    ",C": 0xE4,
    ",c": 0xE5,
    "~N": 0xE6,
    "~n": 0xE7,
    "GA": 0xE8,
    "SI": 0xE9,
    "SM": 0xE9,
    "PI": 0xEA,
    "OM": 0xEB,
    "al": 0xEC,
    "bt": 0xED,
    "ga": 0xEE,
    "dl": 0xEF,
    "ep": 0xF0,
    "si": 0xF1,
    "th": 0xF2,
    "ka": 0xF3,
    "la": 0xF4,
    "mu": 0xF5,
    "rh": 0xF6,
    "pi": 0xF7,
    "ta": 0xF8,
    "ps": 0xF9,
    "ch": 0xFA,
    "om": 0xFB,
    "nu": 0xFC,
    "PN": 0xFD,
    "CN": 0xFE,
    ":-": 0xFF,
}

_X07_UNICODE_KEYS = sorted(X07_UNICODE_MAP, key=len, reverse=True)
_X07_ESCAPE_KEYS = sorted(X07_ESCAPE_MAP, key=len, reverse=True)


def _match_escape_token(text: str) -> tuple[int | None, int]:
    for key in _X07_ESCAPE_KEYS:
        if text.startswith(key):
            return X07_ESCAPE_MAP[key], len(key)
    lower = text.lower()
    for key in _X07_ESCAPE_KEYS:
        if lower.startswith(key.lower()):
            return X07_ESCAPE_MAP[key], len(key)
    if len(text) >= 2 and all(c in '0123456789abcdefABCDEF' for c in text[:2]):
        return int(text[:2], 16), 2
    return None, 0


def x07_encode_text(text: str) -> bytes:
    """Encode a Python string to Canon X-07 text bytes.

    Supports direct Unicode characters (accented latin, Greek, katakana, symbols)
    and the original basX07-style backslash escapes such as \\'e, \\PI or \\A0.
    Unknown characters fall back to '?' so transfers remain robust.
    """
    out = bytearray()
    i = 0
    n = len(text)
    while i < n:
        if text[i] == "\\":
            token, consumed = _match_escape_token(text[i + 1:])
            if token is not None:
                out.append(token)
                i += 1 + consumed
                continue
            out.append(0x5C)
            i += 1
            continue

        matched = False
        for key in _X07_UNICODE_KEYS:
            if text.startswith(key, i):
                out.append(X07_UNICODE_MAP[key])
                i += len(key)
                matched = True
                break
        if matched:
            continue

        ch = text[i]
        code = ord(ch)
        if ch in "\r\n":
            out.append(code)
        elif 0x20 <= code <= 0x7E:
            out.append(code)
        else:
            out.append(ord("?"))
        i += 1
    return bytes(out)


def list_serial_ports():
    ports = [p.device for p in serial.tools.list_ports.comports()]
    if sys.platform == "darwin":
        cu_ports = [p for p in ports if p.startswith("/dev/cu.")]
        if cu_ports:
            return cu_ports
    return ports


def guess_name_in_first_16_bytes(data: bytes, fallback: str) -> str:
    if len(data) < 0x10:
        return fallback

    txt = "".join(chr(b) if 0x20 <= b <= 0x7E else " " for b in data[:0x10])
    tokens = re.findall(r"[A-Za-z0-9][A-Za-z0-9 _\-\[\]]{2,24}", txt)
    return max(tokens, key=len).strip() if tokens else fallback


def build_cas_header_from_filename(path: Path) -> bytes:
    name = (path.stem or "").upper()

    safe = []
    for ch in name:
        if ("A" <= ch <= "Z") or ("0" <= ch <= "9"):
            safe.append(ch)
        else:
            safe.append("_")
    name = "".join(safe)[:6]  # max 6
    name_field = name.encode("ascii", errors="replace").ljust(6, b"\x00")

    return (b"\xD3" * 10) + name_field


class AutoScrollFrame(ttk.Frame):
    def __init__(self, master):
        super().__init__(master)
        self.canvas = tk.Canvas(self, highlightthickness=0)
        self.scroll = ttk.Scrollbar(self, orient="vertical", command=self.canvas.yview)
        self.inner = ttk.Frame(self.canvas)

        self.win = self.canvas.create_window((0, 0), window=self.inner, anchor="nw")
        self.canvas.configure(yscrollcommand=self.scroll.set)

        self.canvas.pack(side="left", fill="both", expand=True)
        self.scroll.pack(side="right", fill="y")

        self.inner.bind("<Configure>", self._sync)
        self.canvas.bind("<Configure>", self._sync)

        self.canvas.bind_all("<MouseWheel>", self._wheel, add="+")
        self.canvas.bind_all("<Button-4>", self._wheel_linux, add="+")
        self.canvas.bind_all("<Button-5>", self._wheel_linux, add="+")

    def _needs_scroll(self) -> bool:
        try:
            inner_h = self.inner.winfo_reqheight()
            canvas_h = self.canvas.winfo_height()
            return inner_h > canvas_h + 2
        except Exception:
            return False

    def _sync(self, _=None):
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        self.canvas.itemconfigure(self.win, width=self.canvas.winfo_width())

        if self._needs_scroll():
            if not self.scroll.winfo_ismapped():
                self.scroll.pack(side="right", fill="y")
            self.canvas.itemconfigure(self.win, height=self.inner.winfo_reqheight())
        else:
            if self.scroll.winfo_ismapped():
                self.scroll.pack_forget()
            self.canvas.itemconfigure(self.win, height=max(1, self.canvas.winfo_height()))

        self.canvas.configure(scrollregion=self.canvas.bbox("all"))

    def _wheel(self, event):
        if self._needs_scroll():
            self.canvas.yview_scroll(-int(event.delta / 120), "units")

    def _wheel_linux(self, event):
        if not self._needs_scroll():
            return
        self.canvas.yview_scroll(-1 if event.num == 4 else 1, "units")


class X07LoaderApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Canon X-07 Serial Fast Loader")
        self.geometry("1040x720")
        self.minsize(920, 620)

        # files
        self.basic_file: Path | None = None
        self.bin_file: Path | None = None
        self.cas_file: Path | None = None

        # job control
        self.cancel_event = threading.Event()
        self.job_lock = threading.Lock()
        self.job_running = False

        # remote keyboard (persistent session)
        self.remote_kbd_on = False
        self.remote_ser: serial.Serial | None = None

        # widgets
        self._transfer_controls: list[ttk.Widget] = []
        self._always_enabled_controls: list[ttk.Widget] = []
        self._kbd_controls: list[ttk.Widget] = []  # enabled only when REMOTE KEYBOARD ON

        self._build_ui()
        self.refresh_ports(initial=True)
        self._log_startup_banner()

    # ---------------- UI ----------------
    def _build_ui(self):
        root = ttk.Frame(self, padding=6)
        root.pack(fill="both", expand=True)

        sf = AutoScrollFrame(root)
        sf.pack(fill="both", expand=True)
        main = sf.inner

        # ---- Serial settings ----
        serial_box = ttk.LabelFrame(main, text="Serial settings", padding=6)
        serial_box.pack(fill="x", pady=(0, 6))

        self.var_port = tk.StringVar(value="")
        self.var_port.trace_add("write", lambda *_: self._save_serial_settings())
        self.var_typing_baud = tk.IntVar(value=4800)   # 8N2
        self.var_xfer_baud = tk.IntVar(value=8000)     # 8N2 loader/runtime

        self.var_char = tk.DoubleVar(value=DEFAULT_CHAR_DELAY_S)
        self.var_line = tk.DoubleVar(value=DEFAULT_LINE_DELAY_S)
        self.var_rtscts = tk.BooleanVar(value=False)

        self.var_rtscts.trace_add("write", lambda *_: self._save_serial_settings())
        self.var_typing_baud.trace_add("write", lambda *_: self._save_serial_settings())
        self.var_xfer_baud.trace_add("write", lambda *_: self._save_serial_settings())
        self.var_char.trace_add("write", lambda *_: self._save_serial_settings())
        self.var_line.trace_add("write", lambda *_: self._save_serial_settings())

        r = ttk.Frame(serial_box)
        r.pack(fill="x")

        ttk.Label(r, text="COM:").pack(side="left")
        self.cbo_port = ttk.Combobox(r, textvariable=self.var_port, width=10, state="readonly")
        self.cbo_port.pack(side="left", padx=(4, 6))

        btn_refresh = ttk.Button(r, text="Refresh", command=self.refresh_ports)
        btn_refresh.pack(side="left", padx=(0, 10))
        self._always_enabled_controls.append(btn_refresh)

        ttk.Label(r, text="Typing baud (8N2):").pack(side="left")
        ttk.Entry(r, textvariable=self.var_typing_baud, width=7).pack(side="left", padx=(4, 10))

        ttk.Label(r, text="Loader baud (8N2):").pack(side="left")
        ttk.Entry(r, textvariable=self.var_xfer_baud, width=7).pack(side="left", padx=(4, 10))

        self.chk_rtscts = ttk.Checkbutton(
            r,
            text="RTS/CTS cable",
            variable=self.var_rtscts,
            command=self._update_handshake_ui,
        )
        self.chk_rtscts.pack(side="left", padx=(10, 10))

        self.lbl_char = ttk.Label(r, text="CHAR(s):")
        self.lbl_char.pack(side="left")
        self.ent_char = ttk.Entry(r, textvariable=self.var_char, width=6)
        self.ent_char.pack(side="left", padx=(3, 8))

        self.lbl_line = ttk.Label(r, text="LINE(s):")
        self.lbl_line.pack(side="left")
        self.ent_line = ttk.Entry(r, textvariable=self.var_line, width=6)
        self.ent_line.pack(side="left", padx=(3, 8))


        # cancel/disable slave row
        r2 = ttk.Frame(serial_box)
        r2.pack(fill="x", pady=(4, 0))
        self.btn_cancel = ttk.Button(r2, text="Cancel transfer", command=self.cancel_current, state="disabled")
        self.btn_cancel.pack(side="left")

        btn_disable_slave = ttk.Button(r2, text="Disable slave (EXEC&HEE33)", command=self.disable_slave_mode)
        btn_disable_slave.pack(side="left", padx=(6, 0))
        self._transfer_controls.append(btn_disable_slave)

        ttk.Label(
            r2,
            text='Note: BASIC TXT and REMOTE KEYBOARD use SLAVE mode (INIT#5,"COM:" then EXEC&HEE1F). ASM uses loader.cas + ASCII fast loader.',
        ).pack(side="left", padx=(10, 0))

        # ---- BASIC ----
        basic_box = ttk.LabelFrame(main, text="BASIC", padding=6)
        basic_box.pack(fill="x", pady=(0, 6))
        cols = ttk.Frame(basic_box)
        cols.pack(fill="x")
        left = ttk.Frame(cols)
        right = ttk.Frame(cols)
        left.pack(side="left", fill="both", expand=True, padx=(0, 6))
        right.pack(side="left", fill="both", expand=True, padx=(6, 0))

        # CAS/K7 (LEFT)
        cas = ttk.LabelFrame(left, text='Cassette stream (.cas/.k7) via LOAD "COM:" / SAVE "COM:" (raw bytes)', padding=6)
        cas.pack(fill="both", expand=True)

        self.lbl_cas = ttk.Label(cas, text="Selected CAS/K7: (none)")
        self.lbl_cas.pack(anchor="w")

        rc = ttk.Frame(cas)
        rc.pack(fill="x", pady=(4, 0))
        btn_pick_cas = ttk.Button(rc, text="Select .cas/.k7…", command=self.pick_cas)
        btn_pick_cas.pack(side="left")
        self._transfer_controls.append(btn_pick_cas)

        btn_inspect = ttk.Button(rc, text="Inspect header", command=self.inspect_cas_header)
        btn_inspect.pack(side="left", padx=(6, 0))
        self._always_enabled_controls.append(btn_inspect)

        btn_send_cas = ttk.Button(rc, text='Send raw (LOAD "COM:")', command=self.send_cas_raw)
        btn_send_cas.pack(side="left", padx=(6, 0))
        self._transfer_controls.append(btn_send_cas)

        btn_recv_cas = ttk.Button(rc, text='Receive raw (SAVE "COM:")', command=self.receive_cas_raw)
        btn_recv_cas.pack(side="left", padx=(6, 0))
        self._transfer_controls.append(btn_recv_cas)

        # TXT/BAS (RIGHT)
        txt = ttk.LabelFrame(right, text="Text listing (.txt/.bas) via SLAVE mode", padding=6)
        txt.pack(fill="both", expand=True)

        self.lbl_basic = ttk.Label(txt, text="Selected BASIC: (none)")
        self.lbl_basic.pack(anchor="w")

        rb = ttk.Frame(txt)
        rb.pack(fill="x", pady=(4, 0))
        btn_pick_basic = ttk.Button(rb, text="Select .txt/.bas…", command=self.pick_basic)
        btn_pick_basic.pack(side="left")
        self._transfer_controls.append(btn_pick_basic)

        btn_send_basic = ttk.Button(rb, text="Send BASIC", command=self.send_basic_file)
        btn_send_basic.pack(side="left", padx=(6, 0))
        self._transfer_controls.append(btn_send_basic)

        # ---- ASM ----
        asm = ttk.LabelFrame(main, text="ASM via loader.cas + ASCII fast loader", padding=6)
        asm.pack(fill="x", pady=(0, 6))

        self.var_loader_addr = tk.StringVar(value=hex(DEFAULT_LOADER_ADDR))
        self.var_addr = tk.StringVar(value=hex(DEFAULT_LOAD_ADDR))
        self.var_loader_addr.trace_add("write", lambda *_: self._save_serial_settings())
        self.var_addr.trace_add("write", lambda *_: self._save_serial_settings())

        self.lbl_bin = ttk.Label(asm, text="Selected ASM binary: (none)")
        self.lbl_bin.pack(anchor="w")

        ra = ttk.Frame(asm)
        ra.pack(fill="x", pady=(4, 0))
        btn_pick_bin = ttk.Button(ra, text="Select bin…", command=self.pick_bin)
        btn_pick_bin.pack(side="left")
        self._transfer_controls.append(btn_pick_bin)

        ttk.Label(ra, text="Loader addr:").pack(side="left", padx=(10, 4))
        ttk.Entry(ra, textvariable=self.var_loader_addr, width=10).pack(side="left")

        ttk.Label(ra, text="ASM addr:").pack(side="left", padx=(10, 4))
        ttk.Entry(ra, textvariable=self.var_addr, width=10).pack(side="left")

        btn_send_loader = ttk.Button(ra, text='Send ASM loader (LOAD"COM:")', command=self.send_fast_loader)
        btn_send_loader.pack(side="left", padx=(10, 0))
        self._transfer_controls.append(btn_send_loader)

        btn_send_asm = ttk.Button(ra, text="Send ASM (loader running)", command=self.send_bin_only)
        btn_send_asm.pack(side="left", padx=(10, 0))
        self._transfer_controls.append(btn_send_asm)

        btn_one_click = ttk.Button(ra, text="One click: send loader + ASM (SLAVE mode)", command=self.send_loader_and_bin)
        btn_one_click.pack(side="left", padx=(10, 0))
        self._transfer_controls.append(btn_one_click)

        # ---- Remote keyboard ----
        kbd = ttk.LabelFrame(main, text='Remote keyboard via SLAVE mode (PC -> X-07)', padding=6)
        kbd.pack(fill="x", pady=(0, 6))

        rk = ttk.Frame(kbd)
        rk.pack(fill="x")
        self.btn_remote_toggle = ttk.Button(rk, text="REMOTE KEYBOARD: OFF", command=self.toggle_remote_keyboard)
        self.btn_remote_toggle.pack(side="left")
        self._always_enabled_controls.append(self.btn_remote_toggle)

        ttk.Label(rk, text='(requires SLAVE mode: INIT#5,"COM:" then EXEC&HEE1F)').pack(side="left", padx=(10, 0))

        row_btns = ttk.Frame(kbd)
        row_btns.pack(fill="x", pady=(6, 0))

        def kbtn(label: str, cmd, width=6):
            b = ttk.Button(row_btns, text=label, width=width, command=cmd)
            b.pack(side="left", padx=(2, 0))
            self._kbd_controls.append(b)
            return b

        kbtn("HOME",  lambda: self._kbd_send_byte(KEY_HOME))
        kbtn("CLR",   lambda: self._kbd_send_byte(KEY_CLR), width=5)
        kbtn("INS",   lambda: self._kbd_send_byte(KEY_INS), width=5)
        kbtn("DEL",   lambda: self._kbd_send_byte(KEY_DEL), width=5)
        kbtn("BREAK", lambda: self._kbd_send_byte(KEY_ON_BREAK), width=6)

        ttk.Separator(row_btns, orient="vertical").pack(side="left", fill="y", padx=8)

        def macro(label: str, text: str, w=7):
            b = ttk.Button(
                row_btns,
                text=label,
                width=w,
                command=lambda: self._kbd_send_bytes(x07_encode_text(text)),
            )
            b.pack(side="left", padx=(2, 0))
            self._kbd_controls.append(b)

        macro("?TIME$",  "?TIME$\r")
        macro("?DATE$",  "?DATE$\r")
        macro("CLOAD",   'CLOAD"\r')
        macro("CSAVE",   'CSAVE"\r')
        macro("LOCATE",  "LOCATE ")
        macro("PRINT",   "PRINT ")
        macro("LIST",    "LIST ")
        macro("RUN",     "RUN\r")
        macro("SLEEP",   "SLEEP")
        macro("CONT",    "CONT\r")

        ttk.Separator(row_btns, orient="vertical").pack(side="left", fill="y", padx=8)

        kbtn("←", lambda: self._kbd_send_byte(KEY_LEFT), width=3)
        kbtn("↑", lambda: self._kbd_send_byte(KEY_UP), width=3)
        kbtn("↓", lambda: self._kbd_send_byte(KEY_DOWN), width=3)
        kbtn("→", lambda: self._kbd_send_byte(KEY_RIGHT), width=3)

        ttk.Label(kbd, text="Type here (sent to X-07 while REMOTE KEYBOARD is ON):").pack(anchor="w", pady=(6, 0))
        self.relay_box = tk.Text(kbd, height=3, wrap="none")
        self.relay_box.pack(fill="x", expand=False)
        self.relay_box.bind("<KeyPress>", self._on_remote_keypress)
        self._kbd_controls.append(self.relay_box)

        # ---- Common progress + Console ----
        bottom = ttk.LabelFrame(main, text="Progress / Console", padding=6)
        bottom.pack(fill="both", expand=True, pady=(0, 0))

        self.progress_var = tk.DoubleVar(value=0)
        self.progress = ttk.Progressbar(bottom, variable=self.progress_var, maximum=100)
        self.progress.pack(fill="x")
        self.progress_label = ttk.Label(bottom, text="Progress: idle")
        self.progress_label.pack(anchor="w", pady=(2, 0))

        self.status_var = tk.StringVar(value="Status: idle")
        ttk.Label(bottom, textvariable=self.status_var).pack(anchor="w")

        self.txt = tk.Text(bottom, height=12, wrap="word")
        self.txt.pack(fill="both", expand=True, pady=(4, 0))

        self._set_keyboard_controls_enabled(False)
        self._update_handshake_ui()

    def _update_handshake_ui(self):
        return

    # ---------------- Logging ----------------
    def _ts(self) -> str:
        return time.strftime("[%H:%M:%S]")

    def log(self, msg: str):
        self.txt.insert("end", f"{self._ts()} {msg}\n")
        self.txt.see("end")
        self.update_idletasks()

    def _log_startup_banner(self):
        self.log("Canon X-07 Serial Fast Loader - ready.")
        self.log('SLAVE mode (BASIC TXT + REMOTE KEYBOARD): INIT#5,"COM:" then EXEC&HEE1F (user guide p.119).')
        self.log('CAS/K7 raw stream: use LOAD"COM:" or SAVE"COM:" on X-07, then send/receive raw bytes on PC.')
        self.log("Exit slave: EXEC&HEE33 (remote) or power cycle.")
        self.log('Enable "RTS/CTS cable" when using a hardware-handshaked cable. Delay fields stay available and are still used for typing / loader send.')
        self.log("")

    # ---------------- UI enabling/disabling ----------------
    def _set_controls_enabled(self, enabled: bool):
        for w in self._transfer_controls:
            try:
                w.configure(state=("normal" if enabled else "disabled"))
            except Exception:
                pass
        for w in self._always_enabled_controls:
            try:
                w.configure(state="normal")
            except Exception:
                pass

    def _set_keyboard_controls_enabled(self, enabled: bool):
        for w in self._kbd_controls:
            try:
                if isinstance(w, tk.Text):
                    w.configure(state=("normal" if enabled else "disabled"))
                else:
                    w.configure(state=("normal" if enabled else "disabled"))
            except Exception:
                pass

    # ---------------- Ports ----------------
    def _config_path(self) -> Path:
        return Path(__file__).with_suffix(".ini")  # x07_loader.ini

    def _load_serial_settings(self) -> dict:
        cfg_path = self._config_path()
        settings = {
            "port": None,
            "rtscts": False,
            "typing_baud": 4800,
            "xfer_baud": 8000,
            "char_delay": DEFAULT_CHAR_DELAY_S,
            "line_delay": DEFAULT_LINE_DELAY_S,
            "loader_addr": hex(DEFAULT_LOADER_ADDR),
            "asm_addr": hex(DEFAULT_LOAD_ADDR),
        }
        if not cfg_path.exists():
            return settings

        cfg = configparser.ConfigParser()
        try:
            cfg.read(cfg_path, encoding="utf-8")
            settings["port"] = cfg.get("serial", "port", fallback="").strip() or None
            settings["rtscts"] = cfg.getboolean("serial", "rtscts", fallback=False)
            settings["typing_baud"] = cfg.getint("serial", "typing_baud", fallback=4800)
            settings["xfer_baud"] = cfg.getint("serial", "xfer_baud", fallback=8000)
            settings["char_delay"] = cfg.getfloat("serial", "char_delay", fallback=DEFAULT_CHAR_DELAY_S)
            settings["line_delay"] = cfg.getfloat("serial", "line_delay", fallback=DEFAULT_LINE_DELAY_S)
            settings["loader_addr"] = cfg.get("serial", "loader_addr", fallback=hex(DEFAULT_LOADER_ADDR))
            settings["asm_addr"] = cfg.get("serial", "asm_addr", fallback=hex(DEFAULT_LOAD_ADDR))
        except Exception:
            pass
        return settings

    def _save_serial_settings(self) -> None:
        cfg_path = self._config_path()
        cfg = configparser.ConfigParser()
        cfg["serial"] = {
            "port": self.var_port.get(),
            "rtscts": str(self.var_rtscts.get()),
            "typing_baud": str(self.var_typing_baud.get()),
            "xfer_baud": str(self.var_xfer_baud.get()),
            "char_delay": str(self.var_char.get()),
            "line_delay": str(self.var_line.get()),
            "loader_addr": self.var_loader_addr.get(),
            "asm_addr": self.var_addr.get(),
        }

        try:
            with cfg_path.open("w", encoding="utf-8") as f:
                cfg.write(f)
        except Exception:
            pass

    def refresh_ports(self, initial=False):
        ports = list_serial_ports()
        self.cbo_port["values"] = ports

        cur = (self.var_port.get() or "").strip()
        settings = self._load_serial_settings()
        self.var_rtscts.set(bool(settings.get("rtscts", False)))
        self.var_typing_baud.set(int(settings.get("typing_baud", 4800)))
        self.var_xfer_baud.set(int(settings.get("xfer_baud", 8000)))
        self.var_char.set(float(settings.get("char_delay", DEFAULT_CHAR_DELAY_S)))
        self.var_line.set(float(settings.get("line_delay", DEFAULT_LINE_DELAY_S)))
        self.var_loader_addr.set(str(settings.get("loader_addr", hex(DEFAULT_LOADER_ADDR))))
        self.var_addr.set(str(settings.get("asm_addr", hex(DEFAULT_LOAD_ADDR))))
        last = (settings.get("port") or "").strip()

        if cur and cur in ports:
            chosen = cur
        elif last and last in ports:
            chosen = last
        else:
            chosen = ports[0] if ports else ""

        self.var_port.set(chosen)
        self._update_handshake_ui()

        if initial:
            if ports:
                self.log(f"[INFO] COM ports loaded: {', '.join(ports[:10])}{'...' if len(ports) > 10 else ''}")
            else:
                self.log("[WARN] No COM ports detected. Plug adapter and click Refresh.")
        else:
            self.log("[INFO] COM ports refreshed.")

    # ---------------- Job control ----------------
    def cancel_current(self):
        self.cancel_event.set()
        self.log("[WARN] Cancel requested...")

    def _job_start(self):
        with self.job_lock:
            if self.job_running:
                raise RuntimeError("A transfer is already running. Cancel it first.")
            self.job_running = True
        self.cancel_event.clear()

        if self.remote_kbd_on:
            self._remote_keyboard_off(reason="[WARN] Remote keyboard disabled during transfer.")

        self.after(0, lambda: self.btn_cancel.config(state="normal"))
        self.after(0, lambda: self._set_controls_enabled(False))
        self.after(0, lambda: self._set_keyboard_controls_enabled(False))

    def _job_end(self):
        with self.job_lock:
            self.job_running = False
        self.after(0, lambda: self.btn_cancel.config(state="disabled"))
        self.after(0, lambda: self._set_controls_enabled(True))
        self.after(0, lambda: self._set_keyboard_controls_enabled(False))

    def _set_status(self, text: str):
        self.after(0, lambda: self.status_var.set(f"Status: {text}"))

    def _set_progress(self, pct: float, label: str):
        def _u():
            self.progress_var.set(max(0.0, min(100.0, pct)))
            self.progress_label.config(text=f"Progress: {label}")
        self.after(0, _u)

    def _run_threaded(self, fn, name: str):
        def runner():
            try:
                self._job_start()
                self.log(f"--- {name} ---")
                fn()
                if self.cancel_event.is_set():
                    self.log(f"[WARN] {name} cancelled.")
            except InterruptedError as e:
                self.log(f"[WARN] {name} cancelled: {e}")
            except Exception as e:
                self.log(f"[ERROR] {name}: {e!r}")
                messagebox.showerror("Error", f"{name}\n\n{e!r}")
            finally:
                self._set_status("idle")
                self._set_progress(0, "idle")
                self._job_end()
        threading.Thread(target=runner, daemon=True).start()

    # ---------------- Serial helpers ----------------
    def _require_port(self) -> str:
        p = (self.var_port.get() or "").strip()
        if not p:
            raise RuntimeError("No COM port selected.")
        return p

    def _open_for_typing(self) -> serial.Serial:
        return serial.Serial(
            self._require_port(),
            int(self.var_typing_baud.get()),
            bytesize=serial.EIGHTBITS,
            parity=serial.PARITY_NONE,
            stopbits=serial.STOPBITS_TWO,
            timeout=0.2,
            xonxoff=False,
            rtscts=self.var_rtscts.get(),
            dsrdtr=False,
        )

    def _open_for_raw(self) -> serial.Serial:
        return serial.Serial(
            self._require_port(),
            int(self.var_xfer_baud.get()),
            bytesize=serial.EIGHTBITS,
            parity=serial.PARITY_NONE,
            stopbits=serial.STOPBITS_TWO,
            timeout=0.2,
            xonxoff=False,
            rtscts=self.var_rtscts.get(),
            dsrdtr=False,
        )


    def _prime_loader_xfer(self, ser: serial.Serial):
        try:
            ser.reset_input_buffer()
        except Exception:
            pass
        try:
            ser.reset_output_buffer()
        except Exception:
            pass

    def _parse_addr(self) -> int:
        s = self.var_addr.get().strip().lower()
        if s.endswith("h"):
            return int(s[:-1], 16)
        return int(s, 0)

    def _parse_loader_addr(self) -> int:
        s = self.var_loader_addr.get().strip().lower()
        if s.endswith("h"):
            return int(s[:-1], 16)
        return int(s, 0)

    def _loader_cas_path(self) -> Path:
        return Path(__file__).with_name(LOADER_CAS_NAME)



    def _build_loader_cas_payload(self) -> bytes:
        template_path = self._loader_cas_path()
        if not template_path.exists():
            raise RuntimeError(f"Missing loader template: {template_path}")

        cas_bytes = template_path.read_bytes()
        if len(cas_bytes) <= CAS_FIXED_BASE:
            raise RuntimeError("Loader template too short.")

        payload = bytearray(cas_bytes[CAS_FIXED_BASE:])
        original_payload_len = len(payload)

        load_addr = self._parse_loader_addr() & 0xFFFF
        end_addr = (load_addr + LOADER_BYTE_COUNT - 1) & 0xFFFF

        # Patch inside the raw payload
        ascii_addr_offsets = [
            (9,  b"&H1800"),  # CLEAR
            (23, b"&H1800"),  # FOR start
            (30, b"&H18A1"),  # FOR end
            (66, b"&H1800"),  # EXEC
        ]
        ascii_addr_values = [
            f"&H{load_addr:04X}".encode("ascii"),
            f"&H{load_addr:04X}".encode("ascii"),
            f"&H{end_addr:04X}".encode("ascii"),
            f"&H{load_addr:04X}".encode("ascii"),
        ]
        for (offset, expected), repl in zip(ascii_addr_offsets, ascii_addr_values):
            if payload[offset:offset + len(expected)] != expected:
                raise RuntimeError(
                    f"Unexpected loader.cas payload format near BASIC address field at offset {offset}."
                )
            if len(repl) != len(expected):
                raise RuntimeError("Internal error: BASIC address replacement length changed.")
            payload[offset:offset + len(expected)] = repl

        data_spans = [
            (78, 125, 16),
            (131, 178, 16),
            (184, 231, 16),
            (237, 284, 16),
            (290, 337, 16),
            (343, 390, 16),
            (396, 443, 16),
            (449, 496, 16),
            (502, 549, 16),
            (555, 602, 16),
            (608, 613, 2),
        ]

        loader_bytes = bytearray()
        for start_, end_, item_count in data_spans:
            chunk = payload[start_:end_].decode("ascii")
            items = chunk.split(",")
            if len(items) != item_count:
                raise RuntimeError(
                    f"Unexpected loader.cas DATA layout at payload {start_}:{end_}; got {len(items)} items."
                )
            loader_bytes.extend(int(x, 16) for x in items)

        if len(loader_bytes) != LOADER_BYTE_COUNT:
            raise RuntimeError(
                f"Unexpected loader.cas DATA byte count: got {len(loader_bytes)}, expected {LOADER_BYTE_COUNT}."
            )

        orig_base = 0x1800
        internal_offsets = {
            "load_addr": 0x42,
            "get_char_filtered": 0x44,
            "read_hex_nibble": 0x5C,
            "read_hex8": 0x84,
            "read_hex16": 0x96,
        }

        def patch_word_sequences(opcode: int, old_addr: int, new_addr: int) -> int:
            old_seq = bytes([opcode, old_addr & 0xFF, (old_addr >> 8) & 0xFF])
            new_seq = bytes([opcode, new_addr & 0xFF, (new_addr >> 8) & 0xFF])
            idx = 0
            count = 0
            while True:
                idx = loader_bytes.find(old_seq, idx)
                if idx < 0:
                    break
                loader_bytes[idx:idx + 3] = new_seq
                idx += 3
                count += 1
            return count

        patch_word_sequences(0x22, orig_base + internal_offsets["load_addr"], load_addr + internal_offsets["load_addr"])
        patch_word_sequences(0x2A, orig_base + internal_offsets["load_addr"], load_addr + internal_offsets["load_addr"])
        patch_word_sequences(0xCD, orig_base + internal_offsets["get_char_filtered"], load_addr + internal_offsets["get_char_filtered"])
        patch_word_sequences(0xCD, orig_base + internal_offsets["read_hex_nibble"], load_addr + internal_offsets["read_hex_nibble"])
        patch_word_sequences(0xCD, orig_base + internal_offsets["read_hex8"], load_addr + internal_offsets["read_hex8"])
        patch_word_sequences(0xCD, orig_base + internal_offsets["read_hex16"], load_addr + internal_offsets["read_hex16"])

        # Patch runtime baud (LD IX,nn). Template uses 8000 baud => DD 21 40 1F
        baud = int(self.var_xfer_baud.get()) & 0xFFFF
        baud_pattern = bytes([0xDD, 0x21, 0x40, 0x1F])
        baud_repl = bytes([0xDD, 0x21, baud & 0xFF, (baud >> 8) & 0xFF])
        idx = loader_bytes.find(baud_pattern)
        if idx < 0:
            raise RuntimeError("Unexpected loader.cas format: loader baud pattern not found.")
        loader_bytes[idx:idx + 4] = baud_repl

        # Rebuild the ASCII DATA bytes back into the payload
        offset = 0
        for start_, end_, item_count in data_spans:
            chunk_bytes = loader_bytes[offset:offset + item_count]
            offset += item_count
            repl = ",".join(f"{b:02X}" for b in chunk_bytes).encode("ascii")
            if len(repl) != (end_ - start_):
                raise RuntimeError("Internal error while rebuilding loader.cas DATA lines.")
            payload[start_:end_] = repl

        if len(payload) != original_payload_len:
            raise RuntimeError("Internal error: loader payload length changed.")

        return bytes(payload)

    def _build_loader_ascii_frame(self, data: bytes) -> bytes:
        if not data:
            raise RuntimeError("Binary is empty.")
        addr = self._parse_addr() & 0xFFFF
        size = len(data)
        if size > 0xFFFF:
            raise RuntimeError("Binary too large for 16-bit ASCII loader header.")
        return ("L" + f"{addr:04X}" + f"{size:04X}" + data.hex().upper()).encode("ascii")

    def _send_loader_cas_raw(self, ser: serial.Serial, *, progress_base: float = 0.0, progress_span: float = 100.0):
        payload = self._build_loader_cas_payload()
        if not payload:
            raise RuntimeError("Loader CAS payload is empty.")

        self.log(f"[INFO] Loader CAS raw payload={len(payload)} bytes.")
        self.log(f"[INFO] Loader load addr: 0x{self._parse_loader_addr():04X}")
        self.log(f"[INFO] Loader runtime baud: {int(self.var_xfer_baud.get())} 8N2")

        self._stream_raw_payload(
            ser,
            payload,
            "Loader CAS",
            progress_base=progress_base,
            progress_span=progress_span,
            chunk_size=64,
            final_rts_drop=True,
        )
        self.log("[INFO] Loader CAS raw transfer complete.")

    def _send_loader_cas_remote(self, ser: serial.Serial, *, progress_base: float = 0.0, progress_span: float = 50.0):
        self.log('[INFO] Remote X-07 side: LOAD"COM:" then EXEC&HEE33:RUN')
        self._type_line(ser, 'LOAD"COM:"')
        pre_stream_delay = max(float(self.var_line.get()), 0.35)
        self.log(f"[INFO] Waiting {pre_stream_delay:.2f}s for X-07 to enter LOAD\"COM:\" receive mode...")
        time.sleep(pre_stream_delay)
        self._send_loader_cas_raw(ser, progress_base=progress_base, progress_span=progress_span)
        post_stream_delay = max(0.25, float(self.var_line.get()))
        time.sleep(post_stream_delay)
        self._type_line(ser, "EXEC&HEE33:RUN")
        self.log("[INFO] Remote loader transferred and started.")


    def _type_line(self, ser: serial.Serial, line: str):
        char_delay = float(self.var_char.get())
        line_delay = float(self.var_line.get())
        for ch in line.encode("ascii", errors="replace"):
            if self.cancel_event.is_set():
                raise InterruptedError("Cancelled during BASIC typing.")
            ser.write(bytes([ch]))
            ser.flush()
            self._tcdrain(ser)
            time.sleep(char_delay)
        ser.write(b"\r")
        ser.flush()
        self._tcdrain(ser)
        time.sleep(line_delay)

    def _tcdrain(self, ser: serial.Serial):
        try:
            fd = ser.fileno()
        except Exception:
            return
        try:
            import termios  # type: ignore
            termios.tcdrain(fd)
        except Exception:
            pass

    def _stream_raw_payload(self, ser: serial.Serial, payload: bytes, label: str,
                            progress_base: float = 0.0, progress_span: float = 100.0,
                            chunk_size: int = 64, final_rts_drop: bool = False):
        total = len(payload)
        if total <= 0:
            return
        chunk_size = max(1, int(chunk_size))
        for offset in range(0, total, chunk_size):
            if self.cancel_event.is_set():
                raise InterruptedError(f"Cancelled during {label} send.")
            chunk = payload[offset:offset + chunk_size]
            ser.write(chunk)
            ser.flush()
            self._tcdrain(ser)
            sent = min(offset + len(chunk), total)
            pct = progress_base + (sent / total) * progress_span
            self._set_progress(pct, f"{label}: {sent}/{total}")
        if final_rts_drop:
            try:
                ser.rts = False
            except Exception:
                pass

    def _send_ascii_frame(self, ser: serial.Serial, frame: bytes, label: str,
                          progress_base: float = 0.0, progress_span: float = 100.0,
                          chunk_size: int = 256):
        total = len(frame)
        if total <= 0:
            return
        chunk_size = max(1, int(chunk_size))
        self._stream_raw_payload(ser, frame, label, progress_base, progress_span, chunk_size)

    # ---------------- File pickers ----------------
    def pick_basic(self):
        p = filedialog.askopenfilename(
            title="Select BASIC program (.txt/.bas)",
            filetypes=[("BASIC", "*.txt *.TXT *.bas *.BAS"), ("All files", "*.*")]
        )
        if not p:
            return
        self.basic_file = Path(p)
        self.lbl_basic.config(text=f"Selected BASIC: {self.basic_file}")
        self.log(f"[OK] BASIC selected: {self.basic_file}")

    def pick_bin(self):
        p = filedialog.askopenfilename(
            title="Select ASM binary (.bin)",
            filetypes=[("Binary", "*.bin *.BIN"), ("All files", "*.*")]
        )
        if not p:
            return
        self.bin_file = Path(p)
        self.lbl_bin.config(text=f"Selected ASM binary: {self.bin_file}")
        self.log(f"[OK] BIN selected: {self.bin_file}")

    def pick_cas(self):
        p = filedialog.askopenfilename(
            title="Select CAS/K7",
            filetypes=[("CAS/K7", "*.cas *.CAS *.k7 *.K7"), ("All files", "*.*")]
        )
        if not p:
            return
        self.cas_file = Path(p)
        self.lbl_cas.config(text=f"Selected CAS/K7: {self.cas_file}")
        self.log(f"[OK] CAS/K7 selected: {self.cas_file}")

    # ---------------- Slave mode helper ----------------
    def disable_slave_mode(self):
        self._run_threaded(self._disable_slave_mode_impl, "Disable slave mode (EXEC&HEE33)")

    def _disable_slave_mode_impl(self):
        self._set_status("sending EXEC&HEE33")
        self._set_progress(0, "sending EXEC&HEE33")
        try:
            with self._open_for_typing() as ser:
                self._type_line(ser, "EXEC&HEE33")
        except (SerialException, OSError) as e:
            self.log(f"[ERROR] Cannot open {self.var_port.get()!r}: {e}")
            return
        self._set_progress(100, "done")
        self.log("[INFO] EXEC&HEE33 sent.")

    # ---------------- BASIC (.txt/.bas) ----------------
    def send_basic_file(self):
        if not self.basic_file or not self.basic_file.exists():
            messagebox.showwarning("Missing BASIC", "Select a BASIC .txt/.bas file first.")
            return
        self._run_threaded(self._send_basic_file_impl, "Send BASIC (.txt/.bas)")

    def _send_basic_file_impl(self):
        self._set_status("typing BASIC (.txt/.bas)")
        self._set_progress(0, "starting...")

        lines = self.basic_file.read_text(encoding="utf-8", errors="replace").splitlines()
        lines = [ln.rstrip("\r\n") for ln in lines if ln.strip() != ""]
        total = max(1, len(lines) + 1)

        try:
            with self._open_for_typing() as ser:
                self.log(f"[INFO] Opened {ser.port} for typing: {ser.baudrate} 8N2.")
                for idx, ln in enumerate(lines, start=1):
                    if self.cancel_event.is_set():
                        raise InterruptedError("Cancelled during BASIC send.")
                    self._type_line(ser, ln)
                    self._set_progress((idx / total) * 100.0, f"BASIC {idx}/{total} lines")

                if self.cancel_event.is_set():
                    raise InterruptedError("Cancelled before EXEC&HEE33.")
                self._type_line(ser, "EXEC&HEE33")
        except (SerialException, OSError) as e:
            self.log(f"[ERROR] Cannot open {self.var_port.get()!r}: {e}")
            return

        self._set_progress(100.0, f"BASIC {total}/{total} lines (done)")
        self.log("[INFO] Sent EXEC&HEE33 after BASIC.")

    # ---------------- Fast loader + ASM ----------------

    def send_fast_loader(self):
        self._run_threaded(self._send_fast_loader_impl, 'Send ASM loader (LOAD"COM:")')

    def _send_fast_loader_impl(self):
        self._set_status("sending loader CAS raw")
        self._set_progress(0, "sending loader CAS...")

        try:
            with self._open_for_typing() as ser:
                self.log(f"[INFO] Opened {ser.port} for loader raw send: {ser.baudrate} 8N2.")
                self._send_loader_cas_raw(ser)
        except (SerialException, OSError) as e:
            self.log(f"[ERROR] Cannot open {self.var_port.get()!r}: {e}")
            return

        self._set_progress(100.0, "loader send done")
        self.log("[INFO] Loader raw send complete.")

    def send_bin_only(self):
        if not self.bin_file or not self.bin_file.exists():
            messagebox.showwarning("Missing BIN", "Select a .bin file first.")
            return
        self._run_threaded(self._send_bin_only_impl, "Send ASM (loader running)")

    def _send_bin_only_impl(self):
        self._set_status("transferring ASM via ASCII loader")
        self._set_progress(0, "building frame...")

        data = self.bin_file.read_bytes()
        frame = self._build_loader_ascii_frame(data)
        total = len(frame)
        primer_len = len(ASM_PRIMER)

        self.log(f"[INFO] BIN size: {len(data)} bytes. Target load: 0x{self._parse_addr():04X}")
        self.log(f"[INFO] Sending ASCII loader frame: {total} chars @ {int(self.var_xfer_baud.get())} 8N2.")
        self.log(f"[INFO] Sending primer first: {primer_len} bytes of ignored non-sync data.")

        try:
            with self._open_for_raw() as ser:
                self._prime_loader_xfer(ser)
                self._stream_raw_payload(ser, ASM_PRIMER, "ASM primer", progress_base=0.0, progress_span=10.0, chunk_size=256)
                self._send_ascii_frame(ser, frame, "ASM frame", progress_base=10.0, progress_span=90.0, chunk_size=256)
        except (SerialException, OSError) as e:
            self.log(f"[ERROR] Cannot open {self.var_port.get()!r}: {e}")
            return

        self._set_progress(100.0, "ASM done")
        self.log("[INFO] ASM transfer done.")

    def send_loader_and_bin(self):
        if not self.bin_file or not self.bin_file.exists():
            messagebox.showwarning("Missing BIN", "Select a .bin file first.")
            return
        self._run_threaded(self._send_loader_and_bin_impl, "One click: loader + ASM")

    def _send_loader_and_bin_impl(self):
        self._set_status("sending loader CAS + ASM")
        self._set_progress(0, "starting...")

        data = self.bin_file.read_bytes()
        frame = self._build_loader_ascii_frame(data)
        total = len(frame)
        primer_len = len(ASM_PRIMER)

        try:
            with self._open_for_typing() as ser:
                self.log(f"[INFO] Opened {ser.port} for one-click: {ser.baudrate} 8N2.")
                self._send_loader_cas_remote(ser, progress_base=0.0, progress_span=50.0)
                self._set_progress(50.0, "loader started")
            time.sleep(POST_LOADER_EXEC_DELAY_S)
            with self._open_for_raw() as ser:
                self._prime_loader_xfer(ser)
                self.log(f"[INFO] Sending primer first: {primer_len} bytes of ignored non-sync data.")
                self._stream_raw_payload(ser, ASM_PRIMER, "ASM primer", progress_base=50.0, progress_span=5.0, chunk_size=256)
                self._send_ascii_frame(ser, frame, "ASM frame", progress_base=55.0, progress_span=45.0, chunk_size=256)
        except (SerialException, OSError) as e:
            self.log(f"[ERROR] Cannot open {self.var_port.get()!r}: {e}")
            return

        self._set_progress(100.0, "one-click done")
        self.log("[INFO] One-click complete.")

    # ---------------- CAS/K7 (RAW BYTES) ----------------
    def inspect_cas_header(self):
        if not self.cas_file or not self.cas_file.exists():
            messagebox.showwarning("Missing CAS/K7", "Select a .cas/.k7 file first.")
            return
        data = self.cas_file.read_bytes()
        name = guess_name_in_first_16_bytes(data, self.cas_file.stem)

        self.log("[INFO] Inspect header:")
        self.log(f"       File: {self.cas_file.name}")
        self.log(f"       Name guess: {name}")
        self.log("       Preview @0x0000:")

        chunk = data[:96]
        hexline = " ".join(f"{b:02X}" for b in chunk[:64])
        asciiline = "".join(chr(b) if 0x20 <= b <= 0x7E else "." for b in chunk[:64])
        self.log(f"       0x0000: {hexline}")
        self.log(f"               {asciiline}")

    def send_cas_raw(self):
        if not self.cas_file or not self.cas_file.exists():
            messagebox.showwarning("Missing CAS/K7", "Select a .cas/.k7 file first.")
            return
        self._run_threaded(self._send_cas_raw_impl, 'Send CAS/K7 raw stream (LOAD"COM:")')

    def _send_cas_raw_impl(self):
        self._set_status("sending CAS/K7 raw bytes")
        self._set_progress(0, "starting...")

        data = self.cas_file.read_bytes()
        base = CAS_FIXED_BASE
        payload = data[base:]
        if not payload:
            raise RuntimeError("Empty payload (base beyond file length).")

        name = guess_name_in_first_16_bytes(data, self.cas_file.stem)
        self.log(f'[INFO] CAS/K7 file: {self.cas_file.name} | name guess: {name}')
        self.log(f'[INFO] Sending raw bytes from base=0x{base:04X} (len={len(payload)}).')
        self.log('[INFO] X-07 side: LOAD"COM:" then press RETURN.')

        total = len(payload)

        try:
            with self._open_for_typing() as ser:
                self._stream_raw_payload(
                    ser,
                    payload,
                    'CAS/K7 send',
                    progress_base=0.0,
                    progress_span=100.0,
                    chunk_size=64,
                    final_rts_drop=True,
                )

        except (SerialException, OSError) as e:
            self.log(f"[ERROR] Cannot open {self.var_port.get()!r}: {e}")
            return

        self._set_progress(100.0, "CAS/K7 send done")
        self.log("[INFO] CAS/K7 raw transfer finished.")

    def receive_cas_raw(self):
        p = filedialog.asksaveasfilename(
            title='Save received stream as (.cas)',
            defaultextension=".cas",
            filetypes=[("CAS", "*.cas"), ("All files", "*.*")]
        )
        if not p:
            return
        out_path = Path(p)
        self._run_threaded(lambda: self._receive_cas_raw_impl(out_path), 'Receive CAS/K7 raw stream (SAVE"COM:")')

    def _receive_cas_raw_impl(self, out_path: Path):
        self._set_status("receiving CAS/K7 raw bytes")
        self._set_progress(0, "waiting for data...")

        self.log(f'[INFO] PC ready to receive into: {out_path}')
        self.log('[INFO] X-07 side: type SAVE"COM:" (optionally with a name) then press RETURN.')
        self.log(f"[INFO] Capture ends after ~{SAVE_IDLE_TIMEOUT_S:.2f}s of inactivity.")

        header = build_cas_header_from_filename(out_path)

        buf = bytearray()
        got_any = False
        last_rx = time.time()

        try:
            with self._open_for_typing() as ser:
                ser.timeout = 0.2
                while True:
                    if self.cancel_event.is_set():
                        raise InterruptedError("Cancelled during CAS/K7 receive.")

                    chunk = ser.read(4096)
                    if chunk:
                        buf.extend(chunk)
                        got_any = True
                        last_rx = time.time()

                        self._set_progress(0.0, f"CAS/K7 recv {len(buf)} bytes")
                    else:
                        if got_any and (time.time() - last_rx) >= SAVE_IDLE_TIMEOUT_S:
                            break

        except InterruptedError:
            if buf:
                try:
                    out_path.write_bytes(header + bytes(buf))
                    self.log(f"[WARN] Partial stream saved ({len(buf)} bytes).")
                except Exception as e:
                    self.log(f"[ERROR] Failed to save partial stream: {e}")
            raise
        except (SerialException, OSError) as e:
            self.log(f"[ERROR] Cannot open {self.var_port.get()!r}: {e}")
            return

        if not buf:
            self.log("[WARN] No data received. Did SAVE\"COM:\" start?")
            self._set_progress(0.0, "no data")
            return

        try:
            out_path.write_bytes(header + bytes(buf))
        except Exception as e:
            self.log(f"[ERROR] Failed to save file: {e}")
            return

        self._set_progress(100.0, f"CAS/K7 recv done ({len(buf)} bytes)")
        self.log(f"[OK] Received {len(buf)} bytes. Saved: {out_path.name}")

    # ---------------- Remote keyboard ----------------
    def toggle_remote_keyboard(self):
        if self.job_running:
            self.log("[WARN] Remote keyboard is disabled during transfers.")
            self._set_status("remote keyboard disabled during transfer")
            return

        if self.remote_kbd_on:
            self._remote_keyboard_off(reason="[OK] REMOTE KEYBOARD: OFF")
        else:
            self._remote_keyboard_on()

    def _remote_keyboard_on(self):
        try:
            self.remote_ser = self._open_for_typing()
        except (SerialException, OSError) as e:
            self.remote_ser = None
            self.remote_kbd_on = False
            self.btn_remote_toggle.config(text="REMOTE KEYBOARD: OFF")
            self._set_keyboard_controls_enabled(False)
            self.log(f"[ERROR] Cannot open {self.var_port.get()!r} for remote keyboard: {e}")
            self._set_status("remote keyboard start failed (check COM port)")
            return

        self.remote_kbd_on = True
        self.btn_remote_toggle.config(text="REMOTE KEYBOARD: ON")
        self._set_keyboard_controls_enabled(True)
        self.log(f"[OK] REMOTE KEYBOARD: ON ({self.remote_ser.port} @ {self.remote_ser.baudrate} 8N2).")
        self.log('[INFO] Remote keyboard requires SLAVE mode (INIT#5,"COM:" then EXEC&HEE1F).')

    def _remote_keyboard_off(self, reason: str = "[OK] REMOTE KEYBOARD: OFF"):
        try:
            if self.remote_ser:
                self.remote_ser.close()
        except Exception:
            pass
        self.remote_ser = None
        self.remote_kbd_on = False
        self.btn_remote_toggle.config(text="REMOTE KEYBOARD: OFF")
        self._set_keyboard_controls_enabled(False)
        self.log(reason)

    def _kbd_send_byte(self, b: int):
        if not self.remote_kbd_on or not self.remote_ser:
            return
        try:
            self.remote_ser.write(bytes([b & 0xFF]))
            self.remote_ser.flush()
        except (SerialException, OSError) as e:
            self.log(f"[ERROR] Remote keyboard write failed: {e}")
            self._remote_keyboard_off(reason="[WARN] REMOTE KEYBOARD stopped (write error).")

    def _kbd_send_bytes(self, data: bytes):
        if not self.remote_kbd_on or not self.remote_ser:
            return
        try:
            self.remote_ser.write(data)
            self.remote_ser.flush()
        except (SerialException, OSError) as e:
            self.log(f"[ERROR] Remote keyboard write failed: {e}")
            self._remote_keyboard_off(reason="[WARN] REMOTE KEYBOARD stopped (write error).")

    def _on_remote_keypress(self, event):
        if not self.remote_kbd_on:
            return "break"

        k = event.keysym

        if k == "Left":
            self._kbd_send_byte(KEY_LEFT)
            return "break"
        if k == "Right":
            self._kbd_send_byte(KEY_RIGHT)
            return "break"
        if k == "Up":
            self._kbd_send_byte(KEY_UP)
            return "break"
        if k == "Down":
            self._kbd_send_byte(KEY_DOWN)
            return "break"
        if k == "Home":
            self._kbd_send_byte(KEY_HOME)
            return "break"
        if k == "Insert":
            self._kbd_send_byte(KEY_INS)
            return "break"
        if k == "Delete":
            self._kbd_send_byte(KEY_DEL)
            return "break"
        if k == "BackSpace":
            self._kbd_send_byte(KEY_LEFT)
            self._kbd_send_byte(KEY_DEL)
            return "break"
        if k in ("Return", "KP_Enter"):
            self._kbd_send_byte(KEY_RETURN)
            return "break"
        if k == "space":
            self._kbd_send_byte(KEY_SPACE)
            return "break"

        ch = event.char
        if ch:
            data = x07_encode_text(ch)
            if data:
                self._kbd_send_bytes(data)
        return "break"


if __name__ == "__main__":
    app = X07LoaderApp()
    app.mainloop()
