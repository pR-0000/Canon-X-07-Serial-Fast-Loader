import argparse
import configparser
from pathlib import Path
import re
import subprocess
import sys
import threading
import time
from typing import Protocol
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

IS_MACOS = sys.platform == "darwin"

try:
    import termios
    HAS_TERMIOS = True
except ImportError:
    termios = None  # type: ignore[assignment]
    HAS_TERMIOS = False

try:
    import serial
    import serial.tools.list_ports
    from serial.serialutil import SerialException
    HAS_PYSERIAL = True
except ImportError:
    HAS_PYSERIAL = False
    if not (IS_MACOS and HAS_TERMIOS):
        subprocess.check_call([sys.executable, "-m", "pip", "install", "pyserial"])
        import serial
        import serial.tools.list_ports
        from serial.serialutil import SerialException
        HAS_PYSERIAL = True
    else:
        serial = None  # type: ignore[assignment]
        SerialException = OSError  # type: ignore[misc,assignment]

DARWIN_TERMIOS_AVAILABLE = False
if IS_MACOS and HAS_TERMIOS:
    try:
        import array
        import fcntl
        import os
        DARWIN_TERMIOS_AVAILABLE = True
    except ImportError:
        DARWIN_TERMIOS_AVAILABLE = False

SERIAL_ERRORS = (SerialException, OSError) if HAS_PYSERIAL else (OSError,)

# ---------- Defaults ----------
DEFAULT_CHAR_DELAY_S = 0.04
DEFAULT_LINE_DELAY_S = 0.20
DEFAULT_LOAD_ADDR = 0x2000
DEFAULT_LOADER_ADDR = 0x1F00
POST_LOADER_EXEC_DELAY_S = 3.0
LOADER_CAS_NAME = "loader.cas"

BASIC_START = 0x0553
CAS_FIXED_BASE = 0x0010  # fixed (validated for SEND)
SAVE_IDLE_TIMEOUT_S = 1.25  # end capture if no bytes for this duration (after any data received)


def settings_config_path() -> Path:
    return Path(__file__).with_suffix(".ini")  # x07_loader.ini


def default_serial_settings() -> dict:
    return {
        "port": None,
        "rtscts": False,
        "typing_baud": 4800,
        "xfer_baud": 8000,
        "char_delay": DEFAULT_CHAR_DELAY_S,
        "line_delay": DEFAULT_LINE_DELAY_S,
        "loader_addr": hex(DEFAULT_LOADER_ADDR),
        "asm_addr": hex(DEFAULT_LOAD_ADDR),
    }


def load_saved_serial_settings() -> dict:
    cfg_path = settings_config_path()
    settings = default_serial_settings()
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


class SerialPortLike(Protocol):
    port: str
    baudrate: int
    timeout: float | None
    rts: bool

    def write(self, data: bytes) -> int:
        ...

    def flush(self) -> None:
        ...

    def read(self, size: int = 1) -> bytes:
        ...

    def close(self) -> None:
        ...

    def fileno(self) -> int:
        ...

    def reset_input_buffer(self) -> None:
        ...

    def reset_output_buffer(self) -> None:
        ...

    def __enter__(self):
        ...

    def __exit__(self, exc_type, exc, tb) -> None:
        ...


def serial_backend_summary() -> str:
    if IS_MACOS and DARWIN_TERMIOS_AVAILABLE and HAS_PYSERIAL:
        return "termios preferred on macOS (pyserial fallback available)"
    if IS_MACOS and DARWIN_TERMIOS_AVAILABLE:
        return "termios on macOS"
    if HAS_PYSERIAL:
        return "pyserial"
    return "unavailable"


if IS_MACOS and DARWIN_TERMIOS_AVAILABLE:
    IOSSIOSPEED = 0x80045402  # _IOW('T', 2, speed_t)

    class MacTermiosSerial:
        def __init__(self, port: str, baudrate: int, *, timeout: float | None = 0.2, rtscts: bool = False):
            self.port = port
            self.baudrate = int(baudrate)
            self._termios_backend = True
            self._timeout = timeout
            self._rtscts = bool(rtscts)
            self._fd: int | None = None
            self._rts = True

            try:
                self._fd = os.open(self.port, os.O_RDWR | os.O_NOCTTY)
                self._configure_port()
            except Exception:
                self.close()
                raise

        def __enter__(self):
            return self

        def __exit__(self, exc_type, exc, tb) -> None:
            self.close()

        def _require_open_fd(self) -> int:
            if self._fd is None:
                raise OSError("Serial port is closed.")
            return self._fd

        @property
        def timeout(self) -> float | None:
            return self._timeout

        @timeout.setter
        def timeout(self, value: float | None) -> None:
            self._timeout = None if value is None else max(0.0, float(value))
            if self._fd is not None:
                self._configure_port()

        @property
        def rts(self) -> bool:
            return self._rts

        @rts.setter
        def rts(self, value: bool) -> None:
            self._rts = bool(value)
            self._set_rts_state()

        def _configure_port(self) -> None:
            fd = self._require_open_fd()
            iflag, oflag, cflag, lflag, ispeed, ospeed, cc = termios.tcgetattr(fd)

            cflag |= termios.CLOCAL | termios.CREAD

            for flag_name in ("ICANON", "ECHO", "ECHOE", "ECHOK", "ECHONL", "ISIG", "IEXTEN"):
                if hasattr(termios, flag_name):
                    lflag &= ~getattr(termios, flag_name)
            for flag_name in ("OPOST", "ONLCR", "OCRNL"):
                if hasattr(termios, flag_name):
                    oflag &= ~getattr(termios, flag_name)
            for flag_name in ("INLCR", "IGNCR", "ICRNL", "IGNBRK", "IXON", "IXOFF", "IXANY",
                              "INPCK", "ISTRIP", "BRKINT", "PARMRK"):
                if hasattr(termios, flag_name):
                    iflag &= ~getattr(termios, flag_name)

            cflag &= ~termios.CSIZE
            cflag |= termios.CS8
            cflag &= ~(termios.PARENB | termios.PARODD)
            cflag |= termios.CSTOPB

            for flow_flag in ("CRTSCTS", "CNEW_RTSCTS"):
                if hasattr(termios, flow_flag):
                    if self._rtscts:
                        cflag |= getattr(termios, flow_flag)
                    else:
                        cflag &= ~getattr(termios, flow_flag)
                    break

            if self._timeout is None:
                vmin = 1
                vtime = 0
            else:
                vmin = 0
                vtime = min(255, max(0, int(round(self._timeout * 10))))
                if self._timeout > 0 and vtime == 0:
                    vtime = 1
            cc[termios.VMIN] = vmin
            cc[termios.VTIME] = vtime

            baud_const = getattr(termios, f"B{self.baudrate}", None)
            custom_baud = None
            if baud_const is None:
                baud_const = getattr(termios, "B38400", None)
                if baud_const is None:
                    raise RuntimeError(f"Unsupported baud rate: {self.baudrate}")
                custom_baud = self.baudrate

            ispeed = baud_const
            ospeed = baud_const
            termios.tcsetattr(fd, termios.TCSANOW, [iflag, oflag, cflag, lflag, ispeed, ospeed, cc])

            if custom_baud is not None:
                fcntl.ioctl(fd, IOSSIOSPEED, array.array("i", [custom_baud]), True)

        def _set_rts_state(self) -> None:
            fd = self._fd
            if fd is None:
                return
            request = getattr(termios, "TIOCMBIS", None) if self._rts else getattr(termios, "TIOCMBIC", None)
            bit = getattr(termios, "TIOCM_RTS", None)
            if request is None or bit is None:
                return
            try:
                fcntl.ioctl(fd, request, array.array("I", [bit]), True)
            except Exception:
                pass

        def fileno(self) -> int:
            return self._require_open_fd()

        def write(self, data: bytes) -> int:
            fd = self._require_open_fd()
            view = memoryview(data)
            total = 0
            while total < len(view):
                total += os.write(fd, view[total:])
            return total

        def flush(self) -> None:
            termios.tcdrain(self._require_open_fd())

        def read(self, size: int = 1) -> bytes:
            return os.read(self._require_open_fd(), max(1, int(size)))

        def reset_input_buffer(self) -> None:
            termios.tcflush(self._require_open_fd(), termios.TCIFLUSH)

        def reset_output_buffer(self) -> None:
            termios.tcflush(self._require_open_fd(), termios.TCOFLUSH)

        def close(self) -> None:
            if self._fd is None:
                return
            try:
                os.close(self._fd)
            finally:
                self._fd = None

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
X07_BYTE_TO_UNICODE = {value: key for key, value in X07_UNICODE_MAP.items()}

X07_BASIC_DETOKEN_TOKENS = (
    "END", "FOR ", "NEXT ", "DATA ",
    "INPUT", "DIM ", "READ ", "LET ",
    "GOTO ", "RUN ", "IF ", "RESTORE ",
    "GOSUB ", "RETURN", "REM", "STOP",
    " ELSE ", "TR", "MOTOR ", "DEFSTR ",
    "DEFINT ", "DEFSNG ", "DEFDBL ", "LINE ",
    "ERROR ", "RESUME ", "OUT ", "ON ",
    "LPRINT", "DEFFN", "POKE ", "PRINT",
    "CONT", "LIST ", "LLIST ", "CLEAR ",
    "CIRCLE", "CONSOLE", "CLS", "COLOR ",
    "EXEC ", "LOCATE ", "PSET", "PRESET",
    "OFF", "SLEEP", "DIR", "DELETE ",
    "FSET ", "PAINT ", "LOAD", "SAVE",
    "INIT", "ERASE ", "BEEP ", "CLOAD",
    "CSAVE", "NEW", "TAB(", " TO ",
    "FN", " USING", "ERL", " ERROR",
    "STRING$", "INSTR", "INKEY$", "INP",
    "VARPTR", "USR", "SNS", "ALM$",
    "DATE$", "TIME$", "START$", "FONT$",
    "KEY$", "SCREEN ", " THEN ", "NOT ",
    " STEP ", "+", "-", "*",
    "/", "^", " AND ", " OR ",
    " XOR ", " EQU ", " MOD ", "\\",
    ">", "=", "<", "SGN",
    "INT", "ABS", "FRE", "POS",
    "SQR", "RND", "LOG", "EXP",
    "COS", "SIN", "TAN", "ATN ",
    "PEEK", "CINT", "CSNG", "CDBL",
    "FIX", "LEN", "HEX$", "STR$",
    "VAL", "ASC", "CHR$", "TKEY",
    "LEFT$", "RIGHT$", "MID$", "CSRLIN",
    "STICK", "STRIG", "POINT", "'",
)

TOKEN_FLAG_TRANSPARENT = 0x01
TOKEN_FLAG_PREPEND_COLON = 0x02
TOKEN_FLAG_PREPEND_REM = 0x04

X07_BASIC_TOKEN_SPECS = (
    ("^", 0xD5, 0),
    ("\\", 0xDB, 0),
    ("XOR", 0xD8, 0),
    ("VARPTR", 0xC4, 0),
    ("VAL", 0xF4, 0),
    ("USR", 0xC5, 0),
    ("USING", 0xBD, 0),
    ("TR", 0x91, 0),
    ("TO", 0xBB, 0),
    ("TKEY", 0xF7, 0),
    ("TIME$", 0xC9, 0),
    ("THEN", 0xCE, 0),
    ("TAN", 0xEA, 0),
    ("TAB(", 0xBA, 0),
    ("STRING$", 0xC0, 0),
    ("STRIG", 0xFD, 0),
    ("STR$", 0xF3, 0),
    ("STOP", 0x8F, 0),
    ("STICK", 0xFC, 0),
    ("STEP", 0xD0, 0),
    ("START$", 0xCA, 0),
    ("SQR", 0xE4, 0),
    ("SNS", 0xC6, 0),
    ("SLEEP", 0xAD, 0),
    ("SIN", 0xE9, 0),
    ("SGN", 0xDF, 0),
    ("SCREEN", 0xCD, 0),
    ("SAVE", 0xB3, 0),
    ("RUN", 0x89, 0),
    ("RND", 0xE5, 0),
    ("RIGHT$", 0xF9, 0),
    ("RETURN", 0x8D, 0),
    ("RESUME", 0x99, 0),
    ("RESTORE", 0x8B, 0),
    ("REM", 0x8E, TOKEN_FLAG_TRANSPARENT),
    ("READ", 0x86, 0),
    ("PSET", 0xAA, 0),
    ("PRINT", 0x9F, 0),
    ("PRESET", 0xAB, 0),
    ("POS", 0xE3, 0),
    ("POKE", 0x9E, 0),
    ("POINT", 0xFE, 0),
    ("PEEK", 0xEC, 0),
    ("PAINT", 0xB1, 0),
    ("OUT", 0x9A, 0),
    ("OR", 0xD7, 0),
    ("ON", 0x9B, 0),
    ("OFF", 0xAC, 0),
    ("NOT", 0xCF, 0),
    ("NEXT", 0x82, 0),
    ("NEW", 0xB9, 0),
    ("MOTOR", 0x92, 0),
    ("MOD", 0xDA, 0),
    ("MID$", 0xFA, 0),
    ("LPRINT", 0x9C, 0),
    ("LOG", 0xE6, 0),
    ("LOCATE", 0xA9, 0),
    ("LOAD", 0xB2, 0),
    ("LLIST", 0xA2, 0),
    ("LIST", 0xA1, 0),
    ("LINE", 0x97, 0),
    ("LET", 0x87, 0),
    ("LEN", 0xF1, 0),
    ("LEFT$", 0xF8, 0),
    ("KEY$", 0xCC, 0),
    ("INT", 0xE0, 0),
    ("INSTR", 0xC1, 0),
    ("INPUT", 0x84, 0),
    ("INP", 0xC3, 0),
    ("INKEY$", 0xC2, 0),
    ("INIT", 0xB4, 0),
    ("IF", 0x8A, 0),
    ("HEX$", 0xF2, 0),
    ("GOTO", 0x88, 0),
    ("GOSUB", 0x8C, 0),
    ("FSET", 0xB0, 0),
    ("FRE", 0xE2, 0),
    ("FOR", 0x81, 0),
    ("FONT$", 0xCB, 0),
    ("FN", 0xBC, 0),
    ("FIX", 0xF0, 0),
    ("EXP", 0xE7, 0),
    ("EXEC", 0xA8, 0),
    ("ERROR", 0xBF, 0),
    ("ERROR", 0x98, 0),
    ("ERL", 0xBE, 0),
    ("ERASE", 0xB5, 0),
    ("EQU", 0xD9, 0),
    ("END", 0x80, 0),
    ("ELSE", 0x90, TOKEN_FLAG_PREPEND_COLON),
    ("DIR", 0xAE, 0),
    ("DIM", 0x85, 0),
    ("DELETE", 0xAF, 0),
    ("DEFSTR", 0x93, 0),
    ("DEFSNG", 0x95, 0),
    ("DEFINT", 0x94, 0),
    ("DEFFN", 0x9D, 0),
    ("DEFDBL", 0x96, 0),
    ("DATE$", 0xC8, 0),
    ("DATA", 0x83, TOKEN_FLAG_TRANSPARENT),
    ("CSRLIN", 0xFB, 0),
    ("CSNG", 0xEE, 0),
    ("CSAVE", 0xB8, 0),
    ("COS", 0xE8, 0),
    ("CONT", 0xA0, 0),
    ("CONSOLE", 0xA5, 0),
    ("COLOR", 0xA7, 0),
    ("CLS", 0xA6, 0),
    ("CLOAD", 0xB7, 0),
    ("CLEAR", 0xA3, 0),
    ("CIRCLE", 0xA4, 0),
    ("CINT", 0xED, 0),
    ("CHR$", 0xF6, 0),
    ("CDBL", 0xEF, 0),
    ("BEEP", 0xB6, 0),
    ("ATN", 0xEB, 0),
    ("ASC", 0xF5, 0),
    ("AND", 0xD6, 0),
    ("ALM$", 0xC7, 0),
    ("ABS", 0xE1, 0),
    (">", 0xDC, 0),
    ("=", 0xDD, 0),
    ("<", 0xDE, 0),
    ("/", 0xD4, 0),
    ("-", 0xD2, 0),
    ("+", 0xD1, 0),
    ("*", 0xD3, 0),
    ("'", 0xFF, TOKEN_FLAG_TRANSPARENT | TOKEN_FLAG_PREPEND_COLON | TOKEN_FLAG_PREPEND_REM),
)


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


def _consume_x07_text_unit(text: str, start: int = 0) -> tuple[int, int]:
    if start >= len(text):
        raise ValueError("Start index beyond end of text.")

    if text[start] == "\\":
        token, consumed = _match_escape_token(text[start + 1:])
        if token is not None:
            return token, 1 + consumed
        return 0x5C, 1

    for key in _X07_UNICODE_KEYS:
        if text.startswith(key, start):
            return X07_UNICODE_MAP[key], len(key)

    ch = text[start]
    code = ord(ch)
    if ch in "\r\n":
        return code, 1
    if 0x20 <= code <= 0x7E:
        return code, 1
    return ord("?"), 1


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
        token, consumed = _consume_x07_text_unit(text, i)
        out.append(token)
        i += consumed
    return bytes(out)


def _normalize_basic_source_line(line: str) -> tuple[int, str] | None:
    line = line.rstrip("\r\n")
    i = 0
    while i < len(line) and line[i] == " ":
        i += 1

    start = i
    while i < len(line) and line[i].isdigit():
        i += 1

    if i == start:
        return None

    line_number = int(line[start:i], 10)
    if line_number <= 0:
        return None

    if i < len(line) and line[i] == ":":
        i += 1
    while i < len(line) and line[i] == " ":
        i += 1

    src = line[i:]
    out: list[str] = []
    transparent = False
    in_string = False

    for ch in src:
        current = "".join(out)
        if not in_string and not transparent and (
            current.endswith("'") or current.endswith("DATA") or current.endswith("REM")
        ):
            transparent = True

        if ch == '"':
            in_string = not in_string

        out.append(ch if (in_string or transparent) else ch.upper())

    return line_number, "".join(out)


def _tokenize_basic_body(text: str) -> bytes:
    out = bytearray()
    i = 0
    in_string = False
    transparent = False

    while i < len(text):
        ch = text[i]
        if ch == '"':
            in_string = not in_string

        if not in_string and not transparent and not ch.isdigit():
            matched = False
            for token_text, token_value, flags in X07_BASIC_TOKEN_SPECS:
                if text.startswith(token_text, i):
                    display_text = X07_BASIC_DETOKEN_TOKENS[token_value - 0x80]
                    if display_text.startswith(" ") and out and out[-1] == ord(" "):
                        out.pop()
                    if flags & TOKEN_FLAG_TRANSPARENT:
                        transparent = True
                    if flags & TOKEN_FLAG_PREPEND_COLON:
                        out.append(ord(":"))
                    if flags & TOKEN_FLAG_PREPEND_REM:
                        out.append(0x8E)
                    out.append(token_value)
                    i += len(token_text)
                    if display_text.endswith(" ") and i < len(text) and text[i] == " ":
                        i += 1
                    matched = True
                    break
            if matched:
                continue

        value, consumed = _consume_x07_text_unit(text, i)
        out.append(value)
        i += consumed

    return bytes(out)


def _parse_basic_source(text: str) -> list[tuple[int, str]]:
    parsed: list[tuple[int, str]] = []
    last_line_number = 0

    for raw_line in text.splitlines():
        normalized = _normalize_basic_source_line(raw_line)
        if normalized is None:
            continue
        line_number, body = normalized
        if line_number > 0xFFFF:
            raise RuntimeError(f"Line number too large: {line_number}")
        if line_number <= last_line_number:
            raise RuntimeError(f"Line numbers must be strictly increasing (got {line_number} after {last_line_number}).")
        parsed.append((line_number, body))
        last_line_number = line_number

    if not parsed:
        raise RuntimeError("No line numbers found.")

    return parsed


def build_tokenized_basic_payload(lines: list[tuple[int, str]]) -> bytes:
    payload = bytearray()
    line_pointer = BASIC_START

    for line_number, body in lines:
        encoded = _tokenize_basic_body(body)
        record_len = 4 + len(encoded) + 1
        line_pointer += record_len
        if line_pointer > 0xFFFF:
            raise RuntimeError("Tokenized BASIC program exceeds 16-bit pointer range.")
        payload.extend((
            line_pointer & 0xFF,
            (line_pointer >> 8) & 0xFF,
            line_number & 0xFF,
            (line_number >> 8) & 0xFF,
        ))
        payload.extend(encoded)
        payload.append(0x00)

    payload.extend(b"\x00" * 11)
    return bytes(payload)


def _decode_x07_text_byte(value: int) -> str:
    if value in X07_BYTE_TO_UNICODE:
        return X07_BYTE_TO_UNICODE[value]
    if 0x20 <= value <= 0x7E:
        return chr(value)
    return f"\\{value:02X}"


def _detokenize_basic_body(content: bytes) -> str:
    out: list[str] = []
    insert_space = False
    last_char = ""
    quoted = False
    quoted_until_eol = False
    colon_pending = False
    remark_pending = False

    def emit_text(text: str) -> None:
        nonlocal insert_space, last_char, quoted, quoted_until_eol

        if text and text[0] == " " and last_char == " ":
            text = text[1:]

        for idx, ch in enumerate(text):
            if ch == '"':
                quoted = (not quoted) or quoted_until_eol
                insert_space = False
            elif ch == "\n":
                quoted = False
                quoted_until_eol = False
                insert_space = False

            if not quoted and ord(ch) < 0x20 and ch != "\n":
                out.append(f"\\{ord(ch):02X}")
            elif quoted:
                out.append(ch)
                insert_space = False
            else:
                if ch == " " and idx == len(text) - 1:
                    insert_space = True
                else:
                    if insert_space and (ch.isdigit() or ch >= "A"):
                        out.append(" ")
                    out.append(ch)
                    insert_space = False
            last_char = ch

    for value in content + b"\x00":
        if colon_pending:
            colon_pending = False
            if value == 0x8E:
                remark_pending = True
                continue
            if value != 0x90:
                emit_text(":")
        elif remark_pending:
            remark_pending = False
            if value != 0xFF:
                emit_text(":REM")
                quoted = True
                quoted_until_eol = True

        if value == ord(":") and not quoted:
            colon_pending = True
            continue

        if value == 0x00:
            emit_text("\n")
            continue

        token_text: str | None = None
        if not quoted and value >= 0x80:
            token_text = X07_BASIC_DETOKEN_TOKENS[value - 0x80]
            if value in (0x83, 0x8E, 0xFF):
                quoted = True
                quoted_until_eol = True

        if token_text is not None:
            emit_text(token_text)
        elif value == 0x20:
            out.append(" ")
            insert_space = False
            last_char = " "
        else:
            emit_text(_decode_x07_text_byte(value))

    return "".join(out).rstrip("\n")


def parse_tokenized_basic_payload(payload: bytes) -> list[tuple[int, bytes]]:
    lines: list[tuple[int, bytes]] = []
    pos = 0

    while pos + 4 <= len(payload):
        next_ptr = payload[pos] | (payload[pos + 1] << 8)
        if next_ptr == 0:
            break

        line_number = payload[pos + 2] | (payload[pos + 3] << 8)
        end = payload.find(b"\x00", pos + 4)
        if end < 0:
            raise RuntimeError("Unexpected BASIC payload format: missing line terminator.")

        expected_next_ptr = BASIC_START + end + 1
        if next_ptr != expected_next_ptr:
            raise RuntimeError(
                f"Unexpected BASIC payload format: inconsistent line pointer 0x{next_ptr:04X} "
                f"(expected 0x{expected_next_ptr:04X})."
            )

        lines.append((line_number, bytes(payload[pos + 4:end])))
        pos = end + 1

    if not lines:
        raise RuntimeError("Unexpected BASIC payload format: no tokenized BASIC lines found.")

    return lines


def detokenize_basic_payload(payload: bytes) -> str:
    lines = parse_tokenized_basic_payload(payload)
    return "\n".join(f"{line_number} {_detokenize_basic_body(content)}" for line_number, content in lines)


def list_serial_ports():
    if IS_MACOS and DARWIN_TERMIOS_AVAILABLE:
        ports = sorted(str(p) for p in Path("/dev").glob("cu.*"))
        if ports:
            return ports

    if HAS_PYSERIAL:
        ports = [p.device for p in serial.tools.list_ports.comports()]
        if IS_MACOS:
            cu_ports = [p for p in ports if p.startswith("/dev/cu.")]
            if cu_ports:
                return cu_ports
        return ports

    return []


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
        self.geometry("1040x710")
        self.minsize(920, 580)

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
        self.remote_ser: SerialPortLike | None = None

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

        btn_send_cas = ttk.Button(rc, text="Send raw", command=self.send_cas_raw)
        btn_send_cas.pack(side="left", padx=(6, 0))
        self._transfer_controls.append(btn_send_cas)

        btn_recv_cas = ttk.Button(rc, text="Receive raw", command=self.receive_cas_raw)
        btn_recv_cas.pack(side="left", padx=(6, 0))
        self._transfer_controls.append(btn_recv_cas)

        btn_cas_to_text = ttk.Button(rc, text="Convert to text", command=self.convert_cas_to_text)
        btn_cas_to_text.pack(side="left", padx=(6, 0))
        self._transfer_controls.append(btn_cas_to_text)

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

        btn_text_to_cas = ttk.Button(rb, text="Convert to CAS/K7", command=self.convert_basic_to_cas)
        btn_text_to_cas.pack(side="left", padx=(6, 0))
        self._transfer_controls.append(btn_text_to_cas)

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

        row_btns = ttk.Frame(kbd)
        row_btns.pack(fill="x", pady=(6, 0))

        def kbtn(label: str, cmd, width=6):
            def run_and_refocus():
                cmd()
                self.after_idle(self._focus_remote_input)

            b = ttk.Button(row_btns, text=label, width=width, command=run_and_refocus)
            b.pack(side="left", padx=(2, 0))
            self._kbd_controls.append(b)
            return b

        kbtn("HOME",  lambda: self._kbd_send_byte(KEY_HOME))
        kbtn("CLR",   lambda: self._kbd_send_byte(KEY_CLR), width=5)
        kbtn("INS",   lambda: self._kbd_send_byte(KEY_INS), width=5)
        kbtn("DEL",   lambda: self._kbd_send_byte(KEY_DEL), width=5)
        kbtn("BREAK", lambda: self._kbd_send_byte(KEY_ON_BREAK), width=6)

        ttk.Label(
            kbd,
            text="Type here (sent to X-07 while REMOTE KEYBOARD is ON):",
        ).pack(anchor="w", pady=(6, 0))
        self.relay_box = tk.Text(kbd, height=1, wrap="none")
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

        self.txt = tk.Text(bottom, height=10, wrap="word")
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
        self.log(f"[INFO] Serial backend: {serial_backend_summary()}.")
        self.log("[INFO] Quick reminders:")
        self.log('       Text listing / Remote keyboard: on X-07, run INIT#5,"COM:" then EXEC&HEE1F.')
        self.log('       Cassette stream PC -> X-07: on X-07, run LOAD"COM:" then click Send raw.')
        self.log('       Cassette stream X-07 -> PC: on X-07, run SAVE"COM:" then click Receive raw.')
        self.log("       ASM transfer: send loader.cas first, then send the ASM binary.")
        self.log("       Cassette files: send uses fixed base 0x0010; receive saves a .cas file with D3 header + raw stream.")
        self.log("       Exit SLAVE mode: EXEC&HEE33 or power cycle.")

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
        return settings_config_path()

    def _load_serial_settings(self) -> dict:
        return load_saved_serial_settings()

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

    def _serial_backend_name(self, ser: SerialPortLike | None = None) -> str:
        if ser is not None:
            return "termios" if getattr(ser, "_termios_backend", False) else "pyserial"
        if IS_MACOS and DARWIN_TERMIOS_AVAILABLE and HAS_PYSERIAL:
            return "termios preferred on macOS (pyserial fallback)"
        if IS_MACOS and DARWIN_TERMIOS_AVAILABLE:
            return "termios"
        if HAS_PYSERIAL:
            return "pyserial"
        return "unknown"

    def _open_serial(self, baud: int) -> SerialPortLike:
        port = self._require_port()
        timeout = 0.2

        if IS_MACOS and DARWIN_TERMIOS_AVAILABLE:
            try:
                return MacTermiosSerial(port, baud, timeout=timeout, rtscts=self.var_rtscts.get())
            except Exception as e:
                if HAS_PYSERIAL:
                    self.log(f"[WARN] termios open failed on {port!r}: {e}. Falling back to pyserial.")
                else:
                    raise RuntimeError(f"Cannot open {port!r} with termios backend: {e}") from e

        if not HAS_PYSERIAL:
            raise RuntimeError("No serial backend available. Install pyserial or use the macOS termios backend.")

        return serial.Serial(
            port,
            int(baud),
            bytesize=serial.EIGHTBITS,
            parity=serial.PARITY_NONE,
            stopbits=serial.STOPBITS_TWO,
            timeout=timeout,
            xonxoff=False,
            rtscts=self.var_rtscts.get(),
            dsrdtr=False,
        )

    def _open_for_typing(self) -> SerialPortLike:
        return self._open_serial(int(self.var_typing_baud.get()))

    def _open_for_raw(self) -> SerialPortLike:
        return self._open_serial(int(self.var_xfer_baud.get()))

    def _prepare_raw_transfer(self, ser: SerialPortLike):
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

        # Parse BASIC tokenized lines dynamically so loader.cas updates stay robust.
        lines = []
        pos = 0
        while pos + 4 <= len(payload):
            next_ptr = payload[pos] | (payload[pos + 1] << 8)
            if next_ptr == 0:
                break
            line_no = payload[pos + 2] | (payload[pos + 3] << 8)
            end = pos + 4
            while end < len(payload) and payload[end] != 0:
                end += 1
            if end >= len(payload):
                raise RuntimeError("Unexpected loader.cas format: unterminated BASIC line in payload.")
            content_start = pos + 4
            content_end = end
            lines.append({
                "line_no": line_no,
                "start": pos,
                "content_start": content_start,
                "content_end": content_end,
            })
            pos = end + 1

        if not lines:
            raise RuntimeError("Unexpected loader.cas format: no BASIC payload lines found.")

        def _line_content(line) -> bytes:
            return bytes(payload[line["content_start"]:line["content_end"]])

        def _find_loader_line(token: int, min_fields: int, label: str):
            matches = []
            for ln in lines:
                content = _line_content(ln)
                if not content or content[0] != token:
                    continue
                fields = list(re.finditer(rb"&H[0-9A-F]{4}", content))
                if len(fields) >= min_fields:
                    matches.append(ln)
            if len(matches) != 1:
                found = [ln["line_no"] for ln in matches]
                raise RuntimeError(
                    f"Unexpected loader.cas format: could not uniquely identify {label} line "
                    f"(matches={found or 'none'})."
                )
            return matches[0]

        def _patch_ascii_field(line, repl_values, label: str):
            content = _line_content(line)
            fields = list(re.finditer(rb"&H[0-9A-F]{4}", content))
            if len(fields) < len(repl_values):
                raise RuntimeError(
                    f"Unexpected loader.cas format: {label} line has only {len(fields)} address field(s)."
                )
            patched = bytearray(content)
            for m, repl in zip(fields, repl_values):
                if len(repl) != len(m.group(0)):
                    raise RuntimeError("Internal error: BASIC address replacement length changed.")
                patched[m.start():m.end()] = repl
            payload[line["content_start"]:line["content_end"]] = patched

        clear_line = _find_loader_line(0xA3, 1, "CLEAR")
        for_line = _find_loader_line(0x81, 2, "FOR")
        exec_line = _find_loader_line(0xA8, 1, "EXEC")

        data_lines = [ln for ln in lines if _line_content(ln)[:1] == b"\x83"]
        if not data_lines:
            raise RuntimeError("Unexpected loader.cas format: DATA lines not found.")

        data_spans = []
        loader_bytes = bytearray()
        for ln in data_lines:
            start_ = ln["content_start"]
            end_ = ln["content_end"]
            if payload[start_] == 0x83:  # DATA token
                start_ += 1
            chunk = payload[start_:end_].decode("ascii")
            items = chunk.split(",") if chunk else []
            data_spans.append((start_, end_, len(items)))
            loader_bytes.extend(int(x, 16) for x in items)

        loader_byte_count = len(loader_bytes)
        if loader_byte_count <= 0:
            raise RuntimeError("Unexpected loader.cas format: no DATA bytes found.")

        def _extract_first_ascii_addr(line, label: str) -> int:
            content = _line_content(line)
            m = re.search(rb"&H([0-9A-F]{4})", content)
            if m is None:
                raise RuntimeError(f"Unexpected loader.cas format: {label} line has no &Hxxxx address field.")
            return int(m.group(1), 16)

        orig_base = _extract_first_ascii_addr(clear_line, "CLEAR")
        orig_end_addr = orig_base + loader_byte_count - 1
        if orig_end_addr > 0xFFFF:
            raise RuntimeError("Unexpected loader.cas format: loader image exceeds 16-bit address space.")

        end_addr = (load_addr + loader_byte_count - 1) & 0xFFFF

        _patch_ascii_field(clear_line, [f"&H{load_addr:04X}".encode("ascii")], "CLEAR")
        _patch_ascii_field(
            for_line,
            [f"&H{load_addr:04X}".encode("ascii"), f"&H{end_addr:04X}".encode("ascii")],
            "FOR",
        )
        _patch_ascii_field(exec_line, [f"&H{load_addr:04X}".encode("ascii")], "EXEC")

        def _is_loader_address(addr: int) -> bool:
            return orig_base <= addr <= orig_end_addr

        def relocate_absolute_operands(opcodes: set[int], *, skip_indexed_loads: bool = False) -> int:
            count = 0
            for idx in range(len(loader_bytes) - 2):
                opcode = loader_bytes[idx]
                if opcode not in opcodes:
                    continue
                if skip_indexed_loads and idx > 0 and loader_bytes[idx - 1] in (0xDD, 0xFD):
                    continue
                addr = loader_bytes[idx + 1] | (loader_bytes[idx + 2] << 8)
                if not _is_loader_address(addr):
                    continue
                relocated = (load_addr + (addr - orig_base)) & 0xFFFF
                loader_bytes[idx + 1] = relocated & 0xFF
                loader_bytes[idx + 2] = (relocated >> 8) & 0xFF
                count += 1
            return count

        relocate_absolute_operands({0x22, 0x2A, 0x32, 0x3A, 0xCD, 0xC3})
        relocate_absolute_operands({0x01, 0x11, 0x21, 0x31}, skip_indexed_loads=True)

        baud = int(self.var_xfer_baud.get()) & 0xFFFF
        baud_pattern = bytes([0xDD, 0x21, 0x40, 0x1F])  # template uses 8000 baud
        baud_repl = bytes([0xDD, 0x21, baud & 0xFF, (baud >> 8) & 0xFF])
        idx = loader_bytes.find(baud_pattern)
        if idx < 0:
            raise RuntimeError("Unexpected loader.cas format: loader baud pattern not found.")
        loader_bytes[idx:idx + 4] = baud_repl

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

    def _send_loader_cas_raw(self, ser: SerialPortLike, *, progress_base: float = 0.0, progress_span: float = 100.0):
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

    def _send_loader_cas_remote(self, ser: SerialPortLike, *, progress_base: float = 0.0, progress_span: float = 50.0):
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


    def _type_line(self, ser: SerialPortLike, line: str):
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

    def _tcdrain(self, ser: SerialPortLike):
        try:
            fd = ser.fileno()
        except Exception:
            return
        if not HAS_TERMIOS:
            return
        try:
            termios.tcdrain(fd)
        except Exception:
            pass

    def _stream_raw_payload(self, ser: SerialPortLike, payload: bytes, label: str,
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

    def _send_ascii_frame(self, ser: SerialPortLike, frame: bytes, label: str,
                          progress_base: float = 0.0, progress_span: float = 100.0,
                          chunk_size: int = 256):
        total = len(frame)
        if total <= 0:
            return
        chunk_size = max(1, int(chunk_size))
        # No additional per-character delay for ASM loader/data transfer.
        self._stream_raw_payload(ser, frame, label, progress_base, progress_span, chunk_size)

    def _build_current_asm_frame(self) -> bytes:
        if not self.bin_file or not self.bin_file.exists():
            raise RuntimeError("Select a .bin file first.")
        data = self.bin_file.read_bytes()
        frame = self._build_loader_ascii_frame(data)
        self.log(f"[INFO] BIN size: {len(data)} bytes. Target load: 0x{self._parse_addr():04X}")
        self.log(f"[INFO] Sending ASCII loader frame: {len(frame)} chars @ {int(self.var_xfer_baud.get())} 8N2.")
        return frame

    def _send_current_asm_frame(self, frame: bytes, *, progress_base: float = 0.0, progress_span: float = 100.0):
        with self._open_for_raw() as ser:
            self.log(f"[INFO] Opened {ser.port} for ASM send: {ser.baudrate} 8N2 ({self._serial_backend_name(ser)}).")
            self._prepare_raw_transfer(ser)
            self._send_ascii_frame(
                ser,
                frame,
                "ASM frame",
                progress_base=progress_base,
                progress_span=progress_span,
                chunk_size=256,
            )

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

    def convert_basic_to_cas(self):
        if not self.basic_file or not self.basic_file.exists():
            messagebox.showwarning("Missing BASIC", "Select a BASIC .txt/.bas file first.")
            return

        p = filedialog.asksaveasfilename(
            title="Save tokenized cassette stream as",
            defaultextension=".cas",
            initialfile=self.basic_file.with_suffix(".cas").name,
            filetypes=[("CAS/K7", "*.cas *.k7"), ("CAS", "*.cas"), ("K7", "*.k7"), ("All files", "*.*")],
        )
        if not p:
            return

        out_path = Path(p)
        self._run_threaded(lambda: self._convert_basic_to_cas_impl(out_path), "Convert BASIC to cassette stream")

    def _convert_basic_to_cas_impl(self, out_path: Path):
        self._set_status("converting BASIC to cassette stream")
        self._set_progress(0, "reading BASIC listing...")

        source = self.basic_file.read_text(encoding="utf-8", errors="replace")
        parsed = _parse_basic_source(source)
        payload = bytearray()
        line_pointer = BASIC_START
        total = len(parsed)

        self.log(f"[INFO] Converting BASIC listing: {self.basic_file.name}")
        self.log(f"[INFO] Output cassette stream: {out_path.name}")

        for idx, (line_number, body) in enumerate(parsed, start=1):
            if self.cancel_event.is_set():
                raise InterruptedError("Cancelled during BASIC tokenization.")

            encoded = _tokenize_basic_body(body)
            record_len = 4 + len(encoded) + 1
            line_pointer += record_len
            if line_pointer > 0xFFFF:
                raise RuntimeError("Tokenized BASIC program exceeds 16-bit pointer range.")

            payload.extend((
                line_pointer & 0xFF,
                (line_pointer >> 8) & 0xFF,
                line_number & 0xFF,
                (line_number >> 8) & 0xFF,
            ))
            payload.extend(encoded)
            payload.append(0x00)

            self._set_progress((idx / total) * 100.0, f"tokenizing BASIC line {idx}/{total}")

        payload.extend(b"\x00" * 11)
        out_path.write_bytes(build_cas_header_from_filename(out_path) + bytes(payload))
        self._set_progress(100.0, "conversion done")
        self.log(f"[OK] Tokenized BASIC saved: {out_path} ({len(payload)} payload bytes, {total} line(s)).")

    def convert_cas_to_text(self):
        if not self.cas_file or not self.cas_file.exists():
            messagebox.showwarning("Missing CAS/K7", "Select a .cas/.k7 file first.")
            return

        p = filedialog.asksaveasfilename(
            title="Save BASIC text listing as",
            defaultextension=".bas",
            initialfile=self.cas_file.with_suffix(".bas").name,
            filetypes=[("BASIC", "*.bas *.txt"), ("BAS", "*.bas"), ("Text", "*.txt"), ("All files", "*.*")],
        )
        if not p:
            return

        out_path = Path(p)
        self._run_threaded(lambda: self._convert_cas_to_text_impl(out_path), "Convert cassette stream to text listing")

    def _convert_cas_to_text_impl(self, out_path: Path):
        self._set_status("converting cassette stream to BASIC text")
        self._set_progress(0, "reading cassette stream...")

        data = self.cas_file.read_bytes()
        if len(data) <= CAS_FIXED_BASE:
            raise RuntimeError("Selected CAS/K7 file is too short.")

        payload = data[CAS_FIXED_BASE:]
        lines = parse_tokenized_basic_payload(payload)
        total = len(lines)
        text_lines: list[str] = []

        self.log(f"[INFO] Converting cassette stream: {self.cas_file.name}")
        self.log(f"[INFO] Output BASIC listing: {out_path.name}")

        for idx, (line_number, content) in enumerate(lines, start=1):
            if self.cancel_event.is_set():
                raise InterruptedError("Cancelled during BASIC detokenization.")
            text_lines.append(f"{line_number} {_detokenize_basic_body(content)}")
            self._set_progress((idx / total) * 100.0, f"detokenizing BASIC line {idx}/{total}")

        out_path.write_text("\n".join(text_lines) + "\n", encoding="utf-8")
        self._set_progress(100.0, "conversion done")
        self.log(f"[OK] BASIC listing saved: {out_path} ({total} line(s)).")

    # ---------------- Slave mode helper ----------------
    def disable_slave_mode(self):
        self._run_threaded(self._disable_slave_mode_impl, "Disable slave mode (EXEC&HEE33)")

    def _disable_slave_mode_impl(self):
        self._set_status("sending EXEC&HEE33")
        self._set_progress(0, "sending EXEC&HEE33")
        try:
            with self._open_for_typing() as ser:
                self._type_line(ser, "EXEC&HEE33")
        except SERIAL_ERRORS as e:
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
                self.log(f"[INFO] Opened {ser.port} for typing: {ser.baudrate} 8N2 ({self._serial_backend_name(ser)}).")
                for idx, ln in enumerate(lines, start=1):
                    if self.cancel_event.is_set():
                        raise InterruptedError("Cancelled during BASIC send.")
                    self._type_line(ser, ln)
                    self._set_progress((idx / total) * 100.0, f"BASIC {idx}/{total} lines")

                if self.cancel_event.is_set():
                    raise InterruptedError("Cancelled before EXEC&HEE33.")
                self._type_line(ser, "EXEC&HEE33")
        except SERIAL_ERRORS as e:
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
                self.log(f"[INFO] Opened {ser.port} for loader raw send: {ser.baudrate} 8N2 ({self._serial_backend_name(ser)}).")
                self._send_loader_cas_raw(ser)
        except SERIAL_ERRORS as e:
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

        frame = self._build_current_asm_frame()

        try:
            self._send_current_asm_frame(frame)
        except SERIAL_ERRORS as e:
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

        frame = self._build_current_asm_frame()

        try:
            with self._open_for_typing() as ser:
                self.log(f"[INFO] Opened {ser.port} for one-click loader stage: {ser.baudrate} 8N2 ({self._serial_backend_name(ser)}).")
                self._send_loader_cas_remote(ser, progress_base=0.0, progress_span=50.0)
                self._set_progress(50.0, "loader started")
            time.sleep(POST_LOADER_EXEC_DELAY_S)
            self._send_current_asm_frame(frame, progress_base=50.0, progress_span=50.0)
        except SERIAL_ERRORS as e:
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
                self.log(f"[INFO] Opened {ser.port} for CAS/K7 send: {ser.baudrate} 8N2 ({self._serial_backend_name(ser)}).")
                self._stream_raw_payload(
                    ser,
                    payload,
                    'CAS/K7 send',
                    progress_base=0.0,
                    progress_span=100.0,
                    chunk_size=64,
                    final_rts_drop=True,
                )

        except SERIAL_ERRORS as e:
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
                self.log(f"[INFO] Opened {ser.port} for CAS/K7 receive: {ser.baudrate} 8N2 ({self._serial_backend_name(ser)}).")
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
        except SERIAL_ERRORS as e:
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
        except SERIAL_ERRORS as e:
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
        self.log(
            f"[OK] REMOTE KEYBOARD: ON ({self.remote_ser.port} @ {self.remote_ser.baudrate} 8N2, "
            f"{self._serial_backend_name(self.remote_ser)})."
        )
        self.log('[INFO] Remote keyboard requires SLAVE mode (INIT#5,"COM:" then EXEC&HEE1F).')
        self.log("[INFO] PC arrow keys are sent when the remote input box has focus.")
        self.after_idle(self._focus_remote_input)

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

    def _focus_remote_input(self):
        if not self.remote_kbd_on:
            return
        try:
            self.relay_box.focus_set()
        except Exception:
            pass

    def _kbd_send_byte(self, b: int):
        if not self.remote_kbd_on or not self.remote_ser:
            return
        try:
            self.remote_ser.write(bytes([b & 0xFF]))
            self.remote_ser.flush()
        except SERIAL_ERRORS as e:
            self.log(f"[ERROR] Remote keyboard write failed: {e}")
            self._remote_keyboard_off(reason="[WARN] REMOTE KEYBOARD stopped (write error).")

    def _kbd_send_bytes(self, data: bytes):
        if not self.remote_kbd_on or not self.remote_ser:
            return
        try:
            self.remote_ser.write(data)
            self.remote_ser.flush()
        except SERIAL_ERRORS as e:
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


class SimpleVar:
    def __init__(self, value):
        self._value = value

    def get(self):
        return self._value

    def set(self, value):
        self._value = value


class X07CliRunner:
    _require_port = X07LoaderApp._require_port
    _serial_backend_name = X07LoaderApp._serial_backend_name
    _open_serial = X07LoaderApp._open_serial
    _open_for_typing = X07LoaderApp._open_for_typing
    _open_for_raw = X07LoaderApp._open_for_raw
    _prepare_raw_transfer = X07LoaderApp._prepare_raw_transfer
    _parse_addr = X07LoaderApp._parse_addr
    _parse_loader_addr = X07LoaderApp._parse_loader_addr
    _loader_cas_path = X07LoaderApp._loader_cas_path
    _build_loader_cas_payload = X07LoaderApp._build_loader_cas_payload
    _build_loader_ascii_frame = X07LoaderApp._build_loader_ascii_frame
    _send_loader_cas_raw = X07LoaderApp._send_loader_cas_raw
    _send_loader_cas_remote = X07LoaderApp._send_loader_cas_remote
    _type_line = X07LoaderApp._type_line
    _tcdrain = X07LoaderApp._tcdrain
    _stream_raw_payload = X07LoaderApp._stream_raw_payload
    _send_ascii_frame = X07LoaderApp._send_ascii_frame
    _build_current_asm_frame = X07LoaderApp._build_current_asm_frame
    _send_current_asm_frame = X07LoaderApp._send_current_asm_frame
    _convert_basic_to_cas_impl = X07LoaderApp._convert_basic_to_cas_impl
    _convert_cas_to_text_impl = X07LoaderApp._convert_cas_to_text_impl
    _disable_slave_mode_impl = X07LoaderApp._disable_slave_mode_impl
    _send_basic_file_impl = X07LoaderApp._send_basic_file_impl
    _send_fast_loader_impl = X07LoaderApp._send_fast_loader_impl
    _send_bin_only_impl = X07LoaderApp._send_bin_only_impl
    _send_loader_and_bin_impl = X07LoaderApp._send_loader_and_bin_impl
    _send_cas_raw_impl = X07LoaderApp._send_cas_raw_impl
    _receive_cas_raw_impl = X07LoaderApp._receive_cas_raw_impl

    def __init__(self, args):
        settings = load_saved_serial_settings()

        self.cancel_event = threading.Event()
        self.basic_file = getattr(args, "input", None)
        self.bin_file = getattr(args, "input", None)
        self.cas_file = getattr(args, "input", None)

        self.var_port = SimpleVar(getattr(args, "port", None) or settings.get("port") or "")
        self.var_rtscts = SimpleVar(bool(getattr(args, "rtscts", settings.get("rtscts", False))))
        self.var_typing_baud = SimpleVar(int(getattr(args, "typing_baud", settings.get("typing_baud", 4800))))
        self.var_xfer_baud = SimpleVar(int(getattr(args, "xfer_baud", settings.get("xfer_baud", 8000))))
        self.var_char = SimpleVar(float(getattr(args, "char_delay", settings.get("char_delay", DEFAULT_CHAR_DELAY_S))))
        self.var_line = SimpleVar(float(getattr(args, "line_delay", settings.get("line_delay", DEFAULT_LINE_DELAY_S))))
        self.var_loader_addr = SimpleVar(str(getattr(args, "loader_addr", settings.get("loader_addr", hex(DEFAULT_LOADER_ADDR)))))
        self.var_addr = SimpleVar(str(getattr(args, "asm_addr", settings.get("asm_addr", hex(DEFAULT_LOAD_ADDR)))))

        self._last_status: str | None = None
        self._last_progress_bucket: int | None = None

    def _ts(self) -> str:
        return time.strftime("[%H:%M:%S]")

    def log(self, msg: str):
        print(f"{self._ts()} {msg}", flush=True)

    def _set_status(self, text: str):
        if text != self._last_status:
            self._last_status = text
            self.log(f"[STATUS] {text}")

    def _set_progress(self, pct: float, label: str):
        pct = max(0.0, min(100.0, float(pct)))
        if pct >= 100.0:
            bucket = 10
        else:
            bucket = int(pct // 10)
        if bucket == self._last_progress_bucket and pct not in (0.0, 100.0):
            return
        self._last_progress_bucket = bucket
        if pct <= 0.0:
            self.log(f"[PROGRESS] {label}")
        elif pct >= 100.0:
            self.log(f"[PROGRESS] 100% - {label}")
        else:
            self.log(f"[PROGRESS] {pct:.0f}% - {label}")

    def inspect_cas_header_cli(self):
        if not self.cas_file or not self.cas_file.exists():
            raise RuntimeError("Select a .cas/.k7 file first.")
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

    def run(self, name: str, fn) -> int:
        try:
            self.log(f"--- {name} ---")
            fn()
            self._set_status("done")
            return 0
        except KeyboardInterrupt:
            self.cancel_event.set()
            self.log(f"[WARN] {name} cancelled by user.")
            return 130
        except InterruptedError as e:
            self.log(f"[WARN] {name} cancelled: {e}")
            return 130
        except Exception as e:
            self.log(f"[ERROR] {name}: {e!r}")
            return 1


def _add_port_arg(parser, settings: dict):
    parser.add_argument(
        "--port",
        default=settings.get("port") or "",
        help='Serial port to use (default: last saved port from x07_loader.ini, if any).',
    )


def _add_rtscts_args(parser, settings: dict):
    group = parser.add_mutually_exclusive_group()
    group.add_argument("--rtscts", dest="rtscts", action="store_true", help="Use RTS/CTS hardware flow control.")
    group.add_argument("--no-rtscts", dest="rtscts", action="store_false", help="Disable RTS/CTS hardware flow control.")
    parser.set_defaults(rtscts=bool(settings.get("rtscts", False)))


def _add_typing_args(parser, settings: dict, *, include_delays: bool = False):
    _add_port_arg(parser, settings)
    _add_rtscts_args(parser, settings)
    parser.add_argument(
        "--typing-baud",
        type=int,
        default=int(settings.get("typing_baud", 4800)),
        help="Baud rate for text/CAS transfers in 8N2 mode.",
    )
    if include_delays:
        parser.add_argument(
            "--char-delay",
            type=float,
            default=float(settings.get("char_delay", DEFAULT_CHAR_DELAY_S)),
            help="Delay in seconds between typed characters.",
        )
        parser.add_argument(
            "--line-delay",
            type=float,
            default=float(settings.get("line_delay", DEFAULT_LINE_DELAY_S)),
            help="Delay in seconds after each typed line.",
        )


def _add_loader_runtime_args(parser, settings: dict, *, include_loader_addr: bool = False, include_asm_addr: bool = False):
    parser.add_argument(
        "--xfer-baud",
        type=int,
        default=int(settings.get("xfer_baud", 8000)),
        help="Loader/runtime baud rate for ASM transfers in 8N2 mode.",
    )
    if include_loader_addr:
        parser.add_argument(
            "--loader-addr",
            default=str(settings.get("loader_addr", hex(DEFAULT_LOADER_ADDR))),
            help="Loader address (for example 0x1F00 or 1F00h).",
        )
    if include_asm_addr:
        parser.add_argument(
            "--asm-addr",
            default=str(settings.get("asm_addr", hex(DEFAULT_LOAD_ADDR))),
            help="ASM load address (for example 0x2000 or 2000h).",
        )


def build_cli_parser() -> argparse.ArgumentParser:
    settings = load_saved_serial_settings()
    parser = argparse.ArgumentParser(
        description="Canon X-07 Serial Fast Loader CLI. Run without arguments to start the GUI.",
    )
    subparsers = parser.add_subparsers(dest="command", required=True)

    subparsers.add_parser("gui", help="Start the graphical interface.")
    subparsers.add_parser("ports", help="List available serial ports.")

    p = subparsers.add_parser("inspect-cas", help="Inspect the header of a .cas/.k7 file.")
    p.add_argument("input", type=Path, help="Input .cas/.k7 file.")

    p = subparsers.add_parser("convert-basic", help="Convert a BASIC .txt/.bas listing to a tokenized .cas/.k7 file.")
    p.add_argument("input", type=Path, help="Input BASIC .txt/.bas file.")
    p.add_argument("output", type=Path, help="Output .cas/.k7 file.")

    p = subparsers.add_parser("convert-cas", help="Convert a tokenized .cas/.k7 file to a BASIC text listing.")
    p.add_argument("input", type=Path, help="Input .cas/.k7 file.")
    p.add_argument("output", type=Path, help="Output .bas/.txt file.")

    p = subparsers.add_parser("disable-slave", help='Send EXEC&HEE33 to leave SLAVE mode.')
    _add_typing_args(p, settings, include_delays=True)

    p = subparsers.add_parser("send-basic", help="Send a BASIC .txt/.bas listing through SLAVE mode.")
    p.add_argument("input", type=Path, help="Input BASIC .txt/.bas file.")
    _add_typing_args(p, settings, include_delays=True)

    p = subparsers.add_parser("send-cas", help='Send a .cas/.k7 raw cassette stream (X-07 side: LOAD"COM:").')
    p.add_argument("input", type=Path, help="Input .cas/.k7 file.")
    _add_typing_args(p, settings, include_delays=False)

    p = subparsers.add_parser("receive-cas", help='Receive a .cas file from the X-07 (X-07 side: SAVE"COM:").')
    p.add_argument("output", type=Path, help="Output .cas file.")
    _add_typing_args(p, settings, include_delays=False)

    p = subparsers.add_parser("send-loader", help='Send loader.cas only (X-07 side: LOAD"COM:").')
    _add_typing_args(p, settings, include_delays=False)
    _add_loader_runtime_args(p, settings, include_loader_addr=True, include_asm_addr=False)

    p = subparsers.add_parser("send-bin", help="Send an ASM .bin file through the already running loader.")
    p.add_argument("input", type=Path, help="Input ASM .bin file.")
    _add_port_arg(p, settings)
    _add_rtscts_args(p, settings)
    _add_loader_runtime_args(p, settings, include_loader_addr=False, include_asm_addr=True)

    p = subparsers.add_parser("send-asm", help='One-click ASM transfer: send loader.cas, then the ASM .bin file (SLAVE mode required).')
    p.add_argument("input", type=Path, help="Input ASM .bin file.")
    _add_typing_args(p, settings, include_delays=True)
    _add_loader_runtime_args(p, settings, include_loader_addr=True, include_asm_addr=True)

    return parser


def run_cli(args) -> int:
    if args.command == "gui":
        app = X07LoaderApp()
        app.mainloop()
        return 0

    if args.command == "ports":
        ports = list_serial_ports()
        if not ports:
            print("No serial port found.", flush=True)
            return 1
        for port in ports:
            print(port, flush=True)
        return 0

    runner = X07CliRunner(args)

    if args.command == "inspect-cas":
        return runner.run("Inspect CAS/K7 header", runner.inspect_cas_header_cli)
    if args.command == "convert-basic":
        return runner.run("Convert BASIC to cassette stream", lambda: runner._convert_basic_to_cas_impl(args.output))
    if args.command == "convert-cas":
        return runner.run("Convert cassette stream to text listing", lambda: runner._convert_cas_to_text_impl(args.output))
    if args.command == "disable-slave":
        return runner.run("Disable SLAVE mode (EXEC&HEE33)", runner._disable_slave_mode_impl)
    if args.command == "send-basic":
        return runner.run("Send BASIC (.txt/.bas)", runner._send_basic_file_impl)
    if args.command == "send-cas":
        return runner.run('Send CAS/K7 raw stream (LOAD"COM:")', runner._send_cas_raw_impl)
    if args.command == "receive-cas":
        return runner.run('Receive CAS/K7 raw stream (SAVE"COM:")', lambda: runner._receive_cas_raw_impl(args.output))
    if args.command == "send-loader":
        return runner.run('Send ASM loader (LOAD"COM:")', runner._send_fast_loader_impl)
    if args.command == "send-bin":
        return runner.run("Send ASM (loader running)", runner._send_bin_only_impl)
    if args.command == "send-asm":
        return runner.run("One click: loader + ASM", runner._send_loader_and_bin_impl)

    raise RuntimeError(f"Unknown command: {args.command}")


def main(argv=None) -> int:
    argv = list(sys.argv[1:] if argv is None else argv)
    force_cli = False
    if argv and argv[0] == "--cli":
        force_cli = True
        argv = argv[1:]

    parser = build_cli_parser()

    if not argv and not force_cli:
        app = X07LoaderApp()
        app.mainloop()
        return 0

    if not argv and force_cli:
        parser.print_help()
        return 0

    args = parser.parse_args(argv)
    return run_cli(args)


if __name__ == "__main__":
    raise SystemExit(main())
