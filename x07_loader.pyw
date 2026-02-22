import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import threading
import time
from pathlib import Path
import sys
import subprocess
import re

try:
    import serial
    import serial.tools.list_ports
    from serial.serialutil import SerialException
except ImportError:
    subprocess.check_call([sys.executable, "-m", "pip", "install", "pyserial"])
    import serial
    import serial.tools.list_ports
    from serial.serialutil import SerialException


# ---------- Defaults ----------
DEFAULT_CHAR_DELAY_S = 0.04
DEFAULT_LINE_DELAY_S = 0.20
DEFAULT_POST_INIT_DELAY_S = 2.0
DEFAULT_BYTE_DELAY_S = 0.02
DEFAULT_LOAD_ADDR = 0x2000

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


def list_serial_ports():
    try:
        return [p.device for p in serial.tools.list_ports.comports()]
    except Exception:
        return []


def guess_name_in_first_16_bytes(data: bytes, fallback: str) -> str:
    """Best-effort: try to read an ASCII-ish name from first 16 bytes."""
    if len(data) < 0x10:
        return fallback

    txt = "".join(chr(b) if 0x20 <= b <= 0x7E else " " for b in data[:0x10])
    tokens = re.findall(r"[A-Za-z0-9][A-Za-z0-9 _\-\[\]]{2,24}", txt)
    return max(tokens, key=len).strip() if tokens else fallback


def build_cas_header_from_filename(path: Path) -> bytes:
    """
    Canon X-07 .CAS header (empirical, per your samples):
      - 10 bytes: 0xD3 repeated
      - 6 bytes : ASCII name (max 6 chars), padded with 0x00
    Total = 16 bytes (0x0010)

    We take the destination filename stem, sanitize to [A-Z0-9_], uppercase, truncate to 6.
    """
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
    """
    Scrollable frame that:
      - shows scrollbar only when needed
      - disables mousewheel scroll when not needed
      - when not scrolling: stretches inner window height to canvas height
        so the console extends to the bottom.
    """
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
        self.var_typing_baud = tk.IntVar(value=4800)   # 8N2
        self.var_xfer_baud = tk.IntVar(value=8000)     # 7E1

        self.var_char = tk.DoubleVar(value=DEFAULT_CHAR_DELAY_S)
        self.var_line = tk.DoubleVar(value=DEFAULT_LINE_DELAY_S)
        self.var_post = tk.DoubleVar(value=DEFAULT_POST_INIT_DELAY_S)
        self.var_byte = tk.DoubleVar(value=DEFAULT_BYTE_DELAY_S)

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

        ttk.Label(r, text="Xfer baud (7E1):").pack(side="left")
        ttk.Entry(r, textvariable=self.var_xfer_baud, width=7).pack(side="left", padx=(4, 10))

        ttk.Label(r, text="CHAR(s):").pack(side="left")
        ttk.Entry(r, textvariable=self.var_char, width=6).pack(side="left", padx=(3, 8))
        ttk.Label(r, text="LINE(s):").pack(side="left")
        ttk.Entry(r, textvariable=self.var_line, width=6).pack(side="left", padx=(3, 8))
        ttk.Label(r, text="PostINIT(s):").pack(side="left")
        ttk.Entry(r, textvariable=self.var_post, width=6).pack(side="left", padx=(3, 8))
        ttk.Label(r, text="Byte(s):").pack(side="left")
        ttk.Entry(r, textvariable=self.var_byte, width=6).pack(side="left", padx=(3, 0))

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
            text='Note: BASIC TXT, ASM and REMOTE KEYBOARD use SLAVE mode (INIT#5,"COM:" then EXEC&HEE1F).',
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

        rc = ttk.Frame(cas); rc.pack(fill="x", pady=(4, 0))
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

        ttk.Label(cas, text="Send base: fixed 0x0010 (validated). Receive: saves as CAS with D3 header + stream.").pack(anchor="w", pady=(4, 0))

        # TXT/BAS (RIGHT)
        txt = ttk.LabelFrame(right, text="Text listing (.txt/.bas) via SLAVE mode", padding=6)
        txt.pack(fill="both", expand=True)

        self.lbl_basic = ttk.Label(txt, text="Selected BASIC: (none)")
        self.lbl_basic.pack(anchor="w")

        rb = ttk.Frame(txt); rb.pack(fill="x", pady=(4, 0))
        btn_pick_basic = ttk.Button(rb, text="Select .txt/.bas…", command=self.pick_basic)
        btn_pick_basic.pack(side="left")
        self._transfer_controls.append(btn_pick_basic)

        btn_send_basic = ttk.Button(rb, text="Send BASIC", command=self.send_basic_file)
        btn_send_basic.pack(side="left", padx=(6, 0))
        self._transfer_controls.append(btn_send_basic)

        # ---- ASM ----
        asm = ttk.LabelFrame(main, text="ASM via SLAVE mode", padding=6)
        asm.pack(fill="x", pady=(0, 6))

        self.var_addr = tk.StringVar(value=hex(DEFAULT_LOAD_ADDR))
        self.var_append_run = tk.BooleanVar(value=True)

        self.lbl_bin = ttk.Label(asm, text="Selected ASM binary: (none)")
        self.lbl_bin.pack(anchor="w")

        ra = ttk.Frame(asm); ra.pack(fill="x", pady=(4, 0))
        btn_pick_bin = ttk.Button(ra, text="Select bin…", command=self.pick_bin)
        btn_pick_bin.pack(side="left")
        self._transfer_controls.append(btn_pick_bin)

        ttk.Label(ra, text="Load addr:").pack(side="left", padx=(10, 4))
        ttk.Entry(ra, textvariable=self.var_addr, width=10).pack(side="left")

        btn_send_loader = ttk.Button(ra, text="Send BASIC fast loader", command=self.send_fast_loader)
        btn_send_loader.pack(side="left", padx=(10, 0))
        self._transfer_controls.append(btn_send_loader)

        ttk.Checkbutton(ra, text="Append RUN", variable=self.var_append_run).pack(side="left", padx=(8, 0))

        btn_send_asm = ttk.Button(ra, text="Send ASM (loader running)", command=self.send_bin_only)
        btn_send_asm.pack(side="left", padx=(10, 0))
        self._transfer_controls.append(btn_send_asm)

        btn_one_click = ttk.Button(ra, text="One click: loader + ASM", command=self.send_loader_and_bin)
        btn_one_click.pack(side="left", padx=(10, 0))
        self._transfer_controls.append(btn_one_click)

        # ---- Remote keyboard ----
        kbd = ttk.LabelFrame(main, text='Remote keyboard via SLAVE mode (PC -> X-07)', padding=6)
        kbd.pack(fill="x", pady=(0, 6))

        rk = ttk.Frame(kbd); rk.pack(fill="x")
        self.btn_remote_toggle = ttk.Button(rk, text="REMOTE KEYBOARD: OFF", command=self.toggle_remote_keyboard)
        self.btn_remote_toggle.pack(side="left")
        self._always_enabled_controls.append(self.btn_remote_toggle)

        ttk.Label(rk, text='(requires SLAVE mode: INIT#5,"COM:" then EXEC&HEE1F)').pack(side="left", padx=(10, 0))

        row_btns = ttk.Frame(kbd); row_btns.pack(fill="x", pady=(6, 0))

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

        def macro(label: str, text: str, w=4):
            b = ttk.Button(row_btns, text=label, width=w,
                           command=lambda: self._kbd_send_bytes(text.encode("ascii", errors="replace")))
            b.pack(side="left", padx=(2, 0))
            self._kbd_controls.append(b)

        macro("F1",  "?TIME$\r")
        macro("F2",  'CLOAD"\r')
        macro("F3",  "LOCATE ")
        macro("F4",  "LIST ")
        macro("F5",  "RUN\r")
        macro("F6",  "?DATE$\r")
        macro("F7",  'CSAVE"\r')
        macro("F8",  "PRINT ")
        macro("F9",  "SLEEP")
        macro("F10", "CONT\r", w=5)

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

    # ---------------- Logging ----------------
    def _ts(self) -> str:
        return time.strftime("[%H:%M:%S]")

    def log(self, msg: str):
        self.txt.insert("end", f"{self._ts()} {msg}\n")
        self.txt.see("end")
        self.update_idletasks()

    def _log_startup_banner(self):
        self.log("Canon X-07 Serial Fast Loader - ready.")
        self.log('SLAVE mode (BASIC TXT + ASM + REMOTE KEYBOARD): INIT#5,"COM:" then EXEC&HEE1F (user guide p.119).')
        self.log('CAS/K7 raw stream: use LOAD"COM:" or SAVE"COM:" on X-07, then send/receive raw bytes on PC.')
        self.log("Exit slave: EXEC&HEE33 (remote) or power cycle.")
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
    def refresh_ports(self, initial=False):
        ports = list_serial_ports()
        self.cbo_port["values"] = ports
        cur = (self.var_port.get() or "").strip()
        if cur in ports:
            self.var_port.set(cur)
        else:
            self.var_port.set(ports[0] if ports else "")
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
            write_timeout=2.0,
            xonxoff=False,
            rtscts=False,
            dsrdtr=False,
        )

    def _switch_to_xfer(self, ser: serial.Serial):
        ser.baudrate = int(self.var_xfer_baud.get())
        ser.bytesize = serial.SEVENBITS
        ser.parity = serial.PARITY_EVEN
        ser.stopbits = serial.STOPBITS_ONE

    def _parse_addr(self) -> int:
        s = self.var_addr.get().strip().lower()
        if s.endswith("h"):
            return int(s[:-1], 16)
        return int(s, 0)

    def _type_line(self, ser: serial.Serial, line: str):
        char_delay = float(self.var_char.get())
        line_delay = float(self.var_line.get())
        for ch in line.encode("ascii", errors="replace"):
            if self.cancel_event.is_set():
                raise InterruptedError("Cancelled during BASIC typing.")
            ser.write(bytes([ch]))
            ser.flush()
            time.sleep(char_delay)
        ser.write(b"\r")
        ser.flush()
        time.sleep(line_delay)

    # ---------------- File pickers ----------------
    def pick_basic(self):
        p = filedialog.askopenfilename(
            title="Select BASIC program (.txt/.bas)",
            filetypes=[("BASIC text", "*.txt *.bas"), ("All files", "*.*")]
        )
        if not p:
            return
        self.basic_file = Path(p)
        self.lbl_basic.config(text=f"Selected BASIC: {self.basic_file}")
        self.log(f"[OK] BASIC selected: {self.basic_file}")

    def pick_bin(self):
        p = filedialog.askopenfilename(
            title="Select ASM binary (.bin)",
            filetypes=[("Binary", "*.bin"), ("All files", "*.*")]
        )
        if not p:
            return
        self.bin_file = Path(p)
        self.lbl_bin.config(text=f"Selected ASM binary: {self.bin_file}")
        self.log(f"[OK] BIN selected: {self.bin_file}")

    def pick_cas(self):
        p = filedialog.askopenfilename(
            title="Select CAS/K7",
            filetypes=[("CAS/K7", "*.cas *.k7"), ("All files", "*.*")]
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
        self._run_threaded(self._send_fast_loader_impl, "Send BASIC fast loader")

    def _build_fast_loader_lines(self) -> list[str]:
        addr = self._parse_addr()
        xfer_baud = int(self.var_xfer_baud.get())
        append_run = self.var_append_run.get()
        lines = [
            "NEW",
            "1EXEC&HEE33",
            f"2Z=&H{addr:04X}",
            f'3INIT#1,"COM:",{xfer_baud},"G',
            "4INPUT#1,N",
            "5FORI=0TO N-1",
            "6INPUT#1,A$",
            "7POKE Z+I,VAL(A$)",
            "8NEXT I",
            "9EXECZ",
        ]
        if append_run:
            lines.append("RUN")
        return lines

    def _send_fast_loader_impl(self):
        self._set_status("typing fast loader")
        loader_lines = self._build_fast_loader_lines()
        total = max(1, len(loader_lines))
        self._set_progress(0, "typing loader...")

        try:
            with self._open_for_typing() as ser:
                for i, l in enumerate(loader_lines, start=1):
                    if self.cancel_event.is_set():
                        raise InterruptedError("Cancelled during loader typing.")
                    self._type_line(ser, l)
                    self._set_progress((i / total) * 100.0, f"Loader {i}/{total} lines")
        except (SerialException, OSError) as e:
            self.log(f"[ERROR] Cannot open {self.var_port.get()!r}: {e}")
            return

        self._set_progress(100.0, "loader done")
        if self.var_append_run.get():
            self.log("[INFO] Loader typed and started (RUN).")
        else:
            self.log("[INFO] Loader typed (no RUN). Start manually if needed.")

    def send_bin_only(self):
        if not self.bin_file or not self.bin_file.exists():
            messagebox.showwarning("Missing BIN", "Select a .bin file first.")
            return
        self._run_threaded(self._send_bin_only_impl, "Send ASM (loader running)")

    def _send_bin_only_impl(self):
        self._set_status("transferring ASM (.bin)")
        self._set_progress(0, "starting...")

        data = self.bin_file.read_bytes()
        n = len(data)
        if n <= 0:
            raise RuntimeError("Binary is empty.")

        addr = self._parse_addr()
        self.log(f"[INFO] BIN size: {n} bytes. Target load: 0x{addr:04X}")

        try:
            with self._open_for_typing() as ser:
                self._switch_to_xfer(ser)
                self.log(f"[INFO] Switched to transfer: {ser.baudrate} 7E1.")
                time.sleep(float(self.var_post.get()))

                ser.write(f"{n}\r".encode("ascii"))
                ser.flush()
                time.sleep(0.01)

                byte_delay = float(self.var_byte.get())
                for i, b in enumerate(data, start=1):
                    if self.cancel_event.is_set():
                        raise InterruptedError("Cancelled during ASM transfer.")
                    ser.write(f"{b}\r".encode("ascii"))
                    ser.flush()
                    if byte_delay > 0:
                        time.sleep(byte_delay)
                    if i == 1 or i == n or (i % 32 == 0):
                        self._set_progress((i / n) * 100.0, f"ASM {i}/{n} bytes")
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
        self._set_status("typing loader + transferring ASM")
        self._set_progress(0, "starting...")

        data = self.bin_file.read_bytes()
        n = len(data)
        if n <= 0:
            raise RuntimeError("Binary is empty.")

        loader_lines = self._build_fast_loader_lines()
        total_loader = max(1, len(loader_lines))

        try:
            with self._open_for_typing() as ser:
                for i, l in enumerate(loader_lines, start=1):
                    if self.cancel_event.is_set():
                        raise InterruptedError("Cancelled during loader typing.")
                    self._type_line(ser, l)
                    self._set_progress((i / total_loader) * 50.0, f"Loader {i}/{total_loader} lines")

                if not self.var_append_run.get():
                    self.log("[WARN] Loader typed without RUN. Enable Append RUN or start manually.")
                    return

                self._switch_to_xfer(ser)
                self.log(f"[INFO] Switched to transfer: {ser.baudrate} 7E1.")
                time.sleep(float(self.var_post.get()))

                ser.write(f"{n}\r".encode("ascii"))
                ser.flush()
                time.sleep(0.01)

                byte_delay = float(self.var_byte.get())
                for i, b in enumerate(data, start=1):
                    if self.cancel_event.is_set():
                        raise InterruptedError("Cancelled during ASM transfer.")
                    ser.write(f"{b}\r".encode("ascii"))
                    ser.flush()
                    if byte_delay > 0:
                        time.sleep(byte_delay)
                    if i == 1 or i == n or (i % 32 == 0):
                        pct = 50.0 + (i / n) * 50.0
                        self._set_progress(pct, f"ASM {i}/{n} bytes")
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
                sent = 0
                CHUNK = 128  # smaller chunks => more reliable on some macOS USB-serial stacks
                while sent < total:
                    if self.cancel_event.is_set():
                        raise InterruptedError("Cancelled during CAS/K7 transfer.")
                    end = min(total, sent + CHUNK)
                    ser.write(payload[sent:end])
                    ser.flush()
                    sent = end
                    if sent == total or (sent % 4096 == 0):
                        self._set_progress((sent / total) * 100.0, f'CAS/K7 send {sent}/{total} bytes')

                # --- Ensure end-of-transfer really reaches the Canon before closing ---
                ser.flush()
                time.sleep(0.25)

                # Force end marker (13 x 0x00) per manual (harmless if already present)
                ser.write(b"\x00" * 13)
                ser.flush()
                time.sleep(0.25)
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

                        if len(buf) == len(chunk) or (len(buf) % 4096 == 0):
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
            self._kbd_send_byte(KEY_LEFT);  return "break"
        if k == "Right":
            self._kbd_send_byte(KEY_RIGHT); return "break"
        if k == "Up":
            self._kbd_send_byte(KEY_UP);    return "break"
        if k == "Down":
            self._kbd_send_byte(KEY_DOWN);  return "break"
        if k == "Home":
            self._kbd_send_byte(KEY_HOME);  return "break"
        if k == "Insert":
            self._kbd_send_byte(KEY_INS);   return "break"
        if k == "Delete":
            self._kbd_send_byte(KEY_DEL);   return "break"
        if k in ("Return", "KP_Enter"):
            self._kbd_send_byte(KEY_RETURN); return "break"
        if k == "space":
            self._kbd_send_byte(KEY_SPACE);  return "break"

        ch = event.char
        if ch:
            b = ord(ch)
            if 0x20 <= b <= 0x7E:
                self._kbd_send_byte(b)
        return "break"


if __name__ == "__main__":
    app = X07LoaderApp()
    app.mainloop()
