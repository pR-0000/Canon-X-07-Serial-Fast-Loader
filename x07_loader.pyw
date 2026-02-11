import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import threading
import time
from pathlib import Path
try:
    import serial
except ImportError:
    subprocess.check_call([sys.executable, "-m", "pip", "install", "pyserial"])
    import serial
import serial.tools.list_ports

# ---------- Defaults ----------
DEFAULT_CHAR_DELAY = 0.04
DEFAULT_LINE_DELAY = 0.20
DEFAULT_POST_INIT_DELAY = 2.0
DEFAULT_BYTE_DELAY = 0.02


def list_serial_ports():
    """Reliable port listing if pyserial tools available, else fallback."""
    try:
        from serial.tools import list_ports
        return [p.device for p in list_ports.comports()]
    except Exception:
        return [f"COM{i}" for i in range(1, 41)]


class X07LoaderApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Canon X-07 Serial Fast Loader")
        self.geometry("1080x760")
        self.minsize(1020, 680)

        self.basic_file: Path | None = None
        self.bin_file: Path | None = None

        # Cancellation / job control
        self.cancel_event = threading.Event()
        self.job_lock = threading.Lock()
        self.job_running = False

        self._build_ui()

        # Force a proper initial refresh AFTER widgets exist
        self.refresh_ports_btn(initial=True)
        self._log_startup_banner()

    # ---------------- UI ----------------
    def _build_ui(self):
        root = ttk.Frame(self, padding=10)
        root.pack(fill="both", expand=True)

        # ==== SERIAL SETTINGS ====
        settings = ttk.LabelFrame(root, text="Serial settings", padding=10)
        settings.pack(fill="x")

        # No default COM port for publishing; leave empty by default
        self.var_port = tk.StringVar(value="")
        self.var_type_baud = tk.IntVar(value=4800)   # BASIC typing (8N2)
        self.var_xfer_baud = tk.IntVar(value=8000)   # transfer (7E1)

        self.var_char_delay = tk.DoubleVar(value=DEFAULT_CHAR_DELAY)
        self.var_line_delay = tk.DoubleVar(value=DEFAULT_LINE_DELAY)
        self.var_post_init_delay = tk.DoubleVar(value=DEFAULT_POST_INIT_DELAY)
        self.var_byte_delay = tk.DoubleVar(value=DEFAULT_BYTE_DELAY)

        row0 = ttk.Frame(settings)
        row0.pack(fill="x")

        ttk.Label(row0, text="COM port:").pack(side="left")
        self.cbo_port = ttk.Combobox(row0, textvariable=self.var_port, width=12, state="readonly")
        self.cbo_port.pack(side="left", padx=(6, 8))

        ttk.Button(row0, text="Refresh", command=self.refresh_ports_btn).pack(side="left", padx=(0, 16))

        ttk.Label(row0, text="Typing baud (8N2):").pack(side="left")
        ttk.Entry(row0, textvariable=self.var_type_baud, width=8).pack(side="left", padx=(6, 16))

        ttk.Label(row0, text="Transfer baud (7E1):").pack(side="left")
        ttk.Entry(row0, textvariable=self.var_xfer_baud, width=8).pack(side="left", padx=(6, 16))

        row1 = ttk.Frame(settings)
        row1.pack(fill="x", pady=(10, 0))

        ttk.Label(row1, text="CHAR_DELAY:").pack(side="left")
        ttk.Entry(row1, textvariable=self.var_char_delay, width=7).pack(side="left", padx=(6, 14))
        ttk.Label(row1, text="LINE_DELAY:").pack(side="left")
        ttk.Entry(row1, textvariable=self.var_line_delay, width=7).pack(side="left", padx=(6, 14))
        ttk.Label(row1, text="Post INIT wait (s):").pack(side="left")
        ttk.Entry(row1, textvariable=self.var_post_init_delay, width=7).pack(side="left", padx=(6, 14))
        ttk.Label(row1, text="Byte delay (s):").pack(side="left")
        ttk.Entry(row1, textvariable=self.var_byte_delay, width=7).pack(side="left", padx=(6, 14))

        # Cancel button (global)
        row_cancel = ttk.Frame(settings)
        row_cancel.pack(fill="x", pady=(10, 0))
        self.btn_cancel = ttk.Button(row_cancel, text="Cancel current transfer", command=self.cancel_current, state="disabled")
        self.btn_cancel.pack(side="left")
        ttk.Label(row_cancel, text="(Stops BASIC/ASM sending loops safely)").pack(side="left", padx=(10, 0))

        # ==== BASIC ====
        files = ttk.LabelFrame(root, text="BASIC program (.txt)", padding=10)
        files.pack(fill="x", pady=(10, 0))

        self.lbl_basic = ttk.Label(files, text="Selected BASIC: (none)")
        self.lbl_basic.pack(anchor="w")

        rowb = ttk.Frame(files)
        rowb.pack(fill="x", pady=(6, 4))
        ttk.Button(rowb, text="Select BASIC .txt…", command=self.pick_basic).pack(side="left")
        ttk.Button(rowb, text="Send BASIC (.txt)", command=self.send_basic_file).pack(side="left", padx=(10, 0))
        ttk.Button(rowb, text="Disable slave mode (EXEC&HEE33)", command=self.disable_slave_mode).pack(side="left", padx=(10, 0))

        # BASIC progress
        self.basic_prog_var = tk.DoubleVar(value=0)
        self.basic_prog = ttk.Progressbar(files, variable=self.basic_prog_var, maximum=100, mode="determinate")
        self.basic_prog.pack(fill="x", pady=(6, 0))
        self.basic_prog_label = ttk.Label(files, text="BASIC progress: idle")
        self.basic_prog_label.pack(anchor="w", pady=(2, 0))

        # ==== FAST LOADER + ASM ====
        xfer = ttk.LabelFrame(root, text="Fast loader + ASM transfer", padding=10)
        xfer.pack(fill="x", pady=(10, 0))

        self.var_addr = tk.StringVar(value="0x1000")
        self.var_append_run = tk.BooleanVar(value=True)

        # --- Requested layout order ---
        # 1) file selector first (with filename shown above)
        self.lbl_bin = ttk.Label(xfer, text="Selected ASM binary (.bin): (none)")
        self.lbl_bin.pack(anchor="w")

        rowbin = ttk.Frame(xfer)
        rowbin.pack(fill="x", pady=(6, 4))
        ttk.Button(rowbin, text="Select bin…", command=self.pick_bin).pack(side="left")

        # 2) load address
        ttk.Label(rowbin, text="Load address (hex):").pack(side="left", padx=(18, 6))
        ttk.Entry(rowbin, textvariable=self.var_addr, width=10).pack(side="left")

        # 3) send loader
        rowl = ttk.Frame(xfer)
        rowl.pack(fill="x", pady=(8, 0))
        ttk.Button(rowl, text="Send BASIC fast loader", command=self.send_fast_loader).pack(side="left")

        # 4) append RUN checkbox
        ttk.Checkbutton(rowl, text="Append RUN (auto start loader)", variable=self.var_append_run).pack(side="left", padx=(14, 0))

        # 5) send ASM / one-click
        rowbtn = ttk.Frame(xfer)
        rowbtn.pack(fill="x", pady=(8, 0))
        ttk.Button(rowbtn, text="Send ASM binary (expects BASIC loader running)", command=self.send_bin_only).pack(side="left")
        ttk.Button(rowbtn, text="Send fast loader + ASM (one click)", command=self.send_loader_and_bin).pack(side="left", padx=(14, 0))

        # ASM progress
        self.asm_prog_var = tk.DoubleVar(value=0)
        self.asm_prog = ttk.Progressbar(xfer, variable=self.asm_prog_var, maximum=100, mode="determinate")
        self.asm_prog.pack(fill="x", pady=(10, 0))
        self.asm_prog_label = ttk.Label(xfer, text="ASM progress: idle")
        self.asm_prog_label.pack(anchor="w", pady=(2, 0))

        # Status line
        self.status_var = tk.StringVar(value="Status: idle")
        ttk.Label(root, textvariable=self.status_var).pack(anchor="w", pady=(10, 0))

        # ==== CONSOLE ====
        console = ttk.LabelFrame(root, text="Console", padding=10)
        console.pack(fill="both", expand=True, pady=(10, 0))

        self.txt = tk.Text(console, height=16, wrap="word")
        self.txt.pack(fill="both", expand=True, side="left")
        scroll = ttk.Scrollbar(console, command=self.txt.yview)
        scroll.pack(fill="y", side="right")
        self.txt.configure(yscrollcommand=scroll.set)

    # ---------------- Ports ----------------
    def refresh_ports_btn(self, initial=False):
        ports = list_serial_ports()
        self.cbo_port["values"] = ports

        current = (self.var_port.get() or "").strip()
        if current and current in ports:
            # keep user selection
            self.var_port.set(current)
        else:
            # no default selection; for initial load keep empty if possible
            if initial:
                self.var_port.set("")
            else:
                # if user hit refresh and their selection disappeared, leave blank
                self.var_port.set("")

        if not initial:
            self.log("[INFO] COM ports refreshed.")
        else:
            if ports:
                self.log(f"[INFO] COM ports loaded: {', '.join(ports[:10])}{'...' if len(ports) > 10 else ''}")
            else:
                self.log("[WARN] No COM ports detected. Plug your USB-serial adapter and click Refresh.")

    # ---------------- Logging ----------------
    def _ts(self) -> str:
        # Always local time: [hh:mm:ss]
        return time.strftime("[%H:%M:%S]")

    def log(self, msg: str):
        self.txt.insert("end", f"{self._ts()} {msg}\n")
        self.txt.see("end")
        self.update_idletasks()

    def _log_startup_banner(self):
        self.log("Canon X-07 Serial Fast Loader - ready.")
        self.log("")
        self.log("=== IMPORTANT (Slave mode reminders) ===")
        self.log("Enter slave mode (minimal typing):")
        self.log('  INIT#5,"COM:')
        self.log('  EXEC&HEE1F')
        self.log("Disable slave mode (remote):")
        self.log("  EXEC&HEE33")
        self.log("---------------------------------------")
        self.log("This tool types BASIC in 8N2, then switches PC side to 7E1 for fast transfer (mode 'G').")
        self.log("")

    # ---------------- File pickers ----------------
    def pick_basic(self):
        p = filedialog.askopenfilename(
            title="Select BASIC program (.txt)",
            filetypes=[("Text file", "*.txt"), ("All files", "*.*")]
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
        self.lbl_bin.config(text=f"Selected ASM binary (.bin): {self.bin_file}")
        self.log(f"[OK] BIN selected: {self.bin_file}")

    # ---------------- Cancellation ----------------
    def cancel_current(self):
        self.cancel_event.set()
        self.log("[WARN] Cancel requested by user...")

    def _job_start(self):
        with self.job_lock:
            if self.job_running:
                raise RuntimeError("A transfer is already running. Cancel it first.")
            self.job_running = True
        self.cancel_event.clear()
        self.after(0, lambda: self.btn_cancel.config(state="normal"))

    def _job_end(self):
        with self.job_lock:
            self.job_running = False
        self.after(0, lambda: self.btn_cancel.config(state="disabled"))

    # ---------------- Progress helpers (thread-safe via after) ----------------
    def _set_status(self, text: str):
        self.after(0, lambda: self.status_var.set(f"Status: {text}"))

    def _set_basic_progress(self, pct: float, label: str):
        def _u():
            self.basic_prog_var.set(max(0.0, min(100.0, pct)))
            self.basic_prog_label.config(text=f"BASIC progress: {label}")
        self.after(0, _u)

    def _set_asm_progress(self, pct: float, label: str):
        def _u():
            self.asm_prog_var.set(max(0.0, min(100.0, pct)))
            self.asm_prog_label.config(text=f"ASM progress: {label}")
        self.after(0, _u)

    # ---------------- Serial helpers ----------------
    def _require_port(self) -> str:
        p = (self.var_port.get() or "").strip()
        if not p:
            raise RuntimeError("No COM port selected. Choose a port in 'Serial settings'.")
        return p

    def _open_for_typing(self) -> serial.Serial:
        # BASIC typing: 8N2
        port = self._require_port()
        return serial.Serial(
            port,
            int(self.var_type_baud.get()),
            bytesize=serial.EIGHTBITS,
            parity=serial.PARITY_NONE,
            stopbits=serial.STOPBITS_TWO,
            timeout=0.2,
            write_timeout=1.0
        )

    def _switch_to_xfer(self, ser: serial.Serial):
        # Transfer: 7E1
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
        char_delay = float(self.var_char_delay.get())
        line_delay = float(self.var_line_delay.get())
        for ch in line.encode("ascii", errors="replace"):
            if self.cancel_event.is_set():
                raise InterruptedError("Cancelled during BASIC typing.")
            ser.write(bytes([ch]))
            ser.flush()
            time.sleep(char_delay)
        ser.write(b"\r")
        ser.flush()
        time.sleep(line_delay)

    # ---------------- Thread wrapper ----------------
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
                self._job_end()
        threading.Thread(target=runner, daemon=True).start()

    # ---------------- Actions ----------------
    def send_basic_file(self):
        if not self.basic_file or not self.basic_file.exists():
            messagebox.showwarning("Missing BASIC", "Please select a BASIC .txt file first.")
            return
        self._run_threaded(self._send_basic_file_impl, "Send BASIC (.txt)")

    def _send_basic_file_impl(self):
        self._set_status("typing BASIC (.txt)")
        self._set_basic_progress(0, "starting...")

        lines = self.basic_file.read_text(encoding="utf-8", errors="replace").splitlines()
        lines = [ln.rstrip("\r\n") for ln in lines if ln.strip() != ""]
        total = max(1, len(lines) + 1)  # +1 for EXEC&HEE33

        with self._open_for_typing() as ser:
            self.log(f"[INFO] Opened {ser.port} for typing: {ser.baudrate} 8N2.")
            for idx, ln in enumerate(lines, start=1):
                if self.cancel_event.is_set():
                    raise InterruptedError("Cancelled during BASIC send.")
                self._type_line(ser, ln)
                pct = (idx / total) * 100.0
                self._set_basic_progress(pct, f"{idx}/{total} lines")

            if self.cancel_event.is_set():
                raise InterruptedError("Cancelled before EXEC&HEE33.")
            self._type_line(ser, "EXEC&HEE33")
            self._set_basic_progress(100.0, f"{total}/{total} lines (EXEC&HEE33 sent)")
            self.log("[INFO] Sent EXEC&HEE33 after BASIC to release slave mode (if active).")

    def disable_slave_mode(self):
        self._run_threaded(self._disable_slave_mode_impl, "Disable slave mode (EXEC&HEE33)")

    def _disable_slave_mode_impl(self):
        self._set_status("sending EXEC&HEE33")
        with self._open_for_typing() as ser:
            self.log(f"[INFO] Opened {ser.port} for typing: {ser.baudrate} 8N2.")
            self._type_line(ser, "EXEC&HEE33")
            self.log("[INFO] EXEC&HEE33 sent.")

    def send_fast_loader(self):
        self._run_threaded(self._send_fast_loader_impl, "Send BASIC fast loader")

    def _build_fast_loader_lines(self) -> list[str]:
        addr = self._parse_addr()
        xfer_baud = int(self.var_xfer_baud.get())
        append_run = self.var_append_run.get()

        lines = [
            "NEW",
            "1 EXEC&HEE33",
            f"2 Z=&H{addr:04X}",
            f'3 INIT#1,"COM:",{xfer_baud},"G',
            "4 INPUT#1,N",
            "5 FORI=0TO N-1",
            "6 INPUT#1,A$",
            "7 POKE Z+I,VAL(A$)",
            "8 NEXT I",
            "9 EXECZ",
        ]
        if append_run:
            lines.append("RUN")
        return lines

    def _send_fast_loader_impl(self):
        self._set_status("typing fast loader")
        self._set_basic_progress(0, "typing loader...")

        loader_lines = self._build_fast_loader_lines()
        total = max(1, len(loader_lines))

        with self._open_for_typing() as ser:
            self.log(f"[INFO] Opened {ser.port} for typing: {ser.baudrate} 8N2.")
            for i, l in enumerate(loader_lines, start=1):
                if self.cancel_event.is_set():
                    raise InterruptedError("Cancelled during loader typing.")
                self._type_line(ser, l)
                self._set_basic_progress((i / total) * 100.0, f"{i}/{total} lines")

            if not self.var_append_run.get():
                if self.cancel_event.is_set():
                    raise InterruptedError("Cancelled before EXEC&HEE33.")
                self._type_line(ser, "EXEC&HEE33")
                self.log("[INFO] Loader typed but NOT started (no RUN). Sent EXEC&HEE33 to release slave mode.")
            else:
                self.log("[INFO] Loader typed and started (RUN). Not sending EXEC&HEE33 now.")

        if self.var_append_run.get():
            self.log("[INFO] X-07 should now be waiting for: N<CR> then N decimal byte lines (7E1).")

    def send_bin_only(self):
        if not self.bin_file or not self.bin_file.exists():
            messagebox.showwarning("Missing BIN", "Please select a .bin file first.")
            return
        self._run_threaded(self._send_bin_only_impl, "Send ASM binary (expects BASIC loader running)")

    def _send_bin_only_impl(self):
        self._set_status("transferring ASM (.bin)")
        self._set_asm_progress(0, "starting...")

        addr = self._parse_addr()
        data = self.bin_file.read_bytes()
        n = len(data)
        if n <= 0:
            raise RuntimeError("Binary is empty.")

        self.log(f"[INFO] BIN size: {n} bytes. Target load address: 0x{addr:04X}")
        self.log("[INFO] This expects the fast loader already running on the X-07.")

        with self._open_for_typing() as ser:
            self.log(f"[INFO] Opened {ser.port} for typing: {ser.baudrate} 8N2.")
            self._switch_to_xfer(ser)
            self.log(f"[INFO] Switched to transfer: {ser.baudrate} 7E1.")
            time.sleep(float(self.var_post_init_delay.get()))

            if self.cancel_event.is_set():
                raise InterruptedError("Cancelled before length send.")

            ser.write(f"{n}\r".encode("ascii"))
            ser.flush()
            time.sleep(0.01)

            byte_delay = float(self.var_byte_delay.get())
            for i, b in enumerate(data, start=1):
                if self.cancel_event.is_set():
                    raise InterruptedError("Cancelled during ASM transfer.")
                ser.write(f"{b}\r".encode("ascii"))
                ser.flush()
                if byte_delay > 0:
                    time.sleep(byte_delay)
                pct = (i / n) * 100.0
                if i == 1 or i == n or (i % 32 == 0):
                    self._set_asm_progress(pct, f"{i}/{n} bytes")

        self._set_asm_progress(100.0, f"{n}/{n} bytes (done)")
        self.log("[INFO] Transfer complete. If loader executes 'EXEC Z', your program should be running now.")

    def send_loader_and_bin(self):
        if not self.bin_file or not self.bin_file.exists():
            messagebox.showwarning("Missing BIN", "Please select a .bin file first.")
            return
        self._run_threaded(self._send_loader_and_bin_impl, "Send fast loader + ASM (one click)")

    def _send_loader_and_bin_impl(self):
        self._set_status("typing loader + transferring ASM")
        self._set_basic_progress(0, "typing loader...")
        self._set_asm_progress(0, "pending...")

        addr = self._parse_addr()
        data = self.bin_file.read_bytes()
        n = len(data)
        if n <= 0:
            raise RuntimeError("Binary is empty.")

        loader_lines = self._build_fast_loader_lines()
        total_loader = max(1, len(loader_lines))

        with self._open_for_typing() as ser:
            self.log(f"[INFO] Opened {ser.port} for typing: {ser.baudrate} 8N2.")

            # 1) Type loader
            for i, l in enumerate(loader_lines, start=1):
                if self.cancel_event.is_set():
                    raise InterruptedError("Cancelled during loader typing.")
                self._type_line(ser, l)
                self._set_basic_progress((i / total_loader) * 100.0, f"{i}/{total_loader} lines")

            if not self.var_append_run.get():
                self.log("[WARN] Loader was typed without RUN. Enable 'Append RUN' or start it manually.")
                return

            if self.cancel_event.is_set():
                raise InterruptedError("Cancelled before transfer phase.")

            # 2) Switch PC side to 7E1
            self._switch_to_xfer(ser)
            self.log(f"[INFO] Switched to transfer: {ser.baudrate} 7E1.")
            time.sleep(float(self.var_post_init_delay.get()))

            if self.cancel_event.is_set():
                raise InterruptedError("Cancelled before length send.")

            # 3) Send N then bytes
            ser.write(f"{n}\r".encode("ascii"))
            ser.flush()
            time.sleep(0.01)

            byte_delay = float(self.var_byte_delay.get())
            for i, b in enumerate(data, start=1):
                if self.cancel_event.is_set():
                    raise InterruptedError("Cancelled during ASM transfer.")
                ser.write(f"{b}\r".encode("ascii"))
                ser.flush()
                if byte_delay > 0:
                    time.sleep(byte_delay)
                pct = (i / n) * 100.0
                if i == 1 or i == n or (i % 32 == 0):
                    self._set_asm_progress(pct, f"{i}/{n} bytes")

        self._set_asm_progress(100.0, f"{n}/{n} bytes (done)")
        self.log("[INFO] One-click complete. X-07 should execute the binary automatically (EXEC Z).")


if __name__ == "__main__":
    app = X07LoaderApp()
    app.mainloop()
