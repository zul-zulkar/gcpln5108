import tkinter as tk
from tkinter import ttk, filedialog, scrolledtext, messagebox
import asyncio
import threading
import sys
import os
import json
import socket
import ssl
import urllib.request
import urllib.error
import subprocess
import time
from datetime import datetime, timedelta

import scrape_fasih
import run_api

SETTINGS_FILE = "fasih_settings.json"

# ── Mode CDP (UI baru / API) ────────────────────────────────────────────────────
CDP_PORT = 9222
CDP_URL = f"http://127.0.0.1:{CDP_PORT}"
_CHROME_PATHS = [
    r"C:\Program Files\Google\Chrome\Application\chrome.exe",
    r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
]
_CDP_PROFILE = os.path.join(os.environ.get("LOCALAPPDATA", os.path.expanduser("~")),
                            "Temp", "claude", "fasih_cdp_profile")
_FASIH_APP_URL = "https://fasih-sm.bps.go.id/app/surveys"


def _find_chrome():
    for p in _CHROME_PATHS:
        if os.path.exists(p):
            return p
    return None


def _cdp_port_open():
    try:
        s = socket.create_connection(("127.0.0.1", CDP_PORT), timeout=3)
        s.close()
        return True
    except OSError:
        return False

# ── VPN helpers ────────────────────────────────────────────────────────────────
_FASIH_HOST = "fasih-sm.bps.go.id"
_FORTICLIENT_PATHS = [
    r"C:\Program Files\Fortinet\FortiClient\FortiSSLVPNclient.exe",
    r"C:\Program Files (x86)\Fortinet\FortiClient\FortiSSLVPNclient.exe",
    r"C:\Program Files\Fortinet\FortiClient\FortiClient.exe",
    r"C:\Program Files (x86)\Fortinet\FortiClient\FortiClient.exe",
]


def _is_vpn_connected():
    # Method 1: HTTPS — respons apapun (termasuk SSL/HTTP error) = server terjangkau
    ctx = ssl.SSLContext(ssl.PROTOCOL_TLS_CLIENT)
    ctx.check_hostname = False
    ctx.verify_mode = ssl.CERT_NONE
    try:
        with urllib.request.urlopen(f"https://{_FASIH_HOST}/", timeout=10, context=ctx):
            pass
        return True
    except urllib.error.HTTPError:
        return True  # HTTP error = server merespons = VPN terhubung
    except ssl.SSLError:
        return True  # SSL error = koneksi mencapai server = VPN terhubung
    except urllib.error.URLError as e:
        # URLError bisa membungkus SSL error — cek reason-nya
        if not isinstance(e.reason, OSError):
            return True
        print(f"[VPN-CHECK] URLError: {e.reason}")
    except Exception as e:
        print(f"[VPN-CHECK] HTTPS gagal: {type(e).__name__}: {e}")

    # Method 2: HTTP plain (port 80) — tanpa SSL, paling compatible di EXE
    try:
        with urllib.request.urlopen(f"http://{_FASIH_HOST}/", timeout=10):
            pass
        return True
    except urllib.error.HTTPError:
        return True
    except Exception:
        pass

    # Method 3: Raw TCP ke port 443
    try:
        s = socket.create_connection((_FASIH_HOST, 443), timeout=10)
        s.close()
        return True
    except OSError as e:
        print(f"[VPN-CHECK] TCP gagal: {e}")
        return False


def _find_forticlient():
    for p in _FORTICLIENT_PATHS:
        if os.path.exists(p):
            return p
    return None


def _get_forticlient_tunnels():
    tunnels = []
    try:
        import winreg
        reg_paths = [
            (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Fortinet\FortiClient\FA_VPN\tunnels"),
            (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\WOW6432Node\Fortinet\FortiClient\FA_VPN\tunnels"),
            (winreg.HKEY_CURRENT_USER,  r"SOFTWARE\Fortinet\FortiClient\FA_VPN\tunnels"),
        ]
        for hive, subkey in reg_paths:
            try:
                with winreg.OpenKey(hive, subkey) as k:
                    i = 0
                    while True:
                        try:
                            tunnels.append(winreg.EnumKey(k, i))
                            i += 1
                        except OSError:
                            break
            except OSError:
                pass
    except Exception:
        pass
    return list(dict.fromkeys(tunnels))


# Proses GUI FortiClient (bukan service) yang perlu dihentikan saat force-reconnect
_FORTICLIENT_GUI_PROCS = [
    "FortiClient.exe",
    "FortiSSLVPNclient.exe",
    "FortiTray.exe",
    "FCHelper64.exe",
    "FCAppDb.exe",
]
# Nama service Windows FortiClient
_FORTICLIENT_SERVICES = ["FortiClient Service", "FortiClientService", "FCLFW"]


def _kill_forticlient_gui():
    """Kill proses GUI/tray FortiClient agar tunnel drop, tanpa menyentuh service."""
    killed_any = False
    for proc in _FORTICLIENT_GUI_PROCS:
        try:
            r = subprocess.run(
                ["taskkill", "/F", "/IM", proc, "/T"],
                stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, timeout=6,
            )
            if r.returncode == 0:
                print(f"[VPN] Proses {proc} dihentikan.")
                killed_any = True
        except Exception:
            pass
    return killed_any


def _restart_forticlient_service():
    """Stop lalu start kembali service FortiClient."""
    for svc in _FORTICLIENT_SERVICES:
        try:
            subprocess.run(["net", "stop", svc], stdout=subprocess.DEVNULL,
                           stderr=subprocess.DEVNULL, timeout=15)
            time.sleep(3)
            r = subprocess.run(["net", "start", svc], stdout=subprocess.DEVNULL,
                               stderr=subprocess.DEVNULL, timeout=15)
            if r.returncode == 0:
                print(f"[VPN] Service '{svc}' di-restart ✓")
                return True
        except Exception:
            pass
    return False


def _wait_vpn_down(max_wait=20):
    """Tunggu hingga socket ke host benar-benar tidak bisa dijangkau (tunnel drop)."""
    for _ in range(max_wait):
        if not _is_vpn_connected():
            return True
        time.sleep(1)
    return False


def _rasdial_connect(tunnel, username, password):
    """Coba connect lewat rasdial (works jika tunnel VPN terdaftar sebagai RAS connection)."""
    try:
        r = subprocess.run(
            ["rasdial", tunnel, username, password],
            stdout=subprocess.PIPE, stderr=subprocess.DEVNULL,
            timeout=60, text=True,
        )
        return r.returncode == 0
    except Exception:
        return False


def _rasdial_disconnect(tunnel):
    """Disconnect tunnel via rasdial."""
    try:
        subprocess.run(
            ["rasdial", tunnel, "/disconnect"],
            stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, timeout=10,
        )
    except Exception:
        pass


def _launch_forticlient_connect(fclient, tunnel, username, password):
    arg_variants = [
        ["--vpnconnect", tunnel, "--username", username, "--password", password],
        ["-vpnconnect", tunnel, "-username", username, "-password", password],
        ["/vpnconnect", f"/VPN:{tunnel}", f"/user:{username}", f"/pwd:{password}"],
    ]
    for args in arg_variants:
        try:
            result = subprocess.run(
                [fclient] + args,
                stdout=subprocess.DEVNULL, stderr=subprocess.PIPE, timeout=5,
            )
            if result.returncode == 0:
                return True
        except subprocess.TimeoutExpired:
            return True
        except Exception:
            pass
    return False


def _vpn_reconnect_auto(fclient, vpn_host, username, password, tunnel_name):
    """
    Coba reconnect VPN secara otomatis tanpa interaksi manual.
    Urutan: rasdial → FortiClient CLI → restart service → FortiClient CLI ulang.
    """
    tunnels = _get_forticlient_tunnels()
    if not tunnels and tunnel_name:
        tunnels = [tunnel_name]

    # Strategi 1: rasdial disconnect + reconnect
    if tunnels:
        print("[VPN] Strategi 1: rasdial…")
        _rasdial_disconnect(tunnels[0])
        time.sleep(2)
        if _rasdial_connect(tunnels[0], username, password):
            print("[VPN] rasdial connect berhasil ✓")
            return True
        print("[VPN] rasdial gagal, lanjut strategi 2…")

    # Strategi 2: kill GUI + FortiClient CLI connect
    print("[VPN] Strategi 2: kill GUI FortiClient → CLI connect…")
    _kill_forticlient_gui()
    _wait_vpn_down(max_wait=15)
    time.sleep(2)

    if fclient:
        basename = os.path.basename(fclient)
        if "SSLVPNclient" in basename:
            try:
                subprocess.Popen(
                    [fclient, "connect", "-s", vpn_host, "-u", username, "-p", password, "--keepalive"],
                    stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
                )
                print("[VPN] Perintah connect (SSLVPNclient) dikirim.")
                return True  # cek koneksi di langkah 3
            except Exception as exc:
                print(f"[VPN] SSLVPNclient gagal: {exc}")
        else:
            for tunnel in (tunnels or []):
                if _launch_forticlient_connect(fclient, tunnel, username, password):
                    print(f"[VPN] CLI connect dikirim untuk tunnel '{tunnel}'.")
                    return True  # cek koneksi di langkah 3

    # Strategi 3: restart service FortiClient lalu CLI connect ulang
    print("[VPN] Strategi 3: restart service FortiClient…")
    if _restart_forticlient_service():
        time.sleep(5)
        if fclient:
            for tunnel in (tunnels or []):
                if _launch_forticlient_connect(fclient, tunnel, username, password):
                    print(f"[VPN] CLI connect (pasca restart service) dikirim untuk tunnel '{tunnel}'.")
                    return True

    print("[VPN] Semua strategi auto-reconnect sudah dicoba.")
    return False  # akan dicek lagi di loop tunggu


def ensure_vpn(vpn_host, username, password, tunnel_name="", force_reconnect=False):
    fclient = _find_forticlient()

    # ── Langkah 1: force disconnect ────────────────────────────────────────────
    if force_reconnect:
        print("[VPN] Force reconnect dimulai…")
        _vpn_reconnect_auto(fclient, vpn_host, username, password, tunnel_name)
    else:
        if _is_vpn_connected():
            print("[VPN] Sudah terhubung ✓")
            return True
        print(f"[VPN] Tidak terhubung ke {_FASIH_HOST}. Mencoba mengaktifkan VPN…")
        if not fclient:
            print("[VPN] FortiClient tidak ditemukan. Tidak bisa auto-connect.")
            return False
        tunnels = _get_forticlient_tunnels()
        if not tunnels and tunnel_name:
            tunnels = [tunnel_name]
        for tunnel in (tunnels or []):
            if _launch_forticlient_connect(fclient, tunnel, username, password):
                print(f"[VPN] Perintah connect dikirim untuk tunnel '{tunnel}'.")
                break

    # ── Langkah 2: tunggu koneksi aktif ───────────────────────────────────────
    max_checks = 18
    for attempt in range(max_checks):
        time.sleep(10)
        elapsed = (attempt + 1) * 10
        print(f"[VPN] Menunggu koneksi VPN… ({elapsed}s / {max_checks * 10}s)")
        if _is_vpn_connected():
            print("[VPN] VPN berhasil terhubung ✓ — melanjutkan scraping…")
            return True

    print(f"[VPN] Gagal terhubung ke VPN setelah {max_checks * 10} detik. Scraping dibatalkan.")
    return False


# ── TextRedirector ─────────────────────────────────────────────────────────────
class TextRedirector:
    """Redirect stdout ke widget log."""
    def __init__(self, widget):
        self.widget = widget

    def write(self, text):
        try:
            self.widget.after(0, self._insert, text)
        except Exception:
            pass

    def _insert(self, text):
        self.widget.configure(state="normal")
        self.widget.insert(tk.END, text)
        self.widget.see(tk.END)
        self.widget.configure(state="disabled")

    def flush(self):
        pass


# ── GUI ────────────────────────────────────────────────────────────────────────
class FasihScraperGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("FASIH Scraper — BPS")
        self.root.geometry("750x680")
        self.root.resizable(True, True)

        self._stop_event  = threading.Event()
        self._thread      = None
        self._pass_entry  = None

        # Auto-run state
        self._sched_lock  = threading.Lock()
        self._sched_timer = None
        self._next_run    = None   # datetime | None

        self._build_ui()
        self._load_settings()

    def _build_ui(self):
        main = ttk.Frame(self.root, padding=14)
        main.pack(fill=tk.BOTH, expand=True)

        # ── Input file ──────────────────────────────────────────────────────
        ttk.Label(main, text="File Petugas:").grid(row=0, column=0, sticky="w", pady=5)
        self.input_var = tk.StringVar()
        ttk.Entry(main, textvariable=self.input_var, width=52).grid(
            row=0, column=1, sticky="ew", padx=6
        )
        ttk.Button(main, text="Browse…", command=self._browse).grid(row=0, column=2, padx=2)

        # ── Username ─────────────────────────────────────────────────────────
        ttk.Label(main, text="Username SSO BPS:").grid(row=1, column=0, sticky="w", pady=5)
        self.username_var = tk.StringVar()
        ttk.Entry(main, textvariable=self.username_var, width=52).grid(
            row=1, column=1, sticky="ew", padx=6
        )

        # ── Password ─────────────────────────────────────────────────────────
        ttk.Label(main, text="Password SSO BPS:").grid(row=2, column=0, sticky="w", pady=5)
        self.password_var = tk.StringVar()
        self._pass_entry = ttk.Entry(
            main, textvariable=self.password_var, show="●", width=40
        )
        self._pass_entry.grid(row=2, column=1, sticky="ew", padx=6)

        self.show_pass_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(
            main,
            text="Tampilkan",
            variable=self.show_pass_var,
            command=self._toggle_pass,
        ).grid(row=2, column=2, padx=2)

        # (UPI/UP3 dihapus — scoping wilayah kini lewat surveyPeriodId di API, bukan input manual)

        # ── Apps Script URL ──────────────────────────────────────────────────
        ttk.Label(main, text="Apps Script URL:").grid(row=4, column=0, sticky="w", pady=5)
        self.sheets_var = tk.StringVar()
        ttk.Entry(main, textvariable=self.sheets_var, width=52).grid(
            row=4, column=1, sticky="ew", padx=6, columnspan=2
        )

        # ── Separator ────────────────────────────────────────────────────────
        ttk.Separator(main, orient="horizontal").grid(
            row=5, column=0, columnspan=3, sticky="ew", pady=(10, 6), padx=2
        )

        # ── Auto VPN ─────────────────────────────────────────────────────────
        vpn_frame = ttk.Frame(main)
        vpn_frame.grid(row=6, column=0, columnspan=3, sticky="ew", padx=6, pady=3)
        self.vpn_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(
            vpn_frame, text="Auto Reconnect VPN", variable=self.vpn_var,
            command=self._on_vpn_toggle,
        ).pack(side=tk.LEFT)
        ttk.Label(vpn_frame, text="Host:").pack(side=tk.LEFT, padx=(14, 4))
        self.vpn_host_var = tk.StringVar()
        ttk.Entry(vpn_frame, textvariable=self.vpn_host_var, width=22).pack(side=tk.LEFT)
        ttk.Label(vpn_frame, text="Tunnel:").pack(side=tk.LEFT, padx=(14, 4))
        self.vpn_tunnel_var = tk.StringVar()
        ttk.Entry(vpn_frame, textvariable=self.vpn_tunnel_var, width=22).pack(side=tk.LEFT)

        # ── Auto-run ─────────────────────────────────────────────────────────
        auto_frame = ttk.Frame(main)
        auto_frame.grid(row=7, column=0, columnspan=3, sticky="ew", padx=6, pady=3)
        self.auto_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(
            auto_frame, text="Auto-run setiap", variable=self.auto_var,
            command=self._on_auto_toggle,
        ).pack(side=tk.LEFT)
        self.auto_hours_var = tk.StringVar(value="2")
        ttk.Spinbox(
            auto_frame, textvariable=self.auto_hours_var,
            from_=0, to=23, width=4, command=self._on_auto_toggle,
        ).pack(side=tk.LEFT, padx=(8, 2))
        ttk.Label(auto_frame, text="jam").pack(side=tk.LEFT)
        self.auto_mins_var = tk.StringVar(value="0")
        ttk.Spinbox(
            auto_frame, textvariable=self.auto_mins_var,
            from_=0, to=59, width=4, command=self._on_auto_toggle,
        ).pack(side=tk.LEFT, padx=(8, 2))
        ttk.Label(auto_frame, text="menit").pack(side=tk.LEFT)
        self.next_run_var = tk.StringVar(value="")
        ttk.Label(
            auto_frame, textvariable=self.next_run_var, foreground="#1a7a3c",
        ).pack(side=tk.LEFT, padx=(14, 0))

        # ── Tombol Run / Stop ────────────────────────────────────────────────
        btn_frame = ttk.Frame(main)
        btn_frame.grid(row=8, column=0, columnspan=3, pady=10, sticky="w")

        self.chrome_btn = ttk.Button(
            btn_frame, text="🌐  Buka Chrome & Login", command=self._launch_chrome
        )
        self.chrome_btn.pack(side=tk.LEFT, padx=(0, 8))

        self.run_btn = ttk.Button(btn_frame, text="▶  Run", command=self._run, width=12)
        self.run_btn.pack(side=tk.LEFT, padx=(0, 8))

        self.stop_btn = ttk.Button(
            btn_frame, text="■  Stop", command=self._stop, width=12, state="disabled"
        )
        self.stop_btn.pack(side=tk.LEFT, padx=(0, 8))

        self.open_btn = ttk.Button(
            btn_frame, text="📂  Buka Folder Hasil", command=self._open_output, state="disabled"
        )
        self.open_btn.pack(side=tk.LEFT, padx=(0, 8))

        self.status_var = tk.StringVar(value="")
        ttk.Label(btn_frame, textvariable=self.status_var, foreground="#555").pack(
            side=tk.LEFT, padx=6
        )

        # ── Log ───────────────────────────────────────────────────────────────
        ttk.Label(main, text="Log:").grid(row=9, column=0, sticky="w", pady=(8, 2))
        self.log = scrolledtext.ScrolledText(
            main, height=18, state="disabled", font=("Consolas", 9), wrap=tk.WORD
        )
        self.log.grid(row=10, column=0, columnspan=3, sticky="nsew")

        main.columnconfigure(1, weight=1)
        main.rowconfigure(10, weight=1)

        self.root.protocol("WM_DELETE_WINDOW", self._on_close)

    # ── Settings ──────────────────────────────────────────────────────────────

    def _load_settings(self):
        try:
            with open(SETTINGS_FILE, encoding="utf-8") as f:
                s = json.load(f)
            if s.get("input_file"):  self.input_var.set(s["input_file"])
            if s.get("username"):    self.username_var.set(s["username"])
            if s.get("sheets_url"):  self.sheets_var.set(s["sheets_url"])
            if s.get("vpn_name"):    self.vpn_host_var.set(s["vpn_name"])
            if s.get("vpn_tunnel"):  self.vpn_tunnel_var.set(s["vpn_tunnel"])
            self.vpn_var.set(bool(s.get("vpn_reconnect", False)))
            self.auto_var.set(bool(s.get("autorun", False)))
            interval = int(s.get("autorun_interval", 120))
            self.auto_hours_var.set(str(interval // 60))
            self.auto_mins_var.set(str(interval % 60))
            if self.auto_var.get():
                self._schedule_next()
        except (FileNotFoundError, json.JSONDecodeError, KeyError):
            pass

    def _save_settings(self):
        interval = self._get_interval_mins()
        s = {
            "input_file":       self.input_var.get().strip(),
            "username":         self.username_var.get().strip(),
            "sheets_url":       self.sheets_var.get().strip(),
            "vpn_reconnect":    self.vpn_var.get(),
            "vpn_name":         self.vpn_host_var.get().strip(),
            "vpn_tunnel":       self.vpn_tunnel_var.get().strip(),
            "autorun":          self.auto_var.get(),
            "autorun_interval": str(interval),
        }
        try:
            with open(SETTINGS_FILE, "w", encoding="utf-8") as f:
                json.dump(s, f, indent=2, ensure_ascii=False)
        except Exception:
            pass

    # ── Helpers ───────────────────────────────────────────────────────────────

    def _get_interval_mins(self):
        try:
            hours = int(self.auto_hours_var.get() or 0)
            mins  = int(self.auto_mins_var.get() or 0)
            total = hours * 60 + mins
            return total if total >= 1 else 120
        except ValueError:
            return 120

    def _browse(self):
        path = filedialog.askopenfilename(
            title="Pilih file daftar petugas",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("Semua file", "*.*")],
        )
        if path:
            self.input_var.set(path)

    def _open_output(self):
        out = os.path.abspath(scrape_fasih.OUTPUT_DIR)
        os.makedirs(out, exist_ok=True)
        os.startfile(out)

    def _launch_chrome(self):
        """Buka Chrome dengan port debug agar scraper bisa nempel (mode CDP).
        User login FASIH manual di jendela ini (lolos WAF)."""
        chrome = _find_chrome()
        if not chrome:
            messagebox.showerror("Chrome", "chrome.exe tidak ditemukan di lokasi standar.")
            return
        if _cdp_port_open():
            messagebox.showinfo(
                "Chrome", "Chrome debug sudah berjalan (port 9222).\n"
                "Pastikan sudah login FASIH, lalu klik Run.")
            return
        try:
            os.makedirs(_CDP_PROFILE, exist_ok=True)
            subprocess.Popen([
                chrome,
                f"--remote-debugging-port={CDP_PORT}",
                f"--user-data-dir={_CDP_PROFILE}",
                _FASIH_APP_URL,
            ])
            messagebox.showinfo(
                "Chrome dibuka",
                "Jendela Chrome dibuka.\n\n"
                "1) LOGIN FASIH manual di jendela itu (sebagai manusia).\n"
                "2) Setelah daftar survei muncul, klik tombol Run.")
        except Exception as e:
            messagebox.showerror("Chrome", f"Gagal membuka Chrome: {e}")

    def _toggle_pass(self):
        self._pass_entry.config(show="" if self.show_pass_var.get() else "●")

    def _log_clear(self):
        self.log.configure(state="normal")
        self.log.delete("1.0", tk.END)
        self.log.configure(state="disabled")

    def _on_vpn_toggle(self):
        self._save_settings()

    def _on_auto_toggle(self):
        self._save_settings()
        if self.auto_var.get():
            self._schedule_next()
        else:
            self._cancel_schedule()

    # ── Auto-run scheduler ────────────────────────────────────────────────────

    def _schedule_next(self):
        with self._sched_lock:
            if self._sched_timer:
                self._sched_timer.cancel()
                self._sched_timer = None
            if not self.auto_var.get():
                self._next_run = None
                self.root.after(0, lambda: self.next_run_var.set(""))
                return
            delay = self._get_interval_mins() * 60
            self._next_run = datetime.now() + timedelta(seconds=delay)
            t = threading.Timer(delay, self._auto_run)
            t.daemon = True
            t.start()
            self._sched_timer = t
            hhmm = self._next_run.strftime("%H:%M")
            self.root.after(0, lambda h=hhmm: self.next_run_var.set(f"⏰ Berikutnya: {h}"))

    def _cancel_schedule(self):
        with self._sched_lock:
            if self._sched_timer:
                self._sched_timer.cancel()
                self._sched_timer = None
            self._next_run = None
        self.next_run_var.set("")

    def _auto_run(self):
        # Jika scraper masih berjalan, tunda 5 menit
        if self._thread and self._thread.is_alive():
            with self._sched_lock:
                self._sched_timer = None
                self._next_run = datetime.now() + timedelta(seconds=300)
                t = threading.Timer(300, self._auto_run)
                t.daemon = True
                t.start()
                self._sched_timer = t
                hhmm = self._next_run.strftime("%H:%M")
                self.root.after(0, lambda h=hhmm: self.next_run_var.set(f"⏰ Berikutnya: {h}"))
            return
        self.root.after(0, self._run_auto)

    def _run_auto(self):
        print(f"\n[AUTO] ▶ Jadwal otomatis dimulai ({datetime.now().strftime('%H:%M:%S')})…\n")
        self._run()

    # ── Run / Stop ────────────────────────────────────────────────────────────

    def _run(self):
        input_file  = self.input_var.get().strip()
        username    = self.username_var.get().strip()
        password    = self.password_var.get().strip()
        sheets_url  = self.sheets_var.get().strip()
        vpn_enabled = self.vpn_var.get()
        vpn_host    = self.vpn_host_var.get().strip()
        vpn_tunnel  = self.vpn_tunnel_var.get().strip()

        if not input_file:
            messagebox.showwarning("Input", "Pilih file daftar petugas terlebih dahulu.")
            return
        if not os.path.exists(input_file):
            messagebox.showerror("File tidak ditemukan", f"{input_file}")
            return
        # Mode CDP: login manual via Chrome debug (bukan auto-login).
        if not _cdp_port_open():
            messagebox.showwarning(
                "Belum login",
                "Chrome debug (port 9222) belum aktif.\n\n"
                "Klik 'Buka Chrome & Login' dulu, login FASIH, baru klik Run.")
            return
        # Username/password hanya diperlukan bila Auto Reconnect VPN aktif.
        if self.vpn_var.get() and (not username or not password):
            messagebox.showwarning("Input", "Username & password SSO diperlukan untuk Auto Reconnect VPN.")
            return

        self._save_settings()
        self._log_clear()
        sys.stdout = TextRedirector(self.log)

        self._stop_event = threading.Event()
        self.run_btn.configure(state="disabled")
        self.stop_btn.configure(state="normal")
        self.open_btn.configure(state="disabled")
        self.status_var.set("Berjalan…")

        self._thread = threading.Thread(
            target=self._run_in_thread,
            args=(input_file, username, password,
                  sheets_url, vpn_enabled, vpn_host, vpn_tunnel, self._stop_event),
            daemon=True,
        )
        self._thread.start()
        self._poll()

    def _run_in_thread(self, input_file, username, password,
                       sheets_url, vpn_enabled, vpn_host, vpn_tunnel, stop_event):
        loop = asyncio.new_event_loop()
        asyncio.set_event_loop(loop)
        try:
            if vpn_enabled and vpn_host:
                if not ensure_vpn(vpn_host, username, password, vpn_tunnel, force_reconnect=True):
                    print("[VPN] Scraping dibatalkan karena VPN tidak terhubung.")
                    return
            if not _cdp_port_open():
                print("[CDP] Chrome debug (port 9222) belum aktif.")
                print("      Klik 'Buka Chrome & Login', login FASIH, lalu Run lagi.")
                return
            # Mode UI baru via API (CDP nempel ke Chrome login-manual)
            loop.run_until_complete(
                run_api.run_pipeline(
                    input_file=input_file,
                    sheets_url=sheets_url or None,
                    send_sheets=True,
                    stop_event=stop_event,
                    cdp_url=CDP_URL,
                )
            )
        except Exception as exc:
            print(f"\n[ERROR] {exc}")
        finally:
            loop.close()

    def _poll(self):
        if self._thread and self._thread.is_alive():
            self.root.after(500, self._poll)
        else:
            self._on_done()

    def _on_done(self):
        sys.stdout = sys.__stdout__
        self.run_btn.configure(state="normal")
        self.stop_btn.configure(state="disabled")
        stopped = self._stop_event.is_set()
        self.status_var.set("Dihentikan." if stopped else "Selesai ✓")
        if not stopped:
            self.open_btn.configure(state="normal")
        # Jadwalkan ulang jika auto aktif dan tidak dihentikan manual
        if self.auto_var.get() and not stopped:
            self._schedule_next()

    def _stop(self):
        self._stop_event.set()
        self.stop_btn.configure(state="disabled")
        self.status_var.set("Menghentikan…")

    def _on_close(self):
        self._stop_event.set()
        self._cancel_schedule()
        self.root.destroy()


if __name__ == "__main__":
    root = tk.Tk()
    FasihScraperGUI(root)
    root.mainloop()
