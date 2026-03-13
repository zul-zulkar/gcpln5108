"""
web_fasih.py — FASIH Scraper Web UI
Jalankan: python web_fasih.py
Buka browser: http://localhost:5000
"""
import sys
import os
import asyncio
import threading
import queue
import json
import socket
import subprocess
import time
from datetime import datetime, timedelta

from flask import Flask, Response, request, jsonify, render_template_string, send_file
from werkzeug.utils import secure_filename
import scrape_fasih

app = Flask(__name__)
UPLOAD_DIR = "uploads"
os.makedirs(UPLOAD_DIR, exist_ok=True)

_lock = threading.Lock()
_state = {"running": False, "stop_event": threading.Event(), "uploaded_file": ""}
_log_queue: queue.Queue = queue.Queue()

# ── Auto-scheduler ─────────────────────────────────────────────────────────────
_sched = {
    "enabled":       False,
    "interval_mins": 120,     # total menit (default 2 jam)
    "next_run":      None,    # datetime | None
    "timer":         None,    # threading.Timer | None
    "params":        {},      # {input_file, username, password}
    "vpn_enabled":   False,
    "vpn_host":      "",
}


def _schedule_next():
    """Pasang timer untuk run berikutnya. Harus dipanggil tanpa _lock."""
    with _lock:
        if not _sched["enabled"] or not _sched["params"]:
            return
        if _sched["timer"]:
            _sched["timer"].cancel()
        delay = _sched["interval_mins"] * 60
        _sched["next_run"] = datetime.now() + timedelta(seconds=delay)
        t = threading.Timer(delay, _auto_run)
        t.daemon = True
        t.start()
        _sched["timer"] = t


def _auto_run():
    """Dipanggil oleh timer — jalankan scraper dengan params terakhir."""
    with _lock:
        if _state["running"]:
            # Sedang berjalan, tunda 5 menit lalu coba lagi
            t = threading.Timer(300, _auto_run)
            t.daemon = True
            t.start()
            _sched["timer"] = t
            _sched["next_run"] = datetime.now() + timedelta(seconds=300)
            return
        params      = _sched["params"].copy()
        vpn_enabled = _sched["vpn_enabled"]
        vpn_host    = _sched["vpn_host"]
        if not params:
            return
        # Bersihkan pesan lama di queue agar stream tidak baca DONE dari run sebelumnya
        while not _log_queue.empty():
            try: _log_queue.get_nowait()
            except Exception: break
        stop_event = threading.Event()
        _state["running"]    = True
        _state["stop_event"] = stop_event

    _log_queue.put(f"\n[AUTO] ▶ Jadwal otomatis dimulai ({datetime.now().strftime('%H:%M:%S')})…\n")

    def _run():
        old_stdout = sys.stdout
        sys.stdout = _QueueWriter()
        # ── VPN ───────────────────────────────────────────────────────────────
        if vpn_enabled and vpn_host:
            if not ensure_vpn(vpn_host, params["username"], params["password"]):
                print("[VPN] Scraping dibatalkan karena VPN tidak terhubung.")
                sys.stdout = old_stdout
                _log_queue.put("\x00DONE\x00")
                with _lock:
                    _state["running"] = False
                _schedule_next()
                return
        # ── Scrape ────────────────────────────────────────────────────────────
        loop = asyncio.new_event_loop()
        asyncio.set_event_loop(loop)
        try:
            loop.run_until_complete(
                scrape_fasih.main_with_stop(
                    stop_event,
                    params["input_file"],
                    params["username"],
                    params["password"],
                    headless=False,
                )
            )
        except Exception as exc:
            print(f"\n[ERROR] {exc}")
        finally:
            loop.close()
            sys.stdout = old_stdout
            _log_queue.put("\x00DONE\x00")
            with _lock:
                _state["running"] = False
            _schedule_next()  # di dalam finally agar selalu terpanggil

    threading.Thread(target=_run, daemon=True).start()


# ── Redirect stdout → queue ────────────────────────────────────────────────────
class _QueueWriter:
    def write(self, text):
        if text:
            _log_queue.put(text)

    def flush(self):
        pass


# ── VPN helpers ────────────────────────────────────────────────────────────────
_FASIH_HOST = "fasih-sm.bps.go.id"
_FORTICLIENT_PATHS = [
    r"C:\Program Files\Fortinet\FortiClient\FortiSSLVPNclient.exe",
    r"C:\Program Files (x86)\Fortinet\FortiClient\FortiSSLVPNclient.exe",
    r"C:\Program Files\Fortinet\FortiClient\FortiClient.exe",
    r"C:\Program Files (x86)\Fortinet\FortiClient\FortiClient.exe",
]


def _is_vpn_connected():
    """Cek apakah host FASIH bisa dijangkau (VPN aktif)."""
    try:
        s = socket.create_connection((_FASIH_HOST, 443), timeout=5)
        s.close()
        return True
    except OSError:
        return False


def _find_forticlient():
    for p in _FORTICLIENT_PATHS:
        if os.path.exists(p):
            return p
    return None


def _get_forticlient_tunnels():
    """
    Baca daftar nama tunnel VPN yang sudah dikonfigurasi di FortiClient
    dari registry Windows.
    """
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
    return list(dict.fromkeys(tunnels))  # deduplicate, preserve order


def _launch_forticlient_connect(fclient, tunnel, username, password):
    """
    Coba connect ke tunnel tertentu menggunakan FortiClient.exe.
    Beberapa versi FortiClient mendukung argumen berbeda; dicoba secara berurutan.
    """
    # Daftar sintaks yang diketahui untuk berbagai versi FortiClient
    arg_variants = [
        ["--vpnconnect", tunnel, "--username", username, "--password", password],
        ["-vpnconnect", tunnel, "-username", username, "-password", password],
        [f"/vpnconnect", f"/VPN:{tunnel}", f"/user:{username}", f"/pwd:{password}"],
    ]
    for args in arg_variants:
        try:
            result = subprocess.run(
                [fclient] + args,
                stdout=subprocess.DEVNULL,
                stderr=subprocess.PIPE,
                timeout=5,
            )
            if result.returncode == 0:
                return True
            # Beberapa versi return non-zero tapi tetap memproses perintah;
            # lanjut tunggu koneksi daripada langsung berhenti.
        except subprocess.TimeoutExpired:
            # Proses masih berjalan → dianggap berhasil dikirim
            return True
        except Exception:
            pass
    return False


def ensure_vpn(vpn_host, username, password):
    """
    Cek koneksi VPN. Jika belum terhubung:
    1. FortiSSLVPNclient.exe → connect via CLI dengan host+credentials.
    2. FortiClient.exe (baru) → baca tunnel dari registry, lalu auto-connect.
    3. Fallback → buka GUI FortiClient, tunggu user connect manual.
    Menunggu maks 3 menit sampai koneksi aktif, lalu lanjutkan scraping.
    """
    if _is_vpn_connected():
        print("[VPN] Sudah terhubung ✓")
        return True

    print(f"[VPN] Tidak terhubung ke {_FASIH_HOST}. Mencoba mengaktifkan VPN…")
    fclient = _find_forticlient()

    if not fclient:
        print("[VPN] FortiClient tidak ditemukan di path standar. Tidak bisa auto-connect.")
        return False

    basename = os.path.basename(fclient)

    if "SSLVPNclient" in basename:
        # ── FortiSSLVPNclient: connect langsung dengan host + credentials ────
        print(f"[VPN] Menggunakan FortiSSLVPNclient CLI: {fclient}")
        try:
            subprocess.Popen(
                [fclient, "connect", "-s", vpn_host, "-u", username, "-p", password, "--keepalive"],
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL,
            )
            print("[VPN] Perintah connect dikirim.")
        except Exception as exc:
            print(f"[VPN] Gagal menjalankan FortiSSLVPNclient: {exc}")

    else:
        # ── FortiClient.exe (versi baru): baca tunnel dari registry ─────────
        tunnels = _get_forticlient_tunnels()
        print(f"[VPN] Menggunakan FortiClient: {fclient}")
        print(f"[VPN] Tunnel ditemukan di registry: {tunnels if tunnels else '(tidak ada)'}")

        connected_cmd = False
        for tunnel in tunnels:
            print(f"[VPN] Mencoba auto-connect tunnel '{tunnel}'…")
            if _launch_forticlient_connect(fclient, tunnel, username, password):
                print(f"[VPN] Perintah connect dikirim untuk tunnel '{tunnel}'.")
                connected_cmd = True
                break

        if not connected_cmd:
            # Fallback: buka GUI, minta user connect manual
            print("[VPN] Tidak bisa auto-connect via CLI. Membuka FortiClient GUI…")
            print("[VPN] Silakan connect VPN secara manual di jendela yang terbuka.")
            try:
                subprocess.Popen([fclient], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
            except Exception as exc:
                print(f"[VPN] Gagal membuka FortiClient: {exc}")

    # Tunggu maks 3 menit (18 × 10 detik)
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


# ── HTML ───────────────────────────────────────────────────────────────────────
HTML = r"""<!DOCTYPE html>
<html lang="id">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>FASIH Scraper</title>
<style>
  *, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }
  body { font-family: 'Segoe UI', sans-serif; background: #f0f2f5; color: #222; }
  .card {
    max-width: 860px; margin: 32px auto; background: #fff;
    border-radius: 10px; box-shadow: 0 2px 12px rgba(0,0,0,.1); overflow: hidden;
  }
  .card-header {
    background: #1a3a5c; color: #fff; padding: 18px 24px;
    font-size: 1.15rem; font-weight: 600; letter-spacing: .3px;
  }
  .card-header span { font-size: .85rem; font-weight: 400; opacity: .7; margin-left: 10px; }
  .card-body { padding: 24px; }
  .form-row { display: flex; align-items: center; margin-bottom: 14px; gap: 10px; }
  .form-row label { width: 120px; font-size: .88rem; color: #555; flex-shrink: 0; }
  .form-row input[type=text],
  .form-row input[type=password] {
    flex: 1; padding: 7px 10px; border: 1px solid #ccc; border-radius: 6px;
    font-size: .93rem; outline: none; transition: border .2s;
  }
  .form-row input:focus { border-color: #1a3a5c; }
  .show-pass { font-size: .82rem; color: #555; cursor: pointer; white-space: nowrap; }
  .show-pass input { margin-right: 4px; cursor: pointer; }
  .btn-row { display: flex; gap: 10px; margin: 20px 0 16px; align-items: center; }
  button {
    padding: 9px 22px; border: none; border-radius: 6px; font-size: .93rem;
    cursor: pointer; font-weight: 600; transition: opacity .15s, transform .1s;
  }
  button:active { transform: scale(.97); }
  #btnRun  { background: #1a7a3c; color: #fff; }
  #btnStop { background: #c0392b; color: #fff; }
  button:disabled { opacity: .45; cursor: not-allowed; transform: none; }
  #statusText { font-size: .85rem; color: #555; }
  .dot { display: inline-block; width: 8px; height: 8px; border-radius: 50%;
         background: #aaa; margin-right: 6px; vertical-align: middle; }
  .dot.running { background: #1a7a3c; animation: pulse 1s infinite; }
  .dot.stopped { background: #c0392b; }
  .dot.done    { background: #1a7a3c; }
  @keyframes pulse { 0%,100%{opacity:1} 50%{opacity:.4} }
  #log {
    background: #0d1117; color: #c9d1d9; font-family: 'Consolas','Courier New',monospace;
    font-size: .82rem; padding: 14px; border-radius: 6px; height: 360px;
    overflow-y: auto; white-space: pre-wrap; word-break: break-all;
  }
  .log-label { font-size: .8rem; color: #888; margin-bottom: 6px; }
  .clear-btn {
    background: none; border: 1px solid #ccc; color: #666; padding: 4px 10px;
    font-size: .78rem; border-radius: 4px; margin-left: auto; font-weight: 400;
  }
  .clear-btn:hover { background: #f5f5f5; }
  #dropZone {
    flex: 1; border: 2px dashed #b0bec5; border-radius: 8px; padding: 18px 14px;
    text-align: center; cursor: pointer; transition: border-color .2s, background .2s;
    background: #fafafa; user-select: none;
  }
  #dropZone:hover, #dropZone.drag-over { border-color: #1a3a5c; background: #eef3f8; }
  #dropZone.ready { border-color: #1a7a3c; background: #f0fff5; }
  #dropIcon { font-size: 1.6rem; margin-bottom: 4px; }
  #dropText { font-size: .9rem; color: #444; }
  #dropSub  { font-size: .78rem; color: #999; margin-top: 3px; }
  #uploadProgress {
    width: 100%; height: 4px; background: #e0e0e0; border-radius: 2px;
    margin-top: 8px; display: none; overflow: hidden;
  }
  #uploadBar { height: 100%; width: 0; background: #1a7a3c; transition: width .3s; }
  .auto-row { display:flex; align-items:center; gap:8px; margin-bottom:12px; flex-wrap:wrap; }
  .auto-row label.lbl { width:120px; font-size:.88rem; color:#555; flex-shrink:0; }
  .auto-row input[type=checkbox] { cursor:pointer; width:15px; height:15px; }
  .auto-row input[type=number] {
    width:52px; padding:5px 6px; border:1px solid #ccc; border-radius:6px;
    font-size:.88rem; text-align:center; outline:none;
  }
  .auto-row input[type=number]:focus { border-color:#1a3a5c; }
  .auto-row input[type=text].vpn-input {
    flex:1; min-width:180px; padding:5px 8px; border:1px solid #ccc; border-radius:6px;
    font-size:.85rem; outline:none;
  }
  .auto-row input[type=text].vpn-input:focus { border-color:#1a3a5c; }
  #nextRunText { font-size:.82rem; color:#1a7a3c; font-weight:500; }
  #vpnStatus { font-size:.8rem; color:#888; }
  .dl-btn {
    display:inline-flex; align-items:center; gap:6px;
    background:#1a3a5c; color:#fff; padding:7px 14px;
    border-radius:6px; font-size:.83rem; font-weight:500;
    text-decoration:none; transition:opacity .15s;
  }
  .dl-btn:hover { opacity:.85; }
</style>
</head>
<body>
<div class="card">
  <div class="card-header">FASIH Scraper <span>BPS — Survey Collection</span></div>
  <div class="card-body">

    <div class="form-row" style="align-items:flex-start;">
      <label style="padding-top:6px;">File Petugas</label>
      <div id="dropZone" onclick="document.getElementById('fileInput').click()" ondragover="onDragOver(event)" ondragleave="onDragLeave(event)" ondrop="onDrop(event)">
        <input type="file" id="fileInput" accept=".xlsx,.xls" style="display:none" onchange="uploadFile(this.files[0])">
        <div id="dropIcon">📂</div>
        <div id="dropText">Klik atau seret file <b>.xlsx</b> ke sini</div>
        <div id="dropSub">Format: daftar_petugas.xlsx</div>
        <div id="uploadProgress"><div id="uploadBar"></div></div>
      </div>
    </div>

    <div class="form-row">
      <label>Username</label>
      <input type="text" id="username" placeholder="Username SSO BPS" autocomplete="username">
    </div>

    <div class="form-row">
      <label>Password</label>
      <input type="password" id="password" placeholder="Password" autocomplete="current-password">
      <label class="show-pass">
        <input type="checkbox" id="showPass" onchange="togglePass()"> Tampilkan
      </label>
    </div>

    <div class="auto-row">
      <label class="lbl">Auto-run</label>
      <input type="checkbox" id="chkAuto" onchange="toggleAuto()">
      <span style="font-size:.88rem;color:#444;">Setiap</span>
      <input type="number" id="autoHours" value="2" min="0" max="23" onchange="toggleAuto()">
      <span style="font-size:.88rem;color:#444;">jam</span>
      <input type="number" id="autoMins" value="0" min="0" max="59" onchange="toggleAuto()">
      <span style="font-size:.88rem;color:#444;">menit</span>
      <span id="nextRunText"></span>
    </div>

    <div class="auto-row">
      <label class="lbl">Auto VPN</label>
      <input type="checkbox" id="chkVpn" onchange="toggleAuto()">
      <input type="text" id="vpnHost" class="vpn-input" placeholder="Host VPN (mis. vpn.bps.go.id)" oninput="saveVpnHost()">
      <span id="vpnStatus"></span>
    </div>

    <div class="btn-row">
      <button id="btnRun" onclick="runScraper()">&#9654; Run</button>
      <button id="btnStop" onclick="stopScraper()" disabled>&#9632; Stop</button>
      <span id="statusText"><span class="dot" id="dot"></span>Siap</span>
    </div>

    <div style="display:flex;align-items:center;margin-bottom:6px;">
      <div class="log-label">Log Output</div>
      <button class="clear-btn" onclick="clearLog()">Bersihkan</button>
    </div>
    <div id="log"></div>

    <div id="downloadSection" style="display:none; margin-top:16px;">
      <div class="log-label" style="margin-bottom:8px;">📥 Unduh Hasil</div>
      <div id="downloadList" style="display:flex; flex-wrap:wrap; gap:8px;"></div>
    </div>

  </div>
</div>

<script>
let evtSource = null;
let uploadedPath = '';

// ── Drag & Drop ──────────────────────────────────────────────────────────────
function onDragOver(e) {
  e.preventDefault();
  document.getElementById('dropZone').classList.add('drag-over');
}
function onDragLeave(e) {
  document.getElementById('dropZone').classList.remove('drag-over');
}
function onDrop(e) {
  e.preventDefault();
  document.getElementById('dropZone').classList.remove('drag-over');
  const file = e.dataTransfer.files[0];
  if (file) uploadFile(file);
}

// ── Upload ────────────────────────────────────────────────────────────────────
function uploadFile(file) {
  if (!file) return;
  if (!/\.(xlsx|xls)$/i.test(file.name)) {
    alert('Hanya file .xlsx atau .xls yang diizinkan.');
    return;
  }

  const zone = document.getElementById('dropZone');
  const prog = document.getElementById('uploadProgress');
  const bar  = document.getElementById('uploadBar');
  zone.classList.remove('ready', 'drag-over');
  document.getElementById('dropIcon').textContent = '⏳';
  document.getElementById('dropText').textContent = 'Mengunggah ' + file.name + '…';
  document.getElementById('dropSub').textContent  = '';
  prog.style.display = 'block';
  bar.style.width = '30%';

  const fd = new FormData();
  fd.append('file', file);

  fetch('/upload', { method: 'POST', body: fd })
    .then(r => r.json())
    .then(data => {
      bar.style.width = '100%';
      setTimeout(() => { prog.style.display = 'none'; bar.style.width = '0'; }, 600);
      if (data.error) {
        document.getElementById('dropIcon').textContent = '❌';
        document.getElementById('dropText').textContent = data.error;
        return;
      }
      uploadedPath = data.path;
      zone.classList.add('ready');
      document.getElementById('dropIcon').textContent = '✅';
      document.getElementById('dropText').textContent = file.name;
      document.getElementById('dropSub').textContent  = (data.rows || '?') + ' petugas ditemukan';
    })
    .catch(e => {
      document.getElementById('dropIcon').textContent = '❌';
      document.getElementById('dropText').textContent = 'Upload gagal: ' + e;
    });
}

// ── Misc ──────────────────────────────────────────────────────────────────────
function togglePass() {
  const p = document.getElementById('password');
  p.type = document.getElementById('showPass').checked ? 'text' : 'password';
}
function setStatus(text, state) {
  document.getElementById('statusText').innerHTML =
    `<span class="dot ${state}"></span>${text}`;
}
function appendLog(text) {
  const log = document.getElementById('log');
  log.textContent += text;
  log.scrollTop = log.scrollHeight;
}
function clearLog() {
  document.getElementById('log').textContent = '';
}

// ── Stream ────────────────────────────────────────────────────────────────────
function startStream() {
  if (evtSource) { evtSource.close(); evtSource = null; }
  evtSource = new EventSource('/stream');
  evtSource.onmessage = e => {
    const data = JSON.parse(e.data);
    if (data.connected) return;
    if (data.done) {
      setStatus(data.stopped ? 'Dihentikan.' : 'Selesai ✓', data.stopped ? 'stopped' : 'done');
      document.getElementById('btnRun').disabled = false;
      document.getElementById('btnStop').disabled = true;
      evtSource.close(); evtSource = null;
      loadDownloads();
      return;
    }
    if (data.text) appendLog(data.text);
  };
  evtSource.onerror = () => {};
}

// ── Run / Stop ────────────────────────────────────────────────────────────────
function runScraper() {
  if (!uploadedPath) { alert('Upload file daftar petugas terlebih dahulu.'); return; }
  const username = document.getElementById('username').value.trim();
  const password = document.getElementById('password').value.trim();
  if (!username || !password) { alert('Username dan password harus diisi.'); return; }

  clearLog();
  startStream();

  const vpnEnabled = document.getElementById('chkVpn').checked;
  const vpnHost    = document.getElementById('vpnHost').value.trim();

  fetch('/run', {
    method: 'POST',
    headers: {'Content-Type': 'application/json'},
    body: JSON.stringify({input_file: uploadedPath, username, password, vpn_enabled: vpnEnabled, vpn_host: vpnHost})
  })
  .then(r => r.json())
  .then(data => {
    if (data.error) {
      appendLog('[ERROR] ' + data.error + '\n');
      setStatus('Error', 'stopped');
      if (evtSource) { evtSource.close(); evtSource = null; }
      return;
    }
    setStatus('Berjalan…', 'running');
    document.getElementById('btnRun').disabled = true;
    document.getElementById('btnStop').disabled = false;
  })
  .catch(e => { appendLog('[ERROR] ' + e + '\n'); });
}

function stopScraper() {
  fetch('/stop', {method: 'POST'}).then(() => {
    setStatus('Menghentikan…', 'running');
    document.getElementById('btnStop').disabled = true;
  });
}

// ── Auto-run ──────────────────────────────────────────────────────────────────
function toggleAuto() {
  const enabled  = document.getElementById('chkAuto').checked;
  const hours    = parseInt(document.getElementById('autoHours').value) || 0;
  const mins     = parseInt(document.getElementById('autoMins').value)  || 0;
  const intervalMins = (hours * 60 + mins) || 120;
  const vpnEnabled = document.getElementById('chkVpn').checked;
  const vpnHost    = document.getElementById('vpnHost').value.trim();
  const username   = document.getElementById('username').value.trim();
  const password   = document.getElementById('password').value.trim();
  fetch('/auto', {
    method: 'POST',
    headers: {'Content-Type': 'application/json'},
    body: JSON.stringify({
      enabled, interval_mins: intervalMins,
      vpn_enabled: vpnEnabled, vpn_host: vpnHost,
      input_file: uploadedPath, username, password
    })
  }).then(r => r.json()).then(syncAutoUI);
}

function syncAutoUI(data) {
  document.getElementById('chkAuto').checked = data.auto_enabled;
  if (data.interval_mins != null) {
    document.getElementById('autoHours').value = Math.floor(data.interval_mins / 60);
    document.getElementById('autoMins').value  = data.interval_mins % 60;
  }
  if (data.vpn_enabled != null) document.getElementById('chkVpn').checked = data.vpn_enabled;
  if (data.vpn_host)   document.getElementById('vpnHost').value = data.vpn_host;
  const el = document.getElementById('nextRunText');
  if (data.auto_enabled && data.next_run) {
    const d = new Date(data.next_run * 1000);
    const hhmm = d.toLocaleTimeString('id-ID', {hour:'2-digit', minute:'2-digit'});
    el.textContent = '⏰ Jadwal berikutnya: ' + hhmm;
  } else {
    el.textContent = data.auto_enabled ? (data.params_ready ? '⏳ Menjadwalkan…' : '⏳ Isi file & kredensial lalu aktifkan auto-run.') : '';
  }

  // Sinkronkan status running (untuk auto-run yang berjalan di background)
  if (data.running && !evtSource) {
    clearLog();
    setStatus('Berjalan… (auto)', 'running');
    document.getElementById('btnRun').disabled = true;
    document.getElementById('btnStop').disabled = false;
    startStream();
  } else if (!data.running && document.getElementById('btnRun').disabled && !evtSource) {
    // Sudah selesai tapi stream tidak aktif (misal: tab baru dibuka saat auto-run selesai)
    setStatus('Selesai ✓', 'done');
    document.getElementById('btnRun').disabled = false;
    document.getElementById('btnStop').disabled = true;
  }
}

function loadDownloads() {
  fetch('/download/list').then(r => r.json()).then(data => {
    const sec  = document.getElementById('downloadSection');
    const list = document.getElementById('downloadList');
    if (!data.files || data.files.length === 0) { sec.style.display = 'none'; return; }
    list.innerHTML = data.files.map(f =>
      `<a class="dl-btn" href="/download/${encodeURIComponent(f)}" download>⬇ ${f}</a>`
    ).join('');
    sec.style.display = 'block';
  }).catch(() => {});
}

function saveVpnHost() {
  localStorage.setItem('vpnHost', document.getElementById('vpnHost').value);
}

function pollStatus() {
  fetch('/status').then(r => r.json()).then(syncAutoUI).catch(() => {});
}
setInterval(pollStatus, 10000);
pollStatus();

// Restore VPN host dari localStorage
const _savedVpn = localStorage.getItem('vpnHost');
if (_savedVpn) document.getElementById('vpnHost').value = _savedVpn;
</script>
</body>
</html>"""


# ── Routes ─────────────────────────────────────────────────────────────────────
@app.route("/")
def index():
    return render_template_string(HTML)


@app.route("/upload", methods=["POST"])
def upload_file():
    if "file" not in request.files:
        return jsonify({"error": "Tidak ada file"}), 400
    f = request.files["file"]
    if not f.filename:
        return jsonify({"error": "Nama file kosong"}), 400
    ext = os.path.splitext(f.filename)[1].lower()
    if ext not in (".xlsx", ".xls"):
        return jsonify({"error": "Hanya file .xlsx/.xls yang diizinkan"}), 400

    filename = secure_filename(f.filename)
    save_path = os.path.join(UPLOAD_DIR, filename)
    f.save(save_path)

    # Count rows to show feedback
    try:
        import openpyxl
        wb = openpyxl.load_workbook(save_path, read_only=True)
        ws = wb.active
        rows = sum(1 for row in ws.iter_rows(min_row=2, values_only=True) if row[0])
        wb.close()
    except Exception:
        rows = None

    return jsonify({"ok": True, "path": save_path, "rows": rows})


@app.route("/run", methods=["POST"])
def run_scraper():
    with _lock:
        if _state["running"]:
            return jsonify({"error": "Scraper sudah berjalan"}), 400

        data = request.get_json() or {}
        input_file = data.get("input_file", "").strip()
        username   = data.get("username",   "").strip()
        password   = data.get("password",   "").strip()

        if not input_file or not os.path.exists(input_file):
            return jsonify({"error": f"File tidak ditemukan: {input_file}"}), 400
        if not username or not password:
            return jsonify({"error": "Username/password kosong"}), 400

        vpn_enabled = bool(data.get("vpn_enabled", False))
        vpn_host    = data.get("vpn_host", "").strip()

        stop_event = threading.Event()
        _state["running"]    = True
        _state["stop_event"] = stop_event
        # Simpan params untuk auto-run
        _sched["params"]      = {"input_file": input_file, "username": username, "password": password}
        _sched["vpn_enabled"] = vpn_enabled
        _sched["vpn_host"]    = vpn_host

    # Drain old logs
    while not _log_queue.empty():
        try: _log_queue.get_nowait()
        except Exception: break

    def _run():
        old_stdout = sys.stdout
        sys.stdout = _QueueWriter()
        # ── VPN ───────────────────────────────────────────────────────────────
        if vpn_enabled and vpn_host:
            if not ensure_vpn(vpn_host, username, password):
                print("[VPN] Scraping dibatalkan karena VPN tidak terhubung.")
                sys.stdout = old_stdout
                _log_queue.put("\x00DONE\x00")
                with _lock:
                    _state["running"] = False
                _schedule_next()
                return
        # ── Scrape ────────────────────────────────────────────────────────────
        loop = asyncio.new_event_loop()
        asyncio.set_event_loop(loop)
        try:
            loop.run_until_complete(
                scrape_fasih.main_with_stop(stop_event, input_file, username, password)
            )
        except Exception as exc:
            print(f"\n[ERROR] {exc}")
        finally:
            loop.close()
            sys.stdout = old_stdout
            _log_queue.put("\x00DONE\x00")
            with _lock:
                _state["running"] = False
            _schedule_next()  # di dalam finally agar selalu terpanggil

    threading.Thread(target=_run, daemon=True).start()
    return jsonify({"ok": True})


@app.route("/stop", methods=["POST"])
def stop_scraper():
    with _lock:
        _state["stop_event"].set()
    return jsonify({"ok": True})


@app.route("/stream")
def stream():
    """Server-Sent Events — kirim log ke browser secara real-time."""
    def generate():
        yield 'data: {"connected":true}\n\n'
        while True:
            try:
                msg = _log_queue.get(timeout=25)
                if msg == "\x00DONE\x00":
                    stopped = _state["stop_event"].is_set()
                    yield f'data: {json.dumps({"done": True, "stopped": stopped})}\n\n'
                    return
                yield f'data: {json.dumps({"text": msg})}\n\n'
            except queue.Empty:
                yield ": ping\n\n"  # keep-alive

    return Response(
        generate(),
        mimetype="text/event-stream",
        headers={"Cache-Control": "no-cache", "X-Accel-Buffering": "no"},
    )


@app.route("/auto", methods=["POST"])
def set_auto():
    data = request.get_json() or {}
    with _lock:
        _sched["enabled"]       = bool(data.get("enabled", False))
        _sched["interval_mins"] = max(1, int(data.get("interval_mins", 120)))
        _sched["vpn_enabled"]   = bool(data.get("vpn_enabled", False))
        _sched["vpn_host"]      = data.get("vpn_host", "").strip()
        # Update params jika credentials dikirim dari form
        input_file = data.get("input_file", "").strip()
        username   = data.get("username",   "").strip()
        password   = data.get("password",   "").strip()
        if input_file and username and password and os.path.exists(input_file):
            _sched["params"] = {"input_file": input_file, "username": username, "password": password}
        if not _sched["enabled"] and _sched["timer"]:
            _sched["timer"].cancel()
            _sched["timer"]    = None
            _sched["next_run"] = None
    if _sched["enabled"] and _sched["params"] and not _state["running"]:
        _schedule_next()
    return _status_json()


@app.route("/status")
def status():
    # Watchdog: jika auto-run aktif tapi timer sudah mati dan tidak sedang running, reschedule
    with _lock:
        enabled     = _sched["enabled"]
        params_ok   = bool(_sched["params"])
        running     = _state["running"]
        timer       = _sched["timer"]
        timer_alive = (timer is not None and timer.is_alive())
    if enabled and params_ok and not running and not timer_alive:
        _schedule_next()
    return _status_json()


def _status_json():
    with _lock:
        nr = _sched["next_run"]
        return jsonify({
            "running":       _state["running"],
            "auto_enabled":  _sched["enabled"],
            "interval_mins": _sched["interval_mins"],
            "vpn_enabled":   _sched["vpn_enabled"],
            "vpn_host":      _sched["vpn_host"],
            "next_run":      nr.timestamp() if nr else None,
            "params_ready":  bool(_sched["params"]),
        })


@app.route("/download/list")
def download_list():
    """Kembalikan daftar file Excel terbaru di folder output/."""
    out_dir = scrape_fasih.OUTPUT_DIR
    if not os.path.isdir(out_dir):
        return jsonify({"files": []})
    files = []
    for fname in os.listdir(out_dir):
        if fname.endswith(".xlsx") and fname.startswith("rekap_fasih"):
            fpath = os.path.join(out_dir, fname)
            files.append({"name": fname, "mtime": os.path.getmtime(fpath)})
    files.sort(key=lambda f: f["mtime"], reverse=True)
    return jsonify({"files": [f["name"] for f in files[:10]]})  # maks 10 terbaru


@app.route("/download/<path:filename>")
def download_file(filename):
    """Unduh satu file Excel dari folder output/."""
    safe = secure_filename(filename)
    fpath = os.path.join(scrape_fasih.OUTPUT_DIR, safe)
    if not os.path.isfile(fpath):
        return "File tidak ditemukan", 404
    return send_file(fpath, as_attachment=True, download_name=safe)


# ── Entry point ────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    import webbrowser
    url = "http://localhost:5000"
    print(f"FASIH Scraper Web UI → {url}")
    webbrowser.open(url)
    app.run(host="127.0.0.1", port=5000, debug=False, threaded=True)
