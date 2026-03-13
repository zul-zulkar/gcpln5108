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
import uuid
from datetime import datetime, timedelta

import openpyxl
from flask import Flask, Response, request, jsonify, render_template_string, send_file
from werkzeug.utils import secure_filename
import scrape_fasih

app = Flask(__name__)
UPLOAD_DIR = "uploads"
os.makedirs(UPLOAD_DIR, exist_ok=True)


# ── Per-thread stdout routing ───────────────────────────────────────────────────
class _ThreadLocalStdout:
    """Routes writes to the current thread's queue; falls back to real stdout."""
    _local = threading.local()

    @classmethod
    def set_queue(cls, q):
        cls._local.queue = q

    @classmethod
    def clear_queue(cls):
        cls._local.queue = None

    def write(self, text):
        q = getattr(self._local, "queue", None)
        if q is not None and text:
            q.put(text)
        else:
            sys.__stdout__.write(text)

    def flush(self):
        pass


_tl_stdout = _ThreadLocalStdout()
sys.stdout = _tl_stdout


# ── Session store ───────────────────────────────────────────────────────────────
_sessions: dict = {}
_sessions_lock = threading.Lock()


def _new_session() -> str:
    sid = str(uuid.uuid4())
    sess = {
        "lock":       threading.Lock(),
        "running":    False,
        "stop_event": threading.Event(),
        "log_queue":  queue.Queue(maxsize=5000),
        "sched": {
            "enabled":       False,
            "interval_mins": 120,
            "next_run":      None,
            "timer":         None,
            "params":        {},
            "vpn_enabled":   False,
            "vpn_host":      "",
        },
    }
    with _sessions_lock:
        _sessions[sid] = sess
    return sid


def _sess(sid: str):
    with _sessions_lock:
        return _sessions.get(sid)


# ── Auto-scheduler (per session) ────────────────────────────────────────────────
def _schedule_next(sid: str):
    sess = _sess(sid)
    if not sess:
        return
    sched = sess["sched"]
    with sess["lock"]:
        if not sched["enabled"] or not sched["params"]:
            return
        if sched["timer"]:
            sched["timer"].cancel()
        delay = sched["interval_mins"] * 60
        sched["next_run"] = datetime.now() + timedelta(seconds=delay)
        t = threading.Timer(delay, _auto_run, args=(sid,))
        t.daemon = True
        t.start()
        sched["timer"] = t


def _auto_run(sid: str):
    sess = _sess(sid)
    if not sess:
        return
    sched = sess["sched"]
    with sess["lock"]:
        if sess["running"]:
            t = threading.Timer(300, _auto_run, args=(sid,))
            t.daemon = True
            t.start()
            sched["timer"] = t
            sched["next_run"] = datetime.now() + timedelta(seconds=300)
            return
        params      = sched["params"].copy()
        vpn_enabled = sched["vpn_enabled"]
        vpn_host    = sched["vpn_host"]
        if not params:
            return
        lq = sess["log_queue"]
        with lq.mutex:
            lq.queue.clear()
        stop_event = threading.Event()
        sess["running"]    = True
        sess["stop_event"] = stop_event

    lq.put(f"\n[AUTO] ▶ Jadwal otomatis dimulai ({datetime.now().strftime('%H:%M:%S')})…\n")
    _launch_scrape_thread(sess, sid, stop_event, lq, params, vpn_enabled, vpn_host)


def _launch_scrape_thread(sess, sid, stop_event, lq, params, vpn_enabled, vpn_host):
    """Shared helper: runs scrape_fasih.main_with_stop in a daemon thread,
    routing stdout to lq, and marking the session not-running when done."""
    username   = params["username"]
    password   = params["password"]
    input_file = params["input_file"]
    sheets_url = params.get("sheets_url", "")
    upi_text   = params.get("upi_text", "")
    up3_text   = params.get("up3_text", "")

    def _run():
        _ThreadLocalStdout.set_queue(lq)
        try:
            if vpn_enabled and vpn_host:
                if not ensure_vpn(vpn_host, username, password):
                    print("[VPN] Scraping dibatalkan karena VPN tidak terhubung.")
                    return
            loop = asyncio.new_event_loop()
            asyncio.set_event_loop(loop)
            try:
                loop.run_until_complete(
                    scrape_fasih.main_with_stop(
                        stop_event, input_file, username, password,
                        sheets_url=sheets_url, upi_text=upi_text, up3_text=up3_text,
                    )
                )
            except Exception as exc:
                print(f"\n[ERROR] {exc}")
            finally:
                loop.close()
        finally:
            _ThreadLocalStdout.clear_queue()
            lq.put("\x00DONE\x00")
            with sess["lock"]:
                sess["running"] = False
            _schedule_next(sid)

    threading.Thread(target=_run, daemon=True).start()


# ── VPN helpers ────────────────────────────────────────────────────────────────
_FASIH_HOST = "fasih-sm.bps.go.id"
_FORTICLIENT_PATHS = [
    r"C:\Program Files\Fortinet\FortiClient\FortiSSLVPNclient.exe",
    r"C:\Program Files (x86)\Fortinet\FortiClient\FortiSSLVPNclient.exe",
    r"C:\Program Files\Fortinet\FortiClient\FortiClient.exe",
    r"C:\Program Files (x86)\Fortinet\FortiClient\FortiClient.exe",
]


def _is_vpn_connected():
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


def _launch_forticlient_connect(fclient, tunnel, username, password):
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
        except subprocess.TimeoutExpired:
            return True
        except Exception:
            pass
    return False


def ensure_vpn(vpn_host, username, password):
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
            print("[VPN] Tidak bisa auto-connect via CLI. Membuka FortiClient GUI…")
            print("[VPN] Silakan connect VPN secara manual di jendela yang terbuka.")
            try:
                subprocess.Popen([fclient], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
            except Exception as exc:
                print(f"[VPN] Gagal membuka FortiClient: {exc}")

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
  .form-row input[type=text].half { flex: 0 0 160px; }
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
  .section-divider {
    font-size:.78rem; color:#aaa; text-transform:uppercase; letter-spacing:.5px;
    margin: 16px 0 10px; border-bottom: 1px solid #eee; padding-bottom: 4px;
  }
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

    <div class="section-divider">Konfigurasi Wilayah & Sheets</div>

    <div class="form-row">
      <label>UPI</label>
      <input type="text" id="upiText" class="half" placeholder="mis. [55]" autocomplete="off" oninput="saveLocal('upiText')">
      <label style="width:auto;margin-left:16px;">UP3</label>
      <input type="text" id="up3Text" class="half" placeholder="mis. [55UTR]" autocomplete="off" oninput="saveLocal('up3Text')">
    </div>

    <div class="form-row">
      <label>Sheets URL</label>
      <input type="text" id="sheetsUrl" placeholder="URL Apps Script Web App (opsional)" autocomplete="off" oninput="saveLocal('sheetsUrl')">
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
      <input type="text" id="vpnHost" class="vpn-input" placeholder="Host VPN (mis. vpn.bps.go.id)" oninput="saveLocal('vpnHost')">
      <span id="vpnStatus"></span>
    </div>

    <div class="btn-row">
      <button id="btnRun" onclick="runScraper()" disabled>&#9654; Run</button>
      <button id="btnStop" onclick="stopScraper()" disabled>&#9632; Stop</button>
      <span id="statusText"><span class="dot" id="dot"></span>Memuat…</span>
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
let sessionId = null;

// ── Session init ──────────────────────────────────────────────────────────────
async function initSession() {
  const saved = localStorage.getItem('sessionId');
  const url = saved ? `/session/new?sid=${saved}` : '/session/new';
  const r = await fetch(url);
  const d = await r.json();
  sessionId = d.session_id;
  localStorage.setItem('sessionId', sessionId);
}

// ── Drag & Drop ───────────────────────────────────────────────────────────────
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
function saveLocal(key) {
  localStorage.setItem(key, document.getElementById(key).value);
}

// ── Stream ────────────────────────────────────────────────────────────────────
function startStream() {
  if (evtSource) { evtSource.close(); evtSource = null; }
  evtSource = new EventSource(`/stream?sid=${sessionId}`);
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
  const sheetsUrl  = document.getElementById('sheetsUrl').value.trim();
  const upiText    = document.getElementById('upiText').value.trim();
  const up3Text    = document.getElementById('up3Text').value.trim();

  fetch('/run', {
    method: 'POST',
    headers: {'Content-Type': 'application/json'},
    body: JSON.stringify({
      session_id: sessionId,
      input_file: uploadedPath, username, password,
      vpn_enabled: vpnEnabled, vpn_host: vpnHost,
      sheets_url: sheetsUrl, upi_text: upiText, up3_text: up3Text
    })
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
    _schedulePoll(2000);
  })
  .catch(e => { appendLog('[ERROR] ' + e + '\n'); });
}

function stopScraper() {
  fetch('/stop', {
    method: 'POST',
    headers: {'Content-Type': 'application/json'},
    body: JSON.stringify({session_id: sessionId})
  }).then(() => {
    setStatus('Menghentikan…', 'running');
    document.getElementById('btnStop').disabled = true;
    _schedulePoll(2000);
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
  const sheetsUrl  = document.getElementById('sheetsUrl').value.trim();
  const upiText    = document.getElementById('upiText').value.trim();
  const up3Text    = document.getElementById('up3Text').value.trim();
  fetch('/auto', {
    method: 'POST',
    headers: {'Content-Type': 'application/json'},
    body: JSON.stringify({
      session_id: sessionId,
      enabled, interval_mins: intervalMins,
      vpn_enabled: vpnEnabled, vpn_host: vpnHost,
      input_file: uploadedPath, username, password,
      sheets_url: sheetsUrl, upi_text: upiText, up3_text: up3Text
    })
  }).then(r => r.json()).then(data => {
    syncAutoUI(data);
    if (data.auto_enabled) _schedulePoll(10000);
  });
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

  if (data.running && !evtSource) {
    clearLog();
    setStatus('Berjalan… (auto)', 'running');
    document.getElementById('btnRun').disabled = true;
    document.getElementById('btnStop').disabled = false;
    startStream();
  } else if (!data.running && document.getElementById('btnRun').disabled && !evtSource) {
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

let _pollTimer = null;
function _schedulePoll(ms) {
  clearTimeout(_pollTimer);
  _pollTimer = setTimeout(pollStatus, ms);
}
function pollStatus() {
  if (!sessionId) return;
  fetch(`/status?sid=${sessionId}`)
    .then(r => r.json())
    .then(data => {
      syncAutoUI(data);
      // fast poll while running, normal poll while auto is on, stop otherwise
      if (data.running)           _schedulePoll(2000);
      else if (data.auto_enabled) _schedulePoll(10000);
      // else: idle — no next poll until user action
    })
    .catch(() => { _schedulePoll(15000); }); // retry on error
}

// ── Init ──────────────────────────────────────────────────────────────────────
document.addEventListener('DOMContentLoaded', async () => {
  await initSession();

  // Restore localStorage
  ['upiText','up3Text','sheetsUrl','vpnHost'].forEach(k => {
    const v = localStorage.getItem(k);
    if (v) document.getElementById(k).value = v;
  });

  // Enable Run button now that session is ready
  setStatus('Siap', '');
  document.getElementById('btnRun').disabled = false;

  loadDownloads();
  pollStatus();
});
</script>
</body>
</html>"""


# ── Routes ─────────────────────────────────────────────────────────────────────
def _parse_run_params(data: dict) -> dict:
    """Extract scraper execution params from request JSON dict."""
    return {
        "input_file": data.get("input_file", "").strip(),
        "username":   data.get("username",   "").strip(),
        "password":   data.get("password",   "").strip(),
        "sheets_url": data.get("sheets_url", "").strip(),
        "upi_text":   data.get("upi_text",   "").strip(),
        "up3_text":   data.get("up3_text",   "").strip(),
    }


@app.route("/")
def index():
    return render_template_string(HTML)


@app.route("/session/new")
def new_session():
    sid = request.args.get("sid", "").strip()
    if sid and _sess(sid) is not None:
        return jsonify({"session_id": sid})
    new_sid = _new_session()
    return jsonify({"session_id": new_sid})


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

    try:
        wb = openpyxl.load_workbook(save_path, read_only=True)
        ws = wb.active
        rows = sum(1 for row in ws.iter_rows(min_row=2, values_only=True) if row[0])
        wb.close()
    except Exception:
        rows = None

    return jsonify({"ok": True, "path": save_path, "rows": rows})


@app.route("/run", methods=["POST"])
def run_scraper():
    data = request.get_json() or {}
    sid  = data.get("session_id", "").strip()
    sess = _sess(sid)
    if sess is None:
        return jsonify({"error": "Session tidak valid, refresh halaman."}), 400

    with sess["lock"]:
        if sess["running"]:
            return jsonify({"error": "Scraper sudah berjalan"}), 400

        params     = _parse_run_params(data)
        input_file = params["input_file"]
        username   = params["username"]
        password   = params["password"]

        if not input_file or not os.path.exists(input_file):
            return jsonify({"error": f"File tidak ditemukan: {input_file}"}), 400
        if not username or not password:
            return jsonify({"error": "Username/password kosong"}), 400

        vpn_enabled = bool(data.get("vpn_enabled", False))
        vpn_host    = data.get("vpn_host", "").strip()

        stop_event = threading.Event()
        sess["running"]    = True
        sess["stop_event"] = stop_event
        sess["sched"]["params"]      = params
        sess["sched"]["vpn_enabled"] = vpn_enabled
        sess["sched"]["vpn_host"]    = vpn_host

    lq = sess["log_queue"]
    with lq.mutex:
        lq.queue.clear()

    _launch_scrape_thread(sess, sid, stop_event, lq, params, vpn_enabled, vpn_host)
    return jsonify({"ok": True})


@app.route("/stop", methods=["POST"])
def stop_scraper():
    data = request.get_json() or {}
    sid  = data.get("session_id", "").strip()
    sess = _sess(sid)
    if sess:
        with sess["lock"]:
            sess["stop_event"].set()
    return jsonify({"ok": True})


@app.route("/stream")
def stream():
    sid  = request.args.get("sid", "").strip()
    sess = _sess(sid)

    def generate():
        yield 'data: {"connected":true}\n\n'
        if not sess:
            yield f'data: {json.dumps({"done": True, "stopped": False})}\n\n'
            return
        lq = sess["log_queue"]
        while True:
            try:
                msg = lq.get(timeout=25)
                if msg == "\x00DONE\x00":
                    stopped = sess["stop_event"].is_set()
                    yield f'data: {json.dumps({"done": True, "stopped": stopped})}\n\n'
                    return
                yield f'data: {json.dumps({"text": msg})}\n\n'
            except queue.Empty:
                yield ": ping\n\n"

    return Response(
        generate(),
        mimetype="text/event-stream",
        headers={"Cache-Control": "no-cache", "X-Accel-Buffering": "no"},
    )


@app.route("/auto", methods=["POST"])
def set_auto():
    data = request.get_json() or {}
    sid  = data.get("session_id", "").strip()
    sess = _sess(sid)
    if sess is None:
        return jsonify({"error": "Session tidak valid"}), 400

    with sess["lock"]:
        sched = sess["sched"]
        sched["enabled"]       = bool(data.get("enabled", False))
        sched["interval_mins"] = max(1, int(data.get("interval_mins", 120)))
        sched["vpn_enabled"]   = bool(data.get("vpn_enabled", False))
        sched["vpn_host"]      = data.get("vpn_host", "").strip()
        params = _parse_run_params(data)
        if params["input_file"] and params["username"] and params["password"] and os.path.exists(params["input_file"]):
            sched["params"] = params
        if not sched["enabled"] and sched["timer"]:
            sched["timer"].cancel()
            sched["timer"]    = None
            sched["next_run"] = None

    if sess["sched"]["enabled"] and sess["sched"]["params"] and not sess["running"]:
        _schedule_next(sid)
    return _status_json(sess)


@app.route("/status")
def status():
    sid  = request.args.get("sid", "").strip()
    sess = _sess(sid)
    if sess is None:
        return jsonify({"running": False, "auto_enabled": False, "interval_mins": 120,
                        "vpn_enabled": False, "vpn_host": "", "next_run": None, "params_ready": False})
    # Watchdog: reschedule if timer died
    with sess["lock"]:
        sched      = sess["sched"]
        enabled    = sched["enabled"]
        params_ok  = bool(sched["params"])
        running    = sess["running"]
        timer      = sched["timer"]
        timer_alive = (timer is not None and timer.is_alive())
    if enabled and params_ok and not running and not timer_alive:
        _schedule_next(sid)
    return _status_json(sess)


def _status_json(sess):
    with sess["lock"]:
        sched = sess["sched"]
        nr    = sched["next_run"]
        return jsonify({
            "running":       sess["running"],
            "auto_enabled":  sched["enabled"],
            "interval_mins": sched["interval_mins"],
            "vpn_enabled":   sched["vpn_enabled"],
            "vpn_host":      sched["vpn_host"],
            "next_run":      nr.timestamp() if nr else None,
            "params_ready":  bool(sched["params"]),
        })


@app.route("/download/list")
def download_list():
    out_dir = scrape_fasih.OUTPUT_DIR
    if not os.path.isdir(out_dir):
        return jsonify({"files": []})
    files = []
    for fname in os.listdir(out_dir):
        if fname.endswith(".xlsx") and fname.startswith("rekap_fasih"):
            fpath = os.path.join(out_dir, fname)
            files.append({"name": fname, "mtime": os.path.getmtime(fpath)})
    files.sort(key=lambda f: f["mtime"], reverse=True)
    return jsonify({"files": [f["name"] for f in files[:10]]})


@app.route("/download/<path:filename>")
def download_file(filename):
    safe  = secure_filename(filename)
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
    app.run(host="0.0.0.0", port=5000, debug=False, threaded=True)
