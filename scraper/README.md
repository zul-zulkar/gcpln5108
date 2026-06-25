# FASIH Scraper — BPS

Tool otomatisasi untuk merekap status **Open**, **Submitted by Pencacah**, dan **Rejected by Admin Kabupaten** per petugas dari sistem [FASIH Survey Collection BPS](https://fasih-sm.bps.go.id), lalu menyimpannya ke **Excel** dan menyinkronkan ke **Google Sheets**.

---

## Cara kerja singkat (alur aktif)

FASIH kini dilindungi **WAF anti-bot** yang memblokir login otomatis Playwright. Solusinya: **login manual sekali** di Chrome (manusia melewati WAF + SSO), lalu skrip menempel ke sesi itu via **CDP** dan membaca **API JSON** UI baru (tab *Dasbor → Rekap Petugas*) — bukan scraping DOM.

```
[Chrome debug + login manual]  ──CDP──►  run_api.py  ──►  Excel (output/)
        port 9222                            │
                                             └──► Google Sheets (Apps Script webhook)
```

Keunggulan alur API: ambil semua petugas sekali jalan, dan **bebas dari bug nilai-basi** yang sempat muncul di UI lama (DOM filter per-petugas).

---

## Prasyarat

- Python 3.10+
- Google Chrome terpasang
- Akun pegawai BPS dengan akses FASIH
- `pip install -r requirements.txt` lalu `playwright install chromium`

---

## Cara Menjalankan (alur baru — direkomendasikan)

### Lewat GUI Desktop (paling mudah)

```bash
python gui_fasih.py
```

1. Klik **🌐 Buka Chrome & Login** → jendela Chrome debug terbuka.
2. **Login FASIH manual** di jendela itu (lewati WAF + SSO sebagai manusia).
3. Klik **Run** → data ditarik via API → Excel + Google Sheets.

### Lewat terminal (CLI)

```bash
# 1. Buka Chrome debug + login manual:
start_fasih_chrome.bat          # lalu login FASIH di jendela yang muncul

# 2. Tarik data:
python run_api.py               # pascabayar + prabayar -> Excel + Google Sheets
python run_api.py --no-sheets   # Excel saja (tidak kirim ke Sheets)
```

> Catatan Windows: gunakan `http://127.0.0.1:9222` (bukan `localhost`) saat menyambung CDP — `localhost` mengarah ke IPv6 `::1` sedangkan Chrome debug bind ke IPv4.

---

## Konfigurasi

### 1. Daftar petugas — `input/daftar_petugas.xlsx`

| Nama          | Email                     |
|---------------|---------------------------|
| Budi Santoso  | budi.santoso@bps.go.id    |
| Ani Rahayu    | ani.rahayu@bps.go.id      |

Baris pertama = header (`Nama`, `Email`). Email harus sama dengan akun SSO BPS.

### 2. Survei & periode — `run_api.py`

```python
SURVEYS_NEW = {
    "pascabayar": {"survey_id": "<UUID-survey>", "period_id": "<UUID-periode>"},
    "prabayar":   {"survey_id": "<UUID-survey>", "period_id": "<UUID-periode>"},
}
```

`survey_id` dan `period_id` adalah dua UUID di URL UI baru:
`/app/surveys/<survey_id>/<period_id>`. Jika `period_id` dikosongkan (`None`),
skrip mencoba menemukannya otomatis dari link Dasbor. `surveyRoleId` (Pencacah)
dan `regionSummaryLevel` dideteksi otomatis.

### 3. Google Sheets (opsional)

Isi `SHEETS_WEBHOOK_URL` di `scrape_fasih.py` dengan URL deploy Apps Script.
Kosongkan (`""`) jika tidak dipakai. Webhook **meng-upsert** per tanggal
(sheet `Ringkasan`) dan per tanggal+email (sheet `Riwayat`), sehingga
menjalankan ulang di hari yang sama otomatis menimpa data lama.

---

## Output

File Excel tersimpan di `output/` dengan nama `rekap_fasih_<timestamp>.xlsx`.

| No | Nama | Email | Open | Submitted by Pencacah | Rejected by Admin Kabupaten |
|----|------|-------|------|-----------------------|-----------------------------|

**Gerbang kualitas:** data hanya dikirim ke Google Sheets bila semua survei berhasil
ditarik lengkap (semua petugas cocok). Bila ada yang gagal, Excel tetap disimpan tapi
pengiriman Sheets dibatalkan — jalankan ulang setelah koneksi stabil.

---

## Struktur Proyek

```
scrape_fasih/
├── run_api.py            # ⭐ Entry point pipeline (CDP/API -> Excel + Sheets)
├── fasih_api.py          # Modul ekstraksi API UI baru (report-progress-by-responsibility)
├── scrape_fasih.py       # Helper bersama: baca petugas, simpan Excel, kirim Sheets
├── gui_fasih.py          # GUI Desktop (Tkinter) — memanggil run_api
├── web_fasih.py          # ⚠ Web UI Flask (alur LAMA auto-login — kini diblokir WAF)
├── rawdata_fasih.py      # Scraper raw-data tabel FASIH (mandiri, output ke rawdata/)
├── start_fasih_chrome.bat# Buka Chrome debug (port 9222) untuk login manual
├── build.bat             # Build EXE via PyInstaller
├── FASIH_Scraper.spec    # Spec PyInstaller (target: gui_fasih.py)
├── playwright_runtime_hook.py
├── requirements.txt
├── fasih_settings.json   # Pengaturan GUI (auto-generate)
├── input/                # ⚠ daftar_petugas.xlsx (TIDAK di-commit)
├── output/               # ⚠ Hasil rekap .xlsx (TIDAK di-commit)
├── rawdata/              # Output rawdata_fasih.py
├── uploads/              # Upload dari web UI
├── archive/              # Skrip lama/usang (mis. inspect_page.py)
├── gcpbi/                # Sub-proyek terpisah: scraper PBI
└── gcpln5108/            # Sub-proyek terpisah: dashboard web (repo git sendiri)
```

---

## Google Sheets Integration

<details>
<summary>Klik untuk lihat langkah setup Apps Script</summary>

1. Buka [Google Sheets](https://sheets.google.com), buat spreadsheet baru.
2. Buat dua sheet: `Ringkasan` dan `Riwayat`.
3. **Extensions → Apps Script** → buat Web App yang menerima POST JSON dari scraper.
4. **Deploy → New Deployment → Web App** (Execute as: *Me*, Who has access: *Anyone*).
5. Salin URL deploy ke `SHEETS_WEBHOOK_URL` di `scrape_fasih.py`.

Format payload **Ringkasan harian**:
```json
{ "tanggal":"13/03/2025","waktu":"08:30",
  "open_pasca":12,"submit_pasca":45,"reject_pasca":2,
  "open_praba":8,"submit_praba":38,"reject_praba":1 }
```

Format payload **Detail per pencacah** (`type:"detail"`):
```json
{ "type":"detail","tanggal":"13/03/2025","waktu":"08:30",
  "rows":[ { "nama":"Budi Santoso","email":"budi.santoso@bps.go.id",
    "open_pasca":3,"submit_pasca":10,"reject_pasca":0,
    "open_praba":2,"submit_praba":8,"reject_praba":0 } ] }
```

</details>

---

## Build EXE (opsional)

```bash
pip install pyinstaller
playwright install chromium
build.bat
```

Hasil di `dist\FASIH_Scraper\` (kirim seluruh folder, bukan hanya `.exe`).
Folder `build/` dan `dist/` adalah artefak regeneratable — aman dihapus dan
sudah dikecualikan di `.gitignore`.

---

## Catatan Keamanan

- **Jangan commit** `input/daftar_petugas.xlsx` (berisi email pegawai) — sudah di-`.gitignore`.
- **Jangan hardcode** username/password SSO di kode — masukkan lewat UI.
  (Skrip lama di `archive/` mungkin masih memuat kredensial plaintext; hapus bila tak dipakai.)
- File `output/` & `rawdata/` dikecualikan karena dapat berisi data personel.

---

## Changelog

### v2.0.0 — 2026-06-25
- **Alur baru CDP + API** menggantikan auto-login (diblokir WAF anti-bot FASIH):
  login manual sekali di Chrome debug, skrip menempel via CDP dan membaca API
  UI baru (`fasih_api.py` + `run_api.py`).
- GUI (`gui_fasih.py`) memakai pipeline baru; tombol **Buka Chrome & Login**;
  input UPI/UP3 dihapus (cakupan wilayah lewat `surveyPeriodId`).
- **Gerbang kualitas** sebelum kirim ke Sheets; retry 504/token-basi.
- Webhook Sheets meng-upsert (rerun di hari sama menimpa data lama).
- Rapikan proyek: hapus artefak build regeneratable & sampah debug, arsipkan skrip usang.

### v1.2.0 — 2026-03-25
- **Auto Reconnect VPN** (`gui_fasih.py`): force disconnect + reconnect FortiClient sebelum scraping.

### v1.1.0
- Antarmuka desktop Tkinter dengan auto-run terjadwal, simpan pengaturan, tombol Stop, build EXE.

### v1.0.0
- Rilis awal: scraper CLI Playwright, login SSO BPS, ekspor Excel, sinkron Google Sheets, web UI Flask.

---

## Lisensi

Internal BPS — tidak untuk distribusi publik.
