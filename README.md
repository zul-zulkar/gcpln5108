# Panduan Lengkap — FASIH Scraper & Dashboard BPS Bali Utara

Sistem ini terdiri dari dua komponen:

| Komponen | Fungsi |
|----------|--------|
| **Aplikasi Scraper (GUI)** | Mengambil data Open/Submitted/Rejected per pencacah dari FASIH secara otomatis |
| **Dashboard Monitoring** | Menampilkan rekap dan grafik tren dari Google Sheets |

---

## Daftar Isi

1. [Persiapan — Buat File Daftar Petugas](#bagian-1--persiapan--buat-file-daftar-petugas)
2. [Persiapan — Siapkan Google Sheets & Apps Script](#bagian-2--persiapan--siapkan-google-sheets--apps-script)
3. [Menjalankan Aplikasi Scraper (GUI)](#bagian-3--menjalankan-aplikasi-scraper-gui)
4. [Membuka & Menggunakan Dashboard](#bagian-4--membuka--menggunakan-dashboard)
5. [Troubleshooting](#troubleshooting)
6. [Untuk Developer — Cara B: Menjalankan via Python](#untuk-developer--cara-b-menjalankan-via-python)

---

## Bagian 1 — Persiapan: Buat File Daftar Petugas

Sebelum menjalankan scraper, siapkan file Excel berisi daftar pencacah.

### Format file

Buat file Excel baru (misalnya `daftar_petugas.xlsx`) dengan format:

| Nama | Email |
|------|-------|
| Budi Santoso | budi.santoso.1234@bps.go.id |
| Ani Rahayu | ani.rahayu.5678@bps.go.id |

**Aturan:**
- Baris pertama = header (`Nama` dan `Email`)
- Kolom `Email` diisi dengan **username SSO BPS** masing-masing pencacah (`nama.nip@bps.go.id`)
- Tidak perlu kolom lain

> **Tip:** Username SSO adalah yang digunakan untuk login ke `sso.bps.go.id`.

---

## Bagian 2 — Persiapan: Siapkan Google Sheets & Apps Script

Ikuti langkah ini **satu kali** sebelum pertama kali menjalankan scraper.

### Langkah 1 — Buat Google Sheets

1. Buka [Google Sheets](https://sheets.google.com) → klik **+** untuk buat spreadsheet baru
2. Beri nama, misalnya: `Rekap FASIH 2025`
3. Buat **tiga tab** (klik **+** di pojok kiri bawah) dengan nama persis:
   - `Utama`
   - `Ringkasan`
   - `Riwayat`

> **Penting:** Nama tab harus persis seperti di atas (huruf kapital di awal). Dashboard membaca tab berdasarkan nama ini.

---

### Langkah 2 — Isi tab Utama (daftar petugas & ULP)

Tab `Utama` digunakan dashboard untuk tombol filter per ULP.

| ULP | Nama | Email |
|-----|------|-------|
| ULP Singaraja | Budi Santoso | budi.santoso.1234@bps.go.id |
| ULP Buleleng | Ani Rahayu | ani.rahayu.5678@bps.go.id |

- Baris pertama = header (`ULP`, `Nama`, `Email`)
- Kolom `Email` harus sama persis dengan yang ada di `daftar_petugas.xlsx`

---

### Langkah 3 — Buat Apps Script

1. Di spreadsheet, klik menu **Extensions → Apps Script**
2. Hapus semua kode yang ada, lalu paste kode berikut:

```javascript
function doPost(e) {
  try {
    const ss   = SpreadsheetApp.getActiveSpreadsheet();
    const data = JSON.parse(e.postData.contents);

    if (data.type === "detail") {
      const ws = ss.getSheetByName("Riwayat") || ss.insertSheet("Riwayat");
      if (ws.getLastRow() === 0) {
        ws.appendRow(["Tanggal","Waktu","Nama","Email",
          "Open Pasca","Submit Pasca","Reject Pasca",
          "Open Praba","Submit Praba","Reject Praba"]);
      }
      (data.rows || []).forEach(r => {
        ws.appendRow([data.tanggal, data.waktu, r.nama, r.email,
          r.open_pasca, r.submit_pasca, r.reject_pasca,
          r.open_praba, r.submit_praba, r.reject_praba]);
      });
    } else {
      const ws = ss.getSheetByName("Ringkasan") || ss.insertSheet("Ringkasan");
      if (ws.getLastRow() === 0) {
        ws.appendRow(["Tanggal","Waktu",
          "Open Pasca","Submit Pasca","Reject Pasca",
          "Open Praba","Submit Praba","Reject Praba"]);
      }
      ws.appendRow([data.tanggal, data.waktu,
        data.open_pasca, data.submit_pasca, data.reject_pasca,
        data.open_praba, data.submit_praba, data.reject_praba]);
    }

    return ContentService.createTextOutput("ok");
  } catch(err) {
    return ContentService.createTextOutput("error: " + err.message);
  }
}
```

3. Klik **💾 Simpan** (Ctrl+S)

---

### Langkah 4 — Deploy sebagai Web App

1. Klik menu **Deploy → New deployment**
2. Klik ikon ⚙️ → pilih **Web app**
3. Isi pengaturan:
   - **Execute as**: `Me`
   - **Who has access**: `Anyone`
4. Klik **Deploy** → **Authorize access** → pilih akun Google → **Allow**
5. Salin **URL** yang muncul:
   ```
   https://script.google.com/macros/s/XXXXX.../exec
   ```

> Simpan URL ini — inilah yang diisi di kolom **URL Apps Script** pada aplikasi scraper.

---

### Langkah 5 — Atur akses baca (untuk dashboard)

1. Klik tombol **Share** di spreadsheet
2. Di bagian **General access**, ubah ke **Anyone with the link → Viewer**
3. Klik **Done**

---

## Bagian 3 — Menjalankan Aplikasi Scraper (GUI)

### Langkah 1 — Salin folder FASIH_Scraper

Minta folder `FASIH_Scraper` dari pengelola sistem (atau salin dari flashdisk/Google Drive).

```
FASIH_Scraper\
├── FASIH_Scraper.exe    ← klik dua kali untuk membuka
├── browsers\            ← Chromium browser (sudah termasuk)
├── _internal\           ← file pendukung (jangan dihapus)
├── fasih_settings.json  ← pengaturan tersimpan (dibuat otomatis)
└── output\              ← hasil Excel tersimpan di sini (dibuat otomatis)
```

> **Penting:** Salin **seluruh folder**, bukan hanya `.exe`.

---

### Langkah 2 — Jalankan aplikasi

Klik dua kali `FASIH_Scraper.exe`.

Jika muncul peringatan Windows SmartScreen: klik **More info** → **Run anyway**.

---

### Langkah 3 — Isi form

| Field | Isi dengan |
|-------|------------|
| **File Petugas** | Klik Browse → pilih `daftar_petugas.xlsx` |
| **Username** | Username SSO BPS Anda (`nama.nip@bps.go.id`) |
| **Password** | Password SSO BPS Anda |
| **UPI** | Kode UPI wilayah, contoh: `[55]` untuk Bali |
| **UP3** | Kode UP3 wilayah, contoh: `[55UTR]` untuk Bali Utara |
| **URL Apps Script** | URL dari Langkah 4 di atas |

> Pengaturan (kecuali password) **tersimpan otomatis** ke `fasih_settings.json` — tidak perlu diisi ulang saat aplikasi dibuka kembali.

---

### Langkah 4 — (Opsional) Aktifkan auto-run harian

Untuk menjalankan scraper otomatis setiap hari pada jam tertentu:

1. Centang **Aktifkan auto-run harian pukul:**
2. Isi jam target, contoh: `07:00`
3. Klik **▶ Run** sekali agar pengaturan tersimpan

Selanjutnya, cukup buka aplikasi sebelum jam yang ditentukan — scraper berjalan otomatis saat jam tiba.

> Aplikasi harus tetap terbuka agar auto-run aktif. Untuk otomatis penuh tanpa membuka manual, gunakan Windows Task Scheduler untuk membuka `FASIH_Scraper.exe` pada waktu sesuai.

---

### Langkah 5 — (Opsional) Aktifkan VPN

Jika jaringan kantor memerlukan VPN:

1. Centang **Reconnect VPN sebelum Run**
2. Isi **Nama Koneksi** — nama koneksi VPN di Windows
3. Isi **User VPN** dan **Pass VPN** jika diperlukan

Mendukung FortiClient dan VPN Windows standar.

---

### Langkah 6 — Klik Run dan tunggu

Klik **▶ Run** — log berjalan real-time di bawah.

Proses untuk 50+ pencacah membutuhkan **20–30 menit**. Jangan tutup aplikasi selama berjalan.

Setelah selesai (status **"Selesai ✓"**):
- Klik **📂 Buka Folder Hasil** → folder `output\` terbuka di Explorer
- File Excel tersimpan di sana:
  ```
  rekap_fasih_pascabayar_20250313_083000.xlsx
  rekap_fasih_prabayar_20250313_083500.xlsx
  ```
- Data juga otomatis terkirim ke Google Sheets (jika URL Apps Script diisi)

---

### Membuat shortcut di Desktop

Klik kanan `FASIH_Scraper.exe` → **Send to → Desktop (create shortcut)**

---

## Bagian 4 — Membuka & Menggunakan Dashboard

### Cara membuka dashboard

**Lokal (offline):** Salin folder `gcpln5108\` ke PC → klik dua kali `index.html`.

**Online (dihosting):** Upload `index.html` dan `config.js` ke server web atau Google Sites.

---

### Mengatur sumber data default (config.js)

File `config.js` menentukan Google Sheets yang dibaca dashboard saat pertama dibuka. Cukup edit file ini satu kali — semua pengguna langsung melihat data yang benar tanpa perlu konfigurasi.

1. Buka `config.js` dengan Notepad
2. Tempel URL lengkap spreadsheet Anda:

```javascript
const CONFIG = {
  DEFAULT_SHEET_URL: 'https://docs.google.com/spreadsheets/d/XXXXX.../edit',
};
```

Cukup salin URL dari address bar browser saat spreadsheet terbuka — tidak perlu mengambil ID secara manual.

3. Simpan `config.js`

---

### Mengganti sumber data dari dalam browser

Tanpa mengedit file, pengguna bisa mengganti sumber data:

1. Klik ikon **⚙️** di pojok kanan atas dashboard
2. Tempel URL spreadsheet Google Sheets
3. Klik **💾 Simpan & Muat**

Pengaturan ini tersimpan di browser (tidak mengubah `config.js`). Klik **↺ Default** untuk kembali ke sumber bawaan.

---

### Cara membaca dashboard

- **Filter ULP** di bagian atas → klik untuk melihat per kantor cabang
- **Kartu ringkasan** → total Open / Submitted / Rejected hari ini
- **Tabel per pencacah** → hijau = sudah submit semua, merah = masih ada open
- **Grafik tren harian** → progress dari awal periode survei
- **Grafik per pencacah** → riwayat historis, bisa difilter per ULP dan per individu
- **Tombol ↑** di pojok kanan bawah → kembali ke atas halaman
- **🔄 Perbarui** di header → muat ulang data sekarang (otomatis setiap 5 menit)

---

## Troubleshooting

**Peringatan "Windows protected your PC"**
→ Klik **More info** → **Run anyway**. EXE belum bersertifikat publisher berbayar — aman.

**Aplikasi langsung tertutup saat dibuka**
→ Pastikan folder `_internal\` dan `browsers\` ada di lokasi yang sama dengan `.exe`.

**Login SSO gagal**
→ Pastikan username format `nama.nip@bps.go.id` dan password sama dengan login ke `sso.bps.go.id`.

**Email petugas tidak ditemukan**
→ Pastikan kolom Email di `daftar_petugas.xlsx` diisi username SSO (bukan email pribadi).

**Proses berhenti di tengah jalan**
→ Klik **▶ Run** lagi. Scraper mengulang dari awal.

**Data di dashboard tidak berubah setelah scraper selesai**
→ Tekan `F5` untuk refresh. Jika masih tidak berubah, pastikan scraper sudah benar-benar selesai.

**Filter ULP tidak muncul di dashboard**
→ Pastikan tab `Utama` sudah diisi dengan kolom `ULP` dan `Email`, lalu refresh dashboard.

**Dashboard kosong / tidak menampilkan data saat pertama dibuka**
→ Edit `config.js`, isi `DEFAULT_SHEET_ID` dengan Sheet ID spreadsheet Anda. Pastikan spreadsheet sudah diatur "Anyone with the link → Viewer".

**Auto-run tidak berjalan**
→ Pastikan aplikasi tetap terbuka dan jam komputer sudah benar.

---

## Untuk Developer — Cara B: Menjalankan via Python

### Langkah 1 — Install Python

Download Python **3.11+** dari [python.org/downloads](https://www.python.org/downloads/). Saat instalasi, **centang "Add Python to PATH"**.

### Langkah 2 — Salin file program

```
fasih_scraper\
├── gui_fasih.py
├── scrape_fasih.py
├── requirements.txt
└── input\
    └── daftar_petugas.xlsx
```

### Langkah 3 — Install library

```
pip install -r requirements.txt
playwright install chromium
```

### Langkah 4 — Jalankan

```
python gui_fasih.py
```

### Membuat ulang EXE dari source code

```
pip install pyinstaller
playwright install chromium
```

Lalu klik dua kali `build.bat` — EXE dan browser dikemas ke `dist\FASIH_Scraper\`.

**Troubleshooting Python:**
- `ModuleNotFoundError` → jalankan ulang `pip install -r requirements.txt`
- `playwright install` gagal → coba `playwright install chromium --with-deps`

---

## Kontak

Untuk pertanyaan teknis, hubungi pengelola sistem di BPS Bali Utara.
