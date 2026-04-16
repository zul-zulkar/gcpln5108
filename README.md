# Panduan — FASIH Scraper & Dashboard BPS Kabupaten Buleleng

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
6. [Changelog](#changelog)

---

## Bagian 1 — Persiapan: Buat File Daftar Petugas

Sebelum menjalankan scraper, siapkan file Excel berisi daftar pencacah.

### Format file

Buat file Excel baru (misalnya `daftar_petugas.xlsx`) dengan format:

| Nama | Email |
|------|-------|
| Budi Santoso | budi.santoso.1234@gmail.com |
| Ani Rahayu | ani.rahayu.5678@gmail.com |

**Aturan:**
- Baris pertama = header (`Nama` dan `Email`)
- Kolom `Email` diisi dengan masing-masing pencacah (`nama.nip@gmail.com`)
- Tidak perlu kolom lain

---

## Bagian 2 — Persiapan: Siapkan Google Sheets & Apps Script

Ikuti langkah ini **satu kali** sebelum pertama kali menjalankan scraper.

### Langkah 1 — Buat Google Sheets

1. Buka [Google Sheets](https://sheets.google.com) → klik **+** untuk buat spreadsheet baru
2. Beri nama, misalnya: `Rekap FASIH 2025`
3. Buat tab berikut (klik **+** di pojok kiri bawah):

| Tab | Wajib? | Fungsi |
|-----|--------|--------|
| `Utama` | ✅ | Daftar petugas & ULP untuk tombol filter dashboard |
| `Ringkasan` | ✅ | Ringkasan harian (diisi Apps Script) |
| `Riwayat` | ✅ | Detail per pencacah per hari (diisi Apps Script) |
| `Target_Prabayar` | Opsional | Target pelanggan prabayar per ULP — jika diisi, dashboard otomatis menampilkan progres vs target |

contoh spreadsheet sumber data ada pada tautan berikut:
https://docs.google.com/spreadsheets/d/1BGXuTJkaOYJKT6Mo0ldR1mmT2wLNRYqyRcmcxOnSCa0/edit?usp=sharing 

> **Penting:** Nama tab harus persis seperti di atas (huruf kapital di awal). Dashboard membaca tab berdasarkan nama ini.

---

### Langkah 1b — (Opsional) Isi tab Target_Prabayar

Buat tab bernama `Target_Prabayar` dengan dua kolom:

| ULP | Target |
|-----|--------|
| ULP Singaraja | 4500 |
| ULP Seririt | 2800 |
| ULP Buleleng | 1200 |

**Aturan:**
- Baris pertama = header; nama kolom bebas asal mengandung kata `ULP` dan `Target` (atau `Jumlah`, `Rekening`, `Nilai`)
- Satu baris per ULP — **tidak perlu baris "Total"**, dashboard menghitung jumlah keseluruhan secara otomatis
- Nama ULP tidak harus sama persis dengan yang ada di tab `Utama` — perbedaan huruf besar/kecil dan spasi diabaikan secara otomatis
- Jika tab ini tidak ada, dashboard tetap berfungsi normal; progres prabayar hanya tidak menampilkan persentase vs target

---

### Langkah 2 — Isi tab Utama (daftar petugas & ULP)

Tab `Utama` digunakan dashboard untuk tombol filter per ULP.

| ULP | Nama | Email |
|-----|------|-------|
| ULP Singaraja | Budi Santoso | bud****oso@gmail.com |
| ULP Buleleng | Ani Rahayu | ani****ayu@yahoo.com |

- Baris pertama = header (`ULP`, `Nama`, `Email`)
- Kolom `Email` harus diisi dalam **format tersamar** yang sama dengan yang dihasilkan Apps Script

> **Cara mudah:** Jalankan scraper sekali, lihat hasil di tab `Riwayat`, lalu salin format email dari sana ke tab `Utama`.

---

### Langkah 3 — Buat Apps Script (Opsional)

Appscript ini akan membantu menambahkan data hasil scraping ke Google sheet secara otomatis sesuai dengan format yang ditentukan. Namun apabila tidak menggunakan Fitur ini, silakan paste hasil scraping secara manual melalui folder output yang akan dijelaskan kemudian.

1. Di spreadsheet, klik menu **Extensions → Apps Script**
2. Hapus semua kode yang ada, lalu paste kode berikut:

```javascript
// Samarkan bagian tengah local-part, tampilkan 3 karakter awal + 3 karakter akhir
// Contoh: budisantoso@gmail.com → bud****oso@gmail.com
//         ani.rahayu@yahoo.com  → ani****ayu@yahoo.com
function maskEmail(email) {
  if (!email || !email.includes('@')) return email;
  const at     = email.indexOf('@');
  const local  = email.slice(0, at);
  const domain = email.slice(at);
  if (local.length <= 6) {
    return local[0] + '*'.repeat(local.length - 1) + domain;
  }
  return local.slice(0, 3) + '*'.repeat(local.length - 6) + local.slice(-3) + domain;
}

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    if (data.type === 'detail') {
      handleDetail(data);
    } else {
      handleRingkasan(data);
    }
    return ContentService.createTextOutput(JSON.stringify({ok: true}))
      .setMimeType(ContentService.MimeType.JSON);
  } catch(err) {
    return ContentService.createTextOutput(JSON.stringify({error: err.message}))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ── Helper: konversi Date object → string "dd/MM/yyyy" ─────────────────────
function _fmt(val) {
  if (val === null || val === undefined || val === '') return '';
  if (val instanceof Date)
    return Utilities.formatDate(val, Session.getScriptTimeZone(), "dd/MM/yyyy");
  return String(val).trim();
}

// ── Tab Ringkasan ──────────────────────────────────────────────────────────────
function handleRingkasan(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Ringkasan');
  if (!sheet) {
    sheet = ss.insertSheet('Ringkasan');
    sheet.appendRow(['Tanggal','Waktu','Open Pasca','Submit Pasca','Reject Pasca',
                     'Open Praba','Submit Praba','Reject Praba']);
    sheet.getRange('A:A').setNumberFormat('@');
  }

  const tanggal = data.tanggal;
  const newRow = [
    tanggal, data.waktu,
    data.open_pasca,   data.submit_pasca, data.reject_pasca,
    data.open_praba,   data.submit_praba, data.reject_praba,
  ];

  // Upsert: cari baris dengan tanggal sama
  const lastRow = sheet.getLastRow();
  if (lastRow >= 2) {
    const dates = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    for (let i = 0; i < dates.length; i++) {
      if (_fmt(dates[i][0]) === tanggal) {
        sheet.getRange(i + 2, 1, 1, newRow.length).setValues([newRow]);
        sheet.getRange(i + 2, 1).setNumberFormat('@');
        return;
      }
    }
  }
  sheet.appendRow(newRow);
  sheet.getRange(sheet.getLastRow(), 1).setNumberFormat('@');
}

// ── Tab Riwayat ─────────────────────────────────────────────────────────────
function handleDetail(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Riwayat');
  if (!sheet) {
    sheet = ss.insertSheet('Riwayat');
    sheet.appendRow(['Tanggal','Nama','Email',
                     'Open Pasca','Submit Pasca','Reject Pasca',
                     'Open Praba','Submit Praba','Reject Praba']);
    sheet.getRange('A:A').setNumberFormat('@');
  }

  const tanggal = data.tanggal;
  const rows    = data.rows || [];

  // Lookup: key = "tanggal|maskedEmail" → nomor baris sheet
  // Kolom: A=Tanggal, B=Nama, C=Email
  const lastRow = sheet.getLastRow();
  const lookup  = {};
  if (lastRow >= 2) {
    const existing = sheet.getRange(2, 1, lastRow - 1, 3).getValues();
    for (let i = 0; i < existing.length; i++) {
      const key = _fmt(existing[i][0]) + '|' + String(existing[i][2]).trim();
      lookup[key] = i + 2;
    }
  }

  for (const r of rows) {
    const email  = String(r.email || '').trim();
    if (!email) continue;
    const masked = maskEmail(email);
    const key    = tanggal + '|' + masked;
    const newRow = [
      tanggal, r.nama || '', masked,
      r.open_pasca   != null ? r.open_pasca   : '',
      r.submit_pasca != null ? r.submit_pasca : '',
      r.reject_pasca != null ? r.reject_pasca : '',
      r.open_praba   != null ? r.open_praba   : '',
      r.submit_praba != null ? r.submit_praba : '',
      r.reject_praba != null ? r.reject_praba : '',
    ];

    if (lookup[key]) {
      sheet.getRange(lookup[key], 1, 1, newRow.length).setValues([newRow]);
      sheet.getRange(lookup[key], 1).setNumberFormat('@');
    } else {
      sheet.appendRow(newRow);
      sheet.getRange(sheet.getLastRow(), 1).setNumberFormat('@');
    }
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

folder `FASIH_Scraper` dari pengelola sistem dapat diakses melalui tautan berikut :
https://drive.google.com/drive/folders/1qM383JjE7mWTXlcTXFEeFw22YkUhwS4K?usp=sharing

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
| **URL Apps Script (Opsional)** | URL dari Langkah 4 di atas |

> Pengaturan (kecuali password) **tersimpan otomatis** ke `fasih_settings.json` — tidak perlu diisi ulang saat aplikasi dibuka kembali.

---

### Langkah 4 — (Opsional) Aktifkan auto-run berkala

Untuk menjalankan scraper secara otomatis setiap X menit:

1. Centang **Aktifkan auto-run setiap:**
2. Isi interval, contoh: `60` untuk setiap 1 jam
3. Klik **▶ Run** sekali agar pengaturan tersimpan

Selanjutnya scraper akan berjalan otomatis setiap interval yang ditentukan selama aplikasi tetap terbuka.

> Untuk otomatis penuh, gunakan Windows Task Scheduler untuk membuka `FASIH_Scraper.exe` pada waktu yang diinginkan.

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

Proses untuk 50+ pencacah membutuhkan **3-5 menit**. Jangan tutup aplikasi selama berjalan.

Setelah selesai (status **"Selesai ✓"**):
- Klik **📂 Buka Folder Hasil** → folder `output\` terbuka di Explorer
- Satu file Excel tersimpan di sana dengan dua sheet:
  ```
  rekap_fasih_20250313_083000.xlsx
    ├── Sheet: Ringkasan   ← total Open/Submit/Reject Pasca & Praba
    └── Sheet: Riwayat     ← detail per pencacah
  ```
- Data juga otomatis terkirim ke Google Sheets (jika URL Apps Script diisi)

---

### Membuat shortcut di Desktop

Klik kanan `FASIH_Scraper.exe` → **Send to → Desktop (create shortcut)**

---

## Bagian 4 — Membuka & Menggunakan Dashboard

### Cara membuka dashboard

**Lokal (offline):** Salin folder `gcpln5108\` ke PC → klik dua kali `index.html`.

Pastikan semua file dan subfolder berada dalam satu folder:
```
gcpln5108\
├── index.html   ← buka file ini
├── app.js
├── config.js
├── css\         ← stylesheet per komponen (jangan dihapus)
└── js\          ← modul JavaScript (jangan dihapus)
```

**Online (dihosting):** Upload seluruh folder di atas ke server web atau GitHub Pages.

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

1. Klik tombol **☰** di pojok kiri atas untuk membuka sidebar
2. Buka bagian **⚙ Settings** di sidebar
3. Tempel URL spreadsheet Google Sheets
4. Klik **💾 Simpan & Muat**

Pengaturan ini tersimpan di browser (tidak mengubah `config.js`). Klik **↺ Default** untuk kembali ke sumber bawaan.

---

### Cara membaca dashboard

- **Tombol ☰** di pojok kiri atas → buka/tutup sidebar navigasi
- **Sidebar** → berisi navigasi antar section, tombol refresh, alert threshold, target prabayar, settings, dan pengaturan tampilan
- **Filter ULP** di bagian atas konten → klik untuk melihat per kantor cabang
- **Kartu ringkasan** → total Open / Submitted / Rejected hari ini
  - Prabayar: jika target sudah diisi, progress bar menampilkan **Submit ÷ Target** (bukan ÷ total pelanggan)
  - Muncul keterangan sisa pelanggan yang belum tercapai, atau tanda ✓ jika target sudah terpenuhi
- **Tabel per pencacah** → hijau = sudah submit semua, merah = masih ada open
- **Filter tanggal** (header kanan atas):
  - Saat filter aktif, tombol berubah warna **kuning/oranye** dan berlabel **"Filter"** agar mudah dikenali
  - Tombol **× Reset** muncul di sebelahnya untuk kembali ke data hari ini
  - Filter ini **otomatis menyinkronkan** batas atas tanggal pada grafik Tren Harian dan Progres per Pencacah — banner oranye muncul di dalam grafik sebagai penanda
- **Grafik tren harian** → progress dari awal periode survei; label persentase prabayar menampilkan **% vs target** jika target sudah ditetapkan (kolom header tabel berubah menjadi **%🎯**)
- **Grafik per pencacah** → riwayat historis, bisa difilter per ULP dan per individu
- **Tombol ↑** di pojok kanan bawah → kembali ke atas halaman

---

### Mengatur target prabayar

Target prabayar menentukan denominasi progress bar dan persentase di grafik tren. Ada tiga cara mengisinya (prioritas dari bawah ke atas):

| Cara | Keterangan | Berlaku untuk |
|------|------------|---------------|
| `config.js` (`DEFAULT_TARGETS`) | Fallback awal sebelum sheet dimuat | Semua pengguna |
| Tab `Target_Prabayar` di spreadsheet | **Sumber utama** — langsung terbaca saat dashboard dibuka | Semua pengguna |
| Sidebar → **Target Prabayar** | Override manual per browser — menimpa nilai sheet untuk ULP yang diubah | Hanya browser tersebut |

Untuk **menghapus override lokal** dan kembali ke nilai sheet, buka sidebar → Target Prabayar → klik **↺ Ke Sheet**.

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
→ Klik **🔄 Refresh** di header. Jika masih tidak berubah, pastikan scraper sudah benar-benar selesai.

**Filter ULP tidak muncul di dashboard**
→ Pastikan tab `Utama` sudah diisi dengan kolom `ULP` dan `Email`, lalu refresh dashboard.

**Dashboard kosong / tidak menampilkan data saat pertama dibuka**
→ Edit `config.js`, isi `DEFAULT_SHEET_URL` dengan URL lengkap spreadsheet Anda. Pastikan spreadsheet sudah diatur "Anyone with the link → Viewer".

**Auto-run tidak berjalan**
→ Pastikan aplikasi tetap terbuka dan centang sudah diaktifkan.

**Target prabayar tidak terbaca dari sheet**
→ Pastikan tab bernama persis `Target_Prabayar` (huruf kapital T dan P). Periksa nama kolom: harus mengandung kata `ULP` dan salah satu dari `Target`, `Jumlah`, `Rekening`, atau `Nilai`. Buka konsol browser (F12 → Console) untuk melihat pesan error `[Target]`.

**Progress bar prabayar masih menampilkan % selesai, bukan % vs target**
→ Klik **🔄 Refresh** untuk memuat ulang semua data termasuk sheet target. Jika tetap tidak berubah, pastikan tab `Target_Prabayar` sudah memiliki data dan spreadsheet masih diatur "Anyone with the link → Viewer".

---

## Changelog

### v1.3.0 — 16 April 2026

**Dashboard**

- **Target Prabayar** — progress bar dan persentase prabayar kini menggunakan `submit ÷ target` (bukan `÷ total pelanggan`) jika target sudah ditetapkan; muncul keterangan sisa pelanggan atau tanda ✓ tercapai
- **Sheet `Target_Prabayar`** — dashboard otomatis membaca target per ULP dari tab baru ini; total keseluruhan dihitung otomatis dari jumlah semua ULP (tidak perlu baris "Total" di sheet)
- **Pencocokan nama ULP fleksibel** — perbedaan huruf besar/kecil dan spasi antara sheet target dan data utama diabaikan secara otomatis
- **Override target via sidebar** — bagian "Target Prabayar" di sidebar memungkinkan pengguna mengubah target per ULP per browser; nilai sheet tetap menjadi basis dan dapat dipulihkan kapan saja
- **Filter tanggal aktif lebih jelas** — saat filter tanggal header aktif, tombol berubah warna kuning/oranye dengan label "Filter" dan animasi denyut; tombol "× Reset" muncul di sebelahnya
- **Sinkronisasi otomatis filter ke grafik tren** — filter tanggal header otomatis membatasi batas atas tanggal pada grafik Tren Harian dan Progres per Pencacah; banner oranye muncul sebagai penanda di dalam grafik
- **Grafik tren — % vs target** — label persentase prabayar di grafik Tren Ringkasan menampilkan persentase terhadap target (bukan completion rate); kolom "%" di tabel ringkasan berubah menjadi "%🎯" saat target aktif

---

### v1.2.0 — 25 Maret 2026

**Scraper (GUI)**
- Fitur **Auto Reconnect VPN** — VPN di-disconnect dan di-reconnect ulang secara otomatis sebelum scraping dimulai, mengatasi kondisi VPN tersambung tapi tidak berfungsi dengan baik
- Proses reconnect berjalan penuh tanpa interaksi manual dengan tiga strategi fallback: `rasdial` → kill proses FortiClient + CLI connect → restart service FortiClient
- Paket distribusi (`FASIH_Scraper/`) dibersihkan — folder Dashboard tidak lagi disertakan

**Dashboard**
- Sidebar kiri yang dapat disembunyikan/ditampilkan, responsif di semua ukuran layar termasuk mobile
- Semua kontrol (refresh, alert, settings, tema, navigasi) dipindahkan dari header ke dalam sidebar
- CSS dipisah menjadi 10 file terorganisir di folder `css/` (variables, layout, sidebar, cards, tables, charts, filters, panels, animations, darkmode)
- Modul JS sidebar dipisahkan ke `js/sidebar.js`

### v1.1.0

**Scraper (GUI)**
- Antarmuka desktop (`FASIH_Scraper.exe`) berbasis Tkinter
- Auto VPN connect saat run pertama kali
- Auto-run terjadwal dengan interval jam/menit
- Simpan pengaturan otomatis ke `fasih_settings.json`
- Tombol Stop untuk membatalkan scraping di tengah jalan

### v1.0.0

- Rilis awal scraper CLI berbasis Playwright
- Login SSO BPS (Keycloak), filter ngx-select Angular
- Ekspor rekap ke Excel (`.xlsx`)
- Sinkronisasi ke Google Sheets via Apps Script webhook
- Dashboard monitoring web (HTML/CSS/JS statis)

---

## Kontak

Untuk pertanyaan teknis, hubungi pengelola sistem di BPS Kabupaten Buleleng.
