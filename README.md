# Panduan Penggunaan — FASIH Scraper & Dashboard BPS Bali Utara

Sistem ini terdiri dari dua bagian:

| Bagian | Fungsi | Akses |
|--------|--------|-------|
| **Scraper** | Mengambil data Open/Submitted/Rejected per pencacah dari FASIH | [s.bps.go.id/5108_dashboardGCPLN](http://s.bps.go.id/5108_dashboardGCPLN) |
| **Dashboard Monitoring** | Menampilkan rekap harian dari Google Sheets | Sama — tab di halaman yang sama |

---

## Bagian 1 — Menjalankan Scraper

### Langkah 1 — Buka aplikasi

Buka browser dan akses:

> **[http://s.bps.go.id/5108_dashboardGCPLN](http://s.bps.go.id/5108_dashboardGCPLN)**

Halaman scraper akan terbuka otomatis.

---

### Langkah 2 — Siapkan file daftar petugas

Buat file Excel (`daftar_petugas.xlsx`) dengan format berikut:

| Nama          | Email                     |
|---------------|---------------------------|
| Budi Santoso  | budi.santoso@bps.go.id    |
| Ani Rahayu    | ani.rahayu@bps.go.id      |

- Baris pertama = header (`Nama` dan `Email`)
- Kolom Email diisi dengan **username SSO BPS** masing-masing petugas (biasanya `nama.nip@bps.go.id`)
- Tidak perlu mengubah apapun selain mengisi nama dan email

---

### Langkah 3 — Isi form dan jalankan

Di halaman scraper:

1. **File Petugas** → klik **Browse** → pilih file `daftar_petugas.xlsx`
2. **Username** → isi dengan username SSO BPS **Anda sendiri** (bukan petugas)
3. **Password** → isi dengan password SSO BPS Anda
4. **UPI** → isi kode UPI wilayah Anda, contoh: `[55]` untuk Bali
5. **UP3** → isi kode UP3 wilayah Anda, contoh: `[55UTR]` untuk Bali Utara
6. **Sheets URL** → isi dengan URL Apps Script milik Anda (lihat [cara mendapatkan URL ini](#cara-setup-google-sheets--apps-script))
7. Klik tombol **▶ Run**

> Gunakan akun SSO BPS Anda sendiri. Akun ini hanya digunakan untuk login ke FASIH, tidak disimpan di mana pun.

> **Catatan:** Setiap pengguna yang membuka halaman ini di browser berbeda berjalan secara **independen** — menjalankan scraper di satu browser tidak memengaruhi pengguna lain.

---

### Langkah 4 — Tunggu proses selesai

- Log berjalan di layar secara real-time
- Proses untuk 50+ petugas membutuhkan sekitar **20–30 menit**
- Jangan tutup browser selama proses berjalan
- Setelah selesai, muncul notifikasi **"Selesai ✓"**

---

### Langkah 5 — Unduh hasil Excel

Setelah selesai, klik tombol **Download Excel** untuk mengunduh file rekap:

```
rekap_fasih_pascabayar_20250313_083000.xlsx
rekap_fasih_prabayar_20250313_083500.xlsx
```

Kolom yang tersedia:

| No | Nama | Email | Open | Submitted by Pencacah | Rejected by Admin Kabupaten |
|----|------|-------|------|-----------------------|-----------------------------|

---

---

## Cara Setup Google Sheets & Apps Script

Ikuti langkah ini **satu kali** sebelum pertama kali menjalankan scraper.

### Langkah 1 — Buat Google Sheets

1. Buka [Google Sheets](https://sheets.google.com) → buat spreadsheet baru
2. Beri nama, misalnya: `Rekap FASIH 2025`
3. Buat **dua sheet** (tab di bawah):
   - Klik **+** di pojok kiri bawah
   - Sheet pertama beri nama: `Ringkasan`
   - Sheet kedua beri nama: `Riwayat`

---

### Langkah 2 — Buat Apps Script

1. Di spreadsheet, klik menu **Extensions → Apps Script**
2. Hapus semua kode yang ada, lalu **paste kode berikut**:

```javascript
function doPost(e) {
  try {
    const ss   = SpreadsheetApp.getActiveSpreadsheet();
    const data = JSON.parse(e.postData.contents);

    if (data.type === "detail") {
      // ── Sheet Riwayat ─────────────────────────────────────────────
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
      // ── Sheet Ringkasan ───────────────────────────────────────────
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

### Langkah 3 — Deploy sebagai Web App

1. Klik menu **Deploy → New deployment**
2. Di samping "Select type", klik ikon ⚙️ → pilih **Web app**
3. Isi pengaturan:
   - **Description**: `FASIH Webhook`
   - **Execute as**: `Me`
   - **Who has access**: `Anyone`
4. Klik **Deploy**
5. Klik **Authorize access** → pilih akun Google Anda → klik **Allow**
6. Salin URL yang muncul — bentuknya:
   ```
   https://script.google.com/macros/s/XXXXX.../exec
   ```

> **Simpan URL ini** — inilah yang diisi di kolom **Sheets URL** pada halaman scraper.

---

### Langkah 4 — Isi URL di scraper

Kembali ke halaman scraper, tempel URL tersebut di kolom **Sheets URL**. URL akan tersimpan otomatis di browser, jadi tidak perlu diisi ulang setiap kali.

---

## Bagian 2 — Membaca Google Sheets

Hasil scraper otomatis tersimpan ke Google Sheets dalam dua sheet:

### Sheet "Ringkasan" — rekap harian

Berisi total per survei per hari:

| Tanggal | Waktu | Open Pasca | Submit Pasca | Reject Pasca | Open Praba | Submit Praba | Reject Praba |
|---------|-------|-----------|--------------|--------------|-----------|--------------|--------------|
| 13/03/2025 | 08:30 | 45 | 210 | 3 | 38 | 195 | 1 |

Gunakan sheet ini untuk memantau **tren harian** progress pengisian.

---

### Sheet "Riwayat" — detail per pencacah

Berisi data lengkap setiap pencacah setiap kali scraper dijalankan:

| Tanggal | Waktu | Nama | Email | Open Pasca | Submit Pasca | Reject Pasca | Open Praba | Submit Praba | Reject Praba |
|---------|-------|------|-------|-----------|--------------|--------------|-----------|--------------|--------------|

Gunakan sheet ini untuk:
- Memantau pencacah mana yang belum submit
- Melihat riwayat perubahan per hari
- Filter/sort berdasarkan nama atau jumlah submit

---

## Bagian 3 — Dashboard Monitoring

Dashboard di halaman yang sama menampilkan visualisasi data dari Google Sheets secara otomatis.

### Cara membaca dashboard

- **Kartu ringkasan** di bagian atas → total Open, Submitted, Rejected hari ini
- **Tabel per pencacah** → status masing-masing petugas (hijau = sudah submit semua, merah = masih ada yang open)
- **Grafik tren** → progress harian dari awal periode survei

### Memperbarui data di dashboard

Dashboard membaca data langsung dari Google Sheets. Data akan otomatis ter-update setelah scraper selesai dijalankan — tidak perlu refresh manual.

---

## Pertanyaan Umum

**Login SSO gagal / salah password?**
Pastikan username dan password sama dengan yang digunakan untuk login ke `sso.bps.go.id`. Jika baru ganti password, gunakan yang terbaru.

**Email petugas tidak ditemukan?**
Pastikan kolom Email di `daftar_petugas.xlsx` diisi dengan username SSO (bukan email pribadi). Contoh: `budi.santoso.1234@bps.go.id`.

**Proses berhenti di tengah jalan?**
Klik **▶ Run** lagi. Scraper akan mengulang dari awal. File Excel akan di-overwrite dengan data terbaru.

**Data di dashboard tidak berubah?**
Refresh halaman browser (`F5`). Jika masih tidak berubah, cek apakah scraper sudah benar-benar selesai (status "Selesai ✓").

---

## Bagian 4 — Menjalankan Scraper di PC Sendiri (Aplikasi Desktop)

Gunakan cara ini jika ingin menjalankan scraper **langsung di PC masing-masing** tanpa bergantung pada server pusat.

---

### Langkah 1 — Install Python

1. Buka [python.org/downloads](https://www.python.org/downloads/)
2. Download Python **3.11** atau lebih baru
3. Saat instalasi, **centang** opsi **"Add Python to PATH"**
4. Klik Install Now

Cek instalasi: buka Command Prompt, ketik:
```
python --version
```
Harus muncul versi Python, misal `Python 3.11.x`.

---

### Langkah 2 — Salin file program

Salin folder ini ke PC Anda (bisa lewat flashdisk, Google Drive, atau email):

```
fasih_scraper/
├── gui_fasih.py        ← aplikasi utama
├── scrape_fasih.py     ← engine scraping
├── requirements.txt    ← daftar library
└── input/
    └── daftar_petugas.xlsx
```

---

### Langkah 3 — Install library

Buka Command Prompt, masuk ke folder program:
```
cd C:\path\ke\fasih_scraper
```

Install semua library:
```
pip install -r requirements.txt
```

Install browser Playwright (hanya perlu dilakukan sekali):
```
playwright install chromium
```

> Jika muncul error `pip tidak dikenal`, coba: `python -m pip install -r requirements.txt`

---

### Langkah 4 — Jalankan aplikasi

```
python gui_fasih.py
```

Akan muncul jendela aplikasi seperti ini:

```
┌─────────────────────────────────────────┐
│  FASIH Scraper — BPS                    │
├─────────────────────────────────────────┤
│ File Petugas: [____________] [Browse…]  │
│ Username:     [____________]            │
│ Password:     [____________] [Tampilkan]│
│ UPI: [____]   UP3: [______]            │
│ Sheets URL:   [________________________]│
│                                         │
│ [▶ Run] [■ Stop] [📂 Buka Folder Hasil]│
│                                         │
│ Log:                                    │
│ ┌─────────────────────────────────────┐ │
│ │ [INFO] 53 petugas loaded...         │ │
│ │ [DONE] pascabayar: 53 petugas       │ │
│ └─────────────────────────────────────┘ │
└─────────────────────────────────────────┘
```

---

### Langkah 5 — Isi form dan jalankan

| Field | Isi dengan |
|-------|-----------|
| **File Petugas** | Klik Browse → pilih `daftar_petugas.xlsx` |
| **Username** | Username SSO BPS Anda (`nama.nip@bps.go.id`) |
| **Password** | Password SSO BPS Anda |
| **UPI** | Kode UPI, contoh: `[55]` untuk Bali |
| **UP3** | Kode UP3, contoh: `[55UTR]` untuk Bali Utara |
| **Sheets URL** | URL Apps Script (lihat [cara setup](#cara-setup-google-sheets--apps-script)) |

Klik **▶ Run** — log berjalan real-time di bawah.

---

### Langkah 6 — Hasil scraping

Setelah selesai (muncul status **"Selesai ✓"**):

- Klik **📂 Buka Folder Hasil** → folder `output/` terbuka di Explorer
- Di dalamnya terdapat file Excel:
  ```
  rekap_fasih_pascabayar_20250313_083000.xlsx
  rekap_fasih_prabayar_20250313_083500.xlsx
  ```
- Data juga otomatis terkirim ke Google Sheets (jika Sheets URL diisi)

---

### Membuat shortcut agar mudah dibuka

1. Klik kanan di Desktop → **New → Shortcut**
2. Isi lokasi: `python C:\path\ke\fasih_scraper\gui_fasih.py`
3. Klik Next → beri nama `FASIH Scraper` → Finish

Atau buat file `jalankan.bat` di folder program:
```bat
@echo off
cd /d %~dp0
python gui_fasih.py
```
Klik dua kali file ini untuk membuka aplikasi.

---

### Troubleshooting

**`ModuleNotFoundError: No module named 'playwright'`**
→ Jalankan ulang: `pip install -r requirements.txt`

**`playwright install` gagal / browser tidak terbuka**
→ Jalankan: `playwright install chromium --with-deps`

**Login SSO gagal**
→ Pastikan username format `nama.nip@bps.go.id` dan password sama dengan login ke `sso.bps.go.id`

**Jendela tidak muncul saat klik Run**
→ Cek apakah file `daftar_petugas.xlsx` sudah dipilih dan path-nya benar

---

## Kontak

Untuk pertanyaan teknis, hubungi pengelola sistem di BPS Bali Utara.
