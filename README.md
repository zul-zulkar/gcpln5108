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
4. Klik tombol **▶ Run**

> Gunakan akun SSO BPS Anda sendiri. Akun ini hanya digunakan untuk login ke FASIH, tidak disimpan di mana pun.

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

## Kontak

Untuk pertanyaan teknis, hubungi pengelola sistem di BPS Bali Utara.
