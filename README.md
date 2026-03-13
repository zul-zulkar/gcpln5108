# FASIH Scraper — BPS

Tool otomatis untuk merekap status pengisian survei FASIH (Open, Submitted, Rejected) per pencacah, lengkap dengan ekspor Excel dan dashboard web.

---

## Yang dibutuhkan

- Windows 10/11
- [Python 3.10+](https://www.python.org/downloads/) — saat install, **centang "Add Python to PATH"**
- Koneksi internet
- Akun SSO BPS

---

## Instalasi (sekali saja)

Buka **Command Prompt** atau **PowerShell**, lalu jalankan satu per satu:

```
pip install -r requirements.txt
playwright install chromium
```

---

## Persiapan file petugas

1. Buka folder `input/`
2. Buat file Excel bernama **`daftar_petugas.xlsx`**
3. Isi dengan format berikut:

| Nama          | Email                     |
|---------------|---------------------------|
| Budi Santoso  | budi.santoso@bps.go.id    |
| Ani Rahayu    | ani.rahayu@bps.go.id      |

- Baris pertama = header (`Nama` dan `Email`)
- Email harus sama persis dengan username SSO BPS

---

## Cara menjalankan

### Pilihan A — Dashboard Web (direkomendasikan)

```
python web_fasih.py
```

Browser akan terbuka otomatis ke `http://localhost:5000`.

1. Klik **Browse** → pilih file `daftar_petugas.xlsx`
2. Isi **Username** dan **Password** SSO BPS
3. Klik **▶ Run**
4. Tunggu hingga selesai — log berjalan di layar

### Pilihan B — Aplikasi Desktop

```
python gui_fasih.py
```

Tampilan jendela akan muncul. Isi file, username, password, lalu klik **▶ Run**.

---

## Hasil

File Excel tersimpan otomatis di folder `output/`:

```
output/rekap_fasih_pascabayar_20250313_083000.xlsx
output/rekap_fasih_prabayar_20250313_083500.xlsx
```

Kolom: No · Nama · Email · Open · Submitted by Pencacah · Rejected by Admin Kabupaten

---

## Pertanyaan umum

**Browser terbuka tapi tidak ada yang terjadi?**
Tunggu, proses login SSO membutuhkan beberapa detik.

**Error "playwright not found"?**
Jalankan ulang: `pip install playwright` lalu `playwright install chromium`

**Email petugas tidak ditemukan di filter?**
Pastikan email di Excel sama persis dengan username SSO (termasuk huruf besar/kecil).
