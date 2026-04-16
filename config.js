// ── Konfigurasi Dashboard FASIH ──────────────────────────────────────────────
// Tempel URL lengkap Google Sheets Anda di bawah.
// Contoh:
//   'https://docs.google.com/spreadsheets/d/1ABC.../edit'
//
// Setelah diedit, simpan file ini. Dashboard akan langsung menggunakan
// spreadsheet tersebut sebagai sumber data default saat pertama dibuka.
// ─────────────────────────────────────────────────────────────────────────────

const CONFIG = {
  // Milik BPS Kabupaten Buleleng
  DEFAULT_SHEET_URL: 'https://docs.google.com/spreadsheets/d/1JP6OcxJx6_thOEnHjy3ITZXpFRzHFvaX5zZzF1ZAjvs/edit',

  // ── Target Prabayar (opsional) ─────────────────────────────────────────
  // Isi angka berikut untuk menetapkan target default jumlah pelanggan prabayar.
  // Pengguna masih bisa mengubahnya dari sidebar (menu "Target Prabayar").
  // Jika sudah diubah via sidebar, nilai localStorage akan digunakan (tidak terpengaruh config ini).
  DEFAULT_TARGETS: {
    praba_total: 0,   // target keseluruhan (0 = belum ditetapkan)
    praba_ulp: {
      // Contoh: 'ULP SINGARAJA': 4500,
      //         'ULP SERIRIT':   2800,
    },
  },
};
