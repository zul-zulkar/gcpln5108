"""Pipeline final FASIH: API (UI baru) -> Excel + Google Sheets.

Prasyarat: Chrome debug (port 9222) sudah dibuka & login manual ke FASIH.
Pakai:
  python run_api.py            # pascabayar + prabayar -> Excel + Sheets
  python run_api.py --no-sheets  # Excel saja (tidak kirim Sheets)
"""
import sys
import asyncio
from playwright.async_api import async_playwright

import scrape_fasih as sf
import fasih_api as fa

CDP_URL = "http://127.0.0.1:9222"

# surveyId (sama dgn URL collect lama). period_id None = ditemukan otomatis.
SURVEYS_NEW = {
    "pascabayar": {"survey_id": "2e31188c-a617-4163-8056-edccf93d8d79",
                   "period_id": "d63e9832-13c6-4ec7-bf5b-59229c2f90f9"},
    "prabayar":   {"survey_id": "2395b67d-d1af-4739-9ef8-c0cc0aa9ce9a",
                   "period_id": "16acea4e-4710-43d1-8b00-eeee589c8b66"},
}


def _norm(email):
    return (email or "").strip().lower()


async def discover_period_id(page, survey_id, tries=3):
    """Temukan surveyPeriodId aktif dari link Dasbor di halaman survei (cadangan)."""
    url = f"https://fasih-sm.bps.go.id/app/surveys/{survey_id}"
    for _ in range(tries):
        if not await fa.prepare_session(page, url, tries=2):
            continue
        pid = await page.evaluate(
            """(sid) => {
                const re = new RegExp('/app/surveys/' + sid + '/([0-9a-f]{8}-[0-9a-f-]{27})');
                for (const a of document.querySelectorAll('a[href]')) {
                    const m = (a.getAttribute('href') || '').match(re);
                    if (m) return m[1];
                }
                return null;
            }""",
            survey_id,
        )
        if pid:
            return pid
        await page.wait_for_timeout(1000)
    return None


def build_results(petugas_list, rekap_by_email):
    """Susun hasil per survei mengikuti daftar petugas input (cocokkan email)."""
    results, hit = [], 0
    for pt in petugas_list:
        r = rekap_by_email.get(_norm(pt["email"]))
        if r:
            hit += 1
            results.append({"nama": pt["nama"], "email": pt["email"],
                            "open": r["open"], "submitted": r["submitted"], "rejected": r["rejected"]})
        else:
            results.append({"nama": pt["nama"], "email": pt["email"],
                            "open": 0, "submitted": 0, "rejected": 0})
    return results, hit


async def run_pipeline(input_file=None, sheets_url=None, send_sheets=True,
                       stop_event=None, cdp_url=CDP_URL):
    """Pipeline inti (bisa dipanggil dari CLI maupun GUI).
    Mengembalikan (all_results, quality)."""
    petugas_list = sf.read_petugas(input_file) if input_file else sf.read_petugas()
    print(f"[INFO] {len(petugas_list)} petugas input")

    async with async_playwright() as p:
        try:
            browser = await p.chromium.connect_over_cdp(cdp_url)
        except Exception as e:
            print(f"[ERROR] Tidak bisa connect ke Chrome debug ({cdp_url}): {e}")
            print("        Buka Chrome debug & login dulu (tombol 'Buka Chrome & Login').")
            return {}, {}
        ctx = browser.contexts[0]
        page = await fa.get_page(ctx)

        all_results = {}
        quality = {}  # name -> True/False (data lengkap & layak kirim)
        for name, cfg in SURVEYS_NEW.items():
            if stop_event and stop_event.is_set():
                print("[STOP] Dihentikan pengguna.")
                break
            survey_id = cfg["survey_id"]
            period_id = cfg["period_id"]
            print(f"\n[SURVEY] {name.upper()} (surveyId={survey_id})")

            if not period_id:
                period_id = await discover_period_id(page, survey_id)
                print(f"  [DISCOVER] surveyPeriodId = {period_id}")
                if not period_id:
                    print(f"  [SKIP] {name}: tidak menemukan periode (akses?).")
                    all_results[name] = []
                    quality[name] = False
                    continue

            dasbor = f"https://fasih-sm.bps.go.id/app/surveys/{survey_id}/{period_id}"
            rendered = await fa.prepare_session(page, dasbor)
            print(f"  [SESSION] dasbor rendered={rendered}")

            role_id = await fa.get_pencacah_role_id(page, survey_id)
            print(f"  [ROLE] Pencacah roleId = {role_id}")
            if not role_id:
                print(f"  [SKIP] {name}: roleId Pencacah tidak ditemukan.")
                all_results[name] = []
                quality[name] = False
                continue

            rekap, ok = await fa.fetch_rekap_petugas(page, period_id, role_id, dasbor)
            rekap_by_email = {_norm(e): v for e, v in rekap.items()}
            print(f"  [API] {len(rekap)} pencacah terambil (lengkap={ok})")

            results, hit = build_results(petugas_list, rekap_by_email)
            print(f"  [MATCH] {hit}/{len(petugas_list)} petugas input cocok di API")
            all_results[name] = results
            # layak kirim bila: pengambilan lengkap, ada data, & semua input cocok
            quality[name] = bool(ok and len(rekap) > 0 and hit == len(petugas_list))

        print("\n[INFO] selesai ambil data (CDP dibiarkan terbuka).")

    # ── Output: Excel selalu disimpan ──
    out_path = sf.save_combined_results(all_results)
    print(f"[EXCEL] {out_path}")

    # ── Gerbang kualitas sebelum kirim ke Sheets ──
    all_good = all(quality.get(n) for n in SURVEYS_NEW)
    print(f"[QUALITY] {quality}  -> layak kirim Sheets: {all_good}")

    if not send_sheets:
        print("[SHEETS] dilewati.")
    elif not all_good:
        bad = [n for n in SURVEYS_NEW if not quality.get(n)]
        print(f"[SHEETS] DIBATALKAN — data tidak lengkap untuk: {bad}. "
              f"Excel tetap tersimpan. Jalankan ulang setelah koneksi stabil.")
    elif not all_results:
        print("[SHEETS] dilewati — tidak ada data (koneksi gagal?).")
    else:
        sheets = sheets_url or sf.SHEETS_WEBHOOK_URL
        sf.append_daily_snapshot(all_results, sheets)
        sf.append_detail_snapshot(all_results, sheets)
        print("[SHEETS] terkirim.")

    return all_results, quality


async def main(send_sheets=True):
    await run_pipeline(send_sheets=send_sheets)


if __name__ == "__main__":
    send = "--no-sheets" not in sys.argv
    asyncio.run(main(send_sheets=send))
