"""Ekstraksi rekap per-petugas FASIH lewat API JSON (anti salah-ambil).

Endpoint: POST /app/api/analytic/api/v2/assignment/report-progress-by-responsibility
Auth: cookie sesi (credentials include) — dipanggil dari dalam page CDP yang sudah login.
"""
import asyncio
import json
from playwright.async_api import async_playwright

CDP_URL = "http://127.0.0.1:9222"

API_ROOT = "https://fasih-sm.bps.go.id/app/api"
EP_REPORT = f"{API_ROOT}/analytic/api/v2/assignment/report-progress-by-responsibility"
EP_ROLES = f"{API_ROOT}/survey/api/v1/survey-roles"

# Status mapping -> kolom rekap
ST_OPEN = "OPEN"
ST_SUBMIT = "SUBMITTED BY Pencacah"
ST_REJECT = "REJECTED BY Admin Kabupaten"

# pascabayar
SURVEY_ID = "2e31188c-a617-4163-8056-edccf93d8d79"
SURVEY_PERIOD_ID = "d63e9832-13c6-4ec7-bf5b-59229c2f90f9"


DASBOR_URL = (
    "https://fasih-sm.bps.go.id/app/surveys/"
    f"{SURVEY_ID}/{SURVEY_PERIOD_ID}"
)


async def get_page(ctx):
    for pg in ctx.pages:
        if "fasih-sm.bps.go.id" in pg.url:
            return pg
    return ctx.pages[0] if ctx.pages else await ctx.new_page()


async def wait_rendered(page, timeout=25):
    """Tunggu app Angular ter-render (lewati skrip anti-bot)."""
    for _ in range(timeout):
        await page.wait_for_timeout(1000)
        ok = await page.evaluate(r"""() => { const t=(document.body.innerText||'').trim();
            return !t.startsWith('(function()') && !t.includes('btTu') && t.length>300; }""")
        if ok:
            return True
    return False


async def prepare_session(page, dasbor_url, tries=3):
    """Reload Dasbor + tunggu render → refresh token agar API mengembalikan JSON (bukan HTML).
    Ulangi beberapa kali kalau render gagal (skrip anti-bot kadang perlu fokus ulang)."""
    for attempt in range(tries):
        await page.bring_to_front()
        try:
            await page.goto(dasbor_url, wait_until="domcontentloaded", timeout=90000)
        except Exception as e:
            print(f"    [render retry {attempt+1}/{tries}] goto err: {e}")
            continue
        await page.bring_to_front()
        if await wait_rendered(page):
            return True
        print(f"    [render retry {attempt+1}/{tries}] belum render")
        await page.wait_for_timeout(1000)
    return False


async def api_post(page, url, body):
    """POST dari dalam page (cookie sesi) + header anti-CSRF X-XSRF-TOKEN dari cookie."""
    return await page.evaluate(
        """async ({url, body}) => {
            const headers = {'Content-Type':'application/json','Accept':'application/json'};
            const m = document.cookie.match(/XSRF-TOKEN=([^;]+)/);
            if (m) headers['X-XSRF-TOKEN'] = decodeURIComponent(m[1]);
            const r = await fetch(url, {
                method: 'POST', headers, credentials: 'include',
                body: JSON.stringify(body),
            });
            let j = null; try { j = await r.json(); } catch(e) {}
            return { status: r.status, json: j };
        }""",
        {"url": url, "body": body},
    )


async def api_get(page, url):
    return await page.evaluate(
        """async (url) => {
            const r = await fetch(url, {headers:{'Accept':'application/json'}, credentials:'include'});
            let j = null; try { j = await r.json(); } catch(e) {}
            return { status: r.status, json: j };
        }""",
        url,
    )


async def get_pencacah_role_id(page, survey_id):
    res = await api_get(page, f"{EP_ROLES}?surveyId={survey_id}")
    for role in (res.get("json") or {}).get("data", []):
        if role.get("isPencacah"):
            return role["id"]
    return None


def _aggregate_user(u):
    """Jumlahkan statusBreakdown semua region -> open/submitted/rejected."""
    agg = {}
    for rs in (u.get("regionSummary") or []):
        for sb in (rs.get("statusBreakdown") or []):
            agg[sb["status"]] = agg.get(sb["status"], 0) + sb.get("count", 0)
    return {
        "open": agg.get(ST_OPEN, 0),
        "submitted": agg.get(ST_SUBMIT, 0),
        "rejected": agg.get(ST_REJECT, 0),
        "total": u.get("total", 0),
        "_all_status": agg,
    }


def _body(period_id, role_id, size, page, level):
    return {
        "surveyPeriodId": period_id,
        "surveyRoleId": role_id,
        "size": size,
        "page": page,
        "search": "",
        "target": "TARGET_ONLY",
        "region": {f"region{i}Id": None for i in range(1, 11)},
        "regionSummaryLevel": level,
    }


async def post_with_retry(page, url, body, dasbor_url=None, tries=4, delay_ms=1000):
    """POST tahan-banting: ulangi saat 504/HTML/empty (server timeout / token basi),
    re-render Dasbor di antara percobaan untuk refresh token."""
    res = {"status": 0, "json": None}
    for attempt in range(tries):
        res = await api_post(page, url, body)
        if res["status"] == 200 and res.get("json"):
            return res
        print(f"    [api retry {attempt+1}/{tries}] status={res['status']} → tunggu & ulang")
        await page.wait_for_timeout(delay_ms)
        if dasbor_url:
            await prepare_session(page, dasbor_url, tries=2)
    return res


async def detect_region_level(page, period_id, role_id, dasbor_url=None):
    """regionSummaryLevel beda tiap survei (kedalaman wilayah beda). Pilih level TERENDAH
    di mana jumlah semua status == total user (data lengkap; level rendah = lebih cepat)."""
    for lvl in (1, 2, 3, 4, 5, 6):
        res = await post_with_retry(page, EP_REPORT, _body(period_id, role_id, 10, 0, lvl), dasbor_url)
        content = (res.get("json") or {}).get("data", {}).get("content", [])
        for u in content:
            if (u.get("regionSummary") or []) and u.get("total"):
                agg = _aggregate_user(u)
                if sum(agg["_all_status"].values()) == u.get("total"):
                    return lvl
    return 1


async def fetch_rekap_petugas(page, survey_period_id, survey_role_id, dasbor_url=None,
                              page_size=10, region_level=None):
    """Ambil SEMUA petugas (paginasi) -> (dict {email: {...}}, ok: bool).
    ok=False bila ada halaman yang gagal setelah retry (data tidak lengkap)."""
    if region_level is None:
        region_level = await detect_region_level(page, survey_period_id, survey_role_id, dasbor_url)
        print(f"  [LEVEL] regionSummaryLevel terpilih = {region_level}")
    out = {}
    ok = True
    pageno = 0
    while True:
        body = _body(survey_period_id, survey_role_id, page_size, pageno, region_level)
        res = await post_with_retry(page, EP_REPORT, body, dasbor_url)
        if res["status"] != 200 or not res.get("json"):
            print(f"[ERR] halaman {pageno} gagal (status={res['status']}) — data TIDAK lengkap")
            ok = False
            break
        await page.wait_for_timeout(150)  # jeda sopan antar-halaman
        data = res["json"].get("data", {})
        content = data.get("content", [])
        for u in content:
            agg = _aggregate_user(u)
            key = u.get("email") or u.get("username")
            if key in out:
                # email dobel (mis. >1 alokasi) → jumlahkan, jangan timpa
                for f in ("open", "submitted", "rejected", "total"):
                    out[key][f] += agg[f]
            else:
                out[key] = {
                    "nama": u.get("fullname") or u.get("username", ""),
                    "email": u.get("email") or u.get("username", ""),
                    "open": agg["open"], "submitted": agg["submitted"],
                    "rejected": agg["rejected"], "total": agg["total"],
                }
        if data.get("last") or not content:
            break
        pageno += 1
    return out, ok


async def main():
    async with async_playwright() as p:
        browser = await p.chromium.connect_over_cdp(CDP_URL)
        ctx = browser.contexts[0]
        page = await get_page(ctx)

        rendered = await prepare_session(page, DASBOR_URL)
        print(f"[SESSION] dasbor rendered={rendered}")

        role_id = await get_pencacah_role_id(page, SURVEY_ID)
        if not role_id:
            role_id = "71bf417e-d2db-48a0-baa1-f9b31d878ace"  # fallback Pencacah (pascabayar)
            print("[ROLE] pakai fallback role id")
        print(f"[ROLE] Pencacah surveyRoleId = {role_id}")

        rekap, ok = await fetch_rekap_petugas(page, SURVEY_PERIOD_ID, role_id, DASBOR_URL)
        print(f"[OK] {len(rekap)} petugas terambil dari API (lengkap={ok})\n")

        print(f"{'Email':<40} {'Open':>6} {'Submit':>7} {'Reject':>7} {'Total':>7}")
        print("-" * 72)
        for email, r in sorted(rekap.items()):
            print(f"{email:<40} {r['open']:>6} {r['submitted']:>7} {r['rejected']:>7} {r['total']:>7}")

        # cross-check vs UI lama
        print("\n=== CROSS-CHECK (vs UI lama) ===")
        for email in ["arissupartha1@gmail.com", "arifsatria1206@gmail.com", "assalamagus801@gmail.com"]:
            r = rekap.get(email)
            print(f"  {email}: {r}" if r else f"  {email}: TIDAK ADA di API")

        with open("output/rekap_via_api.json", "w", encoding="utf-8") as f:
            json.dump(rekap, f, ensure_ascii=False, indent=2)
        print("\n[SAVE] output/rekap_via_api.json")


if __name__ == "__main__":
    asyncio.run(main())
