@echo off
REM ── Buka Chrome dengan port debug untuk scraper FASIH (mode CDP) ──
REM Dobel-klik file ini, lalu LOGIN FASIH manual di jendela yang terbuka.
REM Setelah login, jalankan:  python run_api.py

set "CHROME=C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"
if not exist "%CHROME%" set "CHROME=C:\Program Files\Google\Chrome\Application\chrome.exe"

set "PROFILE=%LOCALAPPDATA%\Temp\claude\fasih_cdp_profile"

start "" "%CHROME%" --remote-debugging-port=9222 --user-data-dir="%PROFILE%" "https://fasih-sm.bps.go.id/app/surveys"

echo.
echo Chrome debug dibuka (port 9222).
echo 1) LOGIN FASIH manual di jendela Chrome yang muncul.
echo 2) Setelah masuk, jalankan:  python run_api.py
echo.
