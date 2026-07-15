@echo off
setlocal EnableDelayedExpansion
cd /d %~dp0

echo ============================================
echo  FASIH Scraper (CDP) -- Build EXE
echo ============================================
echo.

:: -- Cek PyInstaller ---------------------------------------------------------
where pyinstaller >nul 2>&1
if errorlevel 1 (
    echo [ERROR] PyInstaller tidak ditemukan.
    echo Jalankan:  pip install pyinstaller
    pause & exit /b 1
)

:: -- Build EXE ----------------------------------------------------------------
echo [1/1] Membangun EXE dengan PyInstaller...
pyinstaller FASIH_Scraper.spec --clean --noconfirm
if errorlevel 1 (
    echo [ERROR] PyInstaller gagal.
    pause & exit /b 1
)

echo.
echo ============================================
echo  Distribusi siap di:
echo    dist\FASIH_Scraper\
echo.
echo  Struktur paket:
echo    FASIH_Scraper\
echo    +-- FASIH_Scraper.exe
echo    +-- _internal\
echo.
echo  Prasyarat di komputer target:
echo    - Google Chrome terpasang (mode CDP, tidak perlu Chromium bundled)
echo.
echo  Kirim seluruh folder dist\FASIH_Scraper\
echo  (bukan hanya file .exe-nya saja)
echo ============================================
echo.
pause
