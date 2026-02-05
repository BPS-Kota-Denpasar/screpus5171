@echo off
setlocal

REM Pindah ke folder BAT ini (biar path file konsisten)
cd /d "%~dp0"

echo.
echo ==================================================
echo START SCRAPING
echo - Untuk stop aman: buat file STOP.txt di folder ini
echo - Atau (kalau jalan dari terminal): Ctrl+C
echo - Menutup Chrome juga akan stop + autosave
echo ==================================================
echo.

REM Hapus STOP.txt lama kalau ada (opsional, biar tidak langsung berhenti)
if exist "STOP.txt" del /q "STOP.txt" >nul 2>&1

REM Jalankan python (pakai py launcher)
py script.py

echo.
echo Selesai. Tekan tombol apa saja untuk keluar...
pause >nul
endlocal
