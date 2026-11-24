@echo off
title PT. AKAR DAYA - Indihome by Telkomsel - Laporan Rekapitulasi Kinerja (TIKA)
echo ========================================================
echo   PT. AKAR DAYA - Indihome by Telkomsel
echo   Laporan Rekapitulasi Kinerja — Program Blue Device
echo   Rekapitulasi: Caring, Visit Lapangan & Pelanggan Setuju
echo   Uplift Agent: TIKA (Area Pamekasan)
echo   Leader / Supervisor: Bpk. Choirul Umam
echo ========================================================
echo.
echo Membuka dashboard laporan eksekutif...
echo Setiap kali file 'BLUE DEVICE TIKA.xlsx' di-update,
echo dashboard akan membaca data terbaru secara real-time.
echo.

python --version >nul 2>&1
if %ERRORLEVEL% EQU 0 (
    echo Menjalankan local server dan membuka browser...
    start "" http://localhost:8000
    python -m http.server 8000
) else (
    echo Membuka langsung di web browser default...
    start "" index.html
)

pause
