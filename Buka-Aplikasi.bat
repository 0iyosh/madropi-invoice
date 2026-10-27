@echo off
title MADROPI - Invoice Pengiriman Material
cd /d "%~dp0"
echo ========================================================
echo   MADROPI - Sistem Rekapitulasi & Invoice Pengiriman
echo ========================================================
echo.
echo Menjalankan aplikasi dan membuka browser...
echo (Jangan tutup jendela ini selama menggunakan aplikasi)
echo.
node server.js
pause
