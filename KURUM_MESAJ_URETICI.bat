@echo off
chcp 65001 > nul
title NexaScout - AI Kurum Analizi ve Mesaj Uretici
cd /d "%~dp0"
echo ========================================================
echo   NexaScout AI Kurum Analizi ve Mesaj Uretici
echo   Danisman: Yigit Narin (0532 451 40 08)
echo ========================================================
echo.
python KURUM_MESAJ_URETICI.py
pause
