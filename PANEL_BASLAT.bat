@echo off
chcp 65001 >nul
title COLDWELL BANKER VIP - İncek B2B Pazarlama ve Harita Kontrol Paneli

echo ===============================================================================
echo      COLDWELL BANKER VIP - İNCEK TİCARİ MÜLK B2B PAZARLAMA KONTROL PANELİ
echo ===============================================================================
echo.
echo [BİLGİ] Web arayüzü başlatılıyor...
echo [BİLGİ] Web tarayıcınız otomatik olarak http://localhost:5000 adresini açacaktır.
echo.
echo [NOT] Programı kapatmak için bu pencereyi kapatabilir veya CTRL+C tuşlayabilirsiniz.
echo ===============================================================================
echo.

cd /d "%~dp0"
python app.py

pause
