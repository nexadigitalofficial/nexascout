@echo off
chcp 65001 >nul
title Google Maps B2B Kurumsal Veri Toplama ve CRM Sistemi - İncek Pazarlama

:MENU
cls
echo ===============================================================================
echo      COLDWELL BANKER VIP - İNCEK TİCARİ MÜLK B2B VERİ VE PAZARLAMA SİSTEMİ
echo             (LÖSANTE Karşısı Ada 119464 Parsel 13 - Tescilli Ticari YKB)
echo ===============================================================================
echo.
echo [1] 🌟 GÖRSEL WEB KONTROL PANELİNİ VE CRM'İ BAŞLAT (http://localhost:5000)
echo [2] 🗺️ İNTERAKTİF GIS HARİTASINI AÇ (Tarayıcıda Doğrudan Harita)
echo.
echo --- GOOGLE MAPS CANLI VERİ TOPLAMA ---
echo [3] TÜM ÖNERİLEN SEKTÖRLERİ TARA (Sağlık, Eğitim, Savunma, Hukuk, Diplomatik)
echo [4] YALNIZCA SAĞLIK & MEDİKAL SEKTÖRÜNÜ TARA (Tıp Mrk, Poliklinik, Diş, Estetik)
echo [5] YALNIZCA ÖZEL EĞİTİM & KOLEJLERİ TARA (Kolej, Anaokulu, Kampüsler)
echo [6] YALNIZCA SAVUNMA SANAYİİ & BİLİŞİM AR-GE ŞİRKETLERİNİ TARA
echo [7] YALNIZCA PRESTİJ HUKUK BÜROLARI VE BAĞIMSIZ DENETİM ŞİRKETLERİNİ TARA
echo [8] YALNIZCA DİPLOMATİK MİSYON VE ATAŞELİKLERİ TARA
echo [9] ÖZEL KELİME İLE ARA (Örn: 'fizik tedavi İncek', 'siber güvenlik')
echo.
echo --- OTOMATİK TEKLİF ÜRETİMİ ---
echo [10] 📄 TÜM KURUMLARA KİŞİSELLEŞTİRİLMİŞ WORD (.docx) TEKLİF MEKTUBU ÜRET
echo.
echo [11] ÇIKIŞ
echo ===============================================================================
set /p SECIM="Lütfen bir işlem seçiniz [1-11]: "

if "%SECIM%"=="1" goto RUN_WEB
if "%SECIM%"=="2" goto OPEN_MAP
if "%SECIM%"=="3" goto RUN_ALL
if "%SECIM%"=="4" goto RUN_SAGLIK
if "%SECIM%"=="5" goto RUN_EGITIM
if "%SECIM%"=="6" goto RUN_SAVUNMA
if "%SECIM%"=="7" goto RUN_HUKUK
if "%SECIM%"=="8" goto RUN_DIPLOMATIK
if "%SECIM%"=="9" goto RUN_CUSTOM
if "%SECIM%"=="10" goto RUN_DOCX
if "%SECIM%"=="11" goto CIKIS

echo Geçersiz seçim! Lütfen tekrar deneyiniz.
pause
goto MENU

:RUN_WEB
cls
echo [BİLGİ] Web Kontrol Paneli başlatılıyor...
echo [BİLGİ] Tarayıcınız http://localhost:5000 adresinde açılacaktır.
echo [NOT] Kapatmak için bu pencereyi kapatabilirsiniz.
cd /d "%~dp0"
python app.py
goto MENU

:OPEN_MAP
cls
echo [BİLGİ] İnteraktif GIS Haritası tarayıcıda açılıyor...
start "" "%~dp0..\INCEK_TICARI_HEDEF_HARITASI.html"
goto MENU

:RUN_ALL
cls
echo [BİLGİ] Tüm sektörler İncek merkezli taranıyor...
python "%~dp0google_maps_lead_harvester.py" --category all --max-places 6
goto SONUC

:RUN_SAGLIK
cls
echo [BİLGİ] Sağlık ve Medikal sektörü taranıyor...
python "%~dp0google_maps_lead_harvester.py" --category saglik --max-places 8
goto SONUC

:RUN_EGITIM
cls
echo [BİLGİ] Özel Eğitim Kurumları ve Kolejler taranıyor...
python "%~dp0google_maps_lead_harvester.py" --category egitim --max-places 8
goto SONUC

:RUN_SAVUNMA
cls
echo [BİLGİ] Savunma Sanayii ve Teknoloji AR-GE şirketleri taranıyor...
python "%~dp0google_maps_lead_harvester.py" --category savunma_teknoloji --max-places 8
goto SONUC

:RUN_HUKUK
cls
echo [BİLGİ] Prestij Hukuk Büroları ve YMM Denetim şirketleri taranıyor...
python "%~dp0google_maps_lead_harvester.py" --category hukuk_denetim --max-places 8
goto SONUC

:RUN_DIPLOMATIK
cls
echo [BİLGİ] Diplomatik Temsilcilikler ve Ataşelikler taranıyor...
python "%~dp0google_maps_lead_harvester.py" --category diplomatik --max-places 8
goto SONUC

:RUN_CUSTOM
cls
echo ===============================================================================
echo                           ÖZEL SORGULAMA MODU
echo ===============================================================================
echo.
set /p CUSTOM_QUERY="Aramak istediğiniz kelimeyi giriniz (örn: 'diş kliniği İncek'): "
if "%CUSTOM_QUERY%"=="" goto MENU
python "%~dp0google_maps_lead_harvester.py" --query "%CUSTOM_QUERY%" --max-places 10
goto SONUC

:RUN_DOCX
cls
echo [BİLGİ] Tüm kurumlar için kişiselleştirilmiş Word teklifleri üretiliyor...
python -c "import openpyxl; from modules.pitch_generator import generate_pitch_document; wb = openpyxl.load_workbook(r'%~dp0..\GOOGLE_HARITA_TOPLANAN_KURUMLAR.xlsx', data_only=True); ws = wb['TOPLANAN_KURUMLAR']; count=0; leads = [dict(kurum_adi=ws.cell(r,2).value, is_kolu=ws.cell(r,4).value, mesafe_km=ws.cell(r,11).value, hbu_model=ws.cell(r,13).value) for r in range(2, ws.max_row+1) if ws.cell(r,2).value]; [generate_pitch_document(l) for l in leads]; print(f'Toplam {len(leads)} adet Word (.docx) teklif mektubu basıldı!')"
echo.
echo [BAŞARILI] Teklif mektupları şu klasördedir:
echo C:\Users\USER\Desktop\SATTIM\PAZARLAMA\TEKLIF_MEKTUPLARI\
echo.
pause
goto MENU

:SONUC
echo.
echo ===============================================================================
echo [TAMAMLANDI] Veriler başarıyla toplandı, Excel ve İnteraktif Harita güncellendi!
echo Master Excel : C:\Users\USER\Desktop\SATTIM\PAZARLAMA\GOOGLE_HARITA_TOPLANAN_KURUMLAR.xlsx
echo GIS Haritası : C:\Users\USER\Desktop\SATTIM\PAZARLAMA\INCEK_TICARI_HEDEF_HARITASI.html
echo ===============================================================================
echo.
pause
goto MENU

:CIKIS
exit
