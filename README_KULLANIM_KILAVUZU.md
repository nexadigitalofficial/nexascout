# COLDWELL BANKER VIP - İNCEK TİCARİ MÜLK B2B PAZARLAMA VE HARİTA CRM OTOMASYONU
## (Ankara İncek / LÖSANTE Karşısı Ada 119464 Parsel 13 - Tescilli Ticari YKB C278DFHU)

Bu sistem, portföyünüzün **"bölgedeki tek tescilli ticari ruhsatlı mülk"** ve **"LÖSANTE karşısı tekel konumu"** gücünü maksimum satış hızına ve en yüksek fiyata dönüştürmek için tasarlanmış **uçtan uca kurumsal bir B2B Pazarlama, Veri Madenciliği ve Harita İstihbarat Platformudur.**

---

### 🌟 GELİŞTİRİLEN YENİ SÜPER YETENEKLER VE MODÜLLER

```text
C:\Users\USER\Desktop\SATTIM\PAZARLAMA\
│
├── PANEL_BASLAT.bat                    <-- (Çift tıklayınca Web Kontrol Panelini tarayıcınızda açar)
├── INCEK_TICARI_HEDEF_HARITASI.html    <-- (İnteraktif GIS Haritası: Yarıçap çemberleri, renkli pinler)
├── GOOGLE_HARITA_TOPLANAN_KURUMLAR.xlsx<-- (40+ Kurumun 16 sütunlu Master Excel veri tabanı)
├── STRATEJIK_PAZARLAMA_VE_HEDEF_MUSTERI_LISTESI.xlsx <-- (50 VIP kurumsal yatırımcı CRM dosyası)
│
├── TEKLIF_MEKTUPLARI\                  <-- (Her kuruma özel üretilmiş Word (.docx) yatırım teklifleri)
│   ├── Denthaus_Ağız_ve_Diş_VIP_Yatirim_Teklifi.docx
│   ├── DENT_İNCEK_AĞIZ_VE_DİŞ_VIP_Yatirim_Teklifi.docx
│   ├── Maya_Okulları_İncek_VIP_Yatirim_Teklifi.docx
│   ├── Özel_Yeni_Med_Tıp_Merkezi_VIP_Yatirim_Teklifi.docx
│   └── ... (Tüm kurumlar için hazır)
│
└── otomasyon\
    ├── app.py                          <-- (Flask tabanlı modern yerel Web Dashboard sunucusu)
    ├── google_maps_lead_harvester.py   <-- (Playwright Google Maps tarayıcı motoru)
    ├── VERI_TOPLAMA_BASLAT.bat         <-- (11 seçenekli interaktif konsol menüsü)
    ├── config.json                     <-- (Koordinat, yarıçap ve anahtar kelime yapılandırması)
    └── modules\
        ├── enrichment.py               <-- (Derin E-posta, Sosyal Medya ve WhatsApp formatlayıcı)
        ├── map_generator.py            <-- (Folium interaktif GIS harita motoru)
        └── pitch_generator.py          <-- (Yapay zeka şablonlu Word (.docx) mektup üreticisi)
```

---

### 🚀 1. GÖRSEL WEB KONTROL PANELİ (WEB DASHBOARD)

Masaüstünüzdeki **[`PANEL_BASLAT.bat`](file:///C:/Users/USER/Desktop/SATTIM/PAZARLAMA/PANEL_BASLAT.bat)** dosyasına çift tıkladığınızda:
1. Yerel Flask sunucusu başlar (`http://localhost:5000`).
2. Varsayılan internet tarayıcınız (Chrome, Edge vb.) otomatik olarak kontrol panelini açar.

#### Web Panelinde Neler Yapabilirsiniz?
* **📊 Canlı KPI Metrikleri:** Toplam taranan kurum sayısı, Sağlık/Medikal adedi, Kolej/Eğitim adedi, A+ Öncelikli kurumlar ve mülkümüze ortalama mesafe canlı gösterilir.
* **🔎 Gelişmiş Arama & Filtreleme:** DataTables tablosunda tek tıkla sektöre, mesafeye veya kurum adına göre anında arama yapabilirsiniz.
* **💬 Doğrudan WhatsApp Mesajı:** Her satırda bulunan yeşil WhatsApp butonuna tıkladığınızda, o kurumun yetkilisine özel hazırlanmış yatırım mesajı ile WhatsApp Web anında açılır.
* **📄 Kişiselleştirilmiş Word (.docx) Teklifi İndir:** Tek tıkla o kurumun mesafesine ve sektörüne göre yazılmış hazır Word dosyasını indirebilirsiniz.
* **🗺️ Entegre Harita Görüntüleyici:** Harita sekmesinde İncek bölgesindeki tüm kurumların pinlerini, yarıçap çemberlerini ve mülkümüzün altın yıldızını tek ekranda inceleyebilirsiniz.
* **🚀 Yeni Tarama Tetikleme:** Web arayüzündeki form üzerinden istediğiniz yeni bir sektörü veya özel anahtar kelimeyi seçip taramayı webden başlatabilirsiniz.

---

### 🗺️ 2. İNTERAKTİF GIS HARİTA MOTORU (`INCEK_TICARI_HEDEF_HARITASI.html`)

Harita dosyasını tarayıcınızda açtığınızda:
* **⭐ Mülkümüzün Konumu (Altın Yıldız):** LÖSANTE Hastanesi tam karşısında, Ada 119464 Parsel 13 üzerinde kırmızı/altın özel yıldız pini. Tıklandığında mülkün 420 m² kapalı alanı, ticari Yapı Kayıt Belgesi (C278DFHU), 76.5M - 78M TL satış fiyatı ve Yiğit Narin'in iletişim bilgileri açılır.
* **⭕ Yarıçap Etki Çemberleri:**
  * **1 km:** Yürüme mesafesi kuşağı (En yakın sağlık ve eğitim birimleri)
  * **3 km:** Doğrudan İncek Sağlık ve Kolejler aksı
  * **5 km:** Beytepe ve Çayyolu entegrasyon bandı
  * **10 km:** Ankara Çevre Yolu geniş etki alanı
* **🎨 Sektörel Renk Kodlaması:**
  * 🔴 **Kırmızı:** Tıp Merkezleri, Poliklinikler, Hastaneler
  * 🔵 **Mavi:** Ağız ve Diş Sağlığı Polikliniği
  * 🟣 **Mor:** Plastik Cerrahi ve Medikal Estetik
  * 🟠 **Turuncu:** Fizik Tedavi ve Rehabilitasyon
  * 🟢 **Yeşil:** Özel Kolejler ve Anaokulları
  * 🛡️ **Koyu Mavi:** Savunma Sanayii & Teknoloji AR-GE
  * ⚖️ **Koyu Kırmızı:** Prestij Hukuk Büroları
* **Katman Kontrolü (Layer Control):** Sağ üst köşedeki menüden dilediğiniz sektörü açıp kapatabilirsiniz.
* **İşletme Kartları:** Her pine tıklandığında mülke olan mesafe, telefon, web sitesi ve **"WhatsApp Mesajı Başlat"** butonu yer alır.

---

### 📄 3. OTOMATİK KİŞİSELLEŞTİRİLMİŞ WORD (.docx) YATIRIM TEKLİFLERİ

[`TEKLIF_MEKTUPLARI`](file:///C:/Users/USER/Desktop/SATTIM/PAZARLAMA/TEKLIF_MEKTUPLARI/) klasöründe, taranan 40 kurumun her biri için ayrı ayrı üretilmiş Word belgeleri bulunmaktadır.

Mektupların içeriği şablondan ibaret değildir; **o kuruma özel akıllı mantıkla yazılmıştır:**
* **Kurumun Mevcut Konumuna Olan Mesafesi:** *"Mevcut yerleşkenize yalnızca 1.36 km mesafede..."*
* **Sektöre Özel Çözüm Sunumu:**
  * **Diş Kliniği ise:** 4 kata yayılan ıslak hacimler, asansör şaftı, her odaya ünit kurulumu, çift giriş ile hasta/personel ayrımı vurgulanır.
  * **Tıp Merkezi/Cerrahi ise:** LÖSANTE ve eczaneler aksı, tescilli ticari YKB bürokratik muafiyeti, 10 ton su deposu ve kesintisiz elektrik hattı anlatılır.
  * **Kolej/Okul ise:** 398 m² bahçe alanı, 4 katlı serbest mimari, kış bahçesi botanik atölyesi ve servis araçlarına uygun geniş cadde profili vurgulanır.
  * **Savunma/Teknoloji ise:** O-20 Çevre Yolu'na 1 km bağlantı, yüksek güvenlikli istinat duvarları, bodrum kat bağımsız veri merkezi/server altyapısı ve 15 araçlık otopark vurgulanır.
* **Kurumsal Şartlar & İlan Bağlantısı:** Resmi Sahibinden ilan bağlantısı ve yetkili danışman Yiğit Narin (0532 451 40 08) iletişim detayları resmi dille aktarılır.
* **İmza:** Yiğit Narin | Coldwell Banker VIP Real Gayrimenkul A.Ş.

---

### 🔧 4. KONSOL MENÜSÜ İLE KULLANIM (`VERI_TOPLAMA_BASLAT.bat`)

Dilerseniz komut satırı menüsünden de tüm işlemleri yönetebilirsiniz:
* `[1]` Web Kontrol Panelini başlatır (`http://localhost:5000`)
* `[2]` İnteraktif GIS Haritasını doğrudan tarayıcıda açar
* `[3] - [8]` Sektör bazlı canlı tarama yapar
* `[9]` Özel arama kelimesiyle tarama yapar
* `[10]` Tüm taranan kurumlara tek tuşla Word (.docx) teklif mektubu üretir
