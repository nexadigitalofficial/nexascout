# -*- coding: utf-8 -*-
"""
NexaScout Property Strategic Analyzer
Scrape edilen portföy verisini derinlemesine analiz eder:
1. Mülk tipini ve ölçeğini belirler (Ticari Kompleks, Villa, Plaza, vb.)
2. Haksız avantajları (Ruhsat, Köşe parsel, Asansör, Lokasyon, Otopark) filtreler
3. En yüksek satın alma iştahına sahip 3-5 hedef alıcı sektörünü tespit eder
4. Haritadan otomatik toplanacak Google Maps arama sorgularını ve yarıçapını oluşturur
5. Kurumsal teklif motoru (AI Pitch Crafter) için dinamik portföy kartı üretir
"""

import re
from typing import Dict, Any, List

def analyze_scraped_property(raw_portfolio: Dict[str, Any]) -> Dict[str, Any]:
    title = raw_portfolio.get("title", "")
    desc = raw_portfolio.get("description", "")
    specs = raw_portfolio.get("specs", {})
    location = raw_portfolio.get("location", "Ankara")
    url = raw_portfolio.get("url", "")
    
    full_text = f"{title} {desc} {' '.join(str(v) for v in specs.values())}".lower()

    # 1. Mülk Sınıfı & Türü
    asset_class = "Ticari Gayrimenkul"
    if any(w in full_text for w in ["bina", "komple bina", "plaz"]):
        asset_class = "Komple Müstakil Ticari Bina"
    elif any(w in full_text for w in ["villa", "müstakil"]):
        asset_class = "Müstakil Villa / Kurumsal Yönetim Merkezi"
    elif any(w in full_text for w in ["arsa", "tarla"]):
        asset_class = "Geliştirme / Ticari Arsa"
    elif any(w in full_text for w in ["dukkan", "dükkan", "magaza", "mağaza"]):
        asset_class = "Cadde Mağazası / Ticari Dükkan"
    elif any(w in full_text for w in ["ofis", "buro", "büro"]):
        asset_class = "Kurumsal Ofis / Kat"

    # 2. Metrekareler
    kapali_m2 = ""
    m2_matches = re.findall(r"(\d{2,5})\s*(?:m²|m2|metrekare)", full_text)
    if m2_matches:
        kapali_m2 = f"{m2_matches[0]} m²"
    for k, v in specs.items():
        if any(x in k.lower() for x in ["m²", "metrekare", "alan"]):
            kapali_m2 = str(v)
            break

    # 3. Haksız Avantajlar (Unfair Advantages & Hooks)
    advantages = []
    
    # Ruhsat / Yasal Tekel
    if any(w in full_text for w in ["ticari ruhsat", "çift ruhsat", "c278dfhu", "yapı kayıt", "isyeri ruhsat", "çalışma ruhsat"]):
        advantages.append("YASAL TEKEL: Bölgedeki tek TİCARİ ve çalışma ruhsatlı bağımsız yapı")
    elif "ticari" in full_text:
        advantages.append("TİCARİ KULLANIM: Şirket merkezi, klinik ve eğitim açılışına uygun altyapı")

    # Köşe parsel / Cephe
    if any(w in full_text for w in ["kose", "köşe", "çift cephe", "cift cephe"]):
        advantages.append("KÖŞE PARSEL: Çift cepheli kesintisiz tabela değeri, çift bağımsız giriş ve geniş otopark")
    elif "cadde" in full_text:
        advantages.append("CADDE ÜZERİ: Yüksek kurumsal görünürlük ve prestijli erişim aksı")

    # Lokasyon ve Hastane/Kurum Sinerjisi
    if any(w in full_text for w in ["losante", "lösante", "hastane", "sehir hastanesi"]):
        advantages.append("STRATEJİK SAĞLIK AKSI: LÖSANTE ve hastaneler bölgesinde hazır hasta ve ziyaretçi sinerjisi")
    elif any(w in full_text for w in ["incek", "cayyolu", "çayyolu", "beytepe", "bilkent"]):
        advantages.append("PRESTİJLİ LOKASYON: Ankara'nın en değerli kurumsal ve lüks yerleşim koridoru")

    # Asansör & Donatılar
    if any(w in full_text for w in ["asansor", "asansör"]):
        advantages.append("ASANSÖR ALTYAPISI: Katlar arası hazır asansör şaftı mimari boşluğu (sağlık ve VIP ofis standardı)")
    if any(w in full_text for w in ["otopark", "garaj"]):
        advantages.append("GENİŞ OTOPARK: Müstakil kapalı garaj ve misafirler için geniş açık otopark kapasitesi")
    if any(w in full_text for w in ["su deposu", "hidrofor"]):
        advantages.append("KESİNTİSİZ ALTYAPI: Yüksek kapasiteli su deposu ve endüstriyel hidrofor")

    if not advantages:
        advantages.append("Tam müstakil kullanım, özel bahçe ve prestijli mimari tasarım")

    # 4. Hedef Alıcı Personaları ve Sektörleri
    matched_sectors = []
    
    # Sağlık
    if any(w in full_text for w in ["losante", "lösante", "hastane", "klinik", "saglik", "sağlık", "ticari", "asansor", "asansör"]):
        matched_sectors.append({
            "kod": "SAGLIK",
            "ad": "Özel Hastane, Tıp Merkezi & VIP Diş/Estetik Klinikleri",
            "neden": "Bölgedeki tek ticari ruhsatlı müstakil yapı olması ve hazır asansör şaftı sayesinde Sağlık Bakanlığı ruhsatlandırmasına eksiksiz uyum.",
            "queries": ["özel tıp merkezi", "diş polikliniği", "estetik cerrahi kliniği", "fizik tedavi merkezi", "göz tıp merkezi"]
        })

    # Savunma / Teknoloji
    matched_sectors.append({
        "kod": "SAVUNMA_TEKNO",
        "ad": "Savunma Sanayii Tedarikçileri, Siber Güvenlik & Ar-Ge Karargahları",
        "neden": "Ankara Çevre Yolu ve ASELSAN aksına yakınlık, yüksek güvenlikli müstakil bahçe duvarları ve otopark rahatlığı.",
        "queries": ["savunma sanayi", "yazılım arge merkezi", "siber güvenlik firması", "mühendislik şirket merkezi"]
    })

    # Hukuk & Denetim
    matched_sectors.append({
        "kod": "HUKUK_DENETIM",
        "ad": "Büyük Hukuk Büroları, YMM & Bağımsız Denetim Ortaklıkları",
        "neden": "Çukurambar/GOP apartman otopark krizinden uzak, VIP müvekkil kabul salonları ve şömineli prestijli villa ofis konsepti.",
        "queries": ["avukatlık ortaklığı", "hukuk bürosu", "yeminli mali müşavirlik", "bağımsız denetim a.ş."]
    })

    # Eğitim
    if any(w in full_text for w in ["bahceli", "bahçeli", "müstakil", "bina", "ticari"]):
        matched_sectors.append({
            "kod": "EGITIM",
            "ad": "VIP Erken Çocukluk Akademisi, Butik Kolej & Dil Kampüsü",
            "neden": "MEB Standartlar Yönergesi'nin zorunlu kıldığı ticari ruhsat ve müstakil güvenli bahçe şartını tam karşılama.",
            "queries": ["butik kolej", "özel anaokulu", "montessori anaokulu", "yabancı dil akademisi"]
        })

    # Yatırım Fonları & Aile Ofisleri
    matched_sectors.append({
        "kod": "FON_YATIRIM",
        "ad": "Gayrimenkul Yatırım Fonları (GYF), Family Office & Portföy Şirketleri",
        "neden": "Bölge ortalamasının çok altında amortisman süresi, tek kiracılı yüksek kurumsal getiri ve enflasyona karşı yasal tekel koruması.",
        "queries": ["portföy yönetim şirketi", "holding genel merkezi", "yatırım holding", "gayrimenkul yatırım ortaklığı"]
    })

    # 5. Harita Arama Konfigürasyonu
    lat = raw_portfolio.get("lat", 39.845)
    lon = raw_portfolio.get("lon", 32.748)
    radius_km = 6.0

    return {
        "title": title,
        "asset_class": asset_class,
        "kapali_m2": kapali_m2,
        "location": location,
        "lat": lat,
        "lon": lon,
        "radius_km": radius_km,
        "advantages": advantages,
        "matched_sectors": matched_sectors,
        "lead_search_queries": [q for s in matched_sectors for q in s["queries"][:2]],
        "url": url,
        "photo_count": raw_portfolio.get("photo_count", 0),
        "primary_image": raw_portfolio.get("images", [""])[0] if raw_portfolio.get("images") else ""
    }
