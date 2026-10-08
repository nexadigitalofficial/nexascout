# -*- coding: utf-8 -*-
"""
NexaScout Universal Buyer Acquisition Orchestrator
Tek bir Sahibinden veya Coldwell Banker ilan linkini alarak:
1. İlanı canlı scrape eder
2. Özelliklerini ve lokasyonunu otomatik anlar
3. Haksız avantajları ve hedef alıcı sektörlerini eşleştirir
4. Karar verici C-Level teaserlarını, kurumsal e-postalarını ve meslektaş iş birliği mesajlarını üretir
5. Portföy çalışma alanını (Dossier & Metinler) oluşturup çıktıyı sunar
"""

import os
import re
import json
from typing import Dict, Any, Optional

from modules.portfolio_scraper import scrape_any_listing
from modules.property_analyzer import analyze_scraped_property
from modules.ai_pitch_crafter import craft_institution_pitch_suite

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
PORTFOLIOS_DIR = os.path.join(BASE_DIR, "data", "PORTFOYLER")

def run_portfolio_buyer_pipeline(
    url: str,
    broker_name: str = "YİĞİT NARİN",
    broker_phone: str = "0532 451 40 08",
    broker_office: str = "Coldwell Banker VIP Real Gayrimenkul"
) -> Dict[str, Any]:
    """
    Tek bir link ile tüm alıcı bulma sürecini otomatik yürütür.
    """
    # 1. Scrape
    scraped = scrape_any_listing(url)
    
    # 2. Analiz
    analysis = analyze_scraped_property(scraped)

    # 3. Klasörleme
    listing_id = scraped.get("listing_id") or "PORTFOY"
    slug_name = re.sub(r"[^a-zA-Z0-9_\-]", "_", scraped.get("title", "Portfoy")[:40])
    folder_name = f"{listing_id}_{slug_name}".strip("_")
    
    target_dir = os.path.join(PORTFOLIOS_DIR, folder_name)
    os.makedirs(target_dir, exist_ok=True)

    # 4. Avantajlar Metni
    advantages_bullet = "\n".join(f"• {adv}" for adv in analysis["advantages"])

    # 5. Meslektaş Duyuru Metni (Köşe parsel ve portföy avantajlarına göre dinamik)
    meslektas_metni = f"""🚨 DEĞERLİ MESLEKTAŞLARIM
📍 {analysis['location'].upper()}
⭐ {analysis['title'].upper()}
{analysis['asset_class'].upper()} • {analysis['kapali_m2']}

🔑 MÜLKÜN EN BÜYÜK AVANTAJLARI (ALICI KOZLARI):
{advantages_bullet}

🎯 HANGİ MÜŞTERİLERİNİZE SUNABİLİRSİNİZ?:
{chr(10).join(f"✔ {s['ad']} — {s['neden']}" for s in analysis['matched_sectors'])}

🤝 Meslektaş iş birliğine ve ortak çalışmaya tam açığız. Portföyünüzde bu profilde alıcı arayan müşteriniz varsa GETİRİN, BİRLİKTE GÖSTERELİM.

📞 {broker_name} — {broker_phone}
{broker_office}

Müşteriniz varsa, mülk hazır. 🔑
İlan: {url}
"""

    # 6. WhatsApp VIP Yönetici Teaser'ı
    wa_teaser = f"""Merhaba, iyi çalışmalar. Ben {broker_name}, {broker_office} Gayrimenkul ve Yatırım Danışmanıyım.

Kurumunuzun Ankara ve bölge büyüme hedefleri doğrultusunda değerlendirebileceğinizi düşündüğüm çok özel bir mülk için doğrudan şahsınıza ulaşmak istedim.

📍 {analysis['location']}
🏢 {analysis['asset_class']} | {analysis['kapali_m2']}
⭐ Öne Çıkan Değerler:
{chr(10).join(f"• {adv}" for adv in analysis['advantages'][:3])}

Kurumunuz için prestijli yeni hizmet noktası / karargâh olarak kullanıma hazır durumdadır.

İlgili yatırım veya gayrimenkul geliştirme biriminize yönlendirebilirseniz memnuniyet duyarım. Müsait olduğunuzda teknik şartname dosyasını paylaşmak ve yerinde göstermek isterim.

🔗 Resmi İlan Linki:
{url}

📞 {broker_name} — {broker_phone}
"""

    # 7. Kurumsal E-Posta Şablonu
    email_subject = f"Kurumsal Genişleme & Stratejik Yatırım Fırsatı | {analysis['location']}"
    email_body = f"""Sayın Yetkili,

Ben {broker_name}, {broker_office} Gayrimenkul ve Yatırım Danışmanıyım.

Kurumunuzun sektördeki büyüme vizyonunu ve hizmet ağını yakından takip etmekteyiz. Büyüme ve yeni yatırım planlarınız açısından yüksek katma değer sağlayacağını öngördüğümüz, bölgenin en stratejik noktasında yer alan özel bir portföyümüz için size doğrudan ulaşmaktayız.

MÜLK BİLGİLERİ & ÖNE ÇIKAN DEĞERLER:
• Lokasyon: {analysis['location']}
• Mülk Türü: {analysis['asset_class']} ({analysis['kapali_m2']})
{advantages_bullet}

Portföyümüz tek yetki sözleşmesiyle temsil edilmekte olup, kurumunuzun ilgili gayrimenkul geliştirme / yatırım birimine iletilmesini rica ederiz. Uygun göreceğiniz bir zaman diliminde teknik sunum dosyasını aktarmaktan ve yerinde inceleme organize etmekten memnuniyet duyarız.

Resmi İlan Detayı:
{url}

Saygılarımla,

{broker_name}
{broker_office}
İletişim: {broker_phone}
"""

    # 8. Görsel / Afiş Vitrin Metni
    gorsel_metin = f"""🚨 {analysis['location'].upper()}
⭐ {analysis['title'].upper()}

✔ {analysis['asset_class']} ({analysis['kapali_m2']})
{chr(10).join(f"✔ {adv}" for adv in analysis['advantages'][:3])}

💼 Tek Yetkili Danışman: {broker_name}
📱 {broker_phone} | {broker_office}
🔗 {url}
"""

    # 9. Dosyaları Kaydet
    with open(os.path.join(target_dir, "portfolio_data.json"), "w", encoding="utf-8") as f:
        json.dump({
            "scraped": scraped,
            "analysis": analysis,
            "broker": {"name": broker_name, "phone": broker_phone, "office": broker_office}
        }, f, ensure_ascii=False, indent=2)

    with open(os.path.join(target_dir, "MESLEKTAS_IS_BIRLIGI_DUYURUSU.txt"), "w", encoding="utf-8") as f:
        f.write(meslektas_metni)

    with open(os.path.join(target_dir, "VIP_WHATSAPP_TEASER.txt"), "w", encoding="utf-8") as f:
        f.write(wa_teaser)

    with open(os.path.join(target_dir, "KURUMSAL_EPOSTA.txt"), "w", encoding="utf-8") as f:
        f.write(f"KONU: {email_subject}\n\n{email_body}")

    with open(os.path.join(target_dir, "GORSEL_AFIS_METNI.txt"), "w", encoding="utf-8") as f:
        f.write(gorsel_metin)

    # 10. Stratejik Markdown Raporu
    rapor_md = f"""# 🏛️ STRATEJİK PORTFÖY ALICI BULMA VE YATIRIM RAPORU
> **Mülk:** {analysis['title']}  
> **Konum:** {analysis['location']}  
> **Sınıf:** {analysis['asset_class']} ({analysis['kapali_m2']})  
> **Tek Yetkili Danışman:** {broker_name} ({broker_phone}) — {broker_office}  
> **İlan Linki:** {url}  

---

## 📌 1. PORTFÖYÜN HAKSIZ AVANTAJLARI (UNFAIR ADVANTAGES)
{advantages_bullet}

---

## 🎯 2. TESPİT EDİLEN ÖNCELİKLİ ALICI SEKTÖRLERİ
| Sektör / Alıcı Profili | Eşleşme Gerekçesi |
| :--- | :--- |
{chr(10).join(f"| **{s['ad']}** | {s['neden']} |" for s in analysis['matched_sectors'])}

---

## 🔍 3. HARİTADAN HEDEFLENECEK KURUMSAL ARAMA SORGULARI
{chr(10).join(f"- `{q}`" for q in analysis['lead_search_queries'])}

---

## 💬 4. HAZIR MESAJLAŞMA PAKETİ
Tüm metinler `{target_dir}` klasöründe `.txt` olarak oluşturulmuştur.
- [MESLEKTAS_IS_BIRLIGI_DUYURUSU.txt](file:///{target_dir.replace(chr(92), '/')}/MESLEKTAS_IS_BIRLIGI_DUYURUSU.txt)
- [VIP_WHATSAPP_TEASER.txt](file:///{target_dir.replace(chr(92), '/')}/VIP_WHATSAPP_TEASER.txt)
- [KURUMSAL_EPOSTA.txt](file:///{target_dir.replace(chr(92), '/')}/KURUMSAL_EPOSTA.txt)
- [GORSEL_AFIS_METNI.txt](file:///{target_dir.replace(chr(92), '/')}/GORSEL_AFIS_METNI.txt)
"""
    with open(os.path.join(target_dir, "STRATEJIK_RAPOR.md"), "w", encoding="utf-8") as f:
        f.write(rapor_md)

    return {
        "status": "success",
        "url": url,
        "scraped": scraped,
        "analysis": analysis,
        "folder_path": target_dir,
        "outputs": {
            "meslektas_metni": meslektas_metni,
            "wa_teaser": wa_teaser,
            "email_subject": email_subject,
            "email_body": email_body,
            "gorsel_metin": gorsel_metin
        }
    }
