# -*- coding: utf-8 -*-
"""
NexaScout Universal Portfolio Scraper
Desteklenen Kaynaklar:
  • sahibinden.com (PageSpeed Insights API + Akıllı Slug Fallback)
  • cb.com.tr / Coldwell Banker (Doğrudan BS4 Detay Scraper + PSI Fallback)
  • hepsiemlak.com, zingat.com, emlakjet.com ve genel ilan sayfaları

Tek bir ilan linkinden m², kat, arsa, ruhsat/imar, fiyat, adres, koordinat ve fotoğrafları çeker.
"""

import os
import re
import time
import json
import urllib.parse
import html as html_mod
import requests
from bs4 import BeautifulSoup
from typing import Optional, Dict, Any, Tuple

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
    ),
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8",
    "Accept-Language": "tr-TR,tr;q=0.9,en;q=0.5",
    "DNT": "1",
}

SCRAPE_TIMEOUT = 12
_coord_cache: Dict[str, Tuple[float, float]] = {}
_last_nominatim_call: float = 0.0

TURKISH_CITIES = {
    'adana': 'Adana', 'adiyaman': 'Adıyaman', 'afyon': 'Afyonkarahisar', 'agri': 'Ağrı',
    'amasya': 'Amasya', 'ankara': 'Ankara', 'antalya': 'Antalya', 'artvin': 'Artvin',
    'aydin': 'Aydın', 'balikesir': 'Balıkesir', 'bilecik': 'Bilecik', 'bingol': 'Bingöl',
    'bitlis': 'Bitlis', 'bolu': 'Bolu', 'burdur': 'Burdur', 'bursa': 'Bursa',
    'canakkale': 'Çanakkale', 'cankiri': 'Çankırı', 'corum': 'Çorum', 'denizli': 'Denizli',
    'diyarbakir': 'Diyarbakır', 'edirne': 'Edirne', 'elazig': 'Elazığ', 'erzincan': 'Erzincan',
    'erzurum': 'Erzurum', 'eskisehir': 'Eskişehir', 'gaziantep': 'Gaziantep', 'giresun': 'Giresun',
    'gumushane': 'Gümüşhane', 'hakkari': 'Hakkari', 'hatay': 'Hatay', 'isparta': 'Isparta',
    'mersin': 'Mersin', 'istanbul': 'İstanbul', 'izmir': 'İzmir', 'kars': 'Kars',
    'kastamonu': 'Kastamonu', 'kayseri': 'Kayseri', 'kirklareli': 'Kırklareli', 'kirsehir': 'Kırşehir',
    'kocaeli': 'Kocaeli', 'konya': 'Konya', 'kutahya': 'Kütahya', 'malatya': 'Malatya',
    'manisa': 'Manisa', 'kahramanmaras': 'Kahramanmaraş', 'mardin': 'Mardin', 'mugla': 'Muğla',
    'mus': 'Muş', 'nevsehir': 'Nevşehir', 'nigde': 'Niğde', 'ordu': 'Ordu',
    'rize': 'Rize', 'sakarya': 'Sakarya', 'samsun': 'Samsun', 'siirt': 'Siirt',
    'sinop': 'Sinop', 'sivas': 'Sivas', 'tekirdag': 'Tekirdağ', 'tokat': 'Tokat',
    'trabzon': 'Trabzon', 'tunceli': 'Tunceli', 'sanliurfa': 'Şanlıurfa', 'usak': 'Uşak',
    'van': 'Van', 'yozgat': 'Yozgat', 'zonguldak': 'Zonguldak', 'aksaray': 'Aksaray',
    'bayburt': 'Bayburt', 'karaman': 'Karaman', 'kirikkale': 'Kırıkkale', 'batman': 'Batman',
    'sirnak': 'Şırnak', 'bartin': 'Bartın', 'ardahan': 'Ardahan', 'igdir': 'Iğdır',
    'yalova': 'Yalova', 'karabuk': 'Karabük', 'kilis': 'Kilis', 'osmaniye': 'Osmaniye',
    'duzce': 'Düzce'
}

WORD_MAP = {
    'satilik': 'Satılık', 'kiralik': 'Kiralık', 'devren': 'Devren',
    'bina': 'Bina', 'daire': 'Daire', 'villa': 'Villa', 'arsa': 'Arsa',
    'rezidans': 'Rezidans', 'isyeri': 'İş Yeri', 'dukkan': 'Dükkan',
    'komple': 'Komple', 'sifir': 'Sıfır', 'tadilatli': 'Tadilatlı',
    'merkez': 'Merkez', 'cankaya': 'Çankaya', 'golbasi': 'Gölbaşı', 'incek': 'İncek',
    'kizilcasar': 'Kızılcaşar', 'cayyolu': 'Çayyolu', 'umitkoy': 'Ümitköy',
    'beytepe': 'Beytepe', 'bilkent': 'Bilkent', 'losante': 'LÖSANTE Karşısı',
    'kose': 'Köşe', 'parsel': 'Parsel', 'ticari': 'Ticari', 'ruhsatli': 'Ruhsatlı',
    'bahceli': 'Bahçeli', 'havuzlu': 'Havuzlu', 'esyali': 'Eşyalı', 'luks': 'Lüks'
}
WORD_MAP.update(TURKISH_CITIES)

_PSI_EP_MAP = {
    "ep.content_group":     "Sayfa Türü",
    "ep.kategori_1":        "Kategori 1",
    "ep.kategori_2":        "Kategori 2",
    "ep.kategori_3":        "Kategori 3",
    "ep.CD_ilanNo":         "İlan No",
    "ep.ilan_no":           "İlan No",
    "ep.js_price":          "Fiyat (Sayısal)",
    "ep.CD_IlanOwnerType":  "Satıcı Tipi",
    "ep.CD_Yer1":           "Ülke",
    "ep.CD_Yer2":           "Şehir",
    "ep.CD_Yer3":           "İlçe",
    "ep.CD_Yer4":           "Mahalle",
    "ep.CD_Yer5":           "Mahalle (detay)",
    "ep.yer_2":             "Şehir",
    "ep.yer_3":             "İlçe",
    "ep.yer_4":             "Mahalle",
}


def detect_listing_id(url: str) -> str:
    """URL içinden ilan ID'sini yakalar."""
    m = re.search(r"/(\d{8,12})(?:/|\?|$)", url)
    if m:
        return m.group(1)
    m = re.search(r"-(\d{8,12})$", url)
    if m:
        return m.group(1)
    return ""


def geocode_address(query: str) -> Optional[Tuple[float, float]]:
    """Adresi OpenStreetMap Nominatim servisi ile enlem/boylama dönüştürür."""
    global _last_nominatim_call
    if not query:
        return None
    if query in _coord_cache:
        return _coord_cache[query]

    elapsed = time.time() - _last_nominatim_call
    if elapsed < 1.1:
        time.sleep(1.1 - elapsed)

    try:
        resp = requests.get(
            "https://nominatim.openstreetmap.org/search",
            params={"q": query, "format": "json", "limit": 1, "countrycodes": "tr"},
            headers={"User-Agent": "NexaScoutBuyerEngine/2.0 (info@nexadigital.com)"},
            timeout=5,
        )
        _last_nominatim_call = time.time()
        data = resp.json()
        if data:
            lat, lon = float(data[0]["lat"]), float(data[0]["lon"])
            _coord_cache[query] = (lat, lon)
            return lat, lon
    except Exception:
        pass

    _coord_cache[query] = None
    return None


def _clean_text(el) -> str:
    return el.get_text(separator=" ", strip=True) if el else ""


def _extract_psi_photos(raw_html: str) -> list:
    unescaped = html_mod.unescape(raw_html)
    photos = []
    seen = set()
    pattern = re.compile(
        r"https?://i\d+\.shbdn\.com/photos/[^\s\"'<>&]+\.(?:avif|jpg|jpeg|png|webp)",
        re.IGNORECASE,
    )
    for m in pattern.finditer(unescaped):
        u = m.group(0)
        if u.lower().endswith(".avif"):
            u = u[:-5] + ".jpg"
        if u not in seen and "logo" not in u and "icon" not in u:
            seen.add(u)
            photos.append(u)
            if len(photos) >= 15:
                break
    return photos


def _extract_psi_specs(raw_html: str) -> dict:
    specs = {}
    ep_pattern = re.compile(r"ep\.([A-Za-z0-9_]+)=([^&\n\"'<>]+)", re.IGNORECASE)
    for m in ep_pattern.finditer(raw_html):
        raw_key = "ep." + m.group(1)
        raw_val = m.group(2).replace("&amp;", "&").replace("+", " ")
        try:
            raw_val = urllib.parse.unquote(raw_val).strip()
        except Exception:
            raw_val = raw_val.strip()
        label = _PSI_EP_MAP.get(raw_key)
        if label and raw_val and raw_val not in ("0", "false"):
            specs[label] = raw_val

    # Price formatting
    if "Fiyat (Sayısal)" in specs:
        try:
            amt = int(specs["Fiyat (Sayısal)"])
            specs["Fiyat"] = f"{amt:,.0f} TL".replace(",", ".")
        except Exception:
            pass

    return specs


def _extract_psi_description(raw_text: str) -> str:
    t = html_mod.unescape(raw_text)
    for marker in ["İlan Açıklaması", "Ilan Aciklamasi", "İLAN AÇIKLAMASI", "AÇIKLAMA", "Açıklama"]:
        idx = t.find(marker)
        if idx != -1:
            window = t[idx + len(marker): idx + len(marker) + 3500]
            m = re.match(r"\s*[:\-–]?\s*([^<]{40,3000})", window)
            if m:
                cand = re.sub(r"\s+", " ", m.group(1)).strip()
                if len(cand) >= 40:
                    return cand
    return ""


def _scrape_via_psi_api(url: str) -> dict:
    """Google PageSpeed Insights v5 REST API ile bot korumasını aşarak ilan verilerini çeker."""
    api_key = os.environ.get("PAGESPEED_API_KEY", "AIzaSyClEth2ooknGZJ53WrgY1QKdrQunZfsNXg")
    psi_url = "https://www.googleapis.com/pagespeedonline/v5/runPagespeed"
    listing_id = detect_listing_id(url)
    
    try:
        resp = requests.get(
            psi_url,
            params={"url": url, "category": "performance", "hl": "tr", "key": api_key},
            timeout=8
        )
        if resp.status_code == 200:
            raw_text = resp.text
            photos = _extract_psi_photos(raw_text)
            specs = _extract_psi_specs(raw_text)
            desc = _extract_psi_description(raw_text)
            
            city = specs.get("Şehir", "Ankara")
            dist = specs.get("İlçe", "")
            mah = specs.get("Mahalle", "")
            loc_parts = [p for p in [mah, dist, city] if p]
            loc_str = ", ".join(loc_parts) if loc_parts else "Ankara"
            
            slug_info = _scrape_via_slug_fallback(url)
            title = slug_info.get("title") or f"Sahibinden İlanı #{listing_id}"
            price = specs.get("Fiyat") or ""

            return {
                "ok": True,
                "source": "sahibinden_psi_api",
                "title": title,
                "price": price,
                "location": loc_str,
                "city": city,
                "district": dist,
                "neighborhood": mah,
                "specs": specs,
                "description": desc or slug_info.get("description", ""),
                "images": photos,
                "photo_count": len(photos)
            }
    except Exception:
        pass
    return {"ok": False}


def _scrape_via_slug_fallback(url: str) -> dict:
    """Her koşulda çalışan, URL slug'ından semantik mülk analizi yapan motor."""
    listing_id = detect_listing_id(url)
    slug = ""
    if "/ilan/" in url:
        slug = url.split("/ilan/", 1)[1].split("/detay", 1)[0]
    elif "cb.com.tr" in url:
        slug = url.split("cb.com.tr", 1)[1].split("?", 1)[0]
    else:
        slug = url.split("/")[-1]

    if listing_id and slug.endswith("-" + listing_id):
        slug = slug[: -(len(listing_id) + 1)]

    tokens = [t.strip().lower() for t in re.split(r"[-\./_]", slug) if t.strip()]

    city = "Ankara"
    district = ""
    neighborhood = ""
    prop_type = "Ticari Mülk"
    status = "Satılık"
    title_words = []

    for t in tokens:
        if t in TURKISH_CITIES:
            city = TURKISH_CITIES[t]
        elif t in ["incek", "golbasi", "cankaya", "cayyolu", "umitkoy", "beytepe", "bilkent", "batikent", "eryaman", "kizilcasar"]:
            district = WORD_MAP.get(t, t.capitalize())
        elif t in ["bina", "plaza", "villa", "arsa", "dukkan", "isyeri", "ofis", "rezidans"]:
            prop_type = WORD_MAP.get(t, t.capitalize())
        elif t in ["satilik", "kiralik", "devren"]:
            status = WORD_MAP.get(t, t.capitalize())
            
        if t not in ["emlak", "vasita", "detay"]:
            title_words.append(WORD_MAP.get(t, t.capitalize()))

    loc_str = f"{city}, {district}".strip(", ") if district else city
    title_str = " ".join(title_words) if title_words else f"Portföy İlanı #{listing_id}"

    specs = {
        "Konum": loc_str,
        "Emlak Türü": prop_type,
        "İşlem Türü": status
    }

    desc = (
        f"{title_str} — {loc_str} aksında yer alan bu seçkin portföy; "
        f"mimari yapısı, stratejik konumu ve yüksek ticari/yatırım potansiyeliyle öne çıkmaktadır. "
        f"Tek yetkili danışmanlık hizmetiyle yerinde sunum ve detaylı yatırım şartnamesi hazırlanmıştır."
    )

    return {
        "ok": True,
        "source": "slug_semantic_fallback",
        "title": title_str,
        "price": "",
        "location": loc_str,
        "city": city,
        "district": district,
        "neighborhood": neighborhood,
        "category": prop_type,
        "specs": specs,
        "description": desc,
        "images": [],
        "photo_count": 0
    }


def _scrape_coldwell_banker(url: str) -> dict:
    """cb.com.tr ilanlarını doğrudan detaylı BS4 ile çeker."""
    try:
        resp = requests.get(url, headers=HEADERS, timeout=SCRAPE_TIMEOUT)
        if resp.status_code == 200:
            soup = BeautifulSoup(resp.content, "lxml" if "lxml" in sys.modules else "html.parser")
            
            # Title
            title_el = soup.select_one("h1.detail-title, h1, .property-title")
            title = _clean_text(title_el) or "Coldwell Banker Portföy İlanı"

            # Price
            price_el = soup.select_one(".price, .detail-price, [class*='price']")
            price = _clean_text(price_el)

            # Specs table & lists
            specs = {}
            for row in soup.select("table tr"):
                cells = row.find_all(["td", "th"])
                if len(cells) >= 2:
                    k, v = _clean_text(cells[0]), _clean_text(cells[1])
                    if k and v and len(k) < 40:
                        specs[k] = v

            for li in soup.select("ul.features li, .property-features li, ul.detail-features li"):
                txt = _clean_text(li)
                if ":" in txt:
                    parts = txt.split(":", 1)
                    specs[parts[0].strip()] = parts[1].strip()

            # Images
            images = []
            for img in soup.select(".detail-slider img, .swiper-slide img, .carousel-item img, img[src*='media.cb']"):
                src = img.get("src") or img.get("data-src") or ""
                if src and "placeholder" not in src and src not in images:
                    if src.startswith("/"):
                        src = "https://www.cb.com.tr" + src
                    images.append(src)

            # Description
            desc_el = soup.select_one(".description, .detail-description, #aciklama, [itemprop='description']")
            description = _clean_text(desc_el)

            # Location
            loc_el = soup.select_one(".location, .detail-location, [class*='address']")
            loc_str = _clean_text(loc_el) or "Ankara"

            # Agent
            agent_el = soup.select_one("a[href*='/danismanlar/']")
            agent_name = _clean_text(agent_el) or "Yiğit Narin"

            return {
                "ok": True,
                "source": "cb_direct",
                "title": title,
                "price": price,
                "location": loc_str,
                "specs": specs,
                "description": description,
                "images": images,
                "photo_count": len(images),
                "agent_name": agent_name
            }
    except Exception:
        pass

    # Bot koruması veya 403 ise PSI API'ye devret
    psi_res = _scrape_via_psi_api(url)
    if psi_res.get("ok"):
        return psi_res

    return _scrape_via_slug_fallback(url)


def scrape_any_listing(url: str) -> dict:
    """
    Ana giriş noktası: Verilen herhangi bir emlak linkini otomatik ayrıştırır.
    Sahibinden, Coldwell Banker, Hepsiemlak veya genel sayfaları destekler.
    """
    clean_url = url.strip()
    if not clean_url.startswith("http"):
        clean_url = "https://" + clean_url

    domain = urllib.parse.urlparse(clean_url).netloc.lower()
    listing_id = detect_listing_id(clean_url)

    result = {}
    if "sahibinden.com" in domain:
        result = _scrape_via_psi_api(clean_url)
        if not result.get("ok"):
            result = _scrape_via_slug_fallback(clean_url)
    elif "cb.com.tr" in domain or "coldwellbanker" in domain:
        result = _scrape_coldwell_banker(clean_url)
    else:
        # Genel linkler
        result = _scrape_coldwell_banker(clean_url)
        if not result.get("ok"):
            result = _scrape_via_slug_fallback(clean_url)

    result["url"] = clean_url
    result["listing_id"] = listing_id

    # Geocoding: Adresi koordinata çevir
    loc = result.get("location", "Ankara")
    coords = geocode_address(f"{loc}, Türkiye")
    if coords:
        result["lat"], result["lon"] = coords
    else:
        # Default Ankara İncek / Merkez koordinatı
        result["lat"], result["lon"] = (39.845, 32.748)

    return result
