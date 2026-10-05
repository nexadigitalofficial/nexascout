"""
Web Tabanlı B2B Pazarlama ve Harita Kontrol Paneli (Flask Uygulaması)
Ankara İncek / LÖSANTE Karşısı Ticari Mülk Sistemi
Erişim: http://localhost:5000
"""

import os
import sys
import json
import webbrowser
import threading
from flask import Flask, render_template, request, jsonify, send_file, send_from_directory

# Force UTF-8
sys.stdout.reconfigure(encoding="utf-8")

# Add current dir to path
APP_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, APP_DIR)

from modules.enrichment import clean_phone_for_whatsapp, generate_whatsapp_url
from modules.map_generator import generate_interactive_map
from modules.pitch_generator import generate_pitch_document
from google_maps_lead_harvester import scrape_google_maps_places, export_to_excel, load_config

import openpyxl

app = Flask(__name__, template_folder="templates")

# File Paths - Supports both workspace parent dir and standalone clone
LOCAL_DATA_DIR = os.path.join(APP_DIR, "data")
PARENT_DIR = os.path.abspath(os.path.join(APP_DIR, ".."))

if os.path.exists(os.path.join(PARENT_DIR, "GOOGLE_HARITA_TOPLANAN_KURUMLAR.xlsx")):
    DATA_ROOT = PARENT_DIR
elif os.path.exists(os.path.join(LOCAL_DATA_DIR, "GOOGLE_HARITA_TOPLANAN_KURUMLAR.xlsx")):
    DATA_ROOT = LOCAL_DATA_DIR
else:
    DATA_ROOT = PARENT_DIR

MASTER_EXCEL_PATH = os.path.join(DATA_ROOT, "GOOGLE_HARITA_TOPLANAN_KURUMLAR.xlsx")
STRATEGIC_EXCEL_PATH = os.path.join(DATA_ROOT, "STRATEJIK_PAZARLAMA_VE_HEDEF_MUSTERI_LISTESI.xlsx")
MAP_HTML_PATH = os.path.join(DATA_ROOT, "INCEK_TICARI_HEDEF_HARITASI.html")
PROPOSALS_DIR = os.path.join(DATA_ROOT, "TEKLIF_MEKTUPLARI")


def load_leads_from_excel():
    """Master Excel dosyasından toplanan kurumları okur ve zenginleştirir."""
    leads = []
    if not os.path.exists(MASTER_EXCEL_PATH):
        return leads
        
    try:
        wb = openpyxl.load_workbook(MASTER_EXCEL_PATH, data_only=True)
        if "TOPLANAN_KURUMLAR" in wb.sheetnames:
            ws = wb["TOPLANAN_KURUMLAR"]
            for r in range(2, ws.max_row + 1):
                name = ws.cell(row=r, column=2).value
                if not name:
                    continue
                phone = str(ws.cell(row=r, column=7).value or "N/A")
                wa_phone = clean_phone_for_whatsapp(phone)
                
                dist_val = ws.cell(row=r, column=11).value
                try:
                    dist_float = float(dist_val) if dist_val is not None else None
                except Exception:
                    dist_float = None
                    
                leads.append({
                    "kurum_adi": name,
                    "ana_kategori": ws.cell(row=r, column=3).value or "",
                    "is_kolu": ws.cell(row=r, column=4).value or "Kurumsal",
                    "adres": ws.cell(row=r, column=5).value or "N/A",
                    "plus_kodu": ws.cell(row=r, column=6).value or "N/A",
                    "telefon": phone,
                    "telefon_wa": wa_phone or "",
                    "eposta": ws.cell(row=r, column=8).value or "N/A",
                    "web_sitesi": ws.cell(row=r, column=9).value or "",
                    "koordinatlar": ws.cell(row=r, column=10).value or "N/A",
                    "mesafe_km": dist_float,
                    "puan": ws.cell(row=r, column=12).value or "",
                    "hbu_model": ws.cell(row=r, column=13).value or "Ticari Yerleşke",
                    "oncelik": ws.cell(row=r, column=14).value or "B+",
                    "maps_url": ws.cell(row=r, column=15).value or "#",
                    "tarih": ws.cell(row=r, column=16).value or ""
                })
    except Exception as e:
        print(f"Excel okunurken hata: {e}")
        
    return leads

def compute_stats(leads):
    """Kurumlar üzerinden özet analitik metrikleri hesaplar."""
    total = len(leads)
    health_count = sum(1 for l in leads if any(k in (l.get("is_kolu") or "").lower() for k in ["sağlık", "tıp", "onkoloji", "fizik"]))
    edu_count = sum(1 for l in leads if any(k in (l.get("is_kolu") or "").lower() for k in ["eğitim", "kolej", "okul"]))
    dental_count = sum(1 for l in leads if any(k in (l.get("is_kolu") or "").lower() for k in ["diş", "estetik", "cerrahi"]))
    aplus_count = sum(1 for l in leads if l.get("oncelik") == "A+")
    
    valid_dists = [l["mesafe_km"] for l in leads if l.get("mesafe_km") is not None]
    avg_dist = round(sum(valid_dists) / len(valid_dists), 1) if valid_dists else 0.0
    
    return {
        "total_leads": total,
        "health_count": health_count,
        "edu_count": edu_count,
        "dental_count": dental_count,
        "aplus_count": aplus_count,
        "avg_distance": avg_dist
    }

@app.route("/")
def dashboard():
    leads = load_leads_from_excel()
    stats = compute_stats(leads)
    return render_template("dashboard.html", leads=leads, stats=stats)

@app.route("/map")
def serve_map():
    if not os.path.exists(MAP_HTML_PATH):
        leads = load_leads_from_excel()
        generate_interactive_map(leads, MAP_HTML_PATH)
    return send_file(MAP_HTML_PATH)

@app.route("/download/master_excel")
def download_master_excel():
    if os.path.exists(MASTER_EXCEL_PATH):
        return send_file(MASTER_EXCEL_PATH, as_attachment=True)
    return "Dosya bulunamadı", 404

@app.route("/download/strategic_excel")
def download_strategic_excel():
    if os.path.exists(STRATEGIC_EXCEL_PATH):
        return send_file(STRATEGIC_EXCEL_PATH, as_attachment=True)
    return "Dosya bulunamadı", 404

@app.route("/api/generate_pitch/<int:lead_index>")
def api_generate_pitch(lead_index):
    leads = load_leads_from_excel()
    if 0 <= lead_index < len(leads):
        lead = leads[lead_index]
        file_path = generate_pitch_document(lead, PROPOSALS_DIR)
        return send_file(file_path, as_attachment=True)
    return jsonify({"error": "Kurum bulunamadı"}), 404

@app.route("/api/generate_all_pitches")
def api_generate_all_pitches():
    leads = load_leads_from_excel()
    count = 0
    for lead in leads:
        generate_pitch_document(lead, PROPOSALS_DIR)
        count += 1
    return jsonify({
        "status": "success",
        "message": f"Toplam {count} adet kurum için kişiselleştirilmiş Word (.docx) teklif mektubu '{PROPOSALS_DIR}' klasörüne başarıyla üretildi!"
    })

@app.route("/api/scan", methods=["POST"])
def api_start_scan():
    data = request.json or {}
    category = data.get("category", "all")
    custom_query = data.get("query")
    max_places = int(data.get("max_places", 6))
    
    config = load_config()
    search_queries = []
    
    if custom_query and custom_query.strip():
        search_queries = [custom_query.strip()]
    else:
        presets = config.get("category_presets", {})
        if category == "all":
            for cat_key, cat_data in presets.items():
                search_queries.extend(cat_data.get("keywords", []))
        elif category in presets:
            search_queries = presets[category].get("keywords", [])
        else:
            search_queries = [f"{category} İncek Ankara"]
            
    # Run scraper in current context
    anchor_lat = config.get("anchor_property", {}).get("latitude", 39.8245)
    anchor_lon = config.get("anchor_property", {}).get("longitude", 32.7485)
    
    results = scrape_google_maps_places(
        search_queries=search_queries,
        anchor_lat=anchor_lat,
        anchor_lon=anchor_lon,
        max_places_per_query=max_places,
        headless=True
    )
    
    # Export to Excel
    export_to_excel(results, output_path=MASTER_EXCEL_PATH)
    
    # Regenerate Map
    all_leads = load_leads_from_excel()
    generate_interactive_map(all_leads, MAP_HTML_PATH)
    
    return jsonify({
        "status": "success",
        "new_count": len(results),
        "total_count": len(all_leads)
    })

def open_browser():
    time.sleep(1.5)
    webbrowser.open("http://localhost:5000")

if __name__ == "__main__":
    print("\n" + "="*75)
    print("  COLDWELL BANKER VIP - INCEK TICARI MULK B2B PAZARLAMA KONTROL PANELI")
    print("="*75)
    print("  Web Paneli Baslatiliyor: http://localhost:5000")
    print("  Durdurmak icin: CTRL + C tuslarina basiniz.")
    print("="*75 + "\n")
    
    threading.Thread(target=open_browser, daemon=True).start()
    app.run(host="0.0.0.0", port=5000, debug=False)
