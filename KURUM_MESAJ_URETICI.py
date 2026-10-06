# -*- coding: utf-8 -*-
"""
NexaScout - Kurum Analizi ve AI Mesaj Üretici (CLI)
Kullanım:
  python KURUM_MESAJ_URETICI.py "Medicana Sağlık Grubu"
  veya parametresiz çalıştırıp kurum adını yazın.
"""

import sys
import os

# Set UTF-8
sys.stdout.reconfigure(encoding="utf-8")

APP_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, APP_DIR)

from modules.ai_pitch_crafter import craft_institution_pitch_suite

def main():
    print("=" * 75)
    print("  NEXASCOUT AI KURUM ANALİZİ VE STRATEJİK MESAJ ÜRETİCİ")
    print("  Danışman: YİĞİT NARİN (0532 451 40 08) | Coldwell Banker VIP")
    print("  Mülk: Ankara İncek / LÖSANTE Karşısı Tek Ticari Ruhsatlı Müstakil Bina")
    print("=" * 75)
    
    if len(sys.argv) > 1:
        kurum_adi = " ".join(sys.argv[1:])
    else:
        try:
            kurum_adi = input("\nHedef Kurum / Şirket Adını Giriniz (Örn: Medicana, Dent Group, FNSS): ").strip()
        except EOFError:
            kurum_adi = "Medicana Sağlık Grubu"
            
    if not kurum_adi:
        kurum_adi = "Medicana Sağlık Grubu"
        
    print(f"\n[+] '{kurum_adi}' analiz ediliyor ve sektörel kurgu oluşturuluyor...\n")
    
    res = craft_institution_pitch_suite(kurum_adi)
    
    print("-" * 75)
    print(f"📊 TESPİT EDİLEN SEKTÖR & STRATEJİ: {res['sektor']}")
    print(f"🎯 KONUMLAMA: {res['persona_angle']}")
    print("-" * 75)
    
    print("\n" + "=" * 75)
    print("📧 FORMAT 1: KURUMSAL E-POSTA (Yatırım Direktörlüğü & C-Suite)")
    print("=" * 75)
    print(f"KONU: {res['email_subject']}\n")
    print(res['email_body'])
    
    print("\n" + "=" * 75)
    print("💬 FORMAT 2: WHATSAPP VIP YÖNETİCİ TEASER'I")
    print("=" * 75)
    print(res['whatsapp_teaser'])
    
    print("\n" + "=" * 75)
    print("🤝 FORMAT 3: EMLAK GRUBU & MESLEKTAŞ İŞ BİRLİĞİ (VIP Müşteri Getir Formatı)")
    print("=" * 75)
    print(res['emlak_grubu_mesaji'])
    
    print("\n" + "=" * 75)
    print("🎨 FORMAT 4: GÖRSEL / AFİŞ / STORY İÇİN TİPOGRAFİK VİTRİN METNİ")
    print("=" * 75)
    print(res['gorsel_afis_metni'])
    print("=" * 75)
    
    # Save output to a text file for convenience
    safe_name = "".join(c for c in kurum_adi if c.isalnum() or c in (' ', '_', '-')).strip()
    out_dir = os.path.join(APP_DIR, "data", "AI_MESAJLARI")
    os.makedirs(out_dir, exist_ok=True)
    out_file = os.path.join(out_dir, f"{safe_name}_STRATEJIK_MESAJLARI.txt")
    
    with open(out_file, "w", encoding="utf-8") as f:
        f.write(f"KURUM: {kurum_adi}\nSEKTÖR: {res['sektor']}\n\n")
        f.write("=== FORMAT 1: KURUMSAL E-POSTA ===\n")
        f.write(f"KONU: {res['email_subject']}\n\n")
        f.write(res['email_body'] + "\n\n")
        f.write("=== FORMAT 2: WHATSAPP VIP YÖNETİCİ TEASER'I ===\n")
        f.write(res['whatsapp_teaser'] + "\n\n")
        f.write("=== FORMAT 3: EMLAK GRUBU & MESLEKTAŞ İŞ BİRLİĞİ ===\n")
        f.write(res['emlak_grubu_mesaji'] + "\n\n")
        f.write("=== FORMAT 4: GÖRSEL / AFİŞ METNİ ===\n")
        f.write(res['gorsel_afis_metni'] + "\n")
        
    print(f"\n[✓] Tüm metinler kaydedildi: {out_file}")

if __name__ == "__main__":
    main()
