# -*- coding: utf-8 -*-
r"""
NexaScout Proposal DOCX to TXT Batch Converter
Tüm kişiselleştirilmiş .docx teklif mektuplarını okur,
okunaklı ve profesyonel UTF-8 .txt dosyalarına dönüştürür.
"""

import os
import glob
import docx

PROPOSALS_DIR = r"C:\Users\USER\Desktop\SATTIM\PAZARLAMA\TEKLIF_MEKTUPLARI"
TXT_DEDICATED_DIR = r"C:\Users\USER\Desktop\SATTIM\PAZARLAMA\TEKLIF_MEKTUPLARI_TXT"

def convert_single_docx_to_text(docx_path):
    """Tek bir docx dosyasını hiyerarşik ve okunaklı metne dönüştürür."""
    doc = docx.Document(docx_path)
    lines = []
    
    for element in doc.element.body:
        if element.tag.endswith('p'):
            p = docx.text.paragraph.Paragraph(element, doc)
            txt = p.text.strip()
            if txt:
                lines.append(txt)
        elif element.tag.endswith('tbl'):
            t = docx.table.Table(element, doc)
            lines.append("")
            lines.append("=" * 65)
            for row in t.rows:
                if len(row.cells) >= 2:
                    lbl = row.cells[0].text.strip()
                    val = row.cells[1].text.strip()
                    lines.append(f"  • {lbl:<22}: {val}")
            lines.append("=" * 65)
            lines.append("")
            
    return "\n".join(lines)

def run_conversion():
    os.makedirs(PROPOSALS_DIR, exist_ok=True)
    os.makedirs(TXT_DEDICATED_DIR, exist_ok=True)
    
    docx_files = [f for f in glob.glob(os.path.join(PROPOSALS_DIR, "*.docx")) if not os.path.basename(f).startswith("~$")]
    print(f"Toplam {len(docx_files)} adet Word teklif mektubu (.docx) bulundu.")
    
    success_count = 0
    for docx_path in docx_files:
        try:
            base_name = os.path.splitext(os.path.basename(docx_path))[0]
            txt_content = convert_single_docx_to_text(docx_path)
            
            # 1. Aynı klasöre .txt olarak kaydet
            target_same_folder = os.path.join(PROPOSALS_DIR, f"{base_name}.txt")
            with open(target_same_folder, "w", encoding="utf-8") as f:
                f.write(txt_content)
                
            # 2. Özel TXT klasörüne de kaydet
            target_dedicated = os.path.join(TXT_DEDICATED_DIR, f"{base_name}.txt")
            with open(target_dedicated, "w", encoding="utf-8") as f:
                f.write(txt_content)
                
            success_count += 1
        except Exception as e:
            print(f"Hata ({os.path.basename(docx_path)}): {e}")
            
    print(f"\nİşlem Tamamlandı: Toplam {success_count} adet .txt teklif mektubu üretildi!")
    print(f"1. Klasör: {PROPOSALS_DIR}")
    print(f"2. Klasör: {TXT_DEDICATED_DIR}")
    return success_count

if __name__ == "__main__":
    run_conversion()
