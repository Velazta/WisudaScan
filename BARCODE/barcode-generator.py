import os
import pandas as pd
import re
import csv
import qrcode 
from PIL import Image, ImageDraw, ImageFont

# --- KONFIGURASI ---
OUTPUT_FOLDER = 'Hasil_QR_Code_Siswa' 
QR_SIZE = 12 
BORDER_SIZE = 2 

# --- FUNGSI PENCARI FONT ---
def get_font_path_for_text():
    candidate_fonts = [
        'C:\\Windows\\Fonts\\arial.ttf',
        'C:\\Windows\\Fonts\\Arial.ttf',
        'C:\\Windows\\Fonts\\calibri.ttf',
        'C:\\Windows\\Fonts\\segoeui.ttf',
        'C:\\Windows\\Fonts\\tahoma.ttf'
    ]
    for font_path in candidate_fonts:
        if os.path.exists(font_path):
            return font_path
    return None

FONT_PATH = get_font_path_for_text()

# --- FUNGSI DETEKSI KELAS (UPDATE BARU: LEBIH PINTAR) ---
def detect_class_smart(file_path, filename):
    # 1. Cek dari Nama File (Prioritas Utama)
    # Cari pola "XII" diikuti angka, misal "XII 1", "XII-2", "XII_3"
    match = re.search(r'(XII\s*[-_]?\s*\d+)', filename, re.IGNORECASE)
    if match:
        raw = match.group(1).upper()
        return re.sub(r'[- ]', '_', raw) # Ubah spasi/- jadi underscore (XII_1)
    
    # 2. Jika tidak ada di nama file, Cek ISI FILE (Backup)
    try:
        with open(file_path, 'r', errors='ignore') as f:
            # Baca 5 baris pertama saja untuk cari judul
            head_lines = [f.readline() for _ in range(5)]
            full_text = " ".join(head_lines).upper()
            
            # Cari kata "KELAS" diikuti "XII ..."
            match_text = re.search(r'KELAS\s+.*?(XII\s*[-_]?\s*\d+)', full_text)
            if match_text:
                print(f"   [INFO] Kelas ditemukan di dalam isi file: {match_text.group(1)}")
                raw = match_text.group(1)
                return re.sub(r'[- ]', '_', raw)
    except Exception:
        pass

    # 3. Nyerah
    return "KELAS_UMUM"

# --- FUNGSI BACA CSV CERDAS ---
def smart_read_csv(file_path):
    header_row = None
    delimiter = ',' 

    try:
        with open(file_path, 'r', errors='ignore') as f:
            lines = f.readlines()
            
        for i, line in enumerate(lines[:10]):
            # Deteksi Header
            if "NAMA SISWA" in line.upper() or ("NAMA" in line.upper() and "NO" in line.upper()):
                header_row = i
                if line.count(';') > line.count(','):
                    delimiter = ';'
                else:
                    delimiter = ','
                print(f"   [INFO] Header ditemukan di baris ke-{i+1}. Pemisah: '{delimiter}'")
                break
    except Exception as e:
        print(f"   [ERROR] Gagal membuka file: {e}")
        return pd.DataFrame()

    if header_row is not None:
        try:
            df = pd.read_csv(file_path, header=header_row, sep=delimiter)
            return df
        except Exception as e:
            print(f"   [ERROR] Pandas gagal membaca: {e}")
            return pd.DataFrame()
    else:
        print(f"   [SKIP] Tidak ditemukan tulisan 'NAMA SISWA' di 10 baris awal.")
        return pd.DataFrame()

# --- PROGRAM UTAMA ---
def main():
    if not os.path.exists(OUTPUT_FOLDER):
        os.makedirs(OUTPUT_FOLDER)

    current_dir = os.getcwd()
    csv_files = [f for f in os.listdir(current_dir) if f.endswith('.csv')]
    
    if not csv_files:
        print("Tidak ditemukan file CSV di folder ini.")
        return

    total_processed = 0
    print(f"\nDitemukan {len(csv_files)} file CSV. Memulai proses QR Code...\n")

    for file_csv in csv_files:
        print(f"-> Memproses file: {file_csv}...")
        
        # UPDATE: Gunakan fungsi deteksi kelas yang lebih pintar
        kelas_clean = detect_class_smart(file_csv, file_csv)
        
        # Baca Data
        df = smart_read_csv(file_csv)

        # Cari Kolom Target
        target_col = None
        for col in df.columns:
            if "NAMA" in str(col).upper() and "SISWA" in str(col).upper():
                target_col = col 
                break
            elif str(col).strip().upper() == "NAMA":
                target_col = col

        if target_col:
            df = df.dropna(subset=[target_col])
            
            # Cari kolom NO
            no_col = None
            for col in df.columns:
                if str(col).strip().upper() == "NO":
                    no_col = col
                    break
            
            for index, row in df.iterrows():
                nama = str(row[target_col]).strip()
                
                if no_col and pd.notna(row[no_col]):
                    raw_no = str(row[no_col]).replace('.0', '')
                    no_absen = raw_no.zfill(2)
                else:
                    no_absen = str(index + 1).zfill(2)

                if not nama or nama.upper() == "NAMA SISWA" or len(nama) < 3:
                    continue

                # --- GENERATE QR CODE ---
                folder_path = os.path.join(OUTPUT_FOLDER, kelas_clean)
                if not os.path.exists(folder_path):
                    os.makedirs(folder_path)

                # ============================================================
                # [PENTING] FORMAT ISI QR CODE: "NAMA | KELAS | NO"
                # Ini memastikan data tidak tertukar antar kelas
                # ============================================================
                kelas_formatted = kelas_clean.replace('_', ' ') # XII_1 -> XII 1
                qr_data = f"{nama.upper()} | {kelas_formatted} | {no_absen}"
                
                # Nama File Gambar
                safe_filename = "".join([c for c in nama if c.isalnum() or c in (' ', '_')]).strip()
                file_name = f"{no_absen}_{safe_filename.replace(' ', '_')}"
                save_path = os.path.join(folder_path, file_name + '.png')

                try:
                    qr = qrcode.QRCode(
                        version=1,
                        error_correction=qrcode.constants.ERROR_CORRECT_H,
                        box_size=QR_SIZE,
                        border=BORDER_SIZE,
                    )
                    qr.add_data(qr_data)
                    qr.make(fit=True)

                    img = qr.make_image(fill_color="black", back_color="white").convert("RGB")
                    
                    # Tambah Teks Nama
                    if FONT_PATH:
                        try:
                            font_size = 20
                            font = ImageFont.truetype(FONT_PATH, font_size)
                            
                            text_height = 50 
                            new_img_height = img.height + text_height
                            new_img = Image.new('RGB', (img.width, new_img_height), (255, 255, 255))
                            new_img.paste(img, (0, 0))
                            
                            draw = ImageDraw.Draw(new_img)
                            text_bbox = draw.textbbox((0,0), nama, font=font)
                            text_width = text_bbox[2] - text_bbox[0]
                            text_x = (img.width - text_width) / 2
                            text_y = img.height + (text_height - (text_bbox[3]-text_bbox[1])) / 2 - 5
                            draw.text((text_x, text_y), nama, font=font, fill=(0, 0, 0))
                            
                            new_img.save(save_path)
                        except Exception:
                            img.save(save_path)
                    else:
                        img.save(save_path)

                    total_processed += 1
                except Exception as e:
                    print(f"      [GAGAL QR] {nama}: {e}")

            print(f"   [OK] Selesai memproses {file_csv} (Kelas: {kelas_clean})")
        else:
            print(f"   [SKIP] Gagal deteksi kolom NAMA di {file_csv}")

    print(f"\n========================================")
    print(f"PROSES SELESAI! Total QR Code: {total_processed}")
    print(f"Cek folder: {os.path.abspath(OUTPUT_FOLDER)}")
    print(f"========================================")

if __name__ == "__main__":
    main()