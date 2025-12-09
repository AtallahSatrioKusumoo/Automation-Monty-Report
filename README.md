# ----------------------------------------------------
# 1. Setup Dependencies (Library Installation & Import)
# ----------------------------------------------------

!pip install gspread google-auth google-api-python-client pandas

import gspread
import pandas as pd
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
import sys # Digunakan untuk exit jika terjadi kegagalan fatal

# Catatan: Jika menggunakan Google Colab, aktifkan bagian berikut:
# from google.colab import auth
# auth.authenticate_user() 

# ----------------------------------------------------
# 2. Konfigurasi ID Proyek dan Kunci
# ----------------------------------------------------

# Ganti dengan ID Google Sheet dan nama file Service Account yang sebenarnya
JSON_KEY_FILE   = "montly-report.json" 
SHEET_INPUT_ID  = "YOUR_INPUT_SHEET_ID" 
SHEET_REPORT_ID = "YOUR_REPORT_SHEET_ID"
DOC_ID          = "YOUR_DOC_TEMPLATE_ID" 

# ----------------------------------------------------
# 3. Fungsi Otentikasi Service Account
# ----------------------------------------------------

def setup_auth(json_key_file):
    """
    Mengatur otentikasi menggunakan Google Service Account JSON key file.
    Memberikan otorisasi akses ke Google Sheets dan Docs API.
    """
    SCOPES = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/documents",
        "https://www.googleapis.com/auth/drive"
    ]
    
    try:
        # Membuat kredensial dari file kunci rahasia
        creds = Credentials.from_service_account_file(
            json_key_file,
            scopes=SCOPES
        )
        
        # Otorisasi gspread (untuk interaksi Google Sheet)
        gc = gspread.authorize(creds)
        # Inisialisasi service Docs API (untuk keperluan dokumen laporan)
        docs_service = build("docs", "v1", credentials=creds)
        
        print("— Auth Success: Koneksi ke Google Sheets & Docs berhasil diinisialisasi.")
        return gc, docs_service
        
    except FileNotFoundError:
        print(f"— Auth Error: File kunci '{json_key_file}' tidak ditemukan.")
        sys.exit(1) # Hentikan program jika file kunci hilang
    except Exception as e:
        print(f"— Auth Error: Gagal mengotorisasi Service Account. Detail: {e}")
        sys.exit(1)

# ----------------------------------------------------
# 4. Fungsi Ekstraksi Data (E: Extract)
# ----------------------------------------------------

def extract_data(gc_client, sheet_id):
    """
    Mengambil data dari Google Sheet dan mengonversinya menjadi DataFrame Pandas.
    Mengembalikan DataFrame kosong jika terjadi kegagalan.
    """
    try:
        sheet = gc_client.open_by_key(sheet_id).sheet1
        raw_data = sheet.get_all_values()
        
        # Memisahkan header (baris pertama) dari data
        headers = raw_data[0]
        rows = raw_data[1:]
        
        df = pd.DataFrame(rows, columns=headers)
        
        print(f"— Extract Success: {len(df)} baris data berhasil diambil dari Sheet Input.")
        return df
    except Exception as e:
        print(f"— Extract Error: Gagal mengambil data dari Google Sheet ID {sheet_id}. Detail: {e}")
        return pd.DataFrame() 

# ----------------------------------------------------
# 5. Fungsi Load Data (L: Load)
# ----------------------------------------------------

def load_data(gc_client, sheet_id, df_to_upload):
    """
    Membersihkan sheet report tujuan dan mengunggah data DataFrame baru.
    """
    if df_to_upload.empty:
        print("— Load Info: DataFrame kosong. Tidak ada data yang diunggah.")
        return

    try:
        sheet_report = gc_client.open_by_key(sheet_id).sheet1
        
        # 1. Bersihkan Data Lama
        sheet_report.clear()
        
        # 2. Format Data untuk Gspread (Header + Rows)
        data_to_send = [df_to_upload.columns.values.tolist()] + df_to_upload.values.tolist()
        sheet_report.update(data_to_send)
        
        print(f"— Load Success: Sheet Report berhasil diupdate dengan {len(df_to_upload)} baris data baru.")
    except Exception as e:
        print(f"— Load Error: Gagal mengupdate Sheet Report ID {sheet_id}. Detail: {e}")

# ----------------------------------------------------
# 6. Main Flow (Eksekusi Otomasi ETL)
# ----------------------------------------------------

if __name__ == "__main__":
    
    # Inisialisasi Klien
    gc, docs_service = setup_auth(JSON_KEY_FILE)
    print("\n— START: Memulai proses Automation-Monty-Report —")
    
    # 1. Extract (Ambil Data)
    input_df = extract_data(gc, SHEET_INPUT_ID)
    
    if not input_df.empty:
        
        # 2. Transform (Proses Data)
        # Bagian ini adalah tempat untuk menambahkan logika pengolahan data.
        # Contoh: Filtering, Agregasi, Kalkulasi Metrik.
        # Saat ini, menggunakan data input sebagai target_df (Copy Data).
        target_df = input_df.copy() 
        
        print(f"— Transform Info: Data siap diunggah: {len(target_df)} baris.")
        
        # 3. Load (Kirim Data)
        load_data(gc, SHEET_REPORT_ID, target_df)
    
    else:
        print("— STOP: Data input kosong. Proses dihentikan.")
        
    print("\n— END: Proses Otomasi Laporan Selesai —")
