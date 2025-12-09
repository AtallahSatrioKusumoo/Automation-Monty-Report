# automation_report.py

# ----------------------------------------------------
# 1. Setup Dependencies
# ----------------------------------------------------

!pip install gspread google-auth google-api-python-client pandas

import gspread
import pandas as pd
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
import sys
import logging

# Mengatur format logging agar output konsol terlihat rapi dan standar
logging.basicConfig(level=logging.INFO, format='%(levelname)s: %(message)s')


# ----------------------------------------------------
# 2. Configuration (Ganti Nilai ID dan Key Lo di sini)
# ----------------------------------------------------

JSON_KEY_FILE   = "montly-report.json" # Nama file kunci Service Account
SHEET_INPUT_ID  = "YOUR_INPUT_SHEET_ID" # ID Google Sheet Sumber Data
SHEET_REPORT_ID = "YOUR_REPORT_SHEET_ID" # ID Google Sheet Tujuan Laporan
DOC_ID          = "YOUR_DOC_TEMPLATE_ID" # ID Google Docs Template (Jika nanti dipakai)

# Hak Akses (Scopes) yang dibutuhkan untuk script ini
SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/documents",
    "https://www.googleapis.com/auth/drive"
]


# ----------------------------------------------------
# 3. Class Utama: ReportAutomation (Organisasi Code OOP)
# ----------------------------------------------------

class ReportAutomation:
    """
    Class untuk mengelola seluruh alur otomatisasi Extract, Transform, Load (ETL) 
    antara Google Sheets dan API terkait.
    """
    def __init__(self, key_file, input_id, report_id):
        self.key_file = key_file
        self.input_id = input_id
        self.report_id = report_id
        self.gc = None
        self.docs_service = None
        self.target_df = pd.DataFrame() 
        
        # Panggil otentikasi saat class diinisialisasi
        self._setup_auth()

    def _setup_auth(self):
        """Metode privat untuk otentikasi Service Account dan inisialisasi klien."""
        try:
            creds = Credentials.from_service_account_file(self.key_file, scopes=SCOPES)
            self.gc = gspread.authorize(creds)
            self.docs_service = build("docs", "v1", credentials=creds)
            logging.info("Otentikasi Berhasil. Klien API siap digunakan.")
            
        except Exception as e:
            logging.error(f"Gagal mengotorisasi Service Account. Detail: {e}")
            sys.exit(1) # Hentikan eksekusi jika gagal otentikasi

    def extract_data(self):
        """Mengambil data mentah dari Google Sheet input dan mengonversi ke Pandas DataFrame."""
        try:
            sheet = self.gc.open_by_key(self.input_id).sheet1
            raw_data = sheet.get_all_values()
            
            # Memisahkan Header (baris 0) dari Data (baris 1 dst.)
            headers = raw_data[0]
            rows = raw_data[1:]
            
            input_df = pd.DataFrame(rows, columns=headers)
            logging.info(f"EXTRACT: {len(input_df)} baris data berhasil ditarik.")
            return input_df
            
        except Exception as e:
            logging.error(f"Gagal mengambil data dari Sheet Input. Detail: {e}")
            return pd.DataFrame() 

    def transform_data(self, input_df):
        """
        Tempat untuk menambahkan semua logika pengolahan data (Filtering, Agregasi, Kalkulasi).
        Data mentah diubah menjadi data final yang siap dilaporkan.
        """
        if input_df.empty:
            return pd.DataFrame()
            
        # --- LOGIKA TRANSFORMASI DIM
