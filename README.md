# Automation-Monty-Report
================================================================================================================================================================
|                                                                                                                                                              |
|                                                                                                                                                              |
|                                                                                                                                                              |
|                                                                                                                                                              |
|                 /$$$$$$    /$$               /$$ /$$           /$$             /$$$$$$$                                               /$$                    |
|                /$$__  $$  | $$              | $$| $$          | $$            | $$__  $$                                             | $$                    |
|               | $$  \ $$ /$$$$$$    /$$$$$$ | $$| $$  /$$$$$$ | $$$$$$$       | $$  \ $$ /$$$$$$   /$$$$$$  /$$  /$$$$$$   /$$$$$$$ /$$$$$$                  |
|               | $$$$$$$$|_  $$_/   |____  $$| $$| $$ |____  $$| $$__  $$      | $$$$$$$//$$__  $$ /$$__  $$|__/ /$$__  $$ /$$_____/|_  $$_/                  |
|               | $$__  $$  | $$      /$$$$$$$| $$| $$  /$$$$$$$| $$  \ $$      | $$____/| $$  \__/| $$  \ $$ /$$| $$$$$$$$| $$        | $$                    |
|               | $$  | $$  | $$ /$$ /$$__  $$| $$| $$ /$$__  $$| $$  | $$      | $$     | $$      | $$  | $$| $$| $$_____/| $$        | $$ /$$                |
|               | $$  | $$  |  $$$$/|  $$$$$$$| $$| $$|  $$$$$$$| $$  | $$      | $$     | $$      |  $$$$$$/| $$|  $$$$$$$|  $$$$$$$  |  $$$$/                |
|               |__/  |__/   \___/   \_______/|__/|__/ \_______/|__/  |__/      |__/     |__/       \______/ | $$ \_______/ \_______/   \___/                  |
|                                                                                                       /$$  | $$                                              |
|                                                                                                      |  $$$$$$/                                              |
|                                                                                                        \______/                                              |
|                                                                                                                                                              |
|                                                                                                                                                              |
|                                                                                                                                                              |
|                                                                                                                                                              |
|                                                                                                                                                              |
================================================================================================================================================================

!pip install gspread google-auth google-api-python-client pandas

import gspread
import pandas as pd

from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from google.colab import auth
auth.authenticate_user()
docs_service = build("docs", "v1", credentials=creds)


# ================= AUTH SERVICE ACCOUNT =================
SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/documents",
    "https://www.googleapis.com/auth/drive"
]

creds = Credentials.from_service_account_file(
    "montly-report.json",  # file JSON lo
    scopes=SCOPES
)

gc = gspread.authorize(creds)
docs_service = build("docs", "v1", credentials=creds)

print("✅ Auth Google Sheet & Docs berhasil")


SHEET_INPUT_ID  = ""
SHEET_REPORT_ID = ""
DOC_ID          = ""

print("Semua ID siap")

# ===== BACA SHEET INPUT (SAFE MODE) =====
sheet_input = gc.open_by_key(SHEET_INPUT_ID).sheet1
raw_data = sheet_input.get_all_values()

headers = raw_data[0]
rows = raw_data[1:]

df = pd.DataFrame(rows, columns=headers)

print("DATA INPUT:")
display(df)

target_df = df

sheet_report = gc.open_by_key(SHEET_REPORT_ID).sheet1

# HAPUS DATA LAMA
sheet_report.clear()

# KIRIM DATA BARU
sheet_report.update(
    [target_df.columns.values.tolist()] + target_df.values.tolist()
)

print(" Sheet OUTPUT berhasil diupdate!")

