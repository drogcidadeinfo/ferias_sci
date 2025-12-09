import os
import pandas as pd
import re
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build

# =============== CONFIG ===============
INPUT_FOLDER = "downloads"
SPREADSHEET_ID = os.getenv("SPREADSHEET_ID")
SHEET_NAME = os.getenv("SHEET_NAME")
CREDENTIALS_FILE = os.getenv("GSA_CREDENTIALS")
# =====================================

# === Extract filial from filename ===
def extract_filial_from_filename(filename):
    match = re.search(r"FILIAL\s*-\s*(\d+)", filename)
    if match:
        return f"F{match.group(1).zfill(2)}"
    return None

# === Auto-detect delimiter ===
def detect_delimiter(path):
    with open(path, "r", encoding="latin1", errors="ignore") as f:
        sample = f.read(2048)
        return ";" if sample.count(";") > sample.count(",") else ","

# === Load CSV with automatic encoding ===
def load_and_process_file(path):
    filename = os.path.basename(path)
    filial = extract_filial_from_filename(filename)

    if not filial:
        print(f"⚠️ Could not extract filial from file: {filename}")
        return None

    delimiter = detect_delimiter(path)
    encodings = ["utf-8", "latin1", "cp1252"]
    df, last_error = None, None

    for enc in encodings:
        try:
            df = pd.read_csv(path, encoding=enc, sep=delimiter)
            print(f"📄 Loaded {filename} using '{enc}'  delimiter='{delimiter}'")
            break
        except Exception as e:
            last_error = e

    if df is None:
        print(f"❌ Failed to load {filename}: {last_error}")
        return None

    # Normalize headers
    df.columns = [col.replace("\ufeff", "").strip() for col in df.columns]

    # Find "Centro de custo"
    for col in df.columns:
        if col.lower().replace(" ", "") == "centrodecusto":
            df.rename(columns={col: "Filial"}, inplace=True)
            break

    # Ensure Filial exists
    df["Filial"] = filial

    return df

# === Merge all CSVs ===
def merge_all_files():
    all_data = []

    for file in os.listdir(INPUT_FOLDER):
        if file.lower().endswith(".csv"):
            full_path = os.path.join(INPUT_FOLDER, file)
            print(f"📄 Processing {file} ...")
            df = load_and_process_file(full_path)
            if df is not None:
                all_data.append(df)

    if not all_data:
        print("❌ No valid CSV files found.")
        return None

    merged = pd.concat(all_data, ignore_index=True)
    print(f"✅ Merged {len(all_data)} files, total rows: {len(merged)}")
    return merged

# === Upload to Google Sheets ===
def upload_to_google_sheets(df):
    scopes = ["https://www.googleapis.com/auth/spreadsheets"]
    creds = Credentials.from_service_account_file(CREDENTIALS_FILE, scopes=scopes)
    service = build("sheets", "v4", credentials=creds)

    values = [df.columns.tolist()] + df.values.tolist()
    body = {"values": values}

    # Clear sheet
    service.spreadsheets().values().clear(
        spreadsheetId=SPREADSHEET_ID,
        range=SHEET_NAME
    ).execute()

    # Upload
    service.spreadsheets().values().update(
        spreadsheetId=SPREADSHEET_ID,
        range=SHEET_NAME,
        valueInputOption="RAW",
        body=body
    ).execute()

    print("✅ Uploaded successfully to Google Sheets!")

# === MAIN ===
if __name__ == "__main__":
    df = merge_all_files()

    if df is not None:
        print("\n=== PREVIEW ===")
        print(df.head())

        upload_to_google_sheets(df)
