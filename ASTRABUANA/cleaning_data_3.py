import os
import re
import numpy as np
import pandas as pd

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

INPUT_FILE = os.path.join(BASE_DIR, "dataExcel", "raw", "3a. Database Suspense 170726.xlsx")
SHEET_NAME = "Detail Database"
OUTPUT_FILE = os.path.join(BASE_DIR, "dataExcel", "processed", "astrabuana_output_suspend.xlsx")

CEDANT_COL = "CEDANT SHRT NAME"
CEDANT_VALUE = "ASTRABUANA"

ENTITIES_RE = re.compile(r"\b(PT|CV|TBK|PTE|LTD)\b", re.IGNORECASE)
PUNCT_RE = re.compile(r"[.,]")
MULTIPLE_SPACES_RE = re.compile(r"\s+")

# cleaning dasar
def _clean_text_basic(text):
    if pd.isna(text):
        return ""
    text_str = PUNCT_RE.sub("", str(text))
    return MULTIPLE_SPACES_RE.sub(" ", text_str).strip()

# insured
def clean_insured(name):
    """Bersihkan entitas legal (PT/CV/dll), titik/koma, dan spasi dari nama Insured, lalu ubah ke uppercase."""
    if pd.isna(name):
        return ""
    name_str = ENTITIES_RE.sub("", str(name))
    name_str = PUNCT_RE.sub("", name_str)
    return MULTIPLE_SPACES_RE.sub(" ", name_str).strip().upper()

# polis
def clean_polis(value):
    return _clean_text_basic(value)

# slip
def clean_slip(value):
    return _clean_text_basic(value)

# proses data
def process_data(input_file: str, sheet_name: str, output_file: str) -> None:
    if not os.path.exists(input_file):
        print(f"[ERROR] File tidak ditemukan di lokasi: {input_file}")
        return

    df = pd.read_excel(input_file, sheet_name=sheet_name, header=2)
    df.columns = df.columns.str.strip()

    if CEDANT_COL not in df.columns:
        print(f"[ERROR] Kolom '{CEDANT_COL}' tidak ditemukan di Excel!")
        return

    # filter baris ASTRABUANA
    is_astrabuana = df[CEDANT_COL] == CEDANT_VALUE

    insured_cleaned = df["INSURED"].apply(clean_insured)
    # jika ASTRABUANA -> isi hasil clean, jika BUKAN -> isi None 
    df.insert(
        df.columns.get_loc("INSURED") + 1,
        "INSURED_CLEAN",
        np.where(is_astrabuana, insured_cleaned, None)
    )

    polis_cleaned = df["POLIS"].apply(clean_polis)
    df.insert(
        df.columns.get_loc("POLIS") + 1,
        "POLIS_CLEAN",
        np.where(is_astrabuana, polis_cleaned, None)
    )

    slip_cleaned = df["SLIP NO"].apply(clean_slip)
    df.insert(
        df.columns.get_loc("SLIP NO") + 1,
        "SLIP_NO_CLEAN",
        np.where(is_astrabuana, slip_cleaned, None)
    )

    df.to_excel(output_file, index=False)
    print("Pemrosesan Data 3 Selesai!")

if __name__ == "__main__":
    process_data(INPUT_FILE, SHEET_NAME, OUTPUT_FILE)