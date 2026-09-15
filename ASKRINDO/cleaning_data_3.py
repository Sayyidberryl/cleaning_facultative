import os
import re
import numpy as np
import pandas as pd

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

INPUT_FILE = os.path.join(BASE_DIR, "data_excel", "raw", "3b. Database Suspense 150826.xlsx")
SHEET_NAME = "Sheet1"
OUTPUT_FILE = os.path.join(BASE_DIR, "data_excel", "processed", "askrindo_output_suspend.xlsx")

CEDANT_COL = "CEDANT SHRT NAME"
CEDANT_VALUE = "ASKRINDO"

BASIC_PUNCT_RE = re.compile(r"[.,]")
INSURED_PUNCT_RE = re.compile(r"[.,/+\-&]")

# Pola rentang tahun
YEAR_RANGE_RE = re.compile(r"\b(19|20)\d{2}\s*[-–]?\s*(19|20)\d{2}\b|\b(19|20)\d{2}\b")

# Daftar Stopwords
STOPWORDS = [
    r"Pembayaran\s+Premi\s+Fakultatif",
    r"Aset\s+Produktif",
    r"Periode",
    r"PT", r"CV", r"TBK", r"PTE", r"LTD", r"BIT",
    r"an", r"PAR", r"EQ", r"CECR", r"GIT", r"Fullpayment"
]
STOPWORDS_RE = re.compile(r"\b(" + "|".join(STOPWORDS) + r")\b", re.IGNORECASE)

MULTIPLE_SPACES_RE = re.compile(r"\s+")
EXT_RE = re.compile(r"-EXT\(\d+\)", re.IGNORECASE)

EXCEPTIONS = [
    "pembayaran premi soa",
    "pensesian wajib bppdan",
    "pembayaran dan penagihan premi soa"
]

def _clean_text_basic(text):
    if pd.isna(text):
        return ""
    text_str = BASIC_PUNCT_RE.sub("", str(text))
    return MULTIPLE_SPACES_RE.sub(" ", text_str).strip()

def clean_insured(name):
    if pd.isna(name):
        return [""]
    
    original_name = str(name).strip()
    lower_name = original_name.lower()
    
    # Rules 3, 4, 5 (Exceptions)
    if any(exc in lower_name for exc in EXCEPTIONS):
        # PERBAIKAN: Tambahkan .upper() agar output pengecualian juga uppercase
        return [original_name.upper()]
        
    # NEW RULE: Hapus kata "as Owner/Principal"
    temp_name = re.sub(r'\bas\s+owner/principal\b', ' ', original_name, flags=re.IGNORECASE)
    
    # Hapus "(PERSERO)" terlebih dahulu agar tidak ikut ter-breakdown oleh rule dalam kurung
    temp_name = re.sub(r'\(PERSERO\)', ' ', temp_name, flags=re.IGNORECASE)
    
    # NEW RULE: Breakdown berdasarkan tanda ";" ATAU kata "and/or"
    split_parts = re.split(r';|\band/or\b', temp_name, flags=re.IGNORECASE)
    
    raw_parts = []
    for part in split_parts:
        part = part.strip()
        if not part:
            continue
            
        # Hapus teks di dalam kurung (dianggap singkatan/keterangan tambahan,
        # BUKAN insured terpisah) -> misal "(PGE)", "(BPMA)"
        part = re.sub(r'\([^)]+\)', ' ', part)

        raw_parts.append(part)
    
    # Proses Cleaning masing-masing hasil breakdown
    cleaned_parts = []
    for p in raw_parts:
        p = YEAR_RANGE_RE.sub(" ", p)
        p = INSURED_PUNCT_RE.sub(" ", p)
        p = STOPWORDS_RE.sub(" ", p)
        # Bagian ini sudah uppercase dari kode aslinya
        p = MULTIPLE_SPACES_RE.sub(" ", p).strip().upper()
        
        if p:
            cleaned_parts.append(p)
            
    if not cleaned_parts:
        return [""]
        
    return cleaned_parts

def clean_polis(value):
    if pd.isna(value):
        return ""
    text_str = EXT_RE.sub("", str(value))
    return _clean_text_basic(text_str)

# Kolom CERTIFICATE tetap disediakan di output, namun logic pengisiannya
# sengaja dikosongkan karena belum ada kriteria pola yang sesuai.
def clean_certificate(value):
    return ""

# Aturan khusus SLIP ASKRINDO
def clean_slip(value):
    if pd.isna(value):
        return ""
        
    val_str = str(value).strip()
    
    # Deteksi pola khusus: tanda "-" diikuti 1 angka, "/" dan 2 angka (contoh: -1/02)
    if re.search(r"-\d/\d{2}(?!\d)", val_str):
        # 1. Hapus garis miring (/) dan 2 angka setelahnya
        val_str = re.sub(r"/\d{2}(?!\d)", "", val_str)
        # 2. Hapus tanda baca titik (.) dan strip (-)
        val_str = val_str.replace(".", "").replace("-", "")
        # 3. Bersihkan spasi berlebih
        return MULTIPLE_SPACES_RE.sub(" ", val_str).strip()
        
    return MULTIPLE_SPACES_RE.sub(" ", val_str).strip()

def process_data(input_file: str, sheet_name: str, output_file: str) -> None:
    if not os.path.exists(input_file):
        print(f"[ERROR] File tidak ditemukan di lokasi: {input_file}")
        return

    df = pd.read_excel(input_file, sheet_name=sheet_name, header=2)
    df.columns = df.columns.str.strip()

    if CEDANT_COL not in df.columns:
        print(f"[ERROR] Kolom '{CEDANT_COL}' tidak ditemukan di Excel!")
        return

    # Mencegah error jika penulisan nama kolom status berbeda
    status_col_name = "STATUS"
    for col in df.columns:
        if col.lower() == "status":
            status_col_name = col
            break

    total_raw_data = len(df)

    is_askrindo = df[CEDANT_COL] == CEDANT_VALUE
    
    if status_col_name in df.columns:
        is_suspense = df[status_col_name].astype(str).str.lower().str.strip() == "suspense"
        df_filtered = df[is_askrindo & is_suspense].copy()
    else:
        print(f"[WARNING] Kolom '{status_col_name}' tidak ditemukan, filter hanya berdasarkan CEDANT.")
        df_filtered = df[is_askrindo].copy()

    if df_filtered.empty:
        print(f"Tidak ada data untuk {CEDANT_VALUE} (dengan status SUSPENSE) di dalam file Excel.")
        return

    jumlah_sebelum_breakdown = len(df_filtered)

    insured_cleaned_series = df_filtered["INSURED"].apply(clean_insured)

    insured_expanded = pd.DataFrame(
        insured_cleaned_series.tolist(),
        index=df_filtered.index
    )
    insured_expanded.columns = [f"INSURED_CLEAN_{i + 1}" for i in range(insured_expanded.shape[1])]
    jumlah_kolom_insured = insured_expanded.shape[1]

    insert_pos = df_filtered.columns.get_loc("INSURED") + 1
    for offset, col_name in enumerate(insured_expanded.columns):
        df_filtered.insert(insert_pos + offset, col_name, insured_expanded[col_name])

    polis_idx = df_filtered.columns.get_loc("POLIS")

    polis_cleaned = df_filtered["POLIS"].apply(clean_polis)
    df_filtered.insert(polis_idx + 1, "POLIS_CLEAN", polis_cleaned)

    # Kolom CERTIFICATE tetap dibuat, tapi dikosongkan (lihat clean_certificate)
    certificate_values = df_filtered["SLIP NO"].apply(clean_certificate)
    df_filtered.insert(polis_idx + 2, "CERTIFICATE_1", certificate_values)

    slip_cleaned = df_filtered["SLIP NO"].apply(clean_slip)
    df_filtered.insert(df_filtered.columns.get_loc("SLIP NO") + 1, "SLIP_NO_CLEAN", slip_cleaned)

    print(f"Memfilter '{CEDANT_COL}' = '{CEDANT_VALUE}' dan '{status_col_name}' = 'suspense':")
    print(f"-> Total baris file mentah      : {total_raw_data}")
    print(f"-> Baris yang masuk kriteria    : {jumlah_sebelum_breakdown}")
    print(f"-> Kolom INSURED_CLEAN dibuat   : {jumlah_kolom_insured} kolom (INSURED_CLEAN_1..{jumlah_kolom_insured})")

    os.makedirs(os.path.dirname(output_file), exist_ok=True)
    df_filtered.to_excel(output_file, index=False)
    print("\nPemrosesan Data 3 ASKRINDO Selesai!")

if __name__ == "__main__":
    process_data(INPUT_FILE, SHEET_NAME, OUTPUT_FILE)