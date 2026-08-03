import os
import re

import numpy as np
import pandas as pd


INPUT_FILE  = os.path.join("data", "suspend.xlsx")
OUTPUT_FILE = os.path.join("data", "suspend_clean_aca.xlsx")

CEDANT_FILTER_COL   = "CEDANT SHRT NAME"
CEDANT_FILTER_VALUE = "CENTRAL"

# Regex untuk sub-removal dalam nama insured
_INSURED_TAIL_RE = re.compile(
    r"""
    \bAS\b\s+(?:THE\s+)?(?:PRINCIPAL|OFF-TAKER|MAINTENANCE|CONTRACTOR).*
  | \bBEING\s+(?:THE\s+)?(?:PRINCIPAL|OFF-TAKER).*
  | \bAND\s+ALL\s+SUBSIDIARI.*
  | \bINCLUDING\s+ALL\s+SUBSIDIARI.*
  | \bINCLUDING\s+ANY\s+SUBSIDIAR.*
  | \bCOMPRISING\s+OF.*
  | \bINSTALLMENT\b.*
  | \bRELATED\s+COMPANY\b.*
  | \bPURCHASED\s+OR\s+OTHERWISE\b.*
    """,
    re.IGNORECASE | re.VERBOSE,
)

_INSURED_JUNK_WORDS = frozenset({
    "SHANGHAI", "PR OF CHINA", "CHINA", "INDONESIA", "JAKARTA",
    "OFFICERS", "EMPLOYEES", "ALL OTHER CONTRACTORS",
    "SUB-CONTRACTORS", "SUB CONTRACTORS",
    "COMPANIES", "AFFILIATED", "AFFILIATES",
    "CORPORATIONS AND INCLUDING PARTNERSHIP",
    "JOINT VENTURES AND AGREEMENT OR BY LAW",
    "AS THEIR RESPECTIVE INTEREST MAY APPEAR",
    "AS THEIR RESPECTIVE INTERESTS MAY APPEAR",
    "SUBSIDIARY", "SUBSIDIARIES", "ANY SUBSIDIARY COMPANY",
    "RELATED COMPANY",
    "FOR THEIRS RESPECTIVE RIGHTS AND INTEREST",
    "MIGRASI AS400", "THE PRINCIPAL", "PRINCIPAL", "OWNER",
})


# ─────────────────────────────────────────────────────────────────────────────
# CLEAN FUNCTIONS
# ─────────────────────────────────────────────────────────────────────────────

def clean_polis(val) -> list:
    """Hapus suffix numerik di akhir nomor polis (misal: -01, -02/03)."""
    if pd.isna(val):
        return []
    val = str(val).strip()
    if not val:
        return []

    # Strip suffix "-NNN" atau "-NNN/NNN" secara iteratif
    p = val
    while True:
        stripped = re.sub(r"-\d+(?:/\d+)?$", "", p)
        if stripped == p:
            break
        p = stripped

    return [p] if p else []


def clean_slip(val) -> list:
    """Kembalikan nilai slip apa adanya (no transformation needed)."""
    if pd.isna(val):
        return []
    val = str(val).strip()
    return [val] if val else []


def _remove_polis_slip_from_text(text: str, polis_ori, slip_ori) -> str:
    """Hapus nomor polis & slip dari teks insured secara agresif."""
    if pd.notna(polis_ori):
        for token in [str(polis_ori).strip()] + clean_polis(polis_ori):
            if token and token != "-":
                text = text.replace(token, "")

    if pd.notna(slip_ori):
        for token in [str(slip_ori).strip()] + clean_slip(slip_ori):
            if token and token != "-":
                text = text.replace(token, "")

    return text


def _normalize_insured_part(p: str) -> str:
    """Terapkan sub-removal dan normalisasi pada satu bagian nama insured."""
    p = _INSURED_TAIL_RE.sub("", p)
    p = re.sub(r"\bKB\b",    "", p, flags=re.IGNORECASE)
    p = re.sub(r"\bA\.?W\.?\b", "", p, flags=re.IGNORECASE)
    p = re.sub(r"\(\s*\)",   "", p)
    # Trim leading/trailing AND/OR
    p = re.sub(r"^(?:AND|OR)\b\s*", "", p, flags=re.IGNORECASE)
    p = re.sub(r"\s*\b(?:AND|OR)$", "", p, flags=re.IGNORECASE)
    # Trim leading/trailing non-alphanumeric
    p = re.sub(r"^[^a-zA-Z0-9(]+", "", p)
    p = re.sub(r"[^a-zA-Z0-9)]+$", "", p)
    return p.strip()


def _is_valid_insured_part(p: str) -> bool:
    """Return True jika bagian insured layak dipertahankan."""
    if len(p) <= 2:
        return False
    up = p.upper()
    if re.match(r"^[\d\/\-\.]+$", up):
        return False
    if re.search(r"\b(?:NO\.\s*\d+|BUILDING|FLOOR|ROOM|ROAD|STREET|TOWER|KAV\.?|BLOK)\b", up):
        return False
    if re.search(r"\b(?:PLTGU|PLTMH|PLTU|POWER PLANT|COMBINED CYCLE|MW|HYDRO ELECTRIC)\b", up):
        return False
    return up not in _INSURED_JUNK_WORDS


def clean_insured(val, polis_ori, slip_ori) -> list:
    """Bersihkan nama insured dengan menghapus nomor polis/slip dan junk words."""
    if pd.isna(val):
        return []
    val = str(val).strip()
    if not val:
        return []

    val = _remove_polis_slip_from_text(val, polis_ori, slip_ori)

    # Normalisasi sebelum split
    val = re.sub(r"\(\s*[\d\.\/\-]+\s*\)", "", val)
    val = re.sub(r"\b(?:AND|AN|OR)\s*/\s*(?:AND|OR)\b", ",", val, flags=re.IGNORECASE)
    val = re.sub(r"\bAND\s+OR\b", ",", val, flags=re.IGNORECASE)
    val = re.sub(r"\bCO\.,?\s*LTD\.?\b", ",", val, flags=re.IGNORECASE)
    for pattern in [r"\bTBK\.?\b", r"\(PERSERO\)", r"\bPERSERO\b", r"\bLTD\.?\b", r"\(FCI\.\s*I\)"]:
        val = re.sub(pattern, "", val, flags=re.IGNORECASE)

    split_pattern = r"\bQQ\b|/|,|-(?!\s*(?:19|20)\d{2}\b)|\d+\.|\bPT\.?\b|\bCV\.?\b|:|;"
    parts = re.split(split_pattern, val, flags=re.IGNORECASE)

    cleaned = []
    for p in parts:
        p = _normalize_insured_part(p)
        if _is_valid_insured_part(p):
            cleaned.append(p)

    return cleaned


# ─────────────────────────────────────────────────────────────────────────────
# COLUMN EXPANSION HELPER
# ─────────────────────────────────────────────────────────────────────────────

def _expand_clean_columns(
    df: pd.DataFrame,
    all_lists: list,
    ori_col: str,
    prefix: str,
    max_cols: int,
) -> list:
    """
    Tambahkan kolom clean_{prefix}_1 .. N langsung setelah kolom ori.
    Returns daftar nama kolom baru yang ditambahkan.
    """
    added = []
    for i in range(1, max_cols + 1):
        col_name = f"clean {prefix} {i}"
        df[col_name] = [lst[i - 1] if i - 1 < len(lst) else None for lst in all_lists]
        added.append(col_name)
    return added


# ─────────────────────────────────────────────────────────────────────────────
# PROCESS DATA
# ─────────────────────────────────────────────────────────────────────────────

def process_data(input_file: str, output_file: str) -> None:
    print(f"[1/5] Membaca data dari: {input_file} ...")
    df = pd.read_excel(input_file, header=2)
    print(f"      Total baris keseluruhan: {len(df):,}")

    if CEDANT_FILTER_COL not in df.columns:
        print(f"\n[ERROR] Kolom '{CEDANT_FILTER_COL}' tidak ditemukan!")
        print(f"        Kolom tersedia: {list(df.columns)}")
        return

    df = df[df[CEDANT_FILTER_COL] == CEDANT_FILTER_VALUE].copy()
    print(f"[2/5] Filter cedant '{CEDANT_FILTER_VALUE}': {len(df):,} baris ditemukan.")

    if df.empty:
        print("\n[WARN] Tidak ada data setelah filter. Proses dihentikan.")
        return

    # Rename kolom asli → _ori
    rename_map = {"INSURED": "insured_ori", "POLIS": "polis_ori", "SLIP NO": "slip_ori"}
    df.rename(columns={k: v for k, v in rename_map.items() if k in df.columns}, inplace=True)

    print("[3/5] Menjalankan proses cleaning ...")

    all_clean_polis, all_clean_slip, all_clean_ins = [], [], []
    max_polis = max_slip = max_ins = 1

    for _, row in df.iterrows():
        p_ori = row.get("polis_ori", "")
        s_ori = row.get("slip_ori", "")
        i_ori = row.get("insured_ori", "")

        c_polis = clean_polis(p_ori)
        c_slip  = clean_slip(s_ori)
        c_ins   = clean_insured(i_ori, p_ori, s_ori)

        max_polis = max(max_polis, len(c_polis))
        max_slip  = max(max_slip,  len(c_slip))
        max_ins   = max(max_ins,   len(c_ins))

        all_clean_polis.append(c_polis)
        all_clean_slip.append(c_slip)
        all_clean_ins.append(c_ins)

    print("[4/5] Menyusun kolom output ...")

    # Sisipkan kolom clean langsung setelah kolom _ori-nya
    new_columns = []
    for col in df.columns:
        new_columns.append(col)
        if col == "polis_ori":
            new_columns += _expand_clean_columns(df, all_clean_polis, col, "polis",   max_polis)
        elif col == "slip_ori":
            new_columns += _expand_clean_columns(df, all_clean_slip,  col, "slip",    max_slip)
        elif col == "insured_ori":
            new_columns += _expand_clean_columns(df, all_clean_ins,   col, "insured", max_ins)

    df = df[new_columns]

    print(f"[5/5] Menyimpan hasil ke: {output_file} ...")
    df.to_excel(output_file, index=False)

    print(f"\n{'=' * 55}")
    print(f"  [OK] Selesai!")
    print(f"  Total baris output  : {len(df):,}")
    print(f"  Total kolom output  : {len(df.columns)}")
    print(f"  File disimpan di    : {output_file}")
    print(f"{'=' * 55}")


# ─────────────────────────────────────────────────────────────────────────────
# ENTRY POINT
# ─────────────────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    process_data(INPUT_FILE, OUTPUT_FILE)
