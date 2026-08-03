import os
import re

import numpy as np
import pandas as pd


# ─────────────────────────────────────────────────────────────────────────────
# KONFIGURASI
# ─────────────────────────────────────────────────────────────────────────────

INPUT_FILE  = os.path.join("input", "3a. Database Suspense 170726.xlsx")
OUTPUT_FILE = os.path.join("output", "Suspend_Clean_WAHANA.xlsx")
SHEET_NAME  = "Detail Database"   # sheet yang berisi data mentah (row 1-2 kosong/judul, header di row 3)

# Kolom "comp_name" di data ini bernama CEDANT SHRT NAME
COMPNAME_FILTER_COL   = "CEDANT SHRT NAME"
COMPNAME_FILTER_VALUE = "WAHANA T."

# Business rule: kalau hasil breakdown (insured/polis/slip) > 5 bagian,
# tidak usah dipecah per kolom -> digabung lagi jadi satu kolom.
MAX_BREAKDOWN = 5

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
  | \bWPC\s+DATE\s*:.*
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

# Gelar / sapaan yang ikut dihapus dari nama insured (Bapak, Ibu, Ny, Mr, Mrs, Ms, dll)
_INSURED_TITLE_RE = re.compile(
    r"\b(?:BAPAK|IBU|BPK|NY\.?|SDR\.?|SDRI\.?|MR\.?|MRS\.?|MS\.?)\b\s*",
    re.IGNORECASE,
)


# ─────────────────────────────────────────────────────────────────────────────
# CLEAN FUNCTIONS
# ─────────────────────────────────────────────────────────────────────────────

def _cap_breakdown(parts: list, sep: str = "; ") -> list:
    """
    Business rule: jika jumlah bagian hasil breakdown > MAX_BREAKDOWN,
    jangan dipecah per kolom -> gabungkan lagi jadi satu nilai.
    """
    parts = [p for p in parts if p]
    if len(parts) > MAX_BREAKDOWN:
        return [sep.join(parts)]
    return parts


def clean_polis(val) -> list:
    """Hapus suffix numerik di akhir nomor polis (misal: -01, -02/03)."""
    if pd.isna(val):
        return []
    val = str(val).strip()
    if not val or val == "-":
        return []

    # Hapus koma (mis. "000.6005.201.2025.002555.00," -> "000.6005.201.2025.002555.00")
    val = val.replace(",", "")

    # Kalau ada lebih dari satu nomor polis dipisah "+", proses masing-masing
    raw_parts = [p.strip() for p in val.split("+") if p.strip()]
    if not raw_parts:
        return []

    cleaned_parts = []
    for p in raw_parts:
        cur = p
        while True:
            stripped = re.sub(r"-\d+(?:/\d+)?$", "", cur)
            if stripped == cur:
                break
            cur = stripped
        if cur:
            cur = cur.replace(".", "")  # hapus titik dari nomor polis
            cleaned_parts.append(cur)

    return _cap_breakdown(cleaned_parts)


def clean_slip(val) -> list:
    """Kembalikan nilai slip (tanpa titik), pecah kalau ada lebih dari satu (+)."""
    if pd.isna(val):
        return []
    val = str(val).strip()
    if not val or val == "-":
        return []

    raw_parts = [p.strip().replace(".", "") for p in val.split("+") if p.strip()]
    return _cap_breakdown(raw_parts)


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
    p = _INSURED_TITLE_RE.sub("", p)
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
    """Bersihkan nama insured dengan menghapus nomor polis/slip, gelar, dan junk words."""
    if pd.isna(val):
        return []
    val = str(val).strip()
    if not val:
        return []

    val = _remove_polis_slip_from_text(val, polis_ori, slip_ori)

    # Hapus tail info (mis. "AS PRINCIPAL...", "WPC Date : ...") SEBELUM split,
    # karena split_pattern (":" dll) bisa memotong pola ini duluan.
    val = _INSURED_TAIL_RE.sub("", val)

    # Normalisasi sebelum split
    # Hapus semua isi dalam kurung, mis. "(FACULTATIVE)", "(12.34/56)", dll
    val = re.sub(r"\([^)]*\)", "", val)
    val = re.sub(r"\b(?:AND|AN|OR)\s*/\s*(?:AND|OR)\b", ",", val, flags=re.IGNORECASE)
    val = re.sub(r"\bAND\s+OR\b", ",", val, flags=re.IGNORECASE)
    val = re.sub(r"\bCO\.,?\s*LTD\.?\b", ",", val, flags=re.IGNORECASE)
    for pattern in [r"\bTBK\.?\b", r"\(PERSERO\)", r"\bPERSERO\b", r"\bLTD\.?\b", r"\(FCI\.\s*I\)"]:
        val = re.sub(pattern, "", val, flags=re.IGNORECASE)

    # Delimiter pemisah insured: , / QQ and/or & – + (sesuai catatan rules)
    split_pattern = (
        r"\bQQ\b|/|,|&|\+"
        r"|-(?!\s*(?:19|20)\d{2}\b)"
        r"|\d+\.|\bPT\.?\b|\bCV\.?\b|:|;"
    )
    parts = re.split(split_pattern, val, flags=re.IGNORECASE)

    cleaned = []
    for p in parts:
        p = _normalize_insured_part(p)
        
        if _is_valid_insured_part(p):
            cleaned.append(p.upper())

    return _cap_breakdown(cleaned)


# ─────────────────────────────────────────────────────────────────────────────
# COLUMN EXPANSION HELPER
# ─────────────────────────────────────────────────────────────────────────────

def _expand_clean_columns(
    df: pd.DataFrame,
    all_lists: list,
    prefix: str,
    max_cols: int,
) -> list:
    """
    Tambahkan kolom clean_{prefix}_1 .. N langsung setelah kolom ori.
    max_cols otomatis dibatasi MAX_BREAKDOWN karena _cap_breakdown()
    sudah menggabungkan hasil > MAX_BREAKDOWN jadi 1 kolom.
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
    print(f"[1/5] Membaca data dari: {input_file} (sheet: {SHEET_NAME}) ...")
    df = pd.read_excel(input_file, sheet_name=SHEET_NAME, header=2)
    print(f"      Total baris keseluruhan: {len(df):,}")

    if COMPNAME_FILTER_COL not in df.columns:
        print(f"\n[ERROR] Kolom '{COMPNAME_FILTER_COL}' tidak ditemukan!")
        print(f"        Kolom tersedia: {list(df.columns)}")
        return

    df[COMPNAME_FILTER_COL] = df[COMPNAME_FILTER_COL].astype(str).str.strip()
    is_wahana = df[COMPNAME_FILTER_COL] == COMPNAME_FILTER_VALUE
    print(f"[2/5] Filter comp_name '{COMPNAME_FILTER_VALUE}': {is_wahana.sum():,} baris WAHANA dari total {len(df):,} baris.")

    # Rename kolom asli -> _ori
    rename_map = {"INSURED": "insured_ori", "POLIS": "polis_ori", "SLIP NO": "slip_ori"}
    df.rename(columns={k: v for k, v in rename_map.items() if k in df.columns}, inplace=True)

    print("[3/5] Menjalankan proses cleaning hanya untuk baris WAHANA (baris lain dibiarkan apa adanya) ...")

    all_clean_polis, all_clean_slip, all_clean_ins = [], [], []
    max_polis = max_slip = max_ins = 1

    for idx, (_, row) in enumerate(df.iterrows()):
        p_ori = row.get("polis_ori", "")
        s_ori = row.get("slip_ori", "")
        i_ori = row.get("insured_ori", "")

        # Hanya cleaning baris WAHANA, non-WAHANA dikosongkan (dibiarkan apa adanya)
        if is_wahana.iloc[idx]:
            c_polis = clean_polis(p_ori)
            c_slip  = clean_slip(s_ori)
            c_ins   = clean_insured(i_ori, p_ori, s_ori)
        else:
            c_polis, c_slip, c_ins = [], [], []

        max_polis = max(max_polis, len(c_polis))
        max_slip  = max(max_slip,  len(c_slip))
        max_ins   = max(max_ins,   len(c_ins))

        all_clean_polis.append(c_polis)
        all_clean_slip.append(c_slip)
        all_clean_ins.append(c_ins)

    # Safety net: breakdown tidak boleh lebih dari MAX_BREAKDOWN kolom
    max_polis = min(max_polis, MAX_BREAKDOWN)
    max_slip  = min(max_slip,  MAX_BREAKDOWN)
    max_ins   = min(max_ins,   MAX_BREAKDOWN)

    print("[4/5] Menyusun kolom output ...")

    # Sisipkan kolom clean langsung setelah kolom _ori-nya
    new_columns = []
    for col in df.columns:
        new_columns.append(col)
        if col == "polis_ori":
            new_columns += _expand_clean_columns(df, all_clean_polis, "polis",   max_polis)
        elif col == "slip_ori":
            new_columns += _expand_clean_columns(df, all_clean_slip,  "slip",    max_slip)
        elif col == "insured_ori":
            new_columns += _expand_clean_columns(df, all_clean_ins,   "insured", max_ins)

    df = df[new_columns]

    print(f"[5/5] Menyimpan hasil ke: {output_file} ...")
    os.makedirs(os.path.dirname(output_file), exist_ok=True)
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