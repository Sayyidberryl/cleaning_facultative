import os
import re

import numpy as np
import pandas as pd


INPUT_FILE  = os.path.join("input", "1a. Transaksi Facul 01.01.23 - 17.07.26.xlsx")
OUTPUT_FILE = os.path.join("output", "wahana_output_facul.xlsx")

CEDANT_COL   = "COMP_NAME"
CEDANT_VALUE = "PT.ASURANSI WAHANA TATA"

POLIS_COL   = "FAC_POLICY_NO"
SLIP_COL    = "FAC_SLIP"
INSURED_COL = "FAC_INSURED"
BROKER_COL  = "COMP_NAME2"

MAX_SPLIT_COLS = 5

# Polis: kata/prefix yang menyebabkan nilai dibiarkan apa adanya
POLIS_EXCEPTION_RE = re.compile(
    r"""
    MOP\s*MARINE
  | (?:LINE\s*SLIP|LINESLIP)
  | \b(?:P1|P2|P3|P73)\s*CANCEL
  | \b(?:P1|P2|P3|P73)\b
  | \bCANCEL\b
  | PENYELESAIAN
  | HUTANG\s*PIUTANG
    """,
    re.IGNORECASE | re.VERBOSE,
)

# Slip: kata yang menyebabkan nilai dibiarkan apa adanya
SLIP_EXCEPTION_RE = re.compile(
    r"""
    \bSUMMARY\b
  | \bBORDER[OA]\b
  | \bBORDRO\b
  | \bSINGGLESHIPMENT\b
    """,
    re.IGNORECASE | re.VERBOSE,
)

MONTH_NAMES = frozenset({
    "JANUARI", "FEBRUARI", "MARET", "APRIL", "MEI", "JUNI",
    "JULI", "AGUSTUS", "SEPTEMBER", "OKTOBER", "NOVEMBER", "DESEMBER",
    "JANUARY", "FEBRUARY", "MARCH", "MAY", "JUNE", "JULY", "AUGUST",
    "OCTOBER",
})

# =========================
# INSURED CONFIG
# =========================

INSURED_REMOVE_RE = re.compile(
    r"""
    \b(
        PT|CV|TBK|
        PERSERO|
        LTD|PTE|INC|LLC|
        MR|MRS|MS|
        BAPAK|BPK|IBU|NY|
        DR|DRS|DRA|IR|H|HJ
    )\b
    """,
    re.IGNORECASE | re.VERBOSE,
)

INSURED_POLIS_RE = re.compile(
    r"""
    (POLIS\s*NO\.?.*)
    |
    (POLICY\s*NO\.?.*)
    |
    (SLIP\s*NO\.?.*)
    """,
    re.IGNORECASE | re.VERBOSE,
)

INSURED_SPLIT_RE = re.compile(
    r"""
    \s*,\s*
    |
    \s*/\s*
    |
    \s+QQ\s+
    |
    \s+AND/OR\s+
    |
    \s*&\s*
    |
    \s*\+\s*
    """,
    re.IGNORECASE | re.VERBOSE,
)

INSURED_JUNK_WORDS = {
    "",
    "AND",
    "OR",
    "THE",
    "OF",
    "AS"
}

# Noise token dalam slip
_SLIP_NOISE_RE = re.compile(
    r"""
    \b(?:JANUARI|FEBRUARI|MARET|APRIL|MEI|JUNI|JULI|AGUSTUS
        |SEPTEMBER|OKTOBER|NOVEMBER|DESEMBER)\b
  | \b(?:JANUARY|FEBRUARY|MARCH|APRIL|MAY|JUNE|JULY|AUGUST
        |SEPTEMBER|OCTOBER|NOVEMBER|DECEMBER)\b
  | \b(?:IDR|USD|ENG)\b
  | \b(?:P1|P2|P3|P73)\b
  | \bNEW\b
  | \bVARIOUS\b
  | \b20[0-9]{2}\b
    """,
    re.IGNORECASE | re.VERBOSE,
)

_SLIP_TOKEN_RE = re.compile(r"[A-Z0-9][A-Z0-9\-]{6,}", re.IGNORECASE)


# ─────────────────────────────────────────────────────────────────────────────
# HELPERS
# ─────────────────────────────────────────────────────────────────────────────

def _normalize_spaces(text: str) -> str:
    return re.sub(r"\s{2,}", " ", text).strip()


def _is_valid_polis_token(tok: str) -> bool:
    """Token polis valid: ≥5 char, punya digit, bukan TBA (murni/suffix)."""
    tok = tok.strip()
    return (
        len(tok) >= 5
        and not re.search(r"TBA$", tok, re.IGNORECASE)
        and bool(re.search(r"\d", tok))
    )


def _extract_polis_tokens(text: str) -> list:
    """Ekstrak semua token nomor polis valid dari teks."""
    tokens = []
    for block in re.split(r"\s{2,}", text.strip()):
        for tok in re.split(r"\+|\s+", block.strip()):
            tok = tok.strip().strip("-/")
            if _is_valid_polis_token(tok):
                tokens.append(tok)
    return tokens


def _expand_dash_chain(segment: str) -> list:
    """
    Pecah satu segmen berisi base polis + suffix-suffix pendek dipisah dash.

    Contoh:
      '010115001310-1311-1313'
          → ['010115001310', '010115001311', '010115001313']
      '010115001310-010115300373'
          → ['010115001310', '010115300373']  (keduanya base)

    Aturan:
    - Token ≥8 digit → base baru.
    - Token pendek (semua digit, len < len(base)) → suffix: ganti N digit terakhir base.
    - Lainnya → tambahkan apa adanya jika valid.
    """
    results = []
    current_base = None

    for part in segment.split("-"):
        part = part.strip()
        if not part:
            continue

        if len(part) >= 8 and re.search(r"\d", part):
            current_base = part
            results.append(part)
        elif (
            current_base
            and len(part) >= 2
            and re.match(r"^\d+$", part)
            and len(part) < len(current_base)
        ):
            n = len(part)
            results.append(current_base[:-n] + part)
        else:
            if _is_valid_polis_token(part):
                results.append(part)

    return results or ([segment] if _is_valid_polis_token(segment) else [])


def _is_valid_slip_token(tok: str) -> bool:
    """Token slip valid: ≥7 char, punya digit, bukan tahun 4-digit."""
    tok = tok.strip()
    return (
        len(tok) >= 7
        and bool(re.search(r"\d", tok))
        and not re.match(r"^\d{4}$", tok)
    )


def _strip_slip_suffix(tok: str) -> str:
    """Hapus suffix pendek numerik setelah dash (misal: -03, -000059)."""
    m = re.match(r"^(.+?)-(\d{1,6})$", tok)
    if m and len(m.group(1)) > len(m.group(2)):
        return m.group(1)
    return tok


def _extract_slip_tokens(text: str) -> list:
    """Ekstrak semua token nomor slip valid dari teks."""
    results = []
    for block in re.split(r"\s{2,}", text.strip()):
        clean = _SLIP_NOISE_RE.sub(" ", block)
        clean = re.sub(r"^[\s\-/+,]+|[\s\-/+,]+$", "", clean).strip()

        for cand in _SLIP_TOKEN_RE.findall(clean):
            cand = cand.strip("-")
            parts = cand.split("-")
            if len(parts) == 2:
                a, b = parts
                if _is_valid_slip_token(a) and _is_valid_slip_token(b) and abs(len(a) - len(b)) <= 2:
                    results.extend([a, b])
                    continue
            if _is_valid_slip_token(cand):
                results.append(_strip_slip_suffix(cand))

    return results


def _clean_insured_name(name: str) -> str:

    if pd.isna(name):
        return ""

    name = str(name).upper()

    # hapus info polis/slip
    name = INSURED_POLIS_RE.sub("", name)

    # hapus badan usaha dan gelar
    name = INSURED_REMOVE_RE.sub(" ", name)

    # ganti karakter
    name = re.sub(r"[()]", " ", name)
    name = name.replace("/", " ")
    name = name.replace("-", " ")

    # rapikan spasi
    name = _normalize_spaces(name)

    return name

def split_insured(name):

    if pd.isna(name):
        return []

    name = str(name)

    parts = INSURED_SPLIT_RE.split(name)

    hasil = []

    for p in parts:

        p = _clean_insured_name(p)

        if not p:
            continue

        if p in INSURED_JUNK_WORDS:
            continue

        hasil.append(p)

    return hasil


def _cap_or_join(items: list) -> list:
    """Jika jumlah item melebihi MAX_SPLIT_COLS, gabungkan dengan koma."""
    if len(items) > MAX_SPLIT_COLS:
        return [",".join(items)]
    return items


# ─────────────────────────────────────────────────────────────────────────────
# CLEAN POLIS  (facul memiliki tambahan: dot-suffix, ampersand, _expand_dash_chain)
# ─────────────────────────────────────────────────────────────────────────────

def clean_polis(val) -> list:
    if pd.isna(val):
        return []

    val = str(val).strip()
    if not val:
        return []

    # exception
    if POLIS_EXCEPTION_RE.search(val):
        return [_normalize_spaces(val)]

    if re.search(r"\d+TBA\d+", val, re.IGNORECASE):
        return [_normalize_spaces(val)]

    val = _normalize_spaces(val)
    val_upper = val.upper()

    # ===============================
    # RULE 1
    # Hapus suffix /01 /02 /03 dst
    # Contoh:
    # 022117004330/02 -> 022117004330
    # ===============================
    val = re.sub(r"/\d{1,3}$", "", val)

    # ===============================
    # RULE 2
    # Jika hanya angka + titik
    # hapus semua titik
    # ===============================
    if re.fullmatch(r"[0-9.]+", val):
        return [val.replace(".", "")]

    # ===============================
    # RULE 3
    # Jika hanya angka + strip
    # hapus semua strip
    # ===============================
    if re.fullmatch(r"[0-9-]+", val):
        return [val.replace("-", "")]

    # ===============================
    # RULE 4
    # Polis yang ada huruf → tetap pertahankan isi,
    # tetapi hapus titik
    if re.search(r"[A-Z]", val_upper):
        return [val.replace(".", "")]

    # ===============================
    # RULE 5
    # Multi polis &
    # ===============================
    if "&" in val:
        hasil = []

        for p in val.split("&"):
            p = p.strip()

            p = re.sub(r"/\d{1,3}$", "", p)

            if re.fullmatch(r"[0-9.]+", p):
                p = p.replace(".", "")

            elif re.fullmatch(r"[0-9-]+", p):
                p = p.replace("-", "")

            if p:
                hasil.append(p)

        return _cap_or_join(hasil)

    # ===============================
    # HAPUS P1, P2, P3, P4, ...
    # ===============================

    val = re.sub(r"\bP\d+\b", "", val, flags=re.IGNORECASE)

    # Rapikan tanda + yang tersisa
    val = re.sub(r"\+\s*\+", "+", val)
    val = re.sub(r"^\s*\+\s*|\s*\+\s*$", "", val)

    val = _normalize_spaces(val)

    # ===============================
    # RULE 6
    # Multi polis +
    # ===============================
    if "+" in val:
        hasil = []

        for p in val.split("+"):
            p = p.strip()

            p = re.sub(r"/\d{1,3}$", "", p)

            if re.fullmatch(r"[0-9.]+", p):
                p = p.replace(".", "")

            elif re.fullmatch(r"[0-9-]+", p):
                p = p.replace("-", "")

            if p:
                hasil.append(p)

        return _cap_or_join(hasil)

    # ===============================
    # RULE 7
    # Polis biasa
    # ===============================
    return [val]
# ─────────────────────────────────────────────────────────────────────────────
# CLEAN SLIP
# ─────────────────────────────────────────────────────────────────────────────

def clean_slip(val) -> list:
    if pd.isna(val):
        return []

    val = str(val).upper().strip()

    if not val:
        return []

    val = _normalize_spaces(val)

    # ==========================================
    # Rule 1 : Samakan delimiter multi slip
    # ==========================================
    val = re.sub(r"\s*&\s*", "/", val)
    val = re.sub(r"\s*\+\s*", "/", val)

    slips = []

    for s in val.split("/"):
        s = s.strip()

        if not s:
            continue

        # ==========================================
        # Rule 2 : Hapus keterangan di belakang
        # ==========================================
        s = re.sub(
            r"\b("
            r"USD|IDR|SGD|EUR|JPY|AUD|GBP|"
            r"ORI|ORI\.|ORIGINAL|COPY|"
            r"REALISASI|REALIZATION|"
            r"CANCEL|CANCELLED|"
            r"ENDT?|ENDORSEMENT|"
            r"SA|P1|P2|P3|VAR|REVISI|REV"
            r")\b.*$",
            "",
            s,
            flags=re.IGNORECASE,
        )

        # ==========================================
        # Rule 3 : Rapikan spasi
        # ==========================================
        s = _normalize_spaces(s)

        # ==========================================
        # Rule 4 : Hapus karakter di depan/belakang
        # ==========================================
        s = s.strip("-_,.; ")

        # ==========================================
        # Rule 4 : Hapus karakter di depan/belakang
        # ==========================================
        s = s.strip("-_,.; ")

        # Hapus titik di dalam slip
        s = s.replace(".", "")

        if not s:
            continue

        if not s:
            continue

        if s not in slips:
            slips.append(s)

    return _cap_or_join(slips)
# ─────────────────────────────────────────────────────────────────────────────
# CLEAN INSURED
# ─────────────────────────────────────────────────────────────────────────────

def clean_insured(val) -> list:

    if pd.isna(val):
        return []

    val = str(val).strip()

    if not val:
        return []

    # Breakdown sesuai separator
    insureds = split_insured(val)

    cleaned = []

    for ins in insureds:

        ins = _normalize_spaces(ins)

        if not ins:
            continue

        if len(ins) <= 2:
            continue

        if ins.upper() in INSURED_JUNK_WORDS:
            continue

        # huruf kapital
        ins = ins.upper()

        if ins not in cleaned:
            cleaned.append(ins)

    # kalau tidak berhasil dibreakdown
    if not cleaned:

        fallback = _clean_insured_name(val)

        if fallback:
            cleaned.append(fallback.upper())

    return _cap_or_join(cleaned)
# ─────────────────────────────────────────────────────────────────────────────
# MITRA BISNIS
# ─────────────────────────────────────────────────────────────────────────────

def get_mitra_bisnis(comp_name2, comp_name) -> str:
    """
    Tentukan mitra_bisnis:
    - COMP_NAME2 kosong / 'DIRECT' → gunakan COMP_NAME (cedant)
    - Selain itu → gunakan COMP_NAME2 (broker)
    """
    broker = "" if pd.isna(comp_name2) else str(comp_name2).strip()
    if not broker or broker.upper() == "DIRECT":
        return "" if pd.isna(comp_name) else str(comp_name).strip()
    return broker

def clean_business_partners(val):

    if pd.isna(val):
        return ""

    val = str(val).strip().upper()

    # PT. -> PT
    val = re.sub(
        r"\bPT\.\s*",
        "PT ",
        val,
        flags=re.IGNORECASE
    )

    val = _normalize_spaces(val)

    return val

# ─────────────────────────────────────────────────────────────────────────────
# PROCESS DATA
# ─────────────────────────────────────────────────────────────────────────────

def _insert_clean_columns(
    df: pd.DataFrame,
    all_lists: list,
    prefix: str,
    max_cols: int,
) -> list:
    """Buat kolom clean_{prefix}_N dan kembalikan nama-nama kolom baru."""
    added = []
    for i in range(1, max_cols + 1):
        col_name = f"clean {prefix} {i}"
        df[col_name] = [lst[i - 1] if i - 1 < len(lst) else None for lst in all_lists]
        added.append(col_name)
    return added


def _fast_read_excel(path: str, sheet_name=None, header: int = 0) -> pd.DataFrame:
    """Baca file Excel besar lebih cepat lewat openpyxl read_only + iter_rows."""
    import openpyxl
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb[sheet_name] if sheet_name else wb.worksheets[0]
    rows_iter = ws.iter_rows(values_only=True)
    for _ in range(header):
        next(rows_iter)
    cols = next(rows_iter)
    data = list(rows_iter)
    wb.close()
    return pd.DataFrame(data, columns=cols)


def _fast_write_excel(df: pd.DataFrame, path: str) -> None:
    """Tulis DataFrame besar ke xlsx lebih cepat lewat openpyxl write_only mode."""
    import openpyxl
    os.makedirs(os.path.dirname(path), exist_ok=True)

    df2 = df.astype(object).where(pd.notnull(df), None)

    wb = openpyxl.Workbook(write_only=True)
    ws = wb.create_sheet("Sheet1")
    ws.append(list(df2.columns))
    for row in df2.itertuples(index=False, name=None):
        ws.append(row)
    wb.save(path)


def process_data(input_file: str, output_file: str) -> None:
    print(f"[1/5] Membaca data dari: {input_file} ...")
    df = _fast_read_excel(input_file, header=0)
    print(f"      Total baris keseluruhan: {len(df):,}")

    if CEDANT_COL not in df.columns:
        print(f"\n[ERROR] Kolom '{CEDANT_COL}' tidak ditemukan!")
        print(f"        Kolom tersedia: {list(df.columns)}")
        return

    required_cols = [POLIS_COL, SLIP_COL, INSURED_COL, BROKER_COL]
    for col in required_cols:
        if col not in df.columns:
            print(f"\n[ERROR] Kolom '{col}' tidak ditemukan di file!")
            print(f"        Kolom tersedia: {list(df.columns)}")
            return

    df[CEDANT_COL] = df[CEDANT_COL].astype(str).str.strip()
    is_wahana = df[CEDANT_COL] == CEDANT_VALUE
    print(f"[2/5] Filter cedant '{CEDANT_VALUE}': {is_wahana.sum():,} baris WAHANA dari total {len(df):,} baris.")

    # Hitung mitra_bisnis sebelum rename (butuh akses COMP_NAME & COMP_NAME2)
    broker_s = df[BROKER_COL].fillna("").astype(str).str.strip()
    cedant_s = df[CEDANT_COL].fillna("").astype(str).str.strip()
    use_cedant = (broker_s == "") | (broker_s.str.upper() == "DIRECT")
    mitra_values = [
    clean_business_partners(x)
    for x in np.where(use_cedant, cedant_s, broker_s)
]

    df.rename(columns={POLIS_COL: "polis_ori", SLIP_COL: "slip_ori", INSURED_COL: "insured_ori"},
              inplace=True)

    # Sisipkan kolom mitra_bisnis di sebelah kanan BROKER_COL
    insert_pos = list(df.columns).index(BROKER_COL) + 1 if BROKER_COL in df.columns else len(df.columns)
    df.insert(insert_pos, "BUSINESS PARTNERS", mitra_values)

    print("[3/5] Menjalankan proses cleaning hanya untuk baris WAHANA ...")

    n = len(df)
    all_clean_polis = [[] for _ in range(n)]
    all_clean_slip  = [[] for _ in range(n)]
    all_clean_ins   = [[] for _ in range(n)]
    max_polis = max_slip = max_ins = 1

    wahana_idx = np.flatnonzero(is_wahana.to_numpy())
    polis_vals   = df["polis_ori"].to_numpy()
    slip_vals    = df["slip_ori"].to_numpy()
    insured_vals = df["insured_ori"].to_numpy()

    total_w = len(wahana_idx)
    for n_done, pos in enumerate(wahana_idx, 1):
        if n_done % 5_000 == 0:
            print(f"      Progress: {n_done:,} / {total_w:,} baris WAHANA diproses...")

        c_polis = clean_polis(polis_vals[pos])
        c_slip  = clean_slip(slip_vals[pos])
        c_ins   = clean_insured(insured_vals[pos])

        max_polis = max(max_polis, len(c_polis))
        max_slip  = max(max_slip,  len(c_slip))
        max_ins   = max(max_ins,   len(c_ins))

        all_clean_polis[pos] = c_polis
        all_clean_slip[pos]  = c_slip
        all_clean_ins[pos]   = c_ins

    print(f"      Selesai diproses!")
    print(f"      -> Jumlah kolom clean polis  : {max_polis}")
    print(f"      -> Jumlah kolom clean slip   : {max_slip}")
    print(f"      -> Jumlah kolom clean insured: {max_ins}")

    print("[4/5] Menyusun kolom output ...")

    new_columns = []
    for col in df.columns:
        new_columns.append(col)
        if col == "polis_ori":
            new_columns += _insert_clean_columns(df, all_clean_polis, "polis",   max_polis)
        elif col == "slip_ori":
            new_columns += _insert_clean_columns(df, all_clean_slip,  "slip",    max_slip)
        elif col == "insured_ori":
            new_columns += _insert_clean_columns(df, all_clean_ins,   "insured", max_ins)

    df = df[new_columns]

    print(f"[5/5] Menyimpan hasil ke: {output_file} ...")
    _fast_write_excel(df, output_file)

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