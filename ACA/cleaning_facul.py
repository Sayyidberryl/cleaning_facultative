import os
import re

import pandas as pd


INPUT_FILE  = os.path.join("data", "facul.xlsx")
OUTPUT_FILE = os.path.join("data", "facul_clean_aca.xlsx")

CEDANT_COL   = "COMP_NAME"
CEDANT_VALUE = "PT. ASURANSI CENTRAL ASIA"

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

INSURED_JUNK_WORDS = frozenset({
    "PT", "CV", "TBK", "PERSERO", "LTD", "INC", "LLC",
    "AND", "OR", "THE", "OF", "AS",
    "NON FOOD", "DIV",
})

INSURED_SUFFIX_RE = re.compile(
    r"""
    ,?\s*\bTBK\b\s*(?:,?\s*PT\.?)?
  | ,?\s*\bPT\.?\s*$
  | ,?\s*\bCV\.?\s*$
    """,
    re.IGNORECASE | re.VERBOSE,
)

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
    """Hapus suffix badan usaha (TBK, PT, CV) secara iteratif."""
    name = _normalize_spaces(name)
    for _ in range(3):
        cleaned = _normalize_spaces(INSURED_SUFFIX_RE.sub("", name).strip().strip(","))
        if cleaned == name:
            break
        name = cleaned
    return name


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

    if POLIS_EXCEPTION_RE.search(val):
        return [_normalize_spaces(val)]
    if re.search(r"\d+TBA\d+", val, re.IGNORECASE):
        return [_normalize_spaces(val)]

    val_upper = val.upper().strip()

    if re.match(r"^VARIOUS\s*$", val_upper):
        return [val.strip()]
    if re.match(r"^VARIOUS\s*-\s*SEE\s+ATTACH", val_upper):
        return [val.strip()]
    if re.match(r"^TBA\s*$", val_upper):
        return [val.strip()]

    # Pola titik-suffix: "010114003555.3431" → ["010114003555"]
    m = re.match(r"^(\d[\d\-]+)\.\d{1,6}$", val.strip())
    if m:
        base = m.group(1).strip()
        if _is_valid_polis_token(base):
            return [base]

    # Pola S/D: base_polis-suffixS/D suffix2
    if re.search(r"S/D", val, re.IGNORECASE):
        sd_parts = re.split(r"\s{2,}", val.strip())
        left = sd_parts[0] if sd_parts else val
        base_m = re.match(r"^(\d[\d\-]+?)-\d+S/D\d+", left, re.IGNORECASE) or \
                 re.match(r"^(\d[\d\-]+)S/D\d+",       left, re.IGNORECASE)

        results = [base_m.group(1).strip()] if base_m else []
        for part in sd_parts[1:]:
            part = re.sub(r"\+?\s*\bTBA\b\s*", "", part, flags=re.IGNORECASE).strip().strip("+")
            results += [t for t in part.split() if _is_valid_polis_token(t.strip())]

        if results:
            return _cap_or_join(results)

    # Pola TBA di awal diikuti spasi ganda
    if re.match(r"^\s*TBA\s{2,}", val, re.IGNORECASE):
        rest = re.sub(r"^\s*TBA\s+", "", val, flags=re.IGNORECASE).strip()
        tokens = _extract_polis_tokens(rest)
        if tokens:
            return _cap_or_join(tokens)

    # Pola & (ampersand) sebagai pemisah polis lengkap
    if "&" in val:
        blocks_dd  = re.split(r"\s{2,}", val.strip())
        left_block = blocks_dd[0]
        right_blocks = blocks_dd[1:]

        results = []
        for ap in [p.strip() for p in re.split(r"\s*&\s*", left_block) if p.strip()]:
            results.extend(_expand_dash_chain(ap))

        for rb in right_blocks:
            rb_clean = re.sub(r"\+?\s*\bTBA\b\s*", "", rb, flags=re.IGNORECASE).strip()
            results += [t for t in rb_clean.split() if _is_valid_polis_token(t.strip())]

        if results:
            return _cap_or_join(results)

    # Pola + (suffix pendek atau multi polis)
    if "+" in val:
        blocks     = re.split(r"\s{2,}", val.strip())
        left_parts = [p.strip() for p in blocks[0].split("+") if p.strip()
                      and not re.match(r"^TBA$", p.strip(), re.IGNORECASE)]
        right_blocks = blocks[1:]

        if left_parts:
            first      = left_parts[0]
            rest_parts = left_parts[1:]

            # Guard: semua bagian pendek → biarkan apa adanya
            if all(len(p) < 7 for p in left_parts):
                return [val.strip()]

            any_long   = any(len(r) >= 10 for r in rest_parts)
            all_suffix = rest_parts and all(re.match(r"^\d{1,6}$", r) for r in rest_parts)

            if all_suffix and len(first) >= 8 and not any_long:
                results = [first]
                for s in rest_parts:
                    results.append(first[:-len(s)] + s if len(first) > len(s) else first + s)
                for rb in right_blocks:
                    results += [t for t in rb.split() if _is_valid_polis_token(t.strip())]
                return _cap_or_join(results) if results else [first]

            all_tokens = _extract_polis_tokens(val)
            if all_tokens:
                return _cap_or_join(all_tokens)

    # Pola spasi ganda tanpa +
    if re.search(r"\s{2,}", val):
        tokens = _extract_polis_tokens(val)
        if tokens:
            return _cap_or_join(tokens)

    # Pola chain dash (hanya alfanumerik + dash)
    if "-" in val and re.match(r"^[A-Za-z0-9\-]+$", val):
        dash_parts = val.split("-")
        if any(len(p.strip()) >= 8 and re.search(r"\d", p) for p in dash_parts):
            expanded = _expand_dash_chain(val)
            if len(expanded) > 1:
                return expanded

    # Pola breakdown 2-digit suffix: "base 80/81/83"
    m = re.match(r"^(\d[\d\-]+)\s+((?:(?:\d{1,2})|¿)(?:/(?:(?:\d{1,2})|¿))+)\s*$", val)
    if m:
        base = m.group(1).strip()
        results = [
            base[:-2] + s.zfill(2) if len(base) >= 2 else base + s.zfill(2)
            for s in m.group(2).split("/")
            if re.match(r"^\d{1,2}$", s.strip())
        ]
        if results:
            return _cap_or_join(results)

    # Pola umum: cleanup biasa
    val = re.sub(r"\bVARIOUS\b\s*", "", val, flags=re.IGNORECASE).strip()
    val = re.sub(r"\bVAR\b\s*",     "", val, flags=re.IGNORECASE).strip()
    val = val.replace("¿", "")

    parts   = re.split(r"\s*/\s*", val)
    cleaned = [_normalize_spaces(p).strip(" -/") for p in parts if _normalize_spaces(p).strip(" -/")]

    if len(cleaned) > 1 and all(len(c) < 7 for c in cleaned):
        return [val.strip()] if val.strip() else []

    return _cap_or_join(cleaned) if cleaned else []


# ─────────────────────────────────────────────────────────────────────────────
# CLEAN SLIP
# ─────────────────────────────────────────────────────────────────────────────

def clean_slip(val) -> list:
    if pd.isna(val):
        return []
    val = str(val).strip()
    if not val:
        return []

    # Jika blok kiri adalah noise murni, proses blok kanan saja
    blocks = re.split(r"\s{2,}", val.strip())
    if len(blocks) >= 2:
        left = blocks[0]
        left_check = re.sub(
            r"\b(?:SUMMARY|BORDERO?|BORDERA|BORDRO|SINGGLESHIPMENT|VARIOUS)\b",
            "", left, flags=re.IGNORECASE,
        )
        left_check = _SLIP_NOISE_RE.sub("", left_check)
        left_check = re.sub(r"[\s\+\-/]+", "", left_check).strip()
        if not _is_valid_slip_token(left_check):
            tokens = _extract_slip_tokens("  ".join(blocks[1:]))
            if tokens:
                return _cap_or_join(tokens)

    tokens = _extract_slip_tokens(val)
    if tokens:
        return _cap_or_join(tokens)

    # Fallback: kembalikan apa adanya
    cleaned = re.sub(r"\bVARIOUS\b\s*", "", val, flags=re.IGNORECASE)
    cleaned = re.sub(r"^\s*-\s*", "", cleaned).strip()
    cleaned = _normalize_spaces(cleaned)
    return [cleaned] if cleaned else []


# ─────────────────────────────────────────────────────────────────────────────
# CLEAN INSURED
# ─────────────────────────────────────────────────────────────────────────────

def clean_insured(val) -> list:
    if pd.isna(val):
        return []
    val = str(val).strip()
    if not val:
        return []

    cleaned = []
    for p in re.split(r"/|,", val):
        p = _normalize_spaces(p.strip())
        if len(p) <= 2:
            continue
        if p.upper().strip() in INSURED_JUNK_WORDS:
            continue
        if re.match(r"^[^a-zA-Z0-9]+$", p):
            continue

        p_clean = _clean_insured_name(p)
        if not p_clean or len(p_clean) <= 2:
            continue
        if p_clean.upper().strip() in INSURED_JUNK_WORDS:
            continue

        cleaned.append(p_clean)

    if not cleaned:
        fallback = _clean_insured_name(_normalize_spaces(val))
        return [fallback] if fallback else []

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


# ─────────────────────────────────────────────────────────────────────────────
# PROCESS DATA
# ─────────────────────────────────────────────────────────────────────────────

def _insert_clean_columns(
    df: pd.DataFrame,
    all_lists: list,
    prefix: str,
    max_cols: int = MAX_SPLIT_COLS,
) -> list:
    """
    Buat kolom clean_{prefix}_N dan kembalikan nama-nama kolom baru.
    Jumlah kolom dibatasi maksimal MAX_SPLIT_COLS (5).
    Data dengan item > MAX_SPLIT_COLS sudah digabung menjadi 1 string
    oleh _cap_or_join, sehingga hanya mengisi kolom pertama.
    """
    n_cols = min(max_cols, MAX_SPLIT_COLS)
    added = []
    for i in range(1, n_cols + 1):
        col_name = f"clean {prefix} {i}"
        df[col_name] = [lst[i - 1] if i - 1 < len(lst) else None for lst in all_lists]
        added.append(col_name)
    return added


def process_data(input_file: str, output_file: str) -> None:
    print(f"[1/5] Membaca data dari: {input_file} ...")
    df = pd.read_excel(input_file, header=0)
    print(f"      Total baris keseluruhan: {len(df):,}")

    if CEDANT_COL not in df.columns:
        print(f"\n[ERROR] Kolom '{CEDANT_COL}' tidak ditemukan!")
        print(f"        Kolom tersedia: {list(df.columns)}")
        return

    df[CEDANT_COL] = df[CEDANT_COL].astype(str).str.strip()
    df = df[df[CEDANT_COL] == CEDANT_VALUE].copy()
    print(f"[2/5] Filter cedant '{CEDANT_VALUE}': {len(df):,} baris ditemukan.")

    if df.empty:
        print("\n[WARN] Tidak ada data setelah filter. Proses dihentikan.")
        return

    required_cols = [POLIS_COL, SLIP_COL, INSURED_COL, BROKER_COL]
    for col in required_cols:
        if col not in df.columns:
            print(f"\n[ERROR] Kolom '{col}' tidak ditemukan di file!")
            print(f"        Kolom tersedia: {list(df.columns)}")
            return

    # Hitung mitra_bisnis sebelum rename (butuh akses COMP_NAME & COMP_NAME2)
    mitra_values = [
        get_mitra_bisnis(row[BROKER_COL], row[CEDANT_COL])
        for _, row in df.iterrows()
    ]

    df.rename(columns={POLIS_COL: "polis_ori", SLIP_COL: "slip_ori", INSURED_COL: "insured_ori"},
              inplace=True)

    # Sisipkan kolom mitra_bisnis di sebelah kanan BROKER_COL
    insert_pos = list(df.columns).index(BROKER_COL) + 1 if BROKER_COL in df.columns else len(df.columns)
    df.insert(insert_pos, "mitra_bisnis", mitra_values)

    print("[3/5] Menjalankan proses cleaning (mungkin memerlukan beberapa menit)...")

    all_clean_polis, all_clean_slip, all_clean_ins = [], [], []
    max_polis = max_slip = max_ins = 1

    for idx, (_, row) in enumerate(df.iterrows(), 1):
        if idx % 50_000 == 0:
            print(f"      Progress: {idx:,} / {len(df):,} baris diproses...")

        c_polis = clean_polis(row.get("polis_ori", ""))
        c_slip  = clean_slip(row.get("slip_ori", ""))
        c_ins   = clean_insured(row.get("insured_ori", ""))

        max_polis = max(max_polis, len(c_polis))
        max_slip  = max(max_slip,  len(c_slip))
        max_ins   = max(max_ins,   len(c_ins))

        all_clean_polis.append(c_polis)
        all_clean_slip.append(c_slip)
        all_clean_ins.append(c_ins)

    print(f"      Selesai diproses!")
    print(f"      -> Jumlah kolom clean polis  : {max_polis}")
    print(f"      -> Jumlah kolom clean slip   : {max_slip}")
    print(f"      -> Jumlah kolom clean insured: {max_ins}")

    print("[4/5] Menyusun kolom output ...")

    new_columns = []
    for col in df.columns:
        new_columns.append(col)
        if col == "polis_ori":
            new_columns += _insert_clean_columns(df, all_clean_polis, "polis")
        elif col == "slip_ori":
            new_columns += _insert_clean_columns(df, all_clean_slip,  "slip")
        elif col == "insured_ori":
            new_columns += _insert_clean_columns(df, all_clean_ins,   "insured")

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
