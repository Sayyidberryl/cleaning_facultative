import os
import re
from typing import Any, Callable, List, Optional, Tuple

import numpy as np
import pandas as pd

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

INPUT_FILE = os.path.join(BASE_DIR, "data_excel", "raw", "2b. transaksi Osbal 01.01.23 - 17.08.26.xlsx")
SHEET_NAME = "Query result"
OUTPUT_FILE = os.path.join(BASE_DIR, "data_excel", "processed", "askrindo_output_osbal.xlsx")

CEDANT_COL = "CCOS_COMP_NAME"
CEDANT_VALUE = "PT ASURANSI KREDIT INDONESIA"
MAX_HASIL_BREAKDOWN = 5

INSURED_COL = "FAC_INSURED"

FAC_POLIS_CANDIDATES = ["FAC_POLICY_NO", "FAC POLICY NO", "FAC_POLIS", "FAC POLIS"]
CLSDT_POLIS_CANDIDATES = ["CLSDT_POLICY_NO", "CLSDT POLICY NO", "CLSDT_POLIS", "CLSDT POLIS"]
FALLBACK_POLIS_CANDIDATES = ["POLIS", "POLICY_NO", "POLICY NO"]

FAC_SLIP_CANDIDATES = ["FAC_SLIP NO", "FAC_SLIP", "FAC SLIP NO"]
CLSDT_SLIP_CANDIDATES = ["CLSDT_SLIP NO", "CLSDT_SLIP", "CLSDT SLIP NO"]
FALLBACK_SLIP_CANDIDATES = ["SLIP NO"]

POLIS_GARBAGE_RE = re.compile(r"^VARIOU(S)?$", re.IGNORECASE)
POLIS_TRUNCATE_RE = re.compile(r"-(\d)/(\d{2})(?!\d)")
POLIS_SLASH3_TRUNCATE_RE = re.compile(r"/\d{3}(?=[+&/\s]|$)")
POLIS_SPLIT_RE = re.compile(r"[+&/]")
POLIS_FULL_GROUP_MIN = 4

HYPHEN_FULLNUM_SPLIT_RE = re.compile(r"^[0-9]+(?:-[0-9]+)+$")
FULL_NUMBER_MIN_LEN = 6

# --- Regex & konstanta SLIP ---
# Pola 4: SUSPEN
SLIP_SUSPEN_RE = re.compile(
    r"^(?P<base>[0-9.\-]+)/(?P<trunc>\d{2})/(?:SLIP/)?SUSPEN(?P<year>\d{2,4})/?(?:EXCEL)?$",
    re.IGNORECASE,
)

# Pola 2 & 3: (TBA/ )XXXX.XX.XXX.X.XXXXX-X/YY
SLIP_PATTERN_2_3_RE = re.compile(r"^(?:TBA/\s*)?([A-Z0-9.]+)-([A-Z0-9]+)/\d{2}$", re.IGNORECASE)

# Pola 1: XX.XXXXX.XX.XXXXX.XX.X (hanya deretan alfanumerik yang dipisahkan titik)
SLIP_PATTERN_1_RE = re.compile(r"^[A-Z0-9]+(?:\.[A-Z0-9]+)+$", re.IGNORECASE)

# --- Regex & konstanta INSURED ---
INSURED_REMOVE_WORDS_RE = re.compile(
    r"\b(PT|TBK|PERSERO|PELNI|DALAM PELAKSANAAN PADA|PRODUKTIF PADA|STOCK)\b",
    re.IGNORECASE,
)

INSURED_AMP_GROUP_RE = re.compile(r"&\s*GROUP\b", re.IGNORECASE)

BIO_FARMA_MULTI_CLUSTER_RE = re.compile(
    r"BIO\s+FARMA\s*/\s*(?:STOCK\s+)?CLUSTER\s+([A-Z0-9]+)\s*/\s*([A-Z0-9]+)",
    re.IGNORECASE,
)

BIO_FARMA_VAKSIN_RE = re.compile(
    r"BIO\s+FARMA\s*/\s*VAKSIN\s+BIOFARMA\s*/\s*",
    re.IGNORECASE,
)

BIO_FARMA_CLUSTER_RE = r"(BIO FARMA)\s*/\s*((?:ASET\s+)?CLUSTER\s+[A-Z])"

ROMAN_TO_INT = {"IV": 4, "III": 3, "II": 2, "I": 1}
ROMAN_ALT = "IV|III|II|I"

PELINDO_LIST_RE = re.compile(
    r"\b(PELABUHAN\s+INDONESIA|PELINDO)\b\s*"
    rf"((?:(?:{ROMAN_ALT})\s*,\s*)+(?:{ROMAN_ALT}))\b",
    re.IGNORECASE,
)

PELINDO_ROMAN_RE = re.compile(
    rf"\b(PELABUHAN\s+INDONESIA|PELINDO)\b\s*({ROMAN_ALT})\b",
    re.IGNORECASE,
)

PELINDO_REGIONAL_RESULT_RE = re.compile(r"^(PELABUHAN INDONESIA|PELINDO) REGIONAL (\d+)$")


def _polis_digits_only(text: str) -> str:
    return re.sub(r"[^0-9]", "", text)


def _is_pure_3digit(text: str) -> bool:
    return bool(re.fullmatch(r"\d{3}", text.strip()))


def breakdown_polis(value: Any) -> List[str]:
    if pd.isna(value):
        return [value]

    raw = str(value).strip()
    if not raw:
        return [raw]

    # Rule Pengecualian (dibiarkan apa adanya)
    raw_upper = raw.upper()
    if "PENYELESAIAN HUTANG PIUTANG" in raw_upper or "SEE ATTACHMENT" in raw_upper:
        return [raw]

    # Rule P1
    if raw.upper() == "P1":
        return [raw.upper()]
    raw = re.sub(r"^P1[\.\s\-]+", "", raw, flags=re.IGNORECASE)

    # Rule Hyphen Full Number Split
    if HYPHEN_FULLNUM_SPLIT_RE.match(raw):
        segs = raw.split("-")
        if all(len(s) >= FULL_NUMBER_MIN_LEN for s in segs):
            seen_h = set()
            deduped_h = []
            for s in segs:
                if s not in seen_h:
                    seen_h.add(s)
                    deduped_h.append(s)
            if len(deduped_h) <= MAX_HASIL_BREAKDOWN:
                return [h.upper() for h in deduped_h]
            return [raw.upper()]

    # Truncation /000 & -X/YY
    raw_clean = POLIS_SLASH3_TRUNCATE_RE.sub("", raw)
    raw_clean = POLIS_TRUNCATE_RE.sub(lambda m: f"-{m.group(1)}", raw_clean)

    raw_parts = [p.strip() for p in POLIS_SPLIT_RE.split(raw_clean)]
    raw_parts = [p for p in raw_parts if p and not POLIS_GARBAGE_RE.match(p)]

    if not raw_parts:
        return [raw.upper()]

    if len(raw_parts) == 1:
        cleaned = _polis_digits_only(raw_parts[0])
        return [cleaned.upper()] if cleaned else [raw.upper()]

    def _is_full_segment(p: str) -> bool:
        group_count = p.count(".") + 1
        digits_len = len(_polis_digits_only(p))
        return group_count >= POLIS_FULL_GROUP_MIN or digits_len >= 10

    # Rule: 3 digit sebelum angka terakhir
    if _is_full_segment(raw_parts[0]) and len(raw_parts) > 1 and all(
        _is_pure_3digit(p) for p in raw_parts[1:]
    ):
        anchor_digits = _polis_digits_only(raw_parts[0])
        if len(anchor_digits) >= 4:
            base = anchor_digits[:-4]
            last_digit = anchor_digits[-1]
            hasil_3digit = [anchor_digits] + [
                f"{base}{p.strip()}{last_digit}" for p in raw_parts[1:]
            ]
            seen_3 = set()
            deduped_3 = []
            for h in hasil_3digit:
                if h not in seen_3:
                    seen_3.add(h)
                    deduped_3.append(h)
            if len(deduped_3) <= MAX_HASIL_BREAKDOWN:
                return [h.upper() for h in deduped_3]

    has_full_segment = any(_is_full_segment(p) for p in raw_parts)
    if not has_full_segment:
        return [raw.upper()]

    anchor_digits: Optional[str] = None
    hasil: List[str] = []
    for part in raw_parts:
        part_digits = _polis_digits_only(part)
        if not part_digits:
            continue

        if _is_full_segment(part):
            combined = part_digits
            anchor_digits = part_digits
        elif anchor_digits is not None:
            missing_len = len(anchor_digits) - len(part_digits)
            combined = (anchor_digits[:missing_len] + part_digits) if missing_len > 0 else part_digits
        else:
            combined = part_digits

        hasil.append(combined)

    if not hasil:
        return [raw.upper()]

    seen = set()
    deduped = []
    for h in hasil:
        if h not in seen:
            seen.add(h)
            deduped.append(h)

    if len(deduped) > MAX_HASIL_BREAKDOWN:
        return [raw.upper()]

    return [d.upper() for d in deduped]


# Kolom CERTIFICATE tetap disediakan di output, namun logic ekstraksinya
# sengaja dikosongkan karena belum ada kriteria pola yang sesuai.
def extract_certificate(value: Any) -> Any:
    return np.nan


def _slip_suspen_replace(match: "re.Match[str]") -> str:
    base = match.group("base")
    year = match.group("year")
    digits = re.sub(r"[.\-]", "", base)
    tahun = year if len(year) == 4 else f"20{year}"
    return f"{digits} SUSPENSE {tahun}"


def clean_slip(value: Any) -> Any:
    if pd.isna(value):
        return value

    raw = str(value).strip()
    if not raw:
        return raw

    raw_upper = raw.upper()

    # Pengecualian mutlak: biarkan teks utuh tanpa diubah case atau spasinya
    if "SEE ATTACHMENT" in raw_upper or "PENYELESAIAN HUTANG PIUTANG" in raw_upper:
        return raw

    # Pola 4: SUSPEN
    match_suspen = SLIP_SUSPEN_RE.match(raw)
    if match_suspen:
        return _slip_suspen_replace(match_suspen).upper()

    # Pola 2 & 3
    match_2_3 = SLIP_PATTERN_2_3_RE.match(raw)
    if match_2_3:
        # Menghilangkan titik dari base, membuang tanda hubung, dan mengabaikan '/YY'
        base = match_2_3.group(1).replace(".", "")
        tail = match_2_3.group(2)
        return (base + tail).upper()

    # Pola 1
    match_1 = SLIP_PATTERN_1_RE.match(raw)
    if match_1:
        return raw.replace(".", "").upper()

    # Jika tidak ada pola yang cocok, biarkan karakter . / - , utuh
    return raw_upper


def _pelindo_list_replace(match: "re.Match[str]") -> str:
    word = re.sub(r"\s+", " ", match.group(1).upper())
    romans = [r.strip().upper() for r in match.group(2).split(",")]
    entries = []
    for roman in romans:
        num = ROMAN_TO_INT.get(roman)
        if num is None:
            continue
        entries.append(f"{word} REGIONAL {num}")
    if not entries:
        return match.group(0)
    return "/".join(entries)


def _pelindo_roman_replace(match: "re.Match[str]") -> str:
    word = re.sub(r"\s+", " ", match.group(1).upper())
    roman = match.group(2).upper()
    num = ROMAN_TO_INT.get(roman)
    if num is None:
        return match.group(0)
    return f"{word} REGIONAL {num}"


def _resolve_pelindo_priority(hasil: List[str]) -> List[str]:
    numbers_with_pelabuhan = set()
    for h in hasil:
        m = PELINDO_REGIONAL_RESULT_RE.match(h)
        if m and m.group(1) == "PELABUHAN INDONESIA":
            numbers_with_pelabuhan.add(m.group(2))

    result = []
    for h in hasil:
        m = PELINDO_REGIONAL_RESULT_RE.match(h)
        if m and m.group(1) == "PELINDO" and m.group(2) in numbers_with_pelabuhan:
            continue
        result.append(h)
    return result


def clean_insured_part(text: str) -> str:
    if pd.isna(text):
        return ""

    t = str(text)
    t = re.sub(r"\([^)]*\)", " ", t)
    t = re.sub(r"[.,\-]", " ", t)
    t = INSURED_REMOVE_WORDS_RE.sub(" ", t)
    t = INSURED_AMP_GROUP_RE.sub(" ", t)
    t = re.sub(r"\s+", " ", t).strip()
    return t.upper()


def breakdown_insured(value: Any) -> List[str]:
    if pd.isna(value):
        return [value]

    raw = str(value).strip()
    if not raw:
        return [raw]

    multi_cluster_match = BIO_FARMA_MULTI_CLUSTER_RE.search(raw)
    if multi_cluster_match:
        g1 = multi_cluster_match.group(1).upper()
        g2 = multi_cluster_match.group(2).upper()
        return [f"BIO FARMA CLUSTER {g1}", f"BIO FARMA CLUSTER {g2}"]

    raw = BIO_FARMA_VAKSIN_RE.sub("BIO FARMA ", raw)
    raw = PELINDO_LIST_RE.sub(_pelindo_list_replace, raw)
    raw = PELINDO_ROMAN_RE.sub(_pelindo_roman_replace, raw)
    raw = INSURED_AMP_GROUP_RE.sub(" ", raw)

    if re.search(BIO_FARMA_CLUSTER_RE, raw, flags=re.IGNORECASE):
        raw = re.sub(BIO_FARMA_CLUSTER_RE, r"\1 \2", raw, flags=re.IGNORECASE)

    parts_raw = [p.strip() for p in re.split(r"\bQQ\b|/", raw, flags=re.IGNORECASE) if p.strip()]

    hasil = []
    for p in parts_raw:
        p_clean = clean_insured_part(p)
        if p_clean:
            hasil.append(p_clean)

    if not hasil:
        return [raw]

    hasil = _resolve_pelindo_priority(hasil)

    seen = set()
    deduped = []
    for p in hasil:
        if p not in seen:
            seen.add(p)
            deduped.append(p)

    if len(deduped) > MAX_HASIL_BREAKDOWN:
        return [raw]

    return deduped


def _normalize_colname(name: str) -> str:
    return re.sub(r"\s+", " ", str(name).strip()).upper()


def _find_column(df: pd.DataFrame, candidates: List[str]) -> Optional[str]:
    normalized_map = {_normalize_colname(c): c for c in df.columns}
    for cand in candidates:
        key = _normalize_colname(cand)
        if key in normalized_map:
            return normalized_map[key]
    return None


def _find_column_fuzzy(df: pd.DataFrame, must_contain: List[str], must_not_contain: Optional[List[str]] = None) -> Optional[str]:
    must_not_contain = must_not_contain or []
    for col in df.columns:
        name = _normalize_colname(col)
        if all(tok in name for tok in must_contain) and not any(tok in name for tok in must_not_contain):
            return col
    return None


def build_coalesced_source(
    df: pd.DataFrame,
    label: str,
    fac_candidates: List[str],
    clsdt_candidates: List[str],
    fallback_candidates: List[str],
) -> Tuple[Optional[pd.Series], Optional[str]]:
    fac_col = _find_column(df, fac_candidates)
    clsdt_col = _find_column(df, clsdt_candidates)

    keyword_options = ["SLIP"] if "SLIP" in label.upper() else ["POLIS", "POLIC"]

    def _fuzzy_any(must_contain_prefix: List[str], must_not_contain: Optional[List[str]] = None) -> Optional[str]:
        for kw in keyword_options:
            found = _find_column_fuzzy(df, must_contain=must_contain_prefix + [kw], must_not_contain=must_not_contain)
            if found:
                return found
        return None

    if not fac_col:
        fac_col = _fuzzy_any(["FAC"], must_not_contain=["CLSDT"])
        if fac_col:
            print(f"[INFO] {label}: kolom FAC tidak ketemu exact-match, dipakai hasil fuzzy match '{fac_col}'.")
    if not clsdt_col:
        clsdt_col = _fuzzy_any(["CLSDT"])
        if clsdt_col:
            print(f"[INFO] {label}: kolom CLSDT tidak ketemu exact-match, dipakai hasil fuzzy match '{clsdt_col}'.")

    if fac_col and clsdt_col:
        fac_series = df[fac_col]
        clsdt_series = df[clsdt_col]

        clsdt_str = clsdt_series.astype(str).str.strip()
        clsdt_invalid = (
            clsdt_series.isna()
            | clsdt_str.isin(["", "-"])
            | clsdt_str.str.upper().str.contains("SISA OSBAL", regex=False, na=False)
        )

        source = clsdt_series.where(~clsdt_invalid, fac_series)
        print(f"[INFO] {label}: memakai '{clsdt_col}' (utama), fallback ke '{fac_col}' saat CLSDT kosong/invalid ('-', 'SISA OSBAL').")
        return source, fac_col

    if fac_col:
        print(f"[INFO] {label}: kolom CLSDT tidak ditemukan, memakai kolom '{fac_col}' saja.")
        return df[fac_col], fac_col

    fallback_col = _find_column(df, fallback_candidates) or _fuzzy_any([], must_not_contain=["FAC", "CLSDT"])
    if fallback_col:
        print(f"[WARNING] {label}: kolom FAC/CLSDT tidak ditemukan, fallback ke kolom '{fallback_col}'.")
        return df[fallback_col], fallback_col


def _insert_breakdown_columns(
    df: pd.DataFrame,
    source_series: pd.Series,
    insert_after_col: str,
    fn_breakdown: callable,
    prefix: str,
) -> None:
    hasil_breakdown = source_series.apply(fn_breakdown)

    lengths = hasil_breakdown.apply(len)
    max_cols = int(lengths.max()) if not lengths.empty and lengths.max() > 0 else 1

    padded_hasil = hasil_breakdown.apply(lambda lst: lst + [np.nan] * (max_cols - len(lst)))

    df_breakdown = pd.DataFrame(
        padded_hasil.tolist(),
        columns=[f"{prefix}_{i + 1}" for i in range(max_cols)],
        index=df.index,
    )

    if insert_after_col in df.columns:
        posisi = df.columns.get_loc(insert_after_col)
        for i, col in enumerate(df_breakdown.columns):
            df.insert(posisi + 1 + i, col, df_breakdown[col])
    else:
        for col in df_breakdown.columns:
            df[col] = df_breakdown[col]


def process_data(input_file: str, sheet_name: str, output_file: str) -> None:
    if not os.path.exists(input_file):
        print(f"[ERROR] File tidak ditemukan di lokasi: {input_file}")
        return

    df = pd.read_excel(input_file, sheet_name=sheet_name, header=0)
    df.columns = df.columns.str.strip()

    print(f"[DEBUG] Kolom yang terbaca dari sheet '{sheet_name}': {list(df.columns)}")

    if CEDANT_COL not in df.columns:
        print(f"[ERROR] Kolom '{CEDANT_COL}' tidak ditemukan di Excel!")
        print(f"Kolom yang tersedia: {list(df.columns)}")
        return

    is_askrindo = (
        df[CEDANT_COL]
        .astype(str)
        .str.strip()
        .str.replace(r"\s+", " ", regex=True)
        .str.upper()
        == CEDANT_VALUE.upper()
    )
    df_filtered = df[is_askrindo].copy()

    if df_filtered.empty:
        print(f"Tidak ada data untuk {CEDANT_VALUE} di dalam file Excel.")
        return

    jumlah_sebelum = len(df_filtered)
    total_raw_data = len(df)

    print("Menjalankan pembersihan dan breakdown kolom...")

    if INSURED_COL in df_filtered.columns:
        _insert_breakdown_columns(
            df_filtered,
            df_filtered[INSURED_COL],
            INSURED_COL,
            breakdown_insured,
            "INSURED_CLEAN",
        )
    else:
        print(f"[WARNING] Kolom '{INSURED_COL}' tidak ditemukan, breakdown INSURED dilewati.")

    polis_source, polis_anchor_col = build_coalesced_source(
        df_filtered, "POLIS", FAC_POLIS_CANDIDATES, CLSDT_POLIS_CANDIDATES, FALLBACK_POLIS_CANDIDATES
    )
    if polis_source is not None:
        _insert_breakdown_columns(
            df_filtered,
            polis_source,
            polis_anchor_col,
            breakdown_polis,
            "POLIS_CLEAN",
        )

        # Kolom CERTIFICATE tetap dibuat, tapi dikosongkan (lihat extract_certificate)
        cert_series = pd.Series(np.nan, index=df_filtered.index).apply(extract_certificate)

        if "POLIS_CLEAN_1" in df_filtered.columns:
            posisi_cert = df_filtered.columns.get_loc("POLIS_CLEAN_1") + 1
        else:
            posisi_cert = len(df_filtered.columns)
        df_filtered.insert(posisi_cert, "CERTIFICATE_1", cert_series)

    slip_source, slip_anchor_col = build_coalesced_source(
        df_filtered, "SLIP NO", FAC_SLIP_CANDIDATES, CLSDT_SLIP_CANDIDATES, FALLBACK_SLIP_CANDIDATES
    )
    if slip_source is not None:
        posisi = df_filtered.columns.get_loc(slip_anchor_col) + 1 if slip_anchor_col in df_filtered.columns else len(df_filtered.columns)
        df_filtered.insert(posisi, "SLIP_NO_CLEAN", slip_source.apply(clean_slip))

    print(f"Memfilter '{CEDANT_COL}' = '{CEDANT_VALUE}':")
    print(f"-> Total baris file mentah : {total_raw_data}")
    print(f"-> Baris milik {CEDANT_VALUE} : {jumlah_sebelum}")
    print(f"-> Baris setelah di-breakdown (format array kolom): {len(df_filtered)} baris")

    output_dir = os.path.dirname(output_file)
    if output_dir:
        os.makedirs(output_dir, exist_ok=True)
    df_filtered.to_excel(output_file, index=False)
    print("\nPemrosesan Data ASKRINDO Selesai!")


if __name__ == "__main__":
    process_data(INPUT_FILE, SHEET_NAME, OUTPUT_FILE)