import os
import re
from typing import Any, Callable, List, Optional, Tuple

import numpy as np
import pandas as pd

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

INPUT_FILE = os.path.join(
    BASE_DIR, "data_excel", "raw", "1b. Transaksi Facul 01.01.23 - 17.08.2026.xlsx"
)
SHEET_NAME = "Query result"
OUTPUT_FILE = os.path.join(BASE_DIR, "data_excel", "processed", "askrindo_output_facul.xlsx")

CEDANT_COL = "COMP_NAME"
CEDANT_VALUE = "PT ASURANSI KREDIT INDONESIA"
MAX_HASIL_BREAKDOWN = 5

INSURED_COL = "FAC_INSURED"

# Kolom untuk menentukan BUSINESS_PARTNERS
COMP_NAME_1_COL = "COMP_NAME.1"
BUSINESS_PARTNERS_COL = "BUSINESS_PARTNERS"
DIRECT_VALUE = "DIRECT"

# Hanya mendefinisikan kolom target utama
TARGET_POLIS_COL = ["FAC_POLICY_NO"]
TARGET_SLIP_COL = ["FAC_SLIP"]

# POLIS: KONSTANTA & REGEX
POLIS_GARBAGE_RE = re.compile(r"^VARIOU(S)?$", re.IGNORECASE)
POLIS_TRUNCATE_RE = re.compile(r"-(\d)/(\d{2})(?!\d)")
POLIS_SLASH3_TRUNCATE_RE = re.compile(r"/\d{3}(?=[+&/\s]|$)")
POLIS_SPLIT_RE = re.compile(r"[+&/]")
POLIS_FULL_GROUP_MIN = 4


def _is_full_segment(p: str) -> bool:
    group_count = p.count(".") + 1
    digits_len = len(_polis_digits_only(p))
    return group_count >= POLIS_FULL_GROUP_MIN or digits_len >= 10


HYPHEN_FULLNUM_SPLIT_RE = re.compile(r"^[0-9]+(?:-[0-9]+)+$")
FULL_NUMBER_MIN_LEN = 6

POLIS_NO_CLEAN_EXACT = {
    "023+067+059+043+035+055+047+082+074+031+098+079TBA",
    "0104.22.016.4.00002.6+001.3 + 0102.22.016.4 + 0903",
    "389/RAS/V/2019",
    "0103190271001.5-0102190271.5",
    "102+001+202+904+905+106+103+903+301+002+906",
    "01032101140004.8-00003.5-0903.21.1.011.4.0003.5",
    "038.1406.22000000038/000 EX 038.1406.22000000031/0",
    "15609081800062 ( PHO I )",
    "1560908180074(PHO I & II)",
    "15609081800074 (PHO I)",
    "15609081800074 (PHO II)",
    "15609081800062 (PHO II)",
}


def _normalize_polis_exact(text: str) -> str:
    return re.sub(r"\s+", " ", text.strip()).upper()


POLIS_NO_CLEAN_EXACT_NORMALIZED = {_normalize_polis_exact(s) for s in POLIS_NO_CLEAN_EXACT}

# Pola yang bentuk/jumlah grup titik di tiap sisi "+"-nya berbeda sehingga
# tidak lolos cek _full_segment_parts_splittable(), tapi sebenarnya memang
# harus di-breakdown jadi beberapa polis terpisah. Didaftarkan manual di sini.
POLIS_BREAKDOWN_EXACT = {
    "0103.22.004.1.00003.6 + 0102.22.004.1.00004": [
        "0103220041000036",
        "010222004100004",
    ],
}

POLIS_BREAKDOWN_EXACT_NORMALIZED = {
    _normalize_polis_exact(k): v for k, v in POLIS_BREAKDOWN_EXACT.items()
}

POLIS_SD_SUFFIX_RE = re.compile(r"\s*-\s*\d+\s*S\s*/\s*D\s*\d+\s*$", re.IGNORECASE)
POLIS_TRAILING_SLASH_SHORT_RE = re.compile(r"\s*/\s*\d{1,3}\s*$")

POLIS_HYPHEN_SUFFIX_DROP_RE = re.compile(r"^(\d{10,})\s*-\s*\d{1,7}$")
POLIS_SPECIAL_MIN_DIGITS = 8


def _polis_digits_only(text: str) -> str:
    return re.sub(r"[^0-9]", "", text)


def _apply_polis_special_cleaning(raw: str) -> Optional[str]:

    if "+" in raw or "&" in raw:
        return None
    if len(_polis_digits_only(raw)) < POLIS_SPECIAL_MIN_DIGITS:
        return None

    working = raw.strip()
    working = re.sub(r"^P1[\.\s\-]+", "", working, flags=re.IGNORECASE)
    working = POLIS_SD_SUFFIX_RE.sub("", working)
    working = POLIS_TRAILING_SLASH_SHORT_RE.sub("", working)
    working = working.strip()

    match_hyphen = POLIS_HYPHEN_SUFFIX_DROP_RE.match(working)
    if match_hyphen:
        working = match_hyphen.group(1)

    if working == raw:
        return None

    if re.search(r"[^0-9]", working):
        working = _polis_digits_only(working)

    return working or None


def _is_pure_3digit(text: str) -> bool:
    return bool(re.fullmatch(r"\d{3}", text.strip()))


def _normalize_shape_str(text: str) -> str:
    t = POLIS_TRAILING_SLASH_SHORT_RE.sub("", text).strip()
    # satu suffix "-<digit>" tunggal di ujung (check-digit alternatif tanpa
    # titik) dianggap setara dengan grup titik terakhir.
    t = re.sub(r"-(\d+)$", r".\1", t)
    return t


def _dot_group_count(text: str) -> int:
    return _normalize_shape_str(text).count(".") + 1


def _full_segment_parts_splittable(parts: List[str]) -> bool:
    group_counts = {_dot_group_count(p) for p in parts}
    if len(group_counts) > 1:
        return False
    if group_counts == {1}:
        lengths = {len(_polis_digits_only(p)) for p in parts}
        if len(lengths) > 1:
            return False
    return True


POLIS_SIBLING_JUNK_RE = re.compile(r"/\s*[^\d\s].*$")
POLIS_MID_HYPHEN_SPLIT_RE = re.compile(r"^(?P<a>\d+(?:\.\d+)*)-(?P<b>\d+(?:\.\d+)*(?:-\d+)?)$")
POLIS_MID_HYPHEN_MIN_DIGITS = 10


def _polis_sibling_replace(anchor_part: str, sibling_digits: str) -> str:

    anchor_digits = _polis_digits_only(anchor_part)
    length_sibling = len(sibling_digits)
    if length_sibling == 0 or length_sibling > len(anchor_digits):
        return anchor_digits

    dot_groups = anchor_part.split(".")
    last_group_len = len(_polis_digits_only(dot_groups[-1]))
    second_last_group_len = (
        len(_polis_digits_only(dot_groups[-2])) if len(dot_groups) >= 2 else None
    )

    if last_group_len == 1:
        base = anchor_digits[: -(length_sibling + 1)]
        return base + sibling_digits + anchor_digits[-1]

    base = anchor_digits[:-length_sibling]
    return base + sibling_digits


def _apply_polis_sibling_breakdown(raw: str) -> Optional[List[str]]:
    """Terapkan aturan breakdown by ',' '&' '+' dari #breakdown by + - , &.txt."""
    for sep in (",", "&", "+"):
        if sep not in raw:
            continue
        parts = [p.strip() for p in raw.split(sep) if p.strip()]
        if len(parts) < 2:
            continue

        if all(_is_full_segment(p) for p in parts):
            # Semua bagian adalah polis penuh yang berdiri sendiri-sendiri,
            # jadi tetap dipecah walaupun bentuk/jumlah grup titiknya beda.
            hasil = []
            for p in parts:
                cleaned = POLIS_TRAILING_SLASH_SHORT_RE.sub("", p).strip()
                digits = _polis_digits_only(cleaned)
                if digits:
                    hasil.append(digits)
            if not hasil:
                continue
            seen = set()
            deduped = []
            for h in hasil:
                if h not in seen:
                    seen.add(h)
                    deduped.append(h)
            if len(deduped) <= MAX_HASIL_BREAKDOWN:
                return deduped
            continue

        anchor_part = parts[0]
        siblings = parts[1:]
        if not _is_full_segment(anchor_part):
            continue
        if any(_is_full_segment(p) for p in siblings):
            continue

        hasil = [_polis_digits_only(anchor_part)]
        ok = True
        for sib in siblings:
            sib_clean = POLIS_SIBLING_JUNK_RE.sub("", sib).strip()
            sib_digits = _polis_digits_only(sib_clean)
            if not sib_digits:
                ok = False
                break
            hasil.append(_polis_sibling_replace(anchor_part, sib_digits))

        if not ok:
            continue

        seen = set()
        deduped = []
        for h in hasil:
            if h not in seen:
                seen.add(h)
                deduped.append(h)
        if len(deduped) <= MAX_HASIL_BREAKDOWN:
            return deduped

    return None


def _apply_polis_mid_hyphen_split(raw: str) -> Optional[List[str]]:
    """Terapkan aturan breakdown by '-' di tengah untuk 2 polis penuh."""
    if any(sep in raw for sep in ("+", "&", ",", "/")):
        return None

    match = POLIS_MID_HYPHEN_SPLIT_RE.match(raw.strip())
    if not match:
        return None

    a_str, b_str = match.group("a"), match.group("b")
    if not _full_segment_parts_splittable([a_str, b_str]):
        # Bentuk/jumlah digit kedua sisi tidak sepadan -> jangan dipecah.
        return None

    a_digits = _polis_digits_only(a_str)
    b_digits = _polis_digits_only(b_str)
    if len(a_digits) < POLIS_MID_HYPHEN_MIN_DIGITS or len(b_digits) < POLIS_MID_HYPHEN_MIN_DIGITS:
        return None

    if a_digits == b_digits:
        return [a_digits]
    return [a_digits, b_digits]


POLIS_HYPHEN_SHORT_SIBLING_RE = re.compile(r"^\d{1,6}$")


def _apply_polis_hyphen_sibling_breakdown(raw: str) -> Optional[List[str]]:

    if any(sep in raw for sep in ("+", "&", ",", "/")):
        return None

    parts = [p.strip() for p in raw.split("-") if p.strip()]
    if len(parts) < 2:
        return None

    anchor_part = parts[0]
    siblings = parts[1:]
    if not _is_full_segment(anchor_part):
        return None
    if any(_is_full_segment(p) for p in siblings):
        return None
    if not all(POLIS_HYPHEN_SHORT_SIBLING_RE.match(p) for p in siblings):
        return None

    hasil = [_polis_digits_only(anchor_part)]
    for sib in siblings:
        hasil.append(_polis_sibling_replace(anchor_part, sib))

    seen = set()
    deduped = []
    for h in hasil:
        if h not in seen:
            seen.add(h)
            deduped.append(h)

    if len(deduped) <= MAX_HASIL_BREAKDOWN:
        return deduped
    return None


# POLIS: FUNGSI UTAMA BREAKDOWN
def breakdown_polis(value: Any) -> List[str]:
    if pd.isna(value):
        return [value]

    raw = str(value).strip()
    if not raw:
        return [raw]

    if _normalize_polis_exact(raw) in POLIS_NO_CLEAN_EXACT_NORMALIZED:
        return [raw]

    normalized_raw = _normalize_polis_exact(raw)
    if normalized_raw in POLIS_BREAKDOWN_EXACT_NORMALIZED:
        return [s.upper() for s in POLIS_BREAKDOWN_EXACT_NORMALIZED[normalized_raw]]

    raw_upper = raw.upper()
    if "PENYELESAIAN HUTANG PIUTANG" in raw_upper or "SEE ATTACHMENT" in raw_upper:
        return [raw]

    if raw.upper() == "P1":
        return [raw.upper()]

    special_cleaned = _apply_polis_special_cleaning(raw)
    if special_cleaned is not None:
        return [special_cleaned.upper()]

    raw = re.sub(r"^P1[\.\s\-]+", "", raw, flags=re.IGNORECASE)

    sibling_result = _apply_polis_sibling_breakdown(raw)
    if sibling_result is not None:
        return [s.upper() for s in sibling_result]

    mid_hyphen_result = _apply_polis_mid_hyphen_split(raw)
    if mid_hyphen_result is not None:
        return [s.upper() for s in mid_hyphen_result]

    hyphen_sibling_result = _apply_polis_hyphen_sibling_breakdown(raw)
    if hyphen_sibling_result is not None:
        return [s.upper() for s in hyphen_sibling_result]

    if HYPHEN_FULLNUM_SPLIT_RE.match(raw):
        segs = raw.split("-")
        if len(set(len(s) for s in segs)) > 1:
            if all(_is_full_segment(s) for s in segs):
                # Panjang digit beda, tapi tiap sisi memang polis penuh
                # yang berdiri sendiri (>=10 digit) -> tetap dipecah.
                seen_h = set()
                deduped_h = []
                for s in segs:
                    if s not in seen_h:
                        seen_h.add(s)
                        deduped_h.append(s)
                if len(deduped_h) <= MAX_HASIL_BREAKDOWN:
                    return [h.upper() for h in deduped_h]
            # Angka polos tanpa titik dengan jumlah digit berbeda di tiap
            # sisi "-" -> jangan dipecah, biarkan apa adanya.
            return [raw.upper()]
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

    raw_clean = POLIS_SLASH3_TRUNCATE_RE.sub("", raw)
    raw_clean = POLIS_TRUNCATE_RE.sub(lambda m: f"-{m.group(1)}", raw_clean)

    raw_parts = [p.strip() for p in POLIS_SPLIT_RE.split(raw_clean)]
    raw_parts = [p for p in raw_parts if p and not POLIS_GARBAGE_RE.match(p)]

    if not raw_parts:
        return [raw.upper()]

    if len(raw_parts) == 1:
        cleaned = _polis_digits_only(raw_parts[0])
        return [cleaned.upper()] if cleaned else [raw.upper()]

    if (
        _is_full_segment(raw_parts[0])
        and len(raw_parts) > 1
        and all(_is_pure_3digit(p) for p in raw_parts[1:])
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
            combined = (
                (anchor_digits[:missing_len] + part_digits) if missing_len > 0 else part_digits
            )
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


# CERTIFICATE
CERTIFICATE_COL = "CERTIFICATE_1"
CERT_MIN_BASE_DIGITS = 6

# Pola: <no polis dasar, min 6 digit> - <5 atau 6 digit certif>[ S/D <5 atau 6 digit certif>]
# Sesuai PERUBAHAN_DATA_1_.txt:
# - 111100117000061 - 000012          -> CERTIFICATE = 000012
# - 111100118000023-000002            -> CERTIFICATE = 000002
# - 111100117000061-00011 (5 digit)   -> CERTIFICATE = 000011 (dipad jadi 6 digit)
# - 107150418000014 - 000001 S/D 000009 -> CERTIFICATE = 000001 S/D 000009
# Jika pola tidak persis seperti ini (susah dibaca / campur karakter lain), dikosongkan.
CERT_PATTERN_RE = re.compile(
    rf"^(\d{{{CERT_MIN_BASE_DIGITS},}})\s*-\s*(\d{{5,6}})(?:\s*S\s*/\s*D\s*(\d{{5,6}}))?\s*$",
    re.IGNORECASE,
)


def _pad_certificate(digits: str) -> str:
    """Jika hanya 5 digit, tambahkan 1 digit '0' di depan agar jadi 6 digit."""
    if len(digits) == 5:
        return "0" + digits
    return digits


def extract_certificate(value: Any) -> str:
    """Ekstrak nomor sertifikat dari pola polis mentah (FAC_POLICY_NO / polis ori).

    Lihat rules lengkap di PERUBAHAN_DATA_1_.txt. Jika pola tidak dikenali/susah
    dibaca, kembalikan string kosong (kolom CERTIFICATE dikosongkan).
    """
    if pd.isna(value):
        return ""

    raw = str(value).strip()
    if not raw:
        return ""

    match = CERT_PATTERN_RE.match(raw)
    if not match:
        return ""

    cert1, cert2 = match.group(2), match.group(3)
    if cert2:
        return f"{_pad_certificate(cert1)} S/D {_pad_certificate(cert2)}"
    return _pad_certificate(cert1)


def _add_certificate_column(df: pd.DataFrame, source_col: str, insert_after_col: str) -> None:
    """Tambahkan kolom CERTIFICATE (nilai diekstrak dari polis ori) tepat
    setelah kolom `insert_after_col`."""
    certificate_values = df[source_col].apply(extract_certificate)
    posisi = df.columns.get_loc(insert_after_col)
    df.insert(posisi + 1, CERTIFICATE_COL, certificate_values)
    print(f"[INFO] Kolom '{CERTIFICATE_COL}' ditambahkan setelah '{insert_after_col}'.")


# SLIP: KONSTANTA, REGEX & FUNGSI
SLIP_SUSPEN_RE = re.compile(
    r"^(?P<base>[0-9.\-]+)/(?P<trunc>\d{2})/(?:SLIP/)?SUSPEN(?P<year>\d{2,4})/?(?:EXCEL)?$",
    re.IGNORECASE,
)
SLIP_PATTERN_2_3_RE = re.compile(r"^(?:TBA/\s*)?([A-Z0-9.]+)-([A-Z0-9]+)/\d{2}$", re.IGNORECASE)
SLIP_PATTERN_1_RE = re.compile(r"^[A-Z0-9]+(?:\.[A-Z0-9]+)+$", re.IGNORECASE)
SLIP_13DIGIT_TRUNCATE_RE = re.compile(r"^(\d{13})\s*[/\-].*$")
SLIP_14DIGIT_TRUNCATE_RE = re.compile(r"^(\d{14})\s*[/\-].*$")
SLIP_DOTTED_EMBED_RE = re.compile(r"(\d{4}\.\d{2}\.\d{3}\.\d\.\d{5}-\d)\s*/\s*\d{2}")
SLIP_DASH_BREAKDOWN_RE = re.compile(r"^(\d{8,})\s*-\s*(\d{8,})$")
SLIP_LEADING_DIGITS_RE = re.compile(r"^(\d+)")
SLIP_SLASH_SECOND_NUM_RE = re.compile(r"/\s*(\d{13,})")
SLIP_SLASH_SECOND_NUM_MIN_LEADING = 13


def _slip_suspen_replace(match: "re.Match[str]") -> str:
    base = match.group("base")
    year = match.group("year")
    digits = re.sub(r"[.\-]", "", base)
    tahun = year if len(year) == 4 else f"20{year}"
    return f"{digits} SUSPENSE {tahun}"


def _slip_single_clean(raw: str) -> str:
    """Bersihkan satu nilai SLIP (tanpa breakdown oleh '+')."""
    raw = raw.strip()
    if not raw:
        return raw

    raw_upper = raw.upper()

    if "SEE ATTACHMENT" in raw_upper or "PENYELESAIAN HUTANG PIUTANG" in raw_upper:
        return raw

    match_suspen = SLIP_SUSPEN_RE.match(raw)
    if match_suspen:
        return _slip_suspen_replace(match_suspen).upper()

    # Pola berkelompok titik-strip-slash, boleh nyempil di tengah teks lain.
    match_dotted = SLIP_DOTTED_EMBED_RE.search(raw)
    if match_dotted:
        base = match_dotted.group(1)
        return re.sub(r"[.\-]", "", base).upper()

    # 14 digit di awal lalu "/" atau "-" -> sisakan seluruh 14 digitnya.
    match_14 = SLIP_14DIGIT_TRUNCATE_RE.match(raw)
    if match_14:
        return match_14.group(1)

    # 13 digit di awal lalu "/" atau "-" -> sisakan 13 digit awal saja.
    match_13 = SLIP_13DIGIT_TRUNCATE_RE.match(raw)
    if match_13:
        return match_13.group(1)

    match_2_3 = SLIP_PATTERN_2_3_RE.match(raw)
    if match_2_3:
        base = match_2_3.group(1).replace(".", "")
        tail = match_2_3.group(2)
        return (base + tail).upper()

    match_1 = SLIP_PATTERN_1_RE.match(raw)
    if match_1:
        return raw.replace(".", "").upper()

    return raw_upper


SLIP_PLUS_SPLIT_RE = re.compile(r"\s*\+\s*")
SLIP_PLUS_BREAKDOWN_MIN_DIGITS = 13


def _try_slash_second_number_breakdown(raw: str) -> Optional[List[str]]:

    match_lead = SLIP_LEADING_DIGITS_RE.match(raw)
    if not match_lead:
        return None
    leading = match_lead.group(1)
    if len(leading) < SLIP_SLASH_SECOND_NUM_MIN_LEADING:
        return None

    match_second = SLIP_SLASH_SECOND_NUM_RE.search(raw, match_lead.end())
    if not match_second:
        return None
    second = match_second.group(1)
    if second == leading:
        return None
    return [leading, second]


def breakdown_slip(value: Any) -> List[str]:
    if pd.isna(value):
        return [value]

    raw = str(value).strip()
    if not raw:
        return [raw]

    raw_upper = raw.upper()
    if "SEE ATTACHMENT" in raw_upper or "PENYELESAIAN HUTANG PIUTANG" in raw_upper:
        return [raw]

    if "+" in raw:
        parts = [p for p in SLIP_PLUS_SPLIT_RE.split(raw) if p.strip()]
        if len(parts) >= 2:
            digit_lens = [len(re.sub(r"[^0-9]", "", p)) for p in parts]
            if all(dl < SLIP_PLUS_BREAKDOWN_MIN_DIGITS for dl in digit_lens):
                # Pengecualian: semua bagian terlalu pendek -> jangan breakdown.
                return [raw_upper]

            first_clean = _slip_single_clean(parts[0]).upper()
            first_digits = re.sub(r"[^0-9]", "", first_clean)
            hasil = [first_clean]
            for p in parts[1:]:
                p_stripped = p.strip()
                p_digits = re.sub(r"[^0-9]", "", p_stripped)
                if (
                    p_digits
                    and re.fullmatch(r"\d+", p_stripped)
                    and len(p_digits) < len(first_digits)
                ):
                    # sibling pendek -> gantikan N digit terakhir slip pertama
                    combined = first_digits[: -len(p_digits)] + p_digits
                    hasil.append(combined)
                else:
                    hasil.append(_slip_single_clean(p_stripped).upper())

            seen = set()
            deduped = []
            for h in hasil:
                if h not in seen:
                    seen.add(h)
                    deduped.append(h)

            if deduped and len(deduped) <= MAX_HASIL_BREAKDOWN:
                return deduped

        return [_slip_single_clean(raw).upper()]

    # Breakdown by "-" murni: dua angka panjang, tidak ada teks lain.
    match_dash_breakdown = SLIP_DASH_BREAKDOWN_RE.match(raw)
    if match_dash_breakdown:
        return [match_dash_breakdown.group(1), match_dash_breakdown.group(2)]

    # Breakdown karena ada angka kedua (>=13 digit) setelah "/".
    slash_breakdown = _try_slash_second_number_breakdown(raw)
    if slash_breakdown:
        return slash_breakdown

    return [_slip_single_clean(raw).upper()]


def clean_slip(value: Any) -> Any:

    if pd.isna(value):
        return value
    hasil = breakdown_slip(value)
    return hasil[0] if hasil else value


# INSURED: KONSTANTA, REGEX & FUNGSI
INSURED_REMOVE_WORDS_RE = re.compile(
    r"\b(PTE|LTD|KSO|PT|IM2|STOCK|VARIOUS)\b",
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
BIO_FARMA_CLUSTER_RE = r"(BIO FARMA)\s*/\s*(?:STOCK\s+)?((?:ASET\s+)?CLUSTER\s+[A-Z0-9]+)"
ROMAN_TO_INT = {"IV": 4, "III": 3, "II": 2, "I": 1}
ROMAN_ALT = "IV|III|II|I"
PELINDO_LIST_RE = re.compile(
    r"\b(PELABUHAN\s+INDONESIA|PELINDO)\b\s*" rf"((?:(?:{ROMAN_ALT})\s*,\s*)+(?:{ROMAN_ALT}))\b",
    re.IGNORECASE,
)
PELINDO_ROMAN_RE = re.compile(
    rf"\b(PELABUHAN\s+INDONESIA|PELINDO)\b\s*({ROMAN_ALT})\b",
    re.IGNORECASE,
)

JOB_PERTAMINA_RE = re.compile(r"^(JOB\s+PERTAMINA)\s+(.+)$", re.IGNORECASE)

ROMAN_RANGE_RE = re.compile(
    rf"^(?P<prefix>.+?)\s+(?P<start>{ROMAN_ALT})\s*-\s*(?P<end>{ROMAN_ALT})\b(?P<suffix>.*)$",
    re.IGNORECASE,
)


def _roman_range_expand(match: "re.Match[str]") -> Optional[List[str]]:
    prefix = re.sub(r"\s+", " ", match.group("prefix")).strip().upper()
    start = ROMAN_TO_INT.get(match.group("start").upper())
    end = ROMAN_TO_INT.get(match.group("end").upper())
    if not prefix or start is None or end is None or start > end:
        return None
    return [f"{prefix} {n}" for n in range(start, end + 1)]


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


PELINDO_ADJACENT_RE = re.compile(
    r"((?:PELABUHAN\s+INDONESIA|PELINDO)\s+REGIONAL\s+\d+)\s+"
    r"(?=(?:PELABUHAN\s+INDONESIA|PELINDO)\s+REGIONAL\s+\d+)",
    re.IGNORECASE,
)

PELINDO_REGIONAL_RESULT_RE = re.compile(r"^(PELABUHAN INDONESIA|PELINDO) REGIONAL (\d+)$")


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


INSURED_NO_CLEAN_EXACT = {
    "SALATIGA/KATSURA STA 40+409-STA 71+875",
    "INDONESIA FIFA U-17",
    "MEDCO E AND P INDONESIA PACKAGE A - E",
}


def _normalize_insured_exact(text: str) -> str:
    return re.sub(r"\s+", " ", text.strip()).upper()


INSURED_NO_CLEAN_EXACT_NORMALIZED = {_normalize_insured_exact(s) for s in INSURED_NO_CLEAN_EXACT}

INSURED_SPLIT_RE = re.compile(
    r"\bAND/OR\b|\bQQ\b|(?<=\s)-|-(?=\s)|[/\u00bf]",
    re.IGNORECASE,
)

_PAREN_RE = re.compile(r"\([^)]*\)")
_PAREN_PERSERO_RE = re.compile(r"\(\s*PERSERO\s*\)", re.IGNORECASE)


def _protect_parens(text: str) -> Tuple[str, dict]:

    mapping: dict = {}

    def _replace(match: "re.Match[str]") -> str:
        token = f"\ue000{len(mapping)}\ue000"
        mapping[token] = match.group(0)
        return token

    protected = _PAREN_RE.sub(_replace, text)
    return protected, mapping


def _restore_parens(text: str, mapping: dict) -> str:
    for token, original in mapping.items():
        text = text.replace(token, original)
    return text


def clean_insured_part(text: str) -> str:
    if pd.isna(text):
        return ""

    t = str(text)
    t = re.sub(r"[.,]", " ", t)
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

    if _normalize_insured_exact(raw) in INSURED_NO_CLEAN_EXACT_NORMALIZED:
        return [raw]

    raw = re.sub(r"\bP(?=\()", "", raw)

    multi_cluster_match = BIO_FARMA_MULTI_CLUSTER_RE.search(raw)
    if multi_cluster_match:
        g1 = multi_cluster_match.group(1).upper()
        g2 = multi_cluster_match.group(2).upper()
        return [f"BIO FARMA CLUSTER {g1}", f"BIO FARMA CLUSTER {g2}"]

    # "JOB PERTAMINA <sisanya>" -> breakdown 2 insured terpisah
    job_pertamina_match = JOB_PERTAMINA_RE.match(raw)
    if job_pertamina_match:
        raw = f"{job_pertamina_match.group(1)}/{job_pertamina_match.group(2)}"

    # Rentang angka romawi "... I-III" -> breakdown per angka (1,2,3,...)
    range_match = ROMAN_RANGE_RE.match(raw)
    if range_match:
        entries = _roman_range_expand(range_match)
        if entries and len(entries) <= MAX_HASIL_BREAKDOWN:
            return entries

    raw = BIO_FARMA_VAKSIN_RE.sub("BIO FARMA ", raw)
    raw = PELINDO_LIST_RE.sub(_pelindo_list_replace, raw)
    raw = PELINDO_ROMAN_RE.sub(_pelindo_roman_replace, raw)
    raw = PELINDO_ADJACENT_RE.sub(r"\1/", raw)
    raw = INSURED_AMP_GROUP_RE.sub(" ", raw)

    if re.search(BIO_FARMA_CLUSTER_RE, raw, flags=re.IGNORECASE):
        raw = re.sub(BIO_FARMA_CLUSTER_RE, r"\1 \2", raw, flags=re.IGNORECASE)

    # "(PERSERO)" dibuang; isi tanda kurung lainnya (mis. "( STPI - CURUG )")
    # dilindungi utuh dan tidak ikut dibersihkan/dipecah sama sekali.
    raw = _PAREN_PERSERO_RE.sub(" ", raw)
    raw_protected, paren_map = _protect_parens(raw)

    parts_raw = [p.strip() for p in INSURED_SPLIT_RE.split(raw_protected) if p.strip()]

    hasil = []
    for p in parts_raw:
        p_clean = clean_insured_part(p)
        if p_clean:
            p_clean = _restore_parens(p_clean, paren_map)
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


# HELPER KOLOM DATAFRAME
def _normalize_colname(name: str) -> str:
    return re.sub(r"\s+", " ", str(name).strip()).upper()


def _find_column(df: pd.DataFrame, candidates: List[str]) -> Optional[str]:
    normalized_map = {_normalize_colname(c): c for c in df.columns}
    for cand in candidates:
        key = _normalize_colname(cand)
        if key in normalized_map:
            return normalized_map[key]
    return None


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


def _add_business_partners_column(df: pd.DataFrame) -> None:

    if COMP_NAME_1_COL not in df.columns:
        print(
            f"[WARNING] Kolom '{COMP_NAME_1_COL}' tidak ditemukan, kolom '{BUSINESS_PARTNERS_COL}' dilewati."
        )
        return
    if CEDANT_COL not in df.columns:
        print(
            f"[WARNING] Kolom '{CEDANT_COL}' tidak ditemukan, kolom '{BUSINESS_PARTNERS_COL}' dilewati."
        )
        return

    is_direct = df[COMP_NAME_1_COL].astype(str).str.strip().str.upper() == DIRECT_VALUE
    business_partners = np.where(is_direct, df[CEDANT_COL], df[COMP_NAME_1_COL])

    posisi = df.columns.get_loc(COMP_NAME_1_COL)
    df.insert(posisi + 1, BUSINESS_PARTNERS_COL, business_partners)
    print(f"[INFO] Kolom '{BUSINESS_PARTNERS_COL}' ditambahkan setelah '{COMP_NAME_1_COL}'.")


# PIPELINE UTAMA
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
        df[CEDANT_COL].astype(str).str.strip().str.replace(r"\s+", " ", regex=True).str.upper()
        == CEDANT_VALUE.upper()
    )
    df_filtered = df[is_askrindo].copy()

    if df_filtered.empty:
        print(f"Tidak ada data untuk {CEDANT_VALUE} di dalam file Excel.")
        return

    jumlah_sebelum = len(df_filtered)
    total_raw_data = len(df)

    print("Menjalankan pembersihan dan breakdown kolom...")

    # PROSES BUSINESS_PARTNERS - setelah COMP_NAME.1, sebelum FAC_INSURED
    _add_business_partners_column(df_filtered)

    # PROSES INSURED
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

    # PROSES POLIS - LANGSUNG AMBIL DARI FAC_POLICY_NO
    polis_col = _find_column(df_filtered, TARGET_POLIS_COL)
    if polis_col:
        print(f"[INFO] Memproses kolom '{polis_col}' untuk data Polis.")

        _insert_breakdown_columns(
            df_filtered,
            df_filtered[polis_col],
            polis_col,
            breakdown_polis,
            "POLIS_CLEAN",
        )

        # PROSES CERTIFICATE - nilai diambil dari polis ori, tapi kolomnya
        # diletakkan setelah POLIS_CLEAN_1 (lihat PERUBAHAN_DATA_1_.txt)
        _add_certificate_column(df_filtered, polis_col, "POLIS_CLEAN_1")
    else:
        print(f"[WARNING] Kolom target '{TARGET_POLIS_COL[0]}' tidak ditemukan di dataset.")

    # PROSES SLIP - LANGSUNG AMBIL DARI FAC_SLIP
    slip_col = _find_column(df_filtered, TARGET_SLIP_COL)
    if slip_col:
        print(f"[INFO] Memproses kolom '{slip_col}' untuk data Slip.")
        _insert_breakdown_columns(
            df_filtered,
            df_filtered[slip_col],
            slip_col,
            breakdown_slip,
            "SLIP_CLEAN",
        )
    else:
        print(f"[WARNING] Kolom target '{TARGET_SLIP_COL[0]}' tidak ditemukan di dataset.")

    print(f"\nMemfilter '{CEDANT_COL}' = '{CEDANT_VALUE}':")
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