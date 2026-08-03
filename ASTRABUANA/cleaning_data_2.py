import os
import re
from typing import Any, List, Optional

import numpy as np
import pandas as pd


INPUT_FILE = os.path.join("dataExcel", "raw", "2a. Transaksi Osbal 01.01.23 - 17.07.26 Rev.xlsx")
SHEET_NAME = "Sheet0"
OUTPUT_FILE = os.path.join("dataExcel", "processed", "astrabuana_output_osbal.xlsx")

CEDANT_COL = "CCOS_COMP_NAME"
CEDANT_VALUE = "PT ASURANSI ASTRA BUANA"

MINIMAL_DIGIT_NOMOR_POLIS = 11
MINIMAL_PANJANG_KODE = 6
MINIMAL_PANJANG_KODE_RANGE = 8
MAX_HASIL_BREAKDOWN = 5

LABEL_WORDS_RE = re.compile(
    r"""
    \bTBA\b|\bVARIOUS\b|\bVAR\b|\bBORD\b|\bBORDERO\b|\bBORDEROUX\b|\bSUMMARY\b|\bASTRA\b
    | \b(?:JANUARI|FEBRUARI|MARET|APRIL|MEI|JUNI|JULI|AGUSTUS|SEPTEMBER|OKTOBER|NOVEMBER|DESEMBER)\b
    | \b(?:JANUARY|FEBRUARY|MARCH|APRIL|MAY|JUNE|JULY|AUGUST|OCTOBER|DECEMBER)\b
    | \b(?:JAN|FEB|MAR|APR|MEI|JUN|JUL|AGU|AGS|SEP|OKT|NOV|DES|AUG|OCT|DEC)\b
    | \b(?:IDR|USD|JPY|SGD|EUR|GBP)\b
    | \b20[0-9]{2}\b
    """,
    re.IGNORECASE | re.VERBOSE,
)

PLACEHOLDER_TOKENS_RE = re.compile(r"\+?\s*\bP\d{1,3}\b\s*\+?", re.IGNORECASE)
VARIOUS_STANDALONE_RE = re.compile(r"^\s*VARIOUS(?:\s*\([^)]*\))?(?:\s+(?:INCLUDE\s+)?FISHING\s+VESSEL)?\s*$", re.IGNORECASE)
VARIOUS_PREFIX_RE = re.compile(r"^\s*VARIOUS(?:\s*\([^)]*\))?(?:\s+(?:INCLUDE\s+)?FISHING\s+VESSEL)?\s*/\s*", re.IGNORECASE)
VARIOUS_SUFFIX_RE = re.compile(r"(?:\s*/\s*|\s*,\s*PT\s*/\s*)VARIOUS(?:\s*\([^)]*\))?(?:\s+(?:INCLUDE\s+)?FISHING\s+VESSEL)?\s*$", re.IGNORECASE)
BORDERO_RE = re.compile(r"\s*\b(?:BORDERO|BORDEROUX)\b.*$", re.IGNORECASE)
INSURED_ENTITIES_RE = re.compile(r"[,.\s]*\b(PT|CV|TBK|PTE|LTD|PERSERO|PELAYARAN)\b(?!\w)[,.\s]*", re.IGNORECASE)
TRAILING_SINGLE_LETTER_TYPO_RE = re.compile(r"[,.\s]+\b[A-Za-z]\b$", re.IGNORECASE)
BRANCH_ANNOTATION_RE = re.compile(r"\bCAB(?:ANG)?(?:\s*\.\s*|\s+)[A-Z0-9_.]+\b|\bCAB(?:ANG)?\b", re.IGNORECASE)
PAREN_ANNOTATIONS_RE = re.compile(r"\([^)]*KNOWN\s+AS[^)]*\)|\(\s*APARTEMEN\b[^\)]*\)|\(\s*PERSERO\s*\)|\(\s*[A-Z0-9]\s*\)", re.IGNORECASE)
INSURED_KEEP_AS_IS_RE = re.compile(r"FACILITY|LINESLIP|OPEN COVER", re.IGNORECASE)
QQ_RE = re.compile(r"\bQ\.?Q\.?\b", re.IGNORECASE)
SD_RE = re.compile(r"\bS\s*/\s*D\b", re.IGNORECASE)
STRIP_LABEL_CLEAN_RE = re.compile(r"^[\s/\-:,.+]+|[\s/\-:,.+]+$")
TBA_VAR_RE = re.compile(r"\b(TBA|VARIOUS|VAR)\b", re.IGNORECASE)
NON_ALPHANUM_RE = re.compile(r"[^A-Za-z0-9]")
MULTIPLE_SPACES_RE = re.compile(r"\s+")
POLIS_VALID_FORMAT_RE = re.compile(r"\d{11,}")
POLIS_SUFFIX_FORMAT_RE = re.compile(r"\d{7,10}[+,\-]\d{2,4}")
RANGE_PATTERN_RE = re.compile(r"[A-Za-z]{2,8}\d{8,}")
ABBREVIATION_RE = re.compile(r"\b(?:[A-Za-z]\.){2,}")

# Dideteksi otomatis -> mapping manual per kasus
MANUAL_INSURED_SPLIT = {
    "CEMERLANG COKE INDUSTRIAL SDN BHD AMBANK ISLAMIC BERHAD F.T.": [
        "CEMERLANG COKE INDUSTRIAL SDN BHD",
        "AMBANK ISLAMIC BERHAD F.T.",
    ],
}

def _clean_code(text: Any) -> str:
    return NON_ALPHANUM_RE.sub("", str(text)).upper()

def _hapus_kata_label(text: Any) -> str:
    text_str = str(text).upper()
    text_str = LABEL_WORDS_RE.sub("", text_str)
    text_str = PLACEHOLDER_TOKENS_RE.sub("", text_str)
    return text_str

def _rapikan_teks_label(text: Any) -> str:
    text_str = str(text).upper()
    text_str = TBA_VAR_RE.sub("", text_str)
    text_str = PLACEHOLDER_TOKENS_RE.sub("", text_str)
    text_str = STRIP_LABEL_CLEAN_RE.sub("", text_str)
    return MULTIPLE_SPACES_RE.sub(" ", text_str).strip()

def _pisah_jadi_segmen(text: str) -> List[str]:
    penanda_sd = "§SD§"
    teks_terlindungi = SD_RE.sub(penanda_sd, text)
    segmen_mentah = re.split(r"\s{2,}|/", teks_terlindungi)

    segmen_hasil = []
    for segmen in segmen_mentah:
        segmen = segmen.replace(penanda_sd, "S/D").strip()
        if segmen and re.search(r"[A-Za-z0-9]", segmen):
            segmen_hasil.append(segmen)
    return segmen_hasil

def _mengandung_kode_panjang(text: str) -> bool:
    text_clean = _hapus_kata_label(text)
    pattern = rf"[A-Za-z0-9]{{{MINIMAL_PANJANG_KODE},}}"
    return bool(re.search(pattern, text_clean)) and bool(re.search(r"\d", text_clean))

def _mengandung_kata_sd(text: str) -> bool:
    return SD_RE.search(text) is not None

def _sepertinya_pola_range(text: str) -> bool:
    kode_kode = RANGE_PATTERN_RE.findall(text)
    if len(kode_kode) != 2:
        return False

    kode_awal, kode_akhir = kode_kode
    if kode_awal.upper() == kode_akhir.upper():
        return False

    prefix_awal = re.match(r"[A-Za-z]*", kode_awal).group().upper()
    prefix_akhir = re.match(r"[A-Za-z]*", kode_akhir).group().upper()
    if prefix_awal != prefix_akhir:
        return False

    gabungan = f"{kode_awal}-{kode_akhir}".upper()
    return gabungan in re.sub(r"\s+", "", text).upper()

def _berisi_banyak_nomor_utuh(text: str) -> bool:
    if re.search(r"[+,\-]", text):
        return False
    token_token = text.split()
    if len(token_token) < 2:
        return False
    return all(len(re.sub(r"\D", "", t)) >= MINIMAL_DIGIT_NOMOR_POLIS for t in token_token)

def _split_long_digit_token(t_clean: str) -> List[str]:
    L = len(t_clean)
    if L < 22:
        return [t_clean]
    if L % 12 == 0:
        return [t_clean[i : i + 12] for i in range(0, L, 12)]
    if L % 11 == 0:
        return [t_clean[i : i + 11] for i in range(0, L, 11)]

    for split_pos in (11, 12):
        part1 = t_clean[:split_pos]
        part2 = t_clean[split_pos:]
        if len(part2) in (11, 12):
            if part2.startswith(part1[:2]) or part2.startswith(("0", "1", "6", "7", "8", "9")):
                return [part1, part2]

    res = []
    curr = 0
    while curr < L:
        rem = L - curr
        if rem in (11, 12):
            res.append(t_clean[curr:])
            break
        elif rem >= 23:
            res.append(t_clean[curr : curr + 12])
            curr += 12
        elif rem == 22:
            res.append(t_clean[curr : curr + 11])
            curr += 11
        else:
            res.append(t_clean[curr:])
            break
    return res

def _rekonstruksi_nomor_dari_pengulangan(teks_tanpa_label: str, acuan_awal: Optional[str] = None) -> List[str]:
    raw_tokens = [t.strip() for t in re.split(r"[+,\-]", teks_tanpa_label) if t.strip()]

    tokens_expanded = []
    for t in raw_tokens:
        t_clean = _clean_code(t)
        if not t_clean:
            continue
        if t_clean.isdigit() and len(t_clean) >= 22:
            tokens_expanded.extend(_split_long_digit_token(t_clean))
        else:
            tokens_expanded.append(t_clean)

    if not tokens_expanded:
        return []

    hasil = []
    nomor_acuan_terakhir = acuan_awal

    for tok in tokens_expanded:
        if not tok.isdigit():
            hasil.append(tok)
            nomor_acuan_terakhir = tok
            continue

        if len(tok) >= MINIMAL_DIGIT_NOMOR_POLIS:
            hasil.append(tok)
            nomor_acuan_terakhir = tok
        elif len(tok) in (7, 8, 9) and (nomor_acuan_terakhir is None or len(nomor_acuan_terakhir) < 11):
            nomor_acuan_terakhir = tok
        elif nomor_acuan_terakhir is not None:
            if len(nomor_acuan_terakhir) in (7, 8, 9) and (len(nomor_acuan_terakhir) + len(tok) in (11, 12)):
                full_num = nomor_acuan_terakhir + tok
                hasil.append(full_num)
                nomor_acuan_terakhir = full_num
            elif len(tok) < len(nomor_acuan_terakhir):
                reconstructed = nomor_acuan_terakhir[:-len(tok)] + tok
                hasil.append(reconstructed)
                nomor_acuan_terakhir = reconstructed
            else:
                hasil.append(tok)
                nomor_acuan_terakhir = tok
        else:
            hasil.append(tok)
            nomor_acuan_terakhir = tok

    return hasil

def _proses_satu_segmen(text: str, wajib_11_digit: bool, acuan_awal: Optional[str] = None) -> List[str]:
    if not _mengandung_kode_panjang(text):
        return [_rapikan_teks_label(text)]

    if _mengandung_kata_sd(text) or _sepertinya_pola_range(text):
        return [re.sub(r"\s+", " ", text).strip()]

    if _berisi_banyak_nomor_utuh(text):
        hasil = [_clean_code(t) for t in text.split()]
        return hasil if len(hasil) <= MAX_HASIL_BREAKDOWN else [text.strip()]

    if wajib_11_digit:
        has_valid_polis_format = bool(POLIS_VALID_FORMAT_RE.search(text)) or bool(POLIS_SUFFIX_FORMAT_RE.search(text))
        if not has_valid_polis_format:
            kode_pertama = re.split(r"[+,\-]", text)[0]
            if len(re.sub(r"\D", "", kode_pertama)) < MINIMAL_DIGIT_NOMOR_POLIS:
                return [text.strip()]

    teks_tanpa_placeholder = PLACEHOLDER_TOKENS_RE.sub("", text)
    hasil_breakdown = _rekonstruksi_nomor_dari_pengulangan(teks_tanpa_placeholder, acuan_awal)

    if not hasil_breakdown or len(hasil_breakdown) > MAX_HASIL_BREAKDOWN:
        return [text.strip()]

    return hasil_breakdown

def clean_insured(name: Any) -> str:
    if pd.isna(name):
        return ""

    name_str = str(name)
    name_str = PAREN_ANNOTATIONS_RE.sub("", name_str)
    name_str = BRANCH_ANNOTATION_RE.sub("", name_str)
    name_str = TRAILING_SINGLE_LETTER_TYPO_RE.sub("", name_str)
    name_str = INSURED_ENTITIES_RE.sub(" ", name_str)
    
    # merapikan koma dan titik tanpa merusak singkatan
    name_str = re.sub(r",", " ", name_str)
    name_str = re.sub(r"\.\s+$", "", name_str)
    name_str = re.sub(r"^[\s/\-:,.+&;]+|[\s/\-:,.+&;]+$", "", name_str)
    
    cleaned = MULTIPLE_SPACES_RE.sub(" ", name_str).strip()
    return cleaned.upper() if len(cleaned) > 1 else ""

def preprocess_insured_text(raw_str: str) -> str:
    text = VARIOUS_PREFIX_RE.sub("", raw_str)
    text = VARIOUS_SUFFIX_RE.sub("", text)
    text = BORDERO_RE.sub("", text)
    text = PAREN_ANNOTATIONS_RE.sub("", text)
    return text.strip()

def breakdown_insured(value: Any) -> List[str]:
    if pd.isna(value):
        return [value]

    raw_str = str(value).strip()
    if not raw_str:
        return [raw_str]

    if VARIOUS_STANDALONE_RE.match(raw_str):
        return [raw_str.upper()]

    text = preprocess_insured_text(raw_str)
    if not text:
        return [raw_str.upper()]

    manual_split = MANUAL_INSURED_SPLIT.get(text.upper())
    if manual_split:
        return [clean_insured(p) for p in manual_split if clean_insured(p)]

    if INSURED_KEEP_AS_IS_RE.search(text) or "POLYCHEM" in text.upper():
        cleaned = clean_insured(text)
        return [cleaned] if cleaned else [raw_str.upper()]

    split_pattern = r"\bQ\.?Q\.?\b|/|:|\s+-\s+|\s*,\s*"
    cleaned_parts = [clean_insured(p) for p in re.split(split_pattern, text) if clean_insured(p)]

    if len(cleaned_parts) > 5:
        return [raw_str.upper()]

    return cleaned_parts if cleaned_parts else [clean_insured(raw_str)]


def breakdown_polis_atau_slip(value: Any, wajib_11_digit: bool) -> List[str]:
    """Fungsi generic untuk memecah Polis atau Slip."""
    if pd.isna(value):
        return [value]

    raw = str(value).strip().upper()
    if not raw:
        return [raw]
    if not _mengandung_kode_panjang(raw):
        return [raw]

    text = _hapus_kata_label(raw)
    if not text.strip():
        return [raw]

    segmen_segmen = _pisah_jadi_segmen(text)
    if not segmen_segmen:
        return [raw]

    ada_segmen_berkode = any(_mengandung_kode_panjang(s) for s in segmen_segmen)
    semua_hasil = []
    acuan_terakhir = None

    for segmen in segmen_segmen:
        berkode = _mengandung_kode_panjang(segmen)
        if ada_segmen_berkode and not berkode:
            continue  

        hasil_segmen = _proses_satu_segmen(segmen, wajib_11_digit, acuan_terakhir)
        semua_hasil.extend(hasil_segmen)

        if hasil_segmen and re.fullmatch(r"[A-Za-z0-9]+", hasil_segmen[-1]):
            acuan_terakhir = hasil_segmen[-1]

    hasil_final = []
    for item in semua_hasil:
        if not hasil_final or item != hasil_final[-1]:
            hasil_final.append(item)

    return hasil_final if hasil_final and len(hasil_final) <= MAX_HASIL_BREAKDOWN else [raw]

def clean_polis(value: Any) -> Any:
    return value if pd.isna(value) else _rapikan_teks_label(value)

def breakdown_polis(value: Any) -> List[str]:
    return breakdown_polis_atau_slip(value, wajib_11_digit=True)

def clean_slip(value: Any) -> Any:
    return value if pd.isna(value) else _rapikan_teks_label(value)

def breakdown_slip(value: Any) -> List[str]:
    return breakdown_polis_atau_slip(value, wajib_11_digit=False)

def _insert_breakdown_columns(
    df: pd.DataFrame, 
    source_col: str, 
    fn_breakdown: callable, 
    prefix: str, 
    is_astrabuana_mask: pd.Series, 
    max_cols_target: int = MAX_HASIL_BREAKDOWN
) -> None:
    """Mengaplikasikan fungsi breakdown dan memecahnya ke kolom-kolom baru."""
    hasil_breakdown = df.apply(
        lambda row: fn_breakdown(row[source_col]) if is_astrabuana_mask[row.name] else [],
        axis=1
    )
    
    lengths = hasil_breakdown.apply(len)
    max_in_data = int(lengths.max()) if not lengths.empty and lengths.max() > 0 else 1
    max_cols = max(max_in_data, max_cols_target)
    
    padded_hasil = hasil_breakdown.apply(lambda lst: lst + [np.nan] * (max_cols - len(lst)))

    df_breakdown = pd.DataFrame(
        padded_hasil.tolist(),
        columns=[f"{prefix}_{i + 1}" for i in range(max_cols)],
        index=df.index,
    )

    posisi = df.columns.get_loc(source_col)
    for i, col in enumerate(df_breakdown.columns):
        df.insert(posisi + 1 + i, col, df_breakdown[col])

def process_data(input_file: str, sheet_name: str, output_file: str) -> None:
    print(f"Membaca data dari: {input_file} (Sheet: {sheet_name}) ...")

    if not os.path.exists(input_file):
        print(f"[ERROR] File tidak ditemukan di lokasi: {input_file}")
        return

    df = pd.read_excel(input_file, sheet_name=sheet_name, header=0)
    df.columns = df.columns.str.strip()

    if CEDANT_COL not in df.columns:
        print(f"[ERROR] Kolom '{CEDANT_COL}' tidak ditemukan di file!")
        return

    is_astrabuana = df[CEDANT_COL].astype(str).str.strip() == CEDANT_VALUE

    kolom_sisa = [col for col in df.columns if col.startswith(("FAC_INSURED_CLN", "FAC_POLICY_CLEAN", "FAC_SLIP_CLEAN"))]
    if kolom_sisa:
        df.drop(columns=kolom_sisa, inplace=True)

    print("Menjalankan pembersihan dan breakdown kolom...")
    _insert_breakdown_columns(df, "FAC_INSURED", breakdown_insured, "FAC_INSURED_CLN", is_astrabuana)
    _insert_breakdown_columns(df, "FAC_POLICY_NO", breakdown_polis, "FAC_POLICY_NO_CLEAN", is_astrabuana)
    _insert_breakdown_columns(df, "FAC_SLIP", breakdown_slip, "FAC_SLIP_CLEAN", is_astrabuana)

    print(f"Menyimpan hasil ke Excel: {output_file} ...")

    os.makedirs(os.path.dirname(output_file), exist_ok=True)
    df.to_excel(output_file, index=False)

    print("Pemrosesan data 2 selesai!")

if __name__ == "__main__":
    process_data(INPUT_FILE, SHEET_NAME, OUTPUT_FILE)