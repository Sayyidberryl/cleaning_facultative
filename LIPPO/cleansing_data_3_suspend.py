"""
================================================================================
 SCRIPT CLEANSING DATA 3 - SUSPENSE (KHUSUS CEDANT: PT LIPPO GENERAL INSURANCE)
================================================================================
Breakdown & cleansing kolom INSURED, POLIS, SLIP NO untuk key matching antar database.
Scope: baris dengan CEDANT NAME/SHRT NAME mengandung "LIPPO". Kolom asli dipertahankan;
kolom baru (*_CLN_n) disisipkan di sebelah kolom originalnya.
================================================================================
"""

import os
import re
import pandas as pd
from openpyxl import load_workbook

# ==============================================================================
# 0. KONFIGURASI
# ==============================================================================
INPUT_FILE = "Data_3_Suspend.xlsx"
SHEET_NAME = "Detail Database"
HEADER_ROW = 2  # header di baris excel ke-3
OUTPUT_FILE = "[3Agustus2026] lippo_output_suspend.xlsx"
CEDANT_FILTER = "LIPPO"

MAX_BREAKDOWN_CODES = 5  # >5 pecahan polis/slip -> simpan kode pertama saja
MAX_BREAKDOWN_INSURED = 5  # >5 pecahan insured -> jangan dibreakdown, pakai teks original
TEXT_FORMAT_COLUMN_KEYWORDS = ("POLIS", "SLIP", "RECEIPT NO", "CREDIT NOTES", "DETAIL RINCIAN NO")
RE_I = re.IGNORECASE

# ==============================================================================
# 1. CLEANSING INSURED
# ==============================================================================
_TRANSACTION_HEADER_RE = re.compile(r"^\s*(?:Pembayaran|Penagihan)?\s*(?:dan\s+Penagihan)?\s*Premi\s*(?:Fakultatif|SOA|Pensesian)?\s*(?:PAR\s*&?\s*EQ|BIT|GIT|CECR|QS)?\s*", RE_I)
_TRUNCATE_TRIGGER_RE = re.compile(r"\bsubsidiar|\bassociat|\baffiliat|\bfiliated\b|related\s+compan|respective\s+right|\bincluding\s+an[yd]|\bindu?ding\s+any|as\s+the\s+owner|as\s+owners?\b|as\s+main\s+contractor|\(as\s+mortgagee\s+and\s+loss\s+payee\)", RE_I)
_REMOVE_ONLY_RE = re.compile(r"\bas\s+principal\b|\bsebagai\s+principal\b|\bsebagai\s+kontraktor\b", RE_I)
_GENERIC_BOILERPLATE_TAIL_RE = re.compile(r"\bsub\s*-?\s*kontraktor\b.*$|\banak\s+perusahaan\b.*$|\bafiliasi\s+perusahaan\b.*$|\bsemua\b\s*[–-]?\s*$", RE_I)
_TITLE_RE = re.compile(r"\b(?:S\.?H\.?|S\.?I\.?Kom\.?|S\.?E\.?|S\.?T\.?|S\.?Kom\.?|S\.?Psi\.?|M\.?H\.?|M\.?M\.?|M\.?Sc\.?|Ir\.|Drs\.|Dra\.|Prof\.|Dr\.|Tn\.?|Bpk\.?|Bapak|Ibu|Ny\.?|Nyonya|Mr\.?|Mrs\.?|Ms\.?)\b", RE_I)
_LEGAL_ENTITY_RE = re.compile(r"\b(?:PT\.?|P\.T\.?|CV\.?|C\.V\.?|TBK\.?|Tbk\.?|PERSERO|Perseroan\s+Terbatas)\b", RE_I)
_ENTITY_SPLIT_RE = re.compile(r"\b(?:QQ|Q\.Q\.)\b|\band\s*/\s*or\b|\bdan\s*/\s*atau\b|\bin\s+this\s+case\b|\bdalam\s+hal\s+ini\b|\band\b|\bor\b|,", RE_I)
_LEGAL_ENTITY_DASH_SIGNAL_RE = re.compile(r",\s*(?:PT\.?|P\.T\.?|CV\.?|C\.V\.?|TBK\.?|Tbk\.?|PERSERO)\s*[\r\n]*\s*-\s*", RE_I)
_GENERIC_DASH_SEPARATOR_RE = re.compile(r"\s*[\r\n]+\s*-\s*|\s+-\s+")

# Pola G: "ENTITY1 and/or associated... for their/its respective rights... ENTITY2" 
# -> klausa boilerplate di TENGAH kalimat dibuang, bukan di-truncate ke akhir.
_MIDDLE_BOILERPLATE_CLAUSE_RE = re.compile(
    r"(?:\band\s*/\s*or\s+)?(?:associat\w*|subsidiar\w*|affiliat\w*|related\s+compan\w*)"
    r"(?:\s+and\s*/\s*or\s+(?:associat\w*|subsidiar\w*|affiliat\w*|related\s+compan\w*))*"
    r"\s+for\s+(?:their|its)\s+respective\s+rights?\s+and\s+interests?\b", RE_I)

def _truncate_from_first_trigger(text: str) -> str:
    match = _TRUNCATE_TRIGGER_RE.search(text)
    return text[: match.start()] if match else text

def _strip_middle_boilerplate_clause(text: str) -> str:
    match = _MIDDLE_BOILERPLATE_CLAUSE_RE.search(text)
    if not match: return text
    trailing = text[match.end():].strip(" .,;:-")
    if len(trailing) >= 3 and re.search(r"[A-Za-z]{3,}", trailing):
        return f"{text[: match.start()].rstrip()}, {trailing}"
    return text

def _normalize_dash_entity_list(text: str) -> str:
    if _LEGAL_ENTITY_DASH_SIGNAL_RE.search(text): return _GENERIC_DASH_SEPARATOR_RE.sub(", ", text)
    return text

def _strip_transaction_header(text: str) -> str: return _TRANSACTION_HEADER_RE.sub("", text)

# ------------------------- 1b. Pola khusus tambahan (Data 3) -------------------
# A: "X Group of companies comprising of : Y" -> [X GROUP, Y]
# B: "Consists of :\n1. ...\n2. ..." -> list bernomor
# C: "As Per List Attached (Nama1 - Nama2) including any subsidiary" -> split isi kurung by dash
# D: bare AND/OR sbg separator entitas (lihat _ENTITY_SPLIT_RE)
# E: suffix "X group of companies" (tanpa comprising of) -> suffix dibuang, tetap 1 entitas
# F: "X (PT Y)" dgn Y nama company asli -> [X, Y]
# G: "ENTITY1 and/or associated... for their/its respective rights... ENTITY2" (Pola G)
# Pola B/C dicek sebelum truncate boilerplate umum karena trigger truncate
# justru bagian dari pola itu sendiri. Pola A sebelum E supaya "comprising of" aman.

_COMPRISING_OF_RE = re.compile(r"^(?P<head>.*?\bgroup)\s+of\s+companies\s+comprising\s+of\s*:\s*(?P<tail>.+)$", re.IGNORECASE | re.DOTALL)
_CONSISTS_OF_RE = re.compile(r"consists\s+of\s*:", RE_I)
_NUMBERED_ITEM_RE = re.compile(r"\d+\s*\.\s*")
_AS_PER_LIST_RE = re.compile(r"as\s+per\s+list\s+attached\s*\((?P<inner>[^)]+)\)", RE_I)
_GROUP_OF_COMPANIES_SUFFIX_RE = re.compile(r"\s+group\s+of\s+companies\b", RE_I)
_BARE_AND_OR_RE = re.compile(r"\band\b|\bor\b", RE_I)
_PAREN_PT_RE = re.compile(r"\(\s*(?P<inner>(?:PT\.?|P\.T\.?|CV\.?|C\.V\.?)\s+[^)]+)\)", RE_I)
_PAREN_STRIP_PREFIX_RE = re.compile(r"^(?:PT\.?|P\.T\.?|CV\.?|C\.V\.?)\s*", RE_I)

def _is_real_company_name(inner: str) -> bool:
    name = _PAREN_STRIP_PREFIX_RE.sub("", inner).strip()
    words = [w for w in re.split(r"[\s\-]+", name) if w]
    if len(words) < 2: return False
    has_lowercase = any(c.islower() for c in name)
    all_short_upper = all(len(w) <= 3 and w.isupper() for w in words)
    return not (all_short_upper and not has_lowercase)

def _try_pattern_f_paren_pt(text: str):
    m = _PAREN_PT_RE.search(text)
    if not m: return None
    inner = m.group("inner").strip()
    if not _is_real_company_name(inner): return None
    head, tail_after = text[: m.start()].strip(), text[m.end():].strip()
    if tail_after: head = f"{head} {tail_after}".strip()
    return [head, inner]

def _try_pattern_a_comprising_of(text: str):
    m = _COMPRISING_OF_RE.match(text.strip())
    return [m.group("head").strip(), m.group("tail").strip()] if m else None

def _try_pattern_b_consists_of(text: str):
    m = _CONSISTS_OF_RE.search(text)
    if not m: return None
    tail = _truncate_from_first_trigger(text[m.end():])
    items = [seg.strip() for seg in _NUMBERED_ITEM_RE.split(tail) if seg.strip()]
    return items or None

def _try_pattern_c_as_per_list(text: str):
    m = _AS_PER_LIST_RE.search(text)
    if not m: return None
    parts = [p.strip() for p in re.split(r"\s*[\u2013\u2014-]\s*", m.group("inner")) if p.strip()]
    return parts or None

def _strip_group_of_companies_suffix(text: str) -> str: return _GROUP_OF_COMPANIES_SUFFIX_RE.sub("", text)

def _truncate_boilerplate_tail(text: str) -> str:
    text = _truncate_from_first_trigger(text)
    match_generic = _GENERIC_BOILERPLATE_TAIL_RE.search(text)
    if match_generic: text = text[: match_generic.start()]
    return _REMOVE_ONLY_RE.sub(" ", text)

def _strip_titles(text: str) -> str: return _TITLE_RE.sub(" ", text)
def _strip_legal_entity(text: str) -> str: return _LEGAL_ENTITY_RE.sub(" ", text)

_DANGLING_WORD_RE = re.compile(
    r"^\b(?:AS|AN|A|AND|OR|DAN|ATAU|N|QQ|CQ|C\.Q\.|ITS|ALL|THEIR|FOR|THE|OF|RESPECTIVE)\b\s*"
    r"|\s*\b(?:AS|AN|A|AND|OR|DAN|ATAU|N|QQ|CQ|C\.Q\.|ITS|ALL|THEIR|FOR|THE|OF|RESPECTIVE)\b$", RE_I)

def _final_polish(text: str) -> str:
    t = re.sub(r"[().;:\"']", " ", text)
    t = re.sub(r"[-/]", " ", t)
    t = re.sub(r"\s+", " ", t).strip()
    previous = None
    while previous != t:
        previous, t = t, _DANGLING_WORD_RE.sub("", t).strip()
    return t.upper()

def _finalize_segments(raw_parts: list) -> list:
    """Tail-pipeline atas tiap segmen (truncate residu -> strip legal entity -> polish -> dedup)."""
    results = []
    for part in raw_parts:
        for sub in _ENTITY_SPLIT_RE.split(part):
            cleaned = _final_polish(_strip_legal_entity(_strip_group_of_companies_suffix(_truncate_boilerplate_tail(sub))))
            if cleaned and len(cleaned) >= 2 and re.search(r"[A-Z]", cleaned) and cleaned not in results:
                results.append(cleaned)
    return results

def _clean_insured_name_breakdown(text: str) -> list:
    t = _strip_middle_boilerplate_clause(_normalize_dash_entity_list(_strip_titles(_strip_transaction_header(text.strip()))))

    # Pola C/B duluan
    special_segments = _try_pattern_c_as_per_list(t) or _try_pattern_b_consists_of(t)
    if special_segments is not None: return _finalize_segments(special_segments)

    # Pola A sebelum E
    special_segments = _try_pattern_a_comprising_of(t)
    if special_segments is not None: return _finalize_segments(special_segments)

    special_segments = _try_pattern_f_paren_pt(t)
    if special_segments is not None: return _finalize_segments(special_segments)

    return _finalize_segments([_truncate_boilerplate_tail(_strip_group_of_companies_suffix(t))])

def clean_insured_name(text: str) -> list:
    if not isinstance(text, str) or not text.strip() or text.strip().lower() in ("nan", "none", "-"): return []
    original = text.strip()
    results = _clean_insured_name_breakdown(original)
    return [original] if len(results) > MAX_BREAKDOWN_INSURED else results

# ==============================================================================
# 2. CLEANSING POLIS & SLIP NO
# ==============================================================================
_TOK_MAIN_RE = re.compile(r"^\d{13}$")
_TOK_CERT_RE = re.compile(r"^\d{6}$")
_TOK_ENDORSE_RE = re.compile(r"^\d{1,2}/\d{1,2}$")
_TOK_STATUS_RE = re.compile(r"^(?:New|Endorsement|Cancel\s*All|Adjustment|Cancellation|Reinstatement)$", RE_I)

def _classify_token(tok: str) -> str:
    if _TOK_MAIN_RE.match(tok): return "MAIN"
    if _TOK_CERT_RE.match(tok): return "CERT"
    if _TOK_ENDORSE_RE.match(tok): return "ENDORSE"
    if _TOK_STATUS_RE.match(tok): return "STATUS"
    return "OTHER"

def clean_split_code(text: str) -> list:
    if not isinstance(text, str) or not text.strip() or text.strip().lower() in ("nan", "none", "-"): return []
    
    t = re.sub(r"\s+", " ", re.sub(r"[.,]", "", text.strip())).strip()
    classified = [(_classify_token(tok), tok) for tok in (s.strip() for s in t.split("-")) if tok]

    codes, i, n = [], 0, len(classified)
    while i < n:
        kind, tok = classified[i]
        if kind == "MAIN":
            suffix_parts, j = [], i + 1
            while j < n and classified[j][0] in ("CERT", "ENDORSE"):
                suffix_parts.append(classified[j][1])
                j += 1

            followed_by_status = j < n and classified[j][0] == "STATUS"

            if suffix_parts and not followed_by_status:
                combined = tok + "-" + "-".join(suffix_parts)
                if combined not in codes: codes.append(combined)
                if tok not in codes: codes.append(tok)
            elif tok not in codes:
                codes.append(tok)

            if followed_by_status: j += 1
            i = j
            continue
        i += 1

    if not codes: return [t] if t else []
    return codes[:1] if len(codes) > MAX_BREAKDOWN_CODES else codes

# ==============================================================================
# 3. LOAD & FILTER DATA
# ==============================================================================
def load_source(file_path: str, sheet_name: str, header_row: int) -> pd.DataFrame:
    return pd.read_excel(file_path, sheet_name=sheet_name, header=header_row)

def filter_cedant(df_raw: pd.DataFrame, keyword: str) -> pd.DataFrame:
    cedant_cols = [c for c in df_raw.columns if "CEDANT" in str(c).upper()]
    if not cedant_cols: raise ValueError("Kolom CEDANT tidak ditemukan di source file.")

    mask = pd.Series(False, index=df_raw.index)
    for col in cedant_cols: mask |= df_raw[col].astype(str).str.contains(keyword, case=False, na=False)

    df_filtered = df_raw[mask].copy()
    print(f"[INFO] Filter cedant mengandung '{keyword}': {len(df_filtered)} baris ditemukan dari total {len(df_raw)} baris.")
    return df_filtered

# ==============================================================================
# 4. BUILD OUTPUT
# ==============================================================================
def _detect_key_columns(df: pd.DataFrame) -> tuple[str, str, str]:
    col_insured = next((c for c in df.columns if "INSURED" in str(c).upper()), None)
    col_polis = next((c for c in df.columns if str(c).upper().strip() == "POLIS"), None)
    col_slip = next((c for c in df.columns if "SLIP" in str(c).upper()), None)

    if None in (col_insured, col_polis, col_slip):
        raise ValueError(f"Kolom kunci tidak lengkap. INSURED={col_insured}, POLIS={col_polis}, SLIP={col_slip}")
    print(f"[INFO] Kolom terdeteksi -> INSURED='{col_insured}', POLIS='{col_polis}', SLIP='{col_slip}'")
    return col_insured, col_polis, col_slip

def _breakdown_all_rows(df: pd.DataFrame, col_insured: str, col_polis: str, col_slip: str) -> tuple[list, list, list, int, int, int]:
    insured_cln_all, polis_cln_all, slip_cln_all = [], [], []
    max_ins = max_pol = max_slp = 0

    for _, row in df.iterrows():
        ins_val, pol_val, slp_val = row[col_insured], row[col_polis], row[col_slip]
        ins_cln = clean_insured_name(str(ins_val)) if pd.notna(ins_val) else []
        pol_cln = clean_split_code(str(pol_val)) if pd.notna(pol_val) else []
        slp_cln = clean_split_code(str(slp_val)) if pd.notna(slp_val) else []

        insured_cln_all.append(ins_cln)
        polis_cln_all.append(pol_cln)
        slip_cln_all.append(slp_cln)
        max_ins, max_pol, max_slp = max(max_ins, len(ins_cln)), max(max_pol, len(pol_cln)), max(max_slp, len(slp_cln))

    return insured_cln_all, polis_cln_all, slip_cln_all, max_ins, max_pol, max_slp

def build_output(df_filtered: pd.DataFrame) -> pd.DataFrame:
    col_insured, col_polis, col_slip = _detect_key_columns(df_filtered)
    insured_cln_all, polis_cln_all, slip_cln_all, max_ins, max_pol, max_slp = _breakdown_all_rows(df_filtered, col_insured, col_polis, col_slip)
    
    original_cols = list(df_filtered.columns)
    df_filtered = df_filtered.reset_index(drop=True)

    output_cols_data = {}
    for col in original_cols:
        output_cols_data[col] = df_filtered[col].values
        if col == col_insured:
            for i in range(1, max_ins + 1): output_cols_data[f"INSURED_CLN_{i}"] = [lst[i - 1] if i <= len(lst) else "" for lst in insured_cln_all]
        if col == col_polis:
            for i in range(1, max_pol + 1): output_cols_data[f"POLIS_CLN_{i}"] = [lst[i - 1] if i <= len(lst) else "" for lst in polis_cln_all]
        if col == col_slip:
            for i in range(1, max_slp + 1): output_cols_data[f"SLIP_NO_CLN_{i}"] = [lst[i - 1] if i <= len(lst) else "" for lst in slip_cln_all]

    df_master = pd.DataFrame(output_cols_data)
    drop_cols = [c for c in df_master.columns if "_CLN_" in c and df_master[c].astype(str).str.strip().eq("").all()]
    return df_master.drop(columns=drop_cols)

# ==============================================================================
# 5. SIMPAN OUTPUT
# ==============================================================================
def save_with_text_format(df: pd.DataFrame, output_path: str) -> None:
    df.to_excel(output_path, index=False)
    workbook = load_workbook(output_path)
    worksheet = workbook.active
    text_col_indices = [idx for idx, col in enumerate(df.columns, start=1) if any(kw in str(col).upper() for kw in TEXT_FORMAT_COLUMN_KEYWORDS)]

    for col_idx in text_col_indices:
        for row in range(2, worksheet.max_row + 1):
            cell = worksheet.cell(row=row, column=col_idx)
            cell.number_format = "@"
            if cell.value is not None: cell.value = str(cell.value)
    workbook.save(output_path)

# ==============================================================================
# 6. EXECUTION RUNNER
# ==============================================================================
def main() -> None:
    print("=" * 70 + "\n CLEANSING DATA 3 - SUSPENSE | CEDANT: PT LIPPO GENERAL INSURANCE\n" + "=" * 70)
    input_path = INPUT_FILE
    
    if not os.path.exists(input_path):
        candidates = [f for f in os.listdir(".") if f.endswith(".xlsx") and not f.startswith("Hasil_") and not f.startswith("Clean_")]
        if not candidates: raise FileNotFoundError("File Excel sumber tidak ditemukan di folder kerja.")
        input_path = candidates[0]
        print(f"[AUTO-DETECT] File input tidak ditemukan di path default, memakai: '{input_path}'")

    df_raw = load_source(input_path, SHEET_NAME, HEADER_ROW)
    print(f"[INFO] Total baris source: {len(df_raw)}")
    
    df_lippo = filter_cedant(df_raw, CEDANT_FILTER)
    if df_lippo.empty:
        print("[WARNING] Tidak ada baris yang cocok dengan filter cedant. Proses dihentikan.")
        return

    df_hasil = build_output(df_lippo)
    save_with_text_format(df_hasil, OUTPUT_FILE)
    print("=" * 70 + f"\n[SUCCESS] Selesai. Total baris output: {len(df_hasil)}\n[SUCCESS] File hasil: '{OUTPUT_FILE}'\n" + "=" * 70)

if __name__ == "__main__":
    main()