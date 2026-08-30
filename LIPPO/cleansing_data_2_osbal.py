"""
================================================================================
 SCRIPT CLEANSING DATA 2 - OSBAL (KHUSUS CEDANT: PT LIPPO GENERAL INSURANCE)
================================================================================
Dipakai untuk breakdown & cleansing kolom INSURED, POLIS, dan SLIP NO supaya
bisa dipakai sebagai key matching antar database.

CHANGELOG (revisi):
- Noise words diperluas: PENDING SLIP, CN, currency code, "NO :", kode surat.
- SD-range bisa reconstruct suffix pendek (mis. "000010 SD 000012") ke base.
- Angka pendek (<12 digit) yang berdiri sendiri dan gak bisa direkonstruksi DIBUANG.
- "NO :", kode surat, dan koma list di-strip sebagai noise duluan.
- [REVISI BARU] Kalimat narasi panjang di DEPAN nomor polis/slip di-strip 
  sebelum breakdown jalan, selama masih ada kode valid di baliknya.
- [REVISI BARU] Fix urutan kolom output: kolom *_CLN sekarang nempel LANGSUNG
  di sebelah kolom original-nya, bukan numpuk semua di paling ujung kanan.
- [REVISI BARU] Suffix endorsement format "N/M" dan kata "Endorsement" dibuang 
  total dari hasil breakdown, murni nomor polis/slip aja.
================================================================================
"""

import os
import re
import pandas as pd
from openpyxl import load_workbook

# ==============================================================================
# 0. KONFIGURASI
# ==============================================================================
INPUT_FILE = "Data_2_Osbal.xlsx"
SHEET_NAME = "Sheet0"
HEADER_ROW = 0
OUTPUT_FILE = "[3Agustus2026] lippo_output_osbal.xlsx"
CEDANT_FILTER = "LIPPO"

CEDANT_COLUMN = "CCOS_COMP_NAME"
INSURED_COLUMN = "FAC_INSURED"
POLIS_COLUMN = "FAC_POLICY_NO"
SLIP_COLUMN = "FAC_SLIP"

MAX_BREAKDOWN_CODES = 5
TEXT_FORMAT_COLUMN_KEYWORDS = (
    "POLIS", "SLIP", "RECEIPT NO", "CREDIT NOTES", "DETAIL RINCIAN NO",
    "FAC_POLICY_NO", "FAC_SLIP",
)

# ==============================================================================
# 1. CLEANSING INSURED
# ==============================================================================
_TRANSACTION_HEADER_RE = re.compile(
    r"^\s*(?:Pembayaran|Penagihan)?\s*(?:dan\s+Penagihan)?\s*Premi\s*"
    r"(?:Fakultatif|SOA|Pensesian)?\s*(?:PAR\s*&?\s*EQ|BIT|GIT|CECR|QS)?\s*", flags=re.IGNORECASE)
_TRUNCATE_TRIGGER_RE = re.compile(
    r"\bsubsidiar|\bassociat|\baffiliat|\bfiliated\b|related\s+compan|"
    r"respective\s+right|\bincluding\s+an[yd]|\bindu?ding\s+any|as\s+the\s+owner|"
    r"as\s+owners?\b|as\s+main\s+contractor|\(as\s+mortgagee\s+and\s+loss\s+payee\)", flags=re.IGNORECASE)
_REMOVE_ONLY_RE = re.compile(r"\bas\s+principal\b|\bsebagai\s+principal\b|\bsebagai\s+kontraktor\b", flags=re.IGNORECASE)
_ONLY_WORD_RE = re.compile(r"\bONLY\b", flags=re.IGNORECASE)
_GENERIC_BOILERPLATE_TAIL_RE = re.compile(r"\bsub\s*-?\s*kontraktor\b.*$|\banak\s+perusahaan\b.*$|\bafiliasi\s+perusahaan\b.*$|\bdan\s+semua\b\s*[–-]?\s*$", flags=re.IGNORECASE)
_TITLE_RE = re.compile(
    r"\b(?:S\.?H\.?|S\.?I\.?Kom\.?|S\.?E\.?|S\.?T\.?|S\.?Kom\.?|S\.?Psi\.?|"
    r"M\.?H\.?|M\.?M\.?|M\.?Sc\.?|Ir\.|Drs\.|Dra\.|Prof\.|Dr\.|"
    r"Tn\.?|Bpk\.?|Bapak|Ibu|Ny\.?|Nyonya|Mr\.?|Mrs\.?|Ms\.?)\b", flags=re.IGNORECASE)
_LEGAL_ENTITY_RE = re.compile(r"\b(?:PT\.?|P\.T\.?|CV\.?|C\.V\.?|TBK\.?|Tbk\.?|PERSERO|Perseroan\s+Terbatas)\b", flags=re.IGNORECASE)
_ENTITY_SPLIT_RE = re.compile(r"\b(?:QQ|Q\.Q\.)\b|\band\s*/\s*or\b|\bdan\s*/\s*atau\b|\bin\s+this\s+case\b|\bdalam\s+hal\s+ini\b|,|/", flags=re.IGNORECASE)
_LEGAL_ENTITY_DASH_SIGNAL_RE = re.compile(r",\s*(?:PT\.?|P\.T\.?|CV\.?|C\.V\.?|TBK\.?|Tbk\.?|PERSERO)\s*[\r\n]*\s*-\s*", flags=re.IGNORECASE)
_GENERIC_DASH_SEPARATOR_RE = re.compile(r"\s*[\r\n]+\s*-\s*|\s+-\s+")

_TRAILING_PAREN_MERGE_RE = re.compile(r",?\s*(?:PT\.?|P\.T\.?|CV\.?|C\.V\.?|TBK\.?)?\s*\(([^)]+)\)\s*$", flags=re.IGNORECASE)
_INLINE_PAREN_MERGE_RE = re.compile(r",?\s*(?:PT\.?|P\.T\.?|CV\.?|C\.V\.?|TBK\.?)?\s*\(([^)]+)\)", flags=re.IGNORECASE)
_QQ_PAREN_UNWRAP_RE = re.compile(r"\bQQ\s*\(\s*([^)]*)\)", flags=re.IGNORECASE)

def _unwrap_qq_paren(text: str) -> str:
    def _replace(match: "re.Match") -> str:
        inner = re.sub(r"[:;]", " ", match.group(1).strip())
        inner = re.sub(r"\s+", " ", inner).strip()
        return "QQ , " + inner if inner else "QQ"
    return _QQ_PAREN_UNWRAP_RE.sub(_replace, text)

def _merge_trailing_paren(text: str) -> str:
    match = _TRAILING_PAREN_MERGE_RE.search(text)
    if not match: return text
    inner = match.group(1).strip()
    if not inner or len(inner.split()) > 4 or _TRUNCATE_TRIGGER_RE.search(inner):
        return text[: match.start()] + text[match.end():]
    return text[: match.start()] + " " + inner

def _merge_inline_paren(text: str) -> str:
    def _replace(match: "re.Match") -> str:
        inner = match.group(1).strip()
        if not inner or len(inner.split()) > 4 or _TRUNCATE_TRIGGER_RE.search(inner): return ""
        return " " + inner
    return _INLINE_PAREN_MERGE_RE.sub(_replace, text)

_MONTH_WORDS = (
    "JANUARI", "FEBRUARI", "MARET", "APRIL", "MEI", "JUNI", "JULI", "AGUSTUS", "SEPTEMBER", "OKTOBER", "NOVEMBER", "DESEMBER", "DES",
    "JANUARY", "FEBRUARY", "MARCH", "MAY", "JUNE", "JULY", "AUGUST", "OCTOBER", "DECEMBER",
    "JAN", "FEB", "MAR", "APR", "JUN", "JUL", "AUG", "SEP", "SEPT", "OCT", "NOV", "DEC",
)
_DATE_BATCH_TOKEN_RE = re.compile(r"\b(?:" + "|".join(_MONTH_WORDS) + r")\b|\b(?:19|20)\d{2}\b|\bBATCH\s*\d*\b", flags=re.IGNORECASE)

def _strip_date_batch_info(text: str) -> str: return _DATE_BATCH_TOKEN_RE.sub(" ", text)
def _truncate_from_first_trigger(text: str) -> str:
    match = _TRUNCATE_TRIGGER_RE.search(text)
    return text[: match.start()] if match else text

def _normalize_dash_entity_list(text: str) -> str:
    if _LEGAL_ENTITY_DASH_SIGNAL_RE.search(text): return _GENERIC_DASH_SEPARATOR_RE.sub(", ", text)
    return text

def _strip_transaction_header(text: str) -> str: return _TRANSACTION_HEADER_RE.sub("", text)
def _strip_titles(text: str) -> str: return _TITLE_RE.sub(" ", text)
def _strip_legal_entity(text: str) -> str: return _LEGAL_ENTITY_RE.sub(" ", text)

def _truncate_boilerplate_tail(text: str) -> str:
    text = _truncate_from_first_trigger(text)
    match_generic = _GENERIC_BOILERPLATE_TAIL_RE.search(text)
    if match_generic: text = text[: match_generic.start()]
    return _REMOVE_ONLY_RE.sub(" ", text)

def _final_polish(text: str) -> str:
    t = re.sub(r"[().;:\"']", " ", text)
    t = re.sub(r"[-/]", " ", t)
    t = re.sub(r"\s+", " ", t).strip()
    dangling_pattern = (
        r"^\b(?:AS|AN|A|AND|OR|DAN|ATAU|N|QQ|CQ|C\.Q\.|ITS|ALL|THEIR|FOR|THE|OF|RESPECTIVE)\b\s*"
        r"|\s*\b(?:AS|AN|A|AND|OR|DAN|ATAU|N|QQ|CQ|C\.Q\.|ITS|ALL|THEIR|FOR|THE|OF|RESPECTIVE)\b$"
    )
    previous = None
    while previous != t:
        previous = t
        t = re.sub(dangling_pattern, "", t, flags=re.IGNORECASE).strip()
    return t.upper()

_LIPPO_KNOWN_ENTITIES = ["MATAHARI PUTRA PRIMA", "MATAHARI BOSTON DRIGSTORE", "MATAHARI PUSAKA TAMA", "LIPPO GROUP"]
_LIPPO_KNOWN_ENTITIES_SORTED = sorted(_LIPPO_KNOWN_ENTITIES, key=len, reverse=True)
_KNOWN_ENTITY_RE = re.compile("|".join(re.escape(name) for name in _LIPPO_KNOWN_ENTITIES_SORTED))
_ENTITY_SLASH_NORMALIZE_PATTERNS = [(name, re.compile(r"\b" + r"(?:\s*/\s*|\s+)".join(re.escape(w) for w in name.split()) + r"\b", flags=re.IGNORECASE)) for name in _LIPPO_KNOWN_ENTITIES_SORTED if len(name.split()) > 1]

def _normalize_entity_slash_variants(text: str) -> str:
    for canonical, pattern in _ENTITY_SLASH_NORMALIZE_PATTERNS: text = pattern.sub(canonical, text)
    return text

def _split_by_known_entities(text: str):
    pos, n, found = 0, len(text), []
    while pos < n:
        while pos < n and text[pos] == " ": pos += 1
        if pos >= n: break
        match = _KNOWN_ENTITY_RE.match(text, pos)
        if not match: return None
        found.append(match.group(0))
        pos = match.end()
    return found if len(found) > 1 else None

_INSURED_DESCRIPTIVE_RE = re.compile(r"\bVARIOUS\s+INSUREDS?\b|\bACCEPTED\s+BY\b", flags=re.IGNORECASE)

def clean_insured_name(text: str) -> list:
    if not isinstance(text, str) or not text.strip() or text.strip().lower() in ("nan", "none", "-"): return []
    original = text.strip()
    if _INSURED_DESCRIPTIVE_RE.search(original): return [original]

    t = _ONLY_WORD_RE.sub(" ", original)
    t = _unwrap_qq_paren(t)
    t = _merge_trailing_paren(t)
    t = _merge_inline_paren(t)
    t = _strip_transaction_header(t)
    t = _strip_titles(t)
    t = _normalize_dash_entity_list(t)
    t = _normalize_entity_slash_variants(t)
    t = _truncate_boilerplate_tail(t)

    raw_parts = _ENTITY_SPLIT_RE.split(t)
    results = []
    for part in raw_parts:
        cleaned = _final_polish(_strip_date_batch_info(_strip_legal_entity(_truncate_boilerplate_tail(part))))
        if not cleaned or len(cleaned) < 2 or not re.search(r"[A-Z]", cleaned): continue

        known_split = _split_by_known_entities(cleaned)
        if known_split is not None:
            for name in known_split:
                if name not in results: results.append(name)
            continue
        if cleaned not in results: results.append(cleaned)

    return [original] if len(results) > MAX_BREAKDOWN_CODES else results

# ==============================================================================
# 2. CLEANSING POLIS & SLIP NO (REVISI BARU)
# ==============================================================================
_MAIN_LEN_MIN, _MAIN_LEN_MAX = 12, 13
_MAIN_DIGITS_PATTERN = r"\d{" + str(_MAIN_LEN_MIN) + "," + str(_MAIN_LEN_MAX) + "}"
_TOK_MAIN_RE = re.compile("^" + _MAIN_DIGITS_PATTERN + "$")
_TOK_CERT_RE = re.compile(r"^\d{6}$")
_TOK_SUFFIX_REPLACE_RE = re.compile(r"^\d{1,11}$")
_TOK_ENDORSE_RE = re.compile(r"^\d{1,2}[/\x00]\d{1,2}$")
_TOK_PLACEHOLDER_RE = re.compile(r"^P\d{1,2}$", flags=re.IGNORECASE)
_TOK_STATUS_RE = re.compile(r"^(?:New|End(?:orsement)?|Cancel\s*All|Adjustment|Cancellation|Reinstatement)$", flags=re.IGNORECASE)
_VARIOUS_WORD_RE = re.compile(r"\bVAR(?:IOUS)?\b", flags=re.IGNORECASE)
_SD_WORD_RE = re.compile(r"\bSD\b", flags=re.IGNORECASE)
_AMPERSAND_SUFFIX_AMBIGUOUS_RE = re.compile(r"\d+\s*&\s*\d+")
_CURRENCY_RE = re.compile(r"\b(?:IDR|USD|SGD|EUR|JPY|GBP|AUD|MYR)\b", flags=re.IGNORECASE)
_SEE_ATTACHMENT_RE = re.compile(r"\bSEE\s+ATTACHMENT\b", flags=re.IGNORECASE)
_PENDING_SLIP_RE = re.compile(r"\bPENDING\s+SLIP\b", flags=re.IGNORECASE)
_NO_LABEL_RE = re.compile(r"\bNO\s*[:.]", flags=re.IGNORECASE)
_LETTER_REF_CODE_RE = re.compile(r"\b(?:\d{1,6}\s*/\s*)?[A-Z]{1,6}-[A-Z]{1,6}\s*/\s*[IVXLCM]{1,6}\s*/\s*\d{2,4}\b", flags=re.IGNORECASE)
_BATCH_TOKEN_RE = re.compile(r"\bBATCH\s*\d*\b", flags=re.IGNORECASE)
_BORDERO_WORD_RE = re.compile(r"\bBORDERO\b|\bBORD\b", flags=re.IGNORECASE)
_POLIS_MONTH_RE = re.compile(r"\b(?:" + "|".join(_MONTH_WORDS) + r")\.?\s*(?:\d{2,4})?\b", flags=re.IGNORECASE)
_POLIS_DATE_SLASH_RE = re.compile(r"\b\d{1,2}/\d{1,2}/\d{2,4}\b")
_POLIS_YEAR_RE = re.compile(r"\b(?:19|20)\d{2}\b")
_POLIS_CN_RE = re.compile(r"\bCN\b", flags=re.IGNORECASE)
_POLIS_NARRATIVE_WORDS_RE = re.compile(r"\bSANY\b|\bNO\s+SLIP\b|\bNO\s+POLIS\b|\bPOLIS\s+VARIOUS\b", flags=re.IGNORECASE)
_POLIS_STATUS_WORDS_RE = re.compile(r"\bCancel\s*All\b|\bNew\b|\bEnd(?:orsement)?\b|\bAdjustment\b|\bCancellation\b|\bReinstatement\b", flags=re.IGNORECASE)
_EDGE_SEPARATOR_RE = re.compile(r"^[\s\-/,]+|[\s\-/,]+$")

_NARRATIVE_PREFIX_RE = re.compile(
    r"^\s*(?:JANUARI|FEBRUARI|MARET|APRIL|MEI|JUNI|JULI|AGUSTUS|SEPTEMBER|OKTOBER|NOVEMBER|DESEMBER)"
    r"\s+\d{4}\s*-\s*(?:IDR|USD|SGD|EUR|JPY|GBP|AUD|MYR)?\s*/?\s*", flags=re.IGNORECASE)

# [REVISI BARU] Membuang narasi panjang di DEPAN kode
_MAIN_CODE_FIRST_RE = re.compile(_MAIN_DIGITS_PATTERN)

def _strip_leading_narrative(text: str) -> str:
    match = _MAIN_CODE_FIRST_RE.search(text)
    if not match: return text
    prefix = text[: match.start()]
    if re.search(r"[A-Za-z]", prefix): return text[match.start():]
    return text

def _strip_polis_slip_noise(text: str) -> str:
    t = _NARRATIVE_PREFIX_RE.sub("", text)
    t = _LETTER_REF_CODE_RE.sub(" ", t)
    t = _NO_LABEL_RE.sub(" ", t)
    t = _PENDING_SLIP_RE.sub(" ", t)
    t = _SEE_ATTACHMENT_RE.sub(" ", t)
    t = _POLIS_NARRATIVE_WORDS_RE.sub(" ", t)
    t = _POLIS_STATUS_WORDS_RE.sub(" ", t)
    t = _BATCH_TOKEN_RE.sub(" ", t)
    t = _BORDERO_WORD_RE.sub(" ", t)
    t = _POLIS_DATE_SLASH_RE.sub(" ", t)
    t = _POLIS_MONTH_RE.sub(" ", t)
    t = _POLIS_YEAR_RE.sub(" ", t)
    t = _POLIS_CN_RE.sub(" ", t)
    t = _CURRENCY_RE.sub(" ", t)
    t = _EDGE_SEPARATOR_RE.sub("", re.sub(r"\s+", " ", t).strip()).strip()
    t = re.sub(r"\s+", " ", t).strip()
    return _strip_leading_narrative(t)

_STRONG_SPLIT_CAPTURE_RE = re.compile(r"(\+|-|[,/]|\s+)")

def _split_preserving_continuity(t: str):
    parts = _STRONG_SPLIT_CAPTURE_RE.split(t)
    segments, is_soft_list, pending_soft = [], [], False
    for part in parts:
        if part in ("+", "-"):
            pending_soft = True
            continue
        if part is not None and part.strip() == "" and _STRONG_SPLIT_CAPTURE_RE.fullmatch(part or ""): continue
        if part is not None and _STRONG_SPLIT_CAPTURE_RE.fullmatch(part):
            pending_soft = False
            continue
        seg = (part or "").strip()
        if seg:
            segments.append(seg)
            is_soft_list.append(pending_soft)
            pending_soft = False
    return segments, is_soft_list

_ALPHA_WORD_RE = re.compile(r"[A-Za-z][A-Za-z0-9]*")
_ALLOWED_ALPHA_TOKENS = {"NEW", "END", "ENDORSEMENT", "CANCEL", "ALL", "ADJUSTMENT", "CANCELLATION", "REINSTATEMENT", "VARIOUS", "VAR", "SD"}

def _has_descriptive_narrative(text: str) -> bool:
    for word in _ALPHA_WORD_RE.findall(text):
        if _TOK_PLACEHOLDER_RE.match(word) or word.upper() in _ALLOWED_ALPHA_TOKENS: continue
        return True
    return False

def _split_sd_range(text: str):
    if not _SD_WORD_RE.search(text): return None
    parts = _SD_WORD_RE.split(text, maxsplit=1)
    if len(parts) != 2: return "__FALLBACK__"

    left_all_tokens, right_all_tokens = re.findall(r"\d+", parts[0]), re.findall(r"\d+", parts[1])
    if not left_all_tokens or not right_all_tokens: return "__FALLBACK__"

    left_last, right_first = left_all_tokens[-1], right_all_tokens[0]
    if not _TOK_MAIN_RE.match(left_last) or not _TOK_MAIN_RE.match(right_first): return "__FALLBACK__"

    expanded = _expand_numeric_range(left_last, right_first)
    if expanded is None: return "__FALLBACK__"

    extra_valid = [t for t in right_all_tokens[1:] if _TOK_MAIN_RE.match(t) and t not in expanded]
    return expanded + extra_valid

def _expand_numeric_range(main_a: str, main_b: str):
    if len(main_a) != len(main_b) or not (main_a.isdigit() and main_b.isdigit()): return None
    diff_len = next((i for i in range(1, len(main_a) + 1) if main_a[-i] != main_b[-i]), 0)
    if diff_len == 0 or main_a[:-diff_len] != main_b[:-diff_len]: return None
    try:
        start, end = int(main_a[-diff_len:]), int(main_b[-diff_len:])
    except ValueError: return None
    if start > end or (end - start + 1) > MAX_BREAKDOWN_CODES: return None
    return [f"{main_a[:-diff_len]}{str(n).zfill(diff_len)}" for n in range(start, end + 1)]

def _try_reconstruct_suffix23(tokens: list, tokens_is_soft: list):
    if not tokens or not _TOK_MAIN_RE.match(tokens[0]): return None
    if not all(_TOK_MAIN_RE.match(tok) or _TOK_SUFFIX_REPLACE_RE.match(tok) for tok in tokens): return None
    for tok, soft in zip(tokens, tokens_is_soft):
        if _TOK_SUFFIX_REPLACE_RE.match(tok) and not _TOK_MAIN_RE.match(tok) and not soft: return None
    if not any(_TOK_SUFFIX_REPLACE_RE.match(tok) and not _TOK_MAIN_RE.match(tok) for tok in tokens): return None

    codes, current_base = [], None
    for tok in tokens:
        if _TOK_MAIN_RE.match(tok):
            current_base = tok
            if current_base not in codes: codes.append(current_base)
        elif current_base:
            reconstructed = current_base[:-len(tok)] + tok
            if reconstructed not in codes: codes.append(reconstructed)
    return codes

def clean_split_code(text: str) -> list:
    if not isinstance(text, str) or not text.strip() or text.strip().lower() in ("nan", "none", "-"): return []

    original = text.strip()
    t = _strip_polis_slip_noise(original)
    if not t: return [original]
    if _has_descriptive_narrative(t) or _AMPERSAND_SUFFIX_AMBIGUOUS_RE.search(t): return [original]

    sd_result = _split_sd_range(t)
    if sd_result == "__FALLBACK__": return [original]
    if sd_result is not None: return sd_result

    t = re.sub(r"[\u2012\u2013\u2014\u2015]", "-", _VARIOUS_WORD_RE.sub(" ", t))
    t = re.sub(r"(-\s*\d{1,2})\s*/\s*(\d{1,2})\b", lambda m: f"{m.group(1)}\x00{m.group(2)}", t)

    strong_segments, segment_is_soft = _split_preserving_continuity(t)
    raw_tokens = [re.sub(r"[.]", "", seg).strip() for seg in strong_segments]

    filtered_pairs = [(tok, soft) for tok, soft in zip(raw_tokens, segment_is_soft) if tok and not _TOK_PLACEHOLDER_RE.match(tok)]
    if not filtered_pairs: return [original]

    tokens, tokens_is_soft = [p[0] for p in filtered_pairs], [p[1] for p in filtered_pairs]
    reconstructed = _try_reconstruct_suffix23(tokens, tokens_is_soft)
    if reconstructed is not None: return [original] if len(reconstructed) > MAX_BREAKDOWN_CODES else reconstructed

    codes, current_base = [], None
    for tok, soft in zip(tokens, tokens_is_soft):
        # [REVISI BARU] Token ENDORSE murni ('n\x00m') dibuang total, murni nomor polis aja
        if current_base and soft and re.match(r"^\d{1,2}\x00\d{1,2}$", tok): continue

        if current_base and soft and _TOK_SUFFIX_REPLACE_RE.match(tok) and not _TOK_MAIN_RE.match(tok) and len(tok) < len(current_base):
            reconstructed_code = current_base[: -len(tok)] + tok
            if reconstructed_code not in codes: codes.append(reconstructed_code)
            continue

        if "\x00" in tok and not _TOK_MAIN_RE.match(tok):
            base_part_match = re.match("^(" + _MAIN_DIGITS_PATTERN + r")-?(\d{1,2})\x00(\d{1,2})$", tok)
            if base_part_match:
                base_num = base_part_match.group(1)
                current_base = base_num
                # [REVISI BARU] Endorsement sequence (N/M) dibuang
                if base_num not in codes: codes.append(base_num)
                continue

        parts_dash = tok.split("-")
        anchor = parts_dash[0].replace("\x00", "/")
        suffix_parts = [p.replace("\x00", "/") for p in parts_dash[1:]]

        # Buang angka pendek yatim
        if not _TOK_MAIN_RE.match(anchor): continue
        current_base = anchor

        while suffix_parts and _TOK_STATUS_RE.match(suffix_parts[-1]): suffix_parts.pop()

        # [REVISI BARU] Suffix endorsement format N/M dibuang total
        for p in suffix_parts:
            if _TOK_ENDORSE_RE.match(p.replace("/", "\x00")) or re.match(r"^\d{1,2}/\d{1,2}$", p): continue
            if _TOK_SUFFIX_REPLACE_RE.match(p) and len(p) < len(anchor):
                reconstructed_code = anchor[: -len(p)] + p
                if reconstructed_code not in codes: codes.append(reconstructed_code)

        if anchor not in codes: codes.append(anchor)

    if not codes or len(codes) > MAX_BREAKDOWN_CODES: return [original]
    return codes

# ==============================================================================
# 3. LOAD & FILTER DATA
# ==============================================================================
def load_source(file_path: str, sheet_name: str, header_row: int) -> pd.DataFrame:
    return pd.read_excel(file_path, sheet_name=sheet_name, header=header_row)

def filter_cedant(df_raw: pd.DataFrame, keyword: str, cedant_column: str = "") -> pd.DataFrame:
    if cedant_column:
        if cedant_column not in df_raw.columns: raise ValueError(f"Kolom {cedant_column} tidak ada.")
        cedant_cols = [cedant_column]
    else:
        cedant_cols = [c for c in df_raw.columns if "CEDANT" in str(c).upper()]
        if not cedant_cols: raise ValueError("Kolom CEDANT tidak ditemukan lewat auto-detect.")

    mask = pd.Series(False, index=df_raw.index)
    for col in cedant_cols: mask |= df_raw[col].astype(str).str.contains(keyword, case=False, na=False)

    df_filtered = df_raw[mask].copy()
    print(f"[INFO] Filter cedant '{keyword}': {len(df_filtered)} dari total {len(df_raw)} baris.")
    return df_filtered

# ==============================================================================
# 4. BUILD OUTPUT
# ==============================================================================
def _detect_key_columns(df: pd.DataFrame, insured_col: str, polis_col: str, slip_col: str):
    def _resolve(explicit_col: str, keyword_or_exact, label: str):
        if explicit_col:
            if explicit_col not in df.columns: raise ValueError(f"{label}='{explicit_col}' tidak ada.")
            return explicit_col
        if callable(keyword_or_exact): return next((c for c in df.columns if keyword_or_exact(str(c).upper())), None)
        return next((c for c in df.columns if keyword_or_exact in str(c).upper()), None)

    col_insured = _resolve(insured_col, "INSURED", "INSURED_COLUMN")
    col_polis = _resolve(polis_col, lambda c: c.strip() == "POLIS", "POLIS_COLUMN")
    col_slip = _resolve(slip_col, "SLIP", "SLIP_COLUMN")

    if None in (col_insured, col_polis, col_slip): raise ValueError("Kolom kunci tidak lengkap.")
    print(f"[INFO] Kolom terdeteksi -> INSURED='{col_insured}', POLIS='{col_polis}', SLIP='{col_slip}'")
    return col_insured, col_polis, col_slip

def _breakdown_all_rows(df: pd.DataFrame, col_insured: str, col_polis: str, col_slip: str):
    insured_cln_all, polis_cln_all, slip_cln_all, max_ins, max_pol, max_slp = [], [], [], 0, 0, 0
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
    col_insured, col_polis, col_slip = _detect_key_columns(df_filtered, INSURED_COLUMN, POLIS_COLUMN, SLIP_COLUMN)
    insured_cln_all, polis_cln_all, slip_cln_all, max_ins, max_pol, max_slp = _breakdown_all_rows(df_filtered, col_insured, col_polis, col_slip)

    original_cols = list(df_filtered.columns)
    df_filtered = df_filtered.reset_index(drop=True)
    
    # [REVISI BARU] Dict di Python ngikutin urutan insert, jadi urutan insert = urutan kolom final.
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
    print("=" * 70 + "\n CLEANSING DATA 2 - OSBAL | CEDANT: PT LIPPO GENERAL INSURANCE\n" + "=" * 70)
    input_path = INPUT_FILE
    
    if not os.path.exists(input_path):
        candidates = [f for f in os.listdir(".") if f.endswith(".xlsx") and not f.startswith("Hasil_") and not f.startswith("Clean_")]
        if not candidates: raise FileNotFoundError("File Excel sumber tidak ditemukan di folder kerja.")
        input_path = candidates[0]
        print(f"[AUTO-DETECT] Memakai file terdeteksi: '{input_path}'")

    df_raw = load_source(input_path, SHEET_NAME, HEADER_ROW)
    df_lippo = filter_cedant(df_raw, CEDANT_FILTER, CEDANT_COLUMN)

    if not df_lippo.empty:
        df_hasil = build_output(df_lippo)
        save_with_text_format(df_hasil, OUTPUT_FILE)
        print("=" * 70 + f"\n[SUCCESS] Selesai. Output: '{OUTPUT_FILE}' ({len(df_hasil)} baris)\n" + "=" * 70)
    else:
        print("[WARNING] Tidak ada baris yang cocok dengan filter.")

if __name__ == "__main__":
    main()

