"""
================================================================================
 SCRIPT CLEANSING DATA 1 - FACULTATIVE (KHUSUS CEDANT: PT LIPPO GENERAL INSURANCE)
================================================================================
Dipakai untuk breakdown & cleansing kolom INSURED, POLIS, dan SLIP NO supaya
bisa dipakai sebagai key matching antar database.
================================================================================
"""

import os
import re
import pandas as pd
from openpyxl import load_workbook

# ==============================================================================
# 0. KONFIGURASI
# ==============================================================================
INPUT_FILE = "Lippo_INS_Data_1.xlsx"
SHEET_NAME = "Sheet0"
HEADER_ROW = 0
OUTPUT_FILE = "[2Juli2026] lippo_output_facul.xlsx"
CEDANT_FILTER = "LIPPO"

CEDANT_COLUMN = "COMP_NAME"
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
    r"\bsubsidiar|\b(?:and\s*/\s*or|and|its|their|all)\s+associat(?!ion)|&\s*/\s*or\s+associat(?!ion)|\baffiliat|\bfiliated\b|related\s+compan|"
    r"respective\s+right|\bincluding\s+an[yd]|\bindu?ding\s+any|as\s+the\s+owner|"
    r"as\s+owners?\b|as\s+main\s+contractor|\(as\s+mortgagee\s+and\s+loss\s+payee\)", flags=re.IGNORECASE)

def _truncate_from_first_trigger(text: str) -> str:
    match = _TRUNCATE_TRIGGER_RE.search(text)
    return text[: match.start()] if match else text

_REMOVE_ONLY_RE = re.compile(r"\bas\s+principal\b|\bsebagai\s+principal\b|\bsebagai\s+kontraktor\b", flags=re.IGNORECASE)
_GENERIC_BOILERPLATE_TAIL_RE = re.compile(r"\bsub\s*-?\s*kontraktor\b.*$|\banak\s+perusahaan\b.*$|\bafiliasi\s+perusahaan\b.*$|\bdan\s+semua\b\s*[–-]?\s*$", flags=re.IGNORECASE)
_TITLE_RE = re.compile(
    r"\b(?:S\.?H\.?|S\.?I\.?Kom\.?|S\.?E\.?|S\.?T\.?|S\.?Kom\.?|S\.?Psi\.?|"
    r"M\.?H\.?|M\.?M\.?|M\.?Sc\.?|Ir\.|Drs\.|Dra\.|Prof\.|Dr\.|"
    r"Tn\.?|Bpk\.?|Bapak|Ibu|Ny\.?|Nyonya|Mr\.?|Mrs\.?|Ms\.?)\b", flags=re.IGNORECASE)
_LEGAL_ENTITY_RE = re.compile(r"\b(?:PT\.?|P\.T\.?|CV\.?|C\.V\.?|TBK\.?|Tbk\.?|PP\.?|PERSERO|Perseroan\s+Terbatas|LTD\.?)\b", flags=re.IGNORECASE)

# --- PATCH: koma DIHAPUS dari splitter (dulu ada ",", ";", "/" jadi splitter).
# Alasan: koma di data ini hampir selalu dipakai sbg pemisah "NAMA,PT" atau
# rincian nama panjang (mis. nama dinas pemerintahan), BUKAN pemisah 2 entitas
# insured yang berbeda. Pemisah entitas beneran yang dipakai secara konsisten
# di data adalah "/", "QQ", dan "AND/OR". Semicolon dipertahankan.
_ENTITY_SPLIT_RE = re.compile(r"\b(?:QQ|Q\.Q\.)\b|\band\s*/\s*or\b|\bdan\s*/\s*atau\b|\bin\s+this\s+case\b|\bdalam\s+hal\s+ini\b|;|/", flags=re.IGNORECASE)

_INSURED_VARIOUS_RE = re.compile(r"\bVARIOUS\b", flags=re.IGNORECASE)

# --- PATCH: noise "<BULAN> BATCH <angka>" atau "BATCH <angka>" dibuang total
# dari nama insured. Nama bulan HANYA dibuang kalau menempel langsung dengan kata BATCH.
_INSURED_BATCH_NOISE_RE = re.compile(
    r"\b(?:JAN(?:UARI|UARY)?|FEB(?:RUARI|RUARY)?|MAR(?:ET|CH)?|APR(?:IL)?|MEI|MAY|JUN(?:I|E)?|"
    r"JUL(?:I|Y)?|AGU(?:STUS)?|AUG(?:UST)?|SEP(?:TEMBER)?|OKT(?:OBER)?|OCT(?:OBER)?|"
    r"NOV(?:EMBER)?|DES(?:EMBER)?|DEC(?:EMBER)?)?\s*\bBATCH\.?\s*\d*"
    r"(?:\s*(?:JAN(?:UARI|UARY)?|FEB(?:RUARI|RUARY)?|MAR(?:ET|CH)?|APR(?:IL)?|MEI|MAY|JUN(?:I|E)?|"
    r"JUL(?:I|Y)?|AGU(?:STUS)?|AUG(?:UST)?|SEP(?:TEMBER)?|OKT(?:OBER)?|OCT(?:OBER)?|"
    r"NOV(?:EMBER)?|DES(?:EMBER)?|DEC(?:EMBER)?)\b)?"
    r"|\d+(?=\s*(?:JAN(?:UARI|UARY)?|FEB(?:RUARI|RUARY)?|MAR(?:ET|CH)?|APR(?:IL)?|MEI|MAY|JUN(?:I|E)?|"
    r"JUL(?:I|Y)?|AGU(?:STUS)?|AUG(?:UST)?|SEP(?:TEMBER)?|OKT(?:OBER)?|OCT(?:OBER)?|"
    r"NOV(?:EMBER)?|DES(?:EMBER)?|DEC(?:EMBER)?)\b)", flags=re.IGNORECASE)

def _strip_insured_batch_noise(text: str) -> str: return _INSURED_BATCH_NOISE_RE.sub(" ", text)

# --- PATCH: karakter simbol/mojibake yang muncul dari encoding rusak dibuang total
_STRAY_SYMBOL_RE = re.compile(r"[¿¡‽]")
def _strip_stray_symbols(text: str) -> str: return _STRAY_SYMBOL_RE.sub(" ", text)

_PAREN_GROUP_RE = re.compile(r"\([^()]*\)")

# --- PATCH: isi tanda kurung sekarang di-protect PENUH (bukan cuma delimiter split)
# supaya PT/CV/TBK/PERSERO/dsb di dalam kurung TIDAK ikut ke-strip oleh _strip_legal_entity.
_PAREN_STASH: dict[str, str] = {}
_PAREN_TOKEN_RE = re.compile(r"\uE100(\d+)\uE101")

def _protect_parens_content(text: str) -> str:
    def _mask(match: "re.Match") -> str:
        idx = len(_PAREN_STASH)
        token = f"\uE100{idx}\uE101"
        _PAREN_STASH[token] = match.group(0)
        return token
    return _PAREN_GROUP_RE.sub(_mask, text)

def _restore_protected_chars(text: str) -> str:
    return text.replace("\uE0F0", "/").replace("\uE0F1", ",").replace("\uE0F2", ";")

def _restore_protected_parens(text: str) -> str:
    def _unmask(match: "re.Match") -> str:
        token = f"\uE100{match.group(1)}\uE101"
        return _PAREN_STASH.get(token, match.group(0))
    return _PAREN_TOKEN_RE.sub(_unmask, text)

_INSURED_MERGE_OVERRIDES = {
    "APARTEMEN EKSEKUTIF MENTENG,PERHIM.PENGH": ["APARTEMEN EKSEKUTIF MENTENG PERHIM PENGH"],
}
_LEGAL_ENTITY_DASH_SIGNAL_RE = re.compile(r",\s*(?:PT\.?|P\.T\.?|CV\.?|C\.V\.?|TBK\.?|Tbk\.?|PERSERO)\s*[\r\n]*\s*-\s*", flags=re.IGNORECASE)
_GENERIC_DASH_SEPARATOR_RE = re.compile(r"\s*[\r\n]+\s*-\s*|\s+-\s+")

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

# --- PATCH: kurung TIDAK lagi ikut ke-strip di final polish. Isi kurung
# (PRINCIPAL, PERSERO, WTC SERPONG, dst) dipertahankan apa adanya sbg bagian nama.
def _final_polish(text: str) -> str:
    t = re.sub(r"[,;:\"'.]", " ", text)
    t = re.sub(r"[-/]", " ", t)
    t = re.sub(r"\(\s+", "(", t)
    t = re.sub(r"\s+\)", ")", t)
    t = re.sub(r"\s+", " ", t).strip()
    dangling_pattern = (
        r"^\b(?:AS|AN|AND|OR|DAN|ATAU|N|QQ|CQ|C\.Q\.|ITS|ALL|THEIR|FOR|THE|OF|RESPECTIVE)\b\s*"
        r"|\s*\b(?:AS|AN|AND|OR|DAN|ATAU|N|QQ|CQ|C\.Q\.|ITS|ALL|THEIR|FOR|THE|OF|RESPECTIVE)\b$"
    )
    previous = None
    while previous != t:
        previous = t
        t = re.sub(dangling_pattern, "", t, flags=re.IGNORECASE).strip()
    return t.upper()

_LIPPO_KNOWN_ENTITIES = ["MATAHARI PUTRA PRIMA", "MATAHARI BOSTON DRIGSTORE", "MATAHARI PUSAKA TAMA", "MPP LIPPO GROUP", "LIPPO GROUP"]
_LIPPO_KNOWN_ENTITIES_SORTED = sorted(_LIPPO_KNOWN_ENTITIES, key=len, reverse=True)
_KNOWN_ENTITY_RE = re.compile("|".join(re.escape(name) for name in _LIPPO_KNOWN_ENTITIES_SORTED))

def _build_entity_word_pattern(name: str) -> str:
    joiner = r"(?:\s*/\s*|\s+)"
    return joiner.join(re.escape(w) for w in name.split())

_ENTITY_SLASH_NORMALIZE_PATTERNS = [
    (name, re.compile(r"\b" + _build_entity_word_pattern(name) + r"\b", flags=re.IGNORECASE))
    for name in _LIPPO_KNOWN_ENTITIES_SORTED if len(name.split()) > 1
]

def _normalize_entity_slash_variants(text: str) -> str:
    for canonical, pattern in _ENTITY_SLASH_NORMALIZE_PATTERNS: text = pattern.sub(canonical, text)
    return text

def _split_by_known_entities(text: str) -> list | None:
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
    if original in _INSURED_MERGE_OVERRIDES: return _INSURED_MERGE_OVERRIDES[original]
    if _INSURED_DESCRIPTIVE_RE.search(original): return [original]

    _PAREN_STASH.clear()

    t = _strip_transaction_header(original)
    t = _strip_titles(t)
    t = _strip_stray_symbols(t)
    t = _strip_insured_batch_noise(t)
    t = _INSURED_VARIOUS_RE.sub(" ", t)
    t = _normalize_dash_entity_list(t)
    t = _normalize_entity_slash_variants(t)
    t = _truncate_boilerplate_tail(t)
    t = _protect_parens_content(t)

    raw_parts = _ENTITY_SPLIT_RE.split(t)
    results = []
    
    for part in raw_parts:
        part = _restore_protected_chars(part)
        cleaned = _truncate_boilerplate_tail(part)
        cleaned = _strip_legal_entity(cleaned)
        cleaned = _final_polish(cleaned)
        cleaned = _restore_protected_parens(cleaned)
        cleaned = re.sub(r"\s+", " ", cleaned).strip()

        if not cleaned or len(cleaned) < 2 or not re.search(r"[A-Z]", cleaned): continue

        known_split = _split_by_known_entities(cleaned)
        if known_split is not None:
            for name in known_split:
                if name not in results: results.append(name)
            continue

        if cleaned not in results: results.append(cleaned)

    if len(results) > MAX_BREAKDOWN_CODES: return [original]
    return results

# ==============================================================================
# 2. CLEANSING POLIS & SLIP NO  (KHUSUS POLA DATA 1 - FACULTATIVE) -- UNCHANGED
# ==============================================================================
_MAIN_LEN_MIN, _MAIN_LEN_MAX = 11, 16
_MAIN_DIGITS_PATTERN = rf"\d{{{_MAIN_LEN_MIN},{_MAIN_LEN_MAX}}}"
_TOK_MAIN_RE = re.compile(rf"^{_MAIN_DIGITS_PATTERN}$")
_TOK_INDEPENDENT_RE = re.compile(r"^\d{7,16}$")
_TOK_CERT_RE = re.compile(r"^\d{6}$")
_TOK_SUFFIX_REPLACE_RE = re.compile(r"^\d{1,11}$")
_TOK_PLACEHOLDER_RE = re.compile(r"^P\d{1,2}$", flags=re.IGNORECASE)
_TOK_STATUS_RE = re.compile(r"^(?:New|End(?:orsement)?|Cancel\s*All|Adjustment|Cancellation|Reinstatement)$", flags=re.IGNORECASE)
_VARIOUS_WORD_RE = re.compile(r"\bVAR(?:I(?:O(?:U(?:S)?)?)?)?\b", flags=re.IGNORECASE)
_SD_MARKER_RE = re.compile(r"\bS\s*/\s*D\b|\bSD\b", flags=re.IGNORECASE)
_PURE_DECIMAL_RE = re.compile(r"^\d+(?:\.\d+)+$")
_DASH_AMP_AMBIGUOUS_RE = re.compile(rf"{_MAIN_DIGITS_PATTERN}\s*-\s*\d+\s*&\s*\d+")

_ALPHA_WORD_RE = re.compile(r"[A-Za-z][A-Za-z0-9]*")
_ALLOWED_ALPHA_TOKENS = {"NEW", "END", "ENDORSEMENT", "CANCEL", "ALL", "ADJUSTMENT", "CANCELLATION", "REINSTATEMENT", "VARIOUS", "VAR", "SD", "S", "D"}

_ATTACHMENT_NOISE_RE = re.compile(r"\bSEE\s+ATTACHMENT\b|\bATTACHMENT\b", flags=re.IGNORECASE)
_MONTH_YEAR_NOISE_RE = re.compile(
    r"\b(?:JAN(?:UARI|UARY)?|FEB(?:RUARI|RUARY)?|MAR(?:ET|CH)?|APR(?:IL)?|MEI|MAY|JUN(?:I|E)?|"
    r"JUL(?:I|Y)?|AGU(?:STUS)?|AUG(?:UST)?|SEP(?:TEMBER)?|OKT(?:OBER)?|OCT(?:OBER)?|"
    r"NOV(?:EMBER)?|DES(?:EMBER)?|DEC(?:EMBER)?)\s*\.?\s*\d{2,4}\b", flags=re.IGNORECASE)
_CURRENCY_NOISE_RE = re.compile(r"\b(?:USD|IDR|SGD|EUR)\b", flags=re.IGNORECASE)

_ADMIN_LABEL_NOISE_RE = re.compile(
    r"\bDEKL\.?\b"
    r"|\b\d+\s*SLIP\s+DIJADIKAN\s+\d+\b"
    r"|\bID\s*NO\.?\s*[A-Z0-9]*\b"
    r"|(?:\b\d{1,6}\s*/?\s*)?\bSHIPMENT\s*\d*\b"
    r"|(?:\b\d{1,6}\s*/?\s*)?\bBORDERO\b"
    r"|(?:\b\d{1,6}\s*/?\s*)?\bPENDING\s+SLIP\b"
    r"|(?:\b\d{1,6}\s*/?\s*)?\bBATCH\s*\d*\b", flags=re.IGNORECASE)

_PAREN_GROUP_NOISE_RE = re.compile(r"\([^()]*\)")
_DOC_REF_DATE_NOISE_RE = re.compile(r"(?:\b\d{1,6}\s*/\s*)?\b(?:CN|DN)\s*/\s*\d{1,2}\s*/\s*\d{1,2}\s*/\s*\d{2,4}(?:\s*/\s*\d{1,6}\b)?", flags=re.IGNORECASE)
_ENDORSE_HASH_NOISE_RE = re.compile(r"\bE\s*#\s*\d+\b", flags=re.IGNORECASE)
_ROMAN_NUMERAL_RE = r"(?:X{0,3}(?:IX|IV|V?I{0,3}))"
_TREATY_REF_CODE_NOISE_RE = re.compile(
    rf"\b\d{{1,6}}\s*/\s*[A-Z]+-[A-Z]+(?:\s*/\s*{_ROMAN_NUMERAL_RE}\b)?(?:\s*/\s*\d{{4}}\b)?"
    rf"|\b[A-Z]+-[A-Z]+(?:\s*/\s*{_ROMAN_NUMERAL_RE}\b)?(?:\s*/\s*\d{{4}}\b)?", flags=re.IGNORECASE)

def _strip_known_noise_phrases(text: str) -> str:
    text = _PAREN_GROUP_NOISE_RE.sub(" ", text)
    text = _ATTACHMENT_NOISE_RE.sub(" ", text)
    text = _DOC_REF_DATE_NOISE_RE.sub(" ", text)
    text = _TREATY_REF_CODE_NOISE_RE.sub(" ", text)
    text = _MONTH_YEAR_NOISE_RE.sub(" ", text)
    text = _CURRENCY_NOISE_RE.sub(" ", text)
    text = _ADMIN_LABEL_NOISE_RE.sub(" ", text)
    text = _ENDORSE_HASH_NOISE_RE.sub(" ", text)
    text = re.sub(r"[/\-]\s*(?=[/\-]|$)", " ", text)
    return re.sub(r"\s+", " ", text).strip()

_MAIN_SPACE_SUFFIX_RE = re.compile(rf"^({_MAIN_DIGITS_PATTERN})\s+(\d{{1,6}})\s*(?:/\s*VAR[A-Z]*\s*)?$", flags=re.IGNORECASE)

_POLIS_SLIP_OVERRIDES = {
    "APRIL 2026 - 0015/LI-RBU/V/2026 / 1112302600206,1112302600215,1112302600218,1112302600221.": [
        "1112302600206", "1112302600215", "1112302600218", "1112302600221",
    ],
}

def _has_descriptive_narrative(text: str) -> bool:
    for word in _ALPHA_WORD_RE.findall(text):
        if _TOK_PLACEHOLDER_RE.match(word) or word.upper() in _ALLOWED_ALPHA_TOKENS: continue
        return True
    return False

def _strip_various(text: str) -> str: return _VARIOUS_WORD_RE.sub(" ", text)

def _expand_numeric_range(start_str: str, end_str: str, max_count: int = MAX_BREAKDOWN_CODES) -> list | None:
    if len(start_str) != len(end_str) or not (start_str.isdigit() and end_str.isdigit()): return None
    diff_len = next((i for i in range(1, len(start_str) + 1) if start_str[-i] != end_str[-i]), 0)
    if diff_len == 0 or start_str[:-diff_len] != end_str[:-diff_len]: return None
    try:
        start_n, end_n = int(start_str[-diff_len:]), int(end_str[-diff_len:])
    except ValueError: return None
    if start_n > end_n or (end_n - start_n + 1) > max_count: return None
    return [f"{start_str[:-diff_len]}{str(n).zfill(diff_len)}" for n in range(start_n, end_n + 1)]

def _expand_suffix_range(anchor: str, suffix_a: str, suffix_b: str) -> list | None:
    width = max(len(suffix_a), len(suffix_b))
    if width >= len(anchor) or not (suffix_a.isdigit() and suffix_b.isdigit()): return None
    start_n, end_n = int(suffix_a), int(suffix_b)
    if start_n > end_n or (end_n - start_n + 1) > MAX_BREAKDOWN_CODES: return None
    return [f"{anchor[:-width]}{str(n).zfill(width)}" for n in range(start_n, end_n + 1)]

def _split_sd_range(text: str):
    markers = list(_SD_MARKER_RE.finditer(text))
    if not markers: return None
    if len(markers) > 1: return "__FALLBACK__"

    left, right = text[: markers[0].start()], text[markers[0].end() :]
    runs_left, runs_right = re.findall(r"\d+", left), re.findall(r"\d+", right)
    if not runs_left or not runs_right: return "__FALLBACK__"

    last_left, first_right = runs_left[-1], runs_right[0]
    last_left_is_main = _TOK_MAIN_RE.match(last_left) is not None

    if not last_left_is_main and len(runs_left) >= 2 and _TOK_MAIN_RE.match(runs_left[-2]):
        expanded = _expand_suffix_range(runs_left[-2], last_left, first_right)
    elif last_left_is_main and _TOK_MAIN_RE.match(first_right):
        expanded = _expand_numeric_range(last_left, first_right)
    else:
        return "__FALLBACK__"
    return expanded if expanded is not None else "__FALLBACK__"

def clean_split_code(text: str) -> list:
    if not isinstance(text, str) or not text.strip() or text.strip().lower() in ("nan", "none", "-"): return []
    
    original = text.strip()
    space_concat_match = _MAIN_SPACE_SUFFIX_RE.match(original)
    if space_concat_match: return [space_concat_match.group(1) + space_concat_match.group(2)]
    if original in _POLIS_SLIP_OVERRIDES: return _POLIS_SLIP_OVERRIDES[original]
    
    t = _strip_known_noise_phrases(original)
    if _has_descriptive_narrative(t) or _PURE_DECIMAL_RE.match(t) or _DASH_AMP_AMBIGUOUS_RE.search(t): return [original]
    
    sd_result = _split_sd_range(t)
    if sd_result == "__FALLBACK__": return [original]
    if sd_result is not None: return sd_result
    
    t = re.sub(r"[\u2012\u2013\u2014\u2015]", "-", _strip_various(t))
    _had_slash_context = "/" in t
    _delim_split_re = re.compile(r"(\s{2,}|[,+/;&])")
    
    raw_split = [chunk for chunk in _delim_split_re.split(t) if chunk != ""]
    strong_segments, pending_delim = [], None
    for chunk in raw_split:
        if re.fullmatch(r"\s{2,}", chunk):
            pending_delim = ","
            continue
        if chunk in (",", "+", "/", ";", "&"):
            pending_delim = chunk
            continue
        seg = chunk.strip()
        if seg: strong_segments.append((seg, pending_delim))
        pending_delim = None
        
    raw_tokens = []
    for seg, delim in strong_segments:
        seg = re.sub(r"[.]", "", seg).strip()
        sub_parts = [p.strip() for p in seg.split("-") if p.strip()]
        if len(sub_parts) <= 1:
            if sub_parts: raw_tokens.append((sub_parts[0], delim))
            continue
        buffer, buffer_delim = [sub_parts[0]], delim
        for p in sub_parts[1:]:
            if _TOK_MAIN_RE.match(p) or _TOK_PLACEHOLDER_RE.match(p):
                raw_tokens.append(("-".join(buffer), buffer_delim))
                buffer, buffer_delim = [p], "-"
            else:
                buffer.append(p)
        raw_tokens.append(("-".join(buffer), buffer_delim))
        
    tokens = [(tok, d) for tok, d in raw_tokens if not _TOK_PLACEHOLDER_RE.match(tok)]
    if not tokens: return [original]
    
    n_tokens = len(tokens)
    codes, current_base, current_base_is_strict = [], None, False
    
    for i, (tok, delim) in enumerate(tokens):
        effective_delim = delim if delim is not None else (tokens[i + 1][1] if i + 1 < n_tokens else ("/" if _had_slash_context else None))
        
        if (current_base is not None and current_base_is_strict and delim in ("+", "&", "/", ",")
            and _TOK_SUFFIX_REPLACE_RE.match(tok) and not _TOK_MAIN_RE.match(tok) and len(tok) < len(current_base)):
            reconstructed = current_base[: -len(tok)] + tok
            if reconstructed not in codes: codes.append(reconstructed)
            continue
            
        parts = tok.split("-")
        anchor, suffix_parts = parts[0], parts[1:]
        
        if _TOK_MAIN_RE.match(anchor): is_strict = True
        elif not suffix_parts and effective_delim == "/" and _TOK_INDEPENDENT_RE.match(anchor): is_strict = False
        else:
            if anchor.isdigit(): return [original]
            continue
            
        current_base, current_base_is_strict = anchor, is_strict
        
        while suffix_parts and _TOK_STATUS_RE.match(suffix_parts[-1]): suffix_parts.pop()
        
        append_suffix_parts = []
        for p in suffix_parts:
            if _TOK_CERT_RE.match(p):
                append_suffix_parts.append(p)
            elif _TOK_SUFFIX_REPLACE_RE.match(p) and len(p) < len(anchor):
                reconstructed = anchor[: -len(p)] + p
                if reconstructed not in codes: codes.append(reconstructed)
                
        if append_suffix_parts:
            combined = anchor + "-" + "-".join(append_suffix_parts)
            if combined not in codes: codes.append(combined)
            
        if anchor not in codes: codes.append(anchor)
        
    if not codes or len(codes) > MAX_BREAKDOWN_CODES: return [original]
    return codes

# ==============================================================================
# 3. LOAD & FILTER DATA
# ==============================================================================
def load_source(file_path: str, sheet_name: str, header_row: int) -> pd.DataFrame:
    return pd.read_excel(file_path, sheet_name=sheet_name, header=header_row)

def filter_cedant(df_raw: pd.DataFrame, keyword: str, cedant_column: str) -> pd.DataFrame:
    if cedant_column not in df_raw.columns:
        raise ValueError(f"CEDANT_COLUMN='{cedant_column}' tidak ditemukan. Tersedia: {list(df_raw.columns)}")
    mask = df_raw[cedant_column].astype(str).str.contains(keyword, case=False, na=False)
    df_filtered = df_raw[mask].copy()
    print(f"[INFO] Filter cedant (kolom '{cedant_column}') mengandung '{keyword}': {len(df_filtered)} baris ditemukan dari total {len(df_raw)} baris.")
    return df_filtered

# ==============================================================================
# 4. BUILD OUTPUT
# ==============================================================================
def _breakdown_all_rows(df: pd.DataFrame, col_insured: str, col_polis: str, col_slip: str):
    insured_cln_all, polis_cln_all, slip_cln_all = [], [], []
    max_ins, max_pol, max_slp = 0, 0, 0
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
    col_insured, col_polis, col_slip = INSURED_COLUMN, POLIS_COLUMN, SLIP_COLUMN
    for col, label in ((col_insured, "INSURED_COLUMN"), (col_polis, "POLIS_COLUMN"), (col_slip, "SLIP_COLUMN")):
        if col not in df_filtered.columns:
            raise ValueError(f"Kolom {label}='{col}' tidak ditemukan. Tersedia: {list(df_filtered.columns)}")
            
    insured_cln_all, polis_cln_all, slip_cln_all, max_ins, max_pol, max_slp = _breakdown_all_rows(df_filtered, col_insured, col_polis, col_slip)
    original_cols = list(df_filtered.columns)
    df_filtered = df_filtered.reset_index(drop=True)
    
    # Dict di Python mempertahankan urutan insert, jadi kolom CLN ditaruh tepat setelah kolom original
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
    text_col_indices = [idx for idx, col in enumerate(df.columns, start=1) if any(keyword in str(col).upper() for keyword in TEXT_FORMAT_COLUMN_KEYWORDS)]
    
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
    print("=" * 70 + "\n CLEANSING DATA 1 - FACULTATIVE | CEDANT: PT LIPPO GENERAL INSURANCE\n" + "=" * 70)
    input_path = INPUT_FILE
    
    if not os.path.exists(input_path):
        candidates = [f for f in os.listdir(".") if f.endswith(".xlsx") and not f.startswith("Hasil_") and not f.startswith("Clean_")]
        if candidates:
            input_path = candidates[0]
            print(f"[AUTO-DETECT] Memakai file terdeteksi: '{input_path}'")
        else:
            raise FileNotFoundError("File Excel sumber tidak ditemukan di folder kerja.")
            
    df_raw = load_source(input_path, SHEET_NAME, HEADER_ROW)
    print(f"[INFO] Total baris source: {len(df_raw)}")
    
    df_lippo = filter_cedant(df_raw, CEDANT_FILTER, CEDANT_COLUMN)
    if df_lippo.empty:
        print("[WARNING] Tidak ada baris yang cocok dengan filter cedant. Proses dihentikan.")
        return
        
    df_hasil = build_output(df_lippo)
    save_with_text_format(df_hasil, OUTPUT_FILE)
    
    print("=" * 70 + f"\n[SUCCESS] Selesai. Total baris output: {len(df_hasil)}\n[SUCCESS] File hasil: '{OUTPUT_FILE}'\n" + "=" * 70)

if __name__ == "__main__":
    main()