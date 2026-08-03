import os
import re
import pandas as pd
from typing import List, Optional, Union

INPUT_FILE = os.path.join("dataExcel", "raw", "1a. Transaksi Facul 01.01.23 - 17.07.26.xlsx")
SHEET_NAME = "Sheet0"
OUTPUT_FILE = os.path.join("dataExcel", "processed", "astrabuana_output_facul.xlsx")

CEDANT_COL = "COMP_NAME"
CEDANT_VALUE = "PT ASURANSI ASTRA BUANA"

MINIMAL_DIGIT_NOMOR_POLIS = 11
MINIMAL_PANJANG_KODE = 6
MINIMAL_PANJANG_KODE_RANGE = 8
MAX_HASIL_BREAKDOWN = 5
VALID_FULL_LENGTHS_POLIS = (11, 12, 13, 16)
SLIP_SUFFIX_ALLOWED_LENGTHS = {2, 3, 4, 6, 7}

class Patterns:
    AS_PER_LIST = re.compile(r"\bA\s*S\s*P?\s*E?\s*R?\s*L\s*I\s*S\s*T\b", re.IGNORECASE)
    CEDANT_REMOVE = re.compile(r"\b(?:PT\s+)?ASURANSI\s+ASTRA\s+BUANA\b|\bASTRA\s+BUANA\b", re.IGNORECASE)
    PLACEHOLDER_TOKENS = re.compile(r"\+?\s*\bP\d{1,3}\b\s*\+?", re.IGNORECASE)
    SD = re.compile(r"\bS\s*/\s*D\b", re.IGNORECASE)
    STRIP_LABEL_CLEAN = re.compile(r"^[\s/\-:,.+&;]+|[\s/\-:,.+&;]+$")
    TBA_VAR = re.compile(r"\b(TBA|VARIOUS|VAR)\b", re.IGNORECASE)
    NON_ALPHANUM = re.compile(r"[^A-Za-z0-9]")
    MULTIPLE_SPACES = re.compile(r"\s{2,}")
    
    # Policy Specific
    POLIS_VALID_FORMAT = re.compile(r"\d{11,}")
    POLIS_SUFFIX_FORMAT = re.compile(r"\d{7,10}[+,\-]\d{2,4}")
    RANGE_PATTERN = re.compile(r"[A-Za-z]{2,8}\d{8,}")
    MASTER_POLICY_ANNOTATION = re.compile(r"\(\s*MASTER\b", re.IGNORECASE)
    DOT_GROUPED_CODE = re.compile(r"^[A-Za-z0-9]+(?:\.[A-Za-z0-9]+)+$")
    TAGGED_POLIS = re.compile(r"^\d{8,}(?:/[A-Za-z]{1,5})?$")
    INFO_NOTE_PAREN = re.compile(r"\(\s*LANJUTAN\b[^\)]*\)", re.IGNORECASE)
    
    # Dates & Slip Patterns
    _BULAN_ALT = (
        r"(?:JAN(?:UARY|UARI)?|FEB(?:RUARY|RUARI)?|MAR(?:CH|ET)?|APR(?:IL)?|MAY|MEI|"
        r"JUN(?:E|I)?|JUL(?:Y|I)?|AUG(?:UST)?|AGU(?:STUS)?|SEP(?:TEMBER)?|OCT(?:OBER)?|"
        r"OKT(?:OBER)?|NOV(?:EMBER)?|DEC(?:EMBER)?|DES(?:EMBER)?)"
    )
    MONTH_YEAR_PREFIX = re.compile(rf"^(?:{_BULAN_ALT}\s+)?{_BULAN_ALT}\s+\d{{4}}\s*[-/]?\s*", re.IGNORECASE)
    MONTH_YEAR_SUFFIX = re.compile(
        r"\s*[-/]\s*(?:JAN(?:UARY|UARI)?|FEB(?:RUARY|RUARI)?|MAR(?:CH|ET)?|APR(?:IL)?|MAY|MEI|JUN(?:E|I)?|JUL(?:Y|I)?|AUG(?:UST)?|AGU(?:STUS)?|SEP(?:TEMBERI|TEMBER)?|OCT(?:OBER)?|OKT(?:OBER)?|NOV(?:EMBER)?|DEC(?:EMBER)?|DES(?:EMBER)?)\s+\d{4}\s*$",
        re.IGNORECASE
    )
    TRAILING_NOTE = re.compile(r"\s*/\s*(?:P\d+|\bSLIP\b|\bBLM\b|\bBELUM\b|\bDATANG\b).*$", re.IGNORECASE)
    
    _SLIP_MATA_UANG = r"(?:IDR|USD|JPY|SGD|EUR|GBP)"
    _SLIP_KATA_BORDER = r"(?:BORDEREAUX|BORDEROUX|BORDERO|BORDX|BORDR|BORD|BODR|BD)"
    SLIP_CATATAN_EKOR = re.compile(
        rf"\s*/?\s*(?:{_SLIP_MATA_UANG}\s*/\s*)?(?:{_SLIP_KATA_BORDER}\s+)?(?:{_BULAN_ALT}\s+)?{_BULAN_ALT}(?:\.\d{{1,2}}|\s+\d{{4}})?\s*-?\s*(?:{_SLIP_MATA_UANG})?\s*$",
        re.IGNORECASE
    )
    SLIP_MATA_UANG_EKOR = re.compile(rf"\s*/?\s*-?\s*{_SLIP_MATA_UANG}\s*$", re.IGNORECASE)
    
    _KATA_LABEL_UNTUK_SLASH = (
        r"(?:TBA|VAR|VARIOUS|SUMMARY|ASTRA|IDR|USD|JPY|SGD|EUR|GBP"
        r"|JAN|FEB|MAR|APR|MEI|JUN|JUL|AGU|AGS|SEP|OKT|NOV|DES|AUG|OCT|DEC)"
    )
    SLIP_KEEP_AS_IS = re.compile(
        r"(?:^\d{11}\s+\d{2,3}(?:\s+\d{2,3})+$|\bFISHVSL\b|\b\d{11}\s*\+\s*\d{3}\s+\d{4}\b|^\d{25,}$|^\d{2,3}(?:&\d{2,3})+$)",
        re.IGNORECASE
    )
    
    END_TOKEN = re.compile(r"^END\.?\d*$", re.IGNORECASE)
    END_TRAILING_MARK = re.compile(r"[\s/.]+END\b\.?\d*$", re.IGNORECASE)

    # Insured Patterns
    INSURED_ENTITIES = re.compile(r"[,.\s]*\b(PT|CV|TBK|PTE|LTD|PELAYARAN)\b(?!\w)[,.\s]*", re.IGNORECASE)
    INSURED_PERSERO = re.compile(r"[,.\s]*\(\s*PERSERO\s*\)[,.\s]*|\bPERSERO\b", re.IGNORECASE)
    STRIP_INSURED_BOUNDARIES = re.compile(r"^[\s/\\\-:,.+&]+|[\s/\\\-:,.+&]+$")
    BORDERO_TAIL = re.compile(r"\b(?:BORDEROUX|BORDERO|BORD)\b.*$", re.IGNORECASE)
    SLIP_TAIL_VARIOUS = re.compile(r'\s*/\s*VARIOUS\b.*$', re.IGNORECASE)
    SLIP_TAIL_END_DOT = re.compile(r'\.END(?:\.\d+)?(?:\s*\+\s*\d+)*\s*$', re.IGNORECASE)
    SLIP_TAIL_END_PLUS = re.compile(r'(?:\s*\+\s*END\.\d+)+\s*$', re.IGNORECASE)
    SLIP_FIRST_SEGMENT_CODE = re.compile(r'^[A-Za-z]{2,8}\d[A-Za-z0-9\-]{2,}$')
    
    SLIP_KOMPLEKS_CONNECTORS = [
        re.compile(r'\s*;\s*'), re.compile(r'\s*/\s*'),
        re.compile(r'\s*&\s*'), re.compile(r'\s*\+\s*'),
    ]
    SLIP_FIRST_TOKEN_DASH = re.compile(r'^([A-Za-z]{0,8})-?(\d{7,})$')
    STANDALONE_BULAN_TAHUN = re.compile(rf"^{_BULAN_ALT}(?:\s+\d{{4}})?$", re.IGNORECASE)
    SLIP_JUNK_TOKEN = re.compile(r'^(?:END(?:\.\d+)?|VAR|TBA|VARIOUS|AS ATTACHMENT|P\d{1,3})\b', re.IGNORECASE)
    SLIP_UPLOADED_FILES = re.compile(r'\bSEE\s+UPLOADED\s+FILES\b', re.IGNORECASE)
    SLIP_FIRST_TOKEN = re.compile(r'^([A-Za-z]{2,8})(\d{2,})$')
    
    SLIP_PENGULANGAN_CONNECTORS = [
        re.compile(r'\s*/\s*'), re.compile(r'\s*\+\s*'),
        re.compile(r'\s*-\s*'), re.compile(r'\s+'),
    ]
    VARIOUS_PHRASES_PATTERN = r"VARIOUS(?:\s+INCLUDE\s+FISHING\s+VESSEL|\s+FISHING\s+VESSEL|\s*\(\s*LINE\s*SLIP\s*\))?"
    VARIOUS_EXACT = re.compile(rf"^\s*{VARIOUS_PHRASES_PATTERN}\s*$", re.IGNORECASE)
    TARGET_PARENS = re.compile(
        r'\(\s*(?:PERSERO|ALL\s+THE\s+OWNERS\s+OF\s+UNIT\s+APARTEMENT\s+DHARMAWANGSA\s+1|A|B|APARTEMEN\s+GRAND\s+DHIKA\s+BEKASI|KNOWN\s+AS\s+BANK\s+JASA\s+JAKARTA\s*,\s*ASTRA\s+GROUP)\s*\)',
        re.IGNORECASE
    )
    MASK_FAC = re.compile(r'\b(AUTO\s+FACILITY|CARGO\s+FACILITY|LINESLIP)\s*/', re.IGNORECASE)
    
    BORDERO_KEYWORD_PRESENT = re.compile(_SLIP_KATA_BORDER, re.IGNORECASE)
    BORDERO_VOCAB = re.compile(
        rf"\b(?:ASURANSI\s+ASTRA\s+BUANA|ASTRA\s+BUANA|ASTRA|{_SLIP_KATA_BORDER}|FACILITY|SLIP|DI|{_BULAN_ALT}|{_SLIP_MATA_UANG})\b",
        re.IGNORECASE
    )
    BORDERO_DATE_RANGE = re.compile(r"\b\d{6}\s*-\s*\d{6}\b")
    TOKEN_SAMPAH_SPASI = re.compile(r"^(?:VAR|TBA|VARIOUS|END(?:\.\d+)?|P\d{1,3})$", re.IGNORECASE)

def _normalisasi_as_per_list(text: str) -> str:
    return Patterns.AS_PER_LIST.sub("AS PER LIST", text)

def _clean_code(text: str) -> str:
    return Patterns.NON_ALPHANUM.sub("", str(text)).upper()

def _hapus_kata_label(text: str) -> str:
    text = str(text).upper()
    text = Patterns.CEDANT_REMOVE.sub("", text)
    return Patterns.PLACEHOLDER_TOKENS.sub("", text)

def _hapus_single_char_suffix(text: str) -> str:
    pattern = re.compile(r"(?<=[A-Za-z0-9]{3})\s*[/\-.]\s*[A-Za-z0-9]\b(?!\w)")
    prev = ""
    curr = text.strip()
    while curr != prev:
        prev = curr
        curr = pattern.sub("", curr)
    return curr

def _handle_plus_minus_symmetry(text: str) -> str:
    def _replace_match(match):
        left, op, right = match.group(1), match.group(2), match.group(3)
        if len(left) == len(right):
            if left.isdigit() and right.isdigit():
                return f"{left}{op}{right}"
            protected_op = "XXXXPLUSXXXX" if op == "+" else "XXXXMINUSXXXX"
            return f"{left}{protected_op}{right}"
        if re.search(r"[A-Za-z]", left) or re.search(r"[A-Za-z]", right):
            return f"{left}{right}"
        return f"{left}{op}{right}"

    pattern = re.compile(r"([A-Za-z0-9]+)\s*([\+\-])\s*([A-Za-z0-9]+)")
    prev_text = ""
    curr_text = text
    while curr_text != prev_text:
        prev_text = curr_text
        curr_text = pattern.sub(_replace_match, curr_text)
    return curr_text

def _rapikan_teks_label(text: str) -> str:
    text = str(text).upper()
    text = _normalisasi_as_per_list(text)
    text = Patterns.INFO_NOTE_PAREN.sub("", text)

    protect_var = bool(re.match(r"^\s*(TBA|VARIOUS|VAR)\b", text))
    teks_tanpa_tba = text if protect_var else Patterns.TBA_VAR.sub("", text)
    
    teks_tanpa_tba = Patterns.PLACEHOLDER_TOKENS.sub("", teks_tanpa_tba)
    teks_tanpa_tba = Patterns.STRIP_LABEL_CLEAN.sub("", teks_tanpa_tba).strip()

    if not teks_tanpa_tba:
        text = Patterns.STRIP_LABEL_CLEAN.sub("", text)
        return Patterns.MULTIPLE_SPACES.sub(" ", text).strip()

    text_final = text if protect_var else Patterns.TBA_VAR.sub("", text)
    text_final = Patterns.PLACEHOLDER_TOKENS.sub("", text_final)
    text_final = Patterns.STRIP_LABEL_CLEAN.sub("", text_final)
    return Patterns.MULTIPLE_SPACES.sub(" ", text_final).strip()

def _gabung_digit_terpisah(text: str) -> str:
    cleaned = text.strip()
    if re.fullmatch(r"[\d\s]+", cleaned):
        only_digits = re.sub(r"\s+", "", cleaned)
        if len(only_digits) >= 7:
            return only_digits
    return text

def _pre_clean_polis_or_slip(text: str) -> str:
    text = Patterns.PLACEHOLDER_TOKENS.sub("", text)
    text = Patterns.INFO_NOTE_PAREN.sub("", text)
    text = Patterns.TRAILING_NOTE.sub("", text)
    text = Patterns.MONTH_YEAR_PREFIX.sub("", text)
    text = _hapus_single_char_suffix(text)
    text = _gabung_digit_terpisah(text)

    pattern_dup = re.compile(r"\b([A-Za-z0-9]+)\s*[-/]\s*\1\b", re.IGNORECASE)
    prev = ""
    while prev != text:
        prev = text
        text = pattern_dup.sub(r"\1", text)
        
    return _handle_plus_minus_symmetry(text).strip()

def _pisah_jadi_segmen(text: str) -> List[str]:
    penanda_sd = "§SD§"
    teks_terlindungi = Patterns.SD.sub(penanda_sd, text)
    teks_terlindungi = re.sub(
        rf"/(?!{Patterns._KATA_LABEL_UNTUK_SLASH}\b)([A-Za-z]{{2,5}})\b",
        r"§SLASH§\1",
        teks_terlindungi,
        flags=re.IGNORECASE,
    )

    segmen_mentah = re.split(r"\s{2,}|/", teks_terlindungi)
    segmen_hasil = []
    for segmen in segmen_mentah:
        segmen = segmen.replace(penanda_sd, "S/D").replace("§SLASH§", "/").strip()
        if segmen != "" and re.search(r"[A-Za-z0-9]", segmen):
            segmen_hasil.append(segmen)
    return segmen_hasil

def hapus_kurung(text: str) -> str:
    return Patterns.TARGET_PARENS.sub(' ', text)

def clean_insured(name: Union[str, float]) -> Union[str, float]:
    if pd.isna(name):
        return name

    text = str(name).strip()
    text = Patterns.BORDERO_TAIL.sub("", text)
    text = hapus_kurung(text)
    text = Patterns.INSURED_PERSERO.sub(" ", text)
    text = Patterns.INSURED_ENTITIES.sub(" ", text)
    text = re.sub(r"\s*,\s*", " ", text)
    text = re.sub(r"\s*/\s*(?=[A-Za-z\s]+DIVISION)", " - ", text, flags=re.IGNORECASE)
    text = Patterns.STRIP_INSURED_BOUNDARIES.sub("", text)
    return Patterns.MULTIPLE_SPACES.sub(" ", text).strip()

def breakdown_insured(value: Union[str, float]) -> List:
    if pd.isna(value):
        return [value]

    original_text = str(value).strip()
    if not original_text or Patterns.VARIOUS_EXACT.fullmatch(original_text):
        return [original_text]

    text = Patterns.BORDERO_TAIL.sub("", original_text).strip()

    if re.search(r"POLYCHEM", text, re.IGNORECASE):
        cln = clean_insured(text)
        return [cln] if cln else [original_text]

    text_no_parens = hapus_kurung(text)
    text_masked = Patterns.MASK_FAC.sub(r'\1_SLASH_', text_no_parens)

    SPLIT_RE = re.compile(r'\s*(?:/|,|\bQQ\b|\bOR\b)\s*|\s+-\s+', re.IGNORECASE)
    raw_parts = SPLIT_RE.split(text_masked)
    cleaned_parts = []
    
    for p in raw_parts:
        p = p.replace('_SLASH_', '/')
        if Patterns.VARIOUS_EXACT.fullmatch(p.strip()):
            continue
        cln = clean_insured(p)
        if cln:
            cleaned_parts.append(cln)

    if len(cleaned_parts) > MAX_HASIL_BREAKDOWN:
        return [original_text]

    if not cleaned_parts:
        cln_all = clean_insured(original_text)
        return [cln_all] if cln_all else [original_text]

    return cleaned_parts

def _hapus_end_trailing(text: str) -> str:
    if not re.search(r"\bEND\b", text, re.IGNORECASE):
        return text

    stripped_full = text.strip()
    if Patterns.END_TOKEN.match(stripped_full):
        return text

    segmen_plus = re.split(r"(\+)", text)
    token_token = segmen_plus[0::2]

    idx_mulai_end = None
    for i, tok in enumerate(token_token):
        tok_bersih = tok.strip()
        if not tok_bersih:
            continue
        if Patterns.END_TOKEN.match(tok_bersih):
            idx_mulai_end = i
            break
        m = Patterns.END_TRAILING_MARK.search(tok_bersih)
        if m:
            token_token[i] = tok_bersih[: m.start()]
            idx_mulai_end = i + 1
            break

    if idx_mulai_end is None:
        return text

    token_token = token_token[:idx_mulai_end]
    hasil = "+".join(t for t in token_token if t.strip() != "")
    return hasil if hasil.strip() != "" else text

def _strip_bordero_vocab(text: str) -> str:
    probe = Patterns.BORDERO_DATE_RANGE.sub(" ", text)
    probe = Patterns.BORDERO_VOCAB.sub(" ", probe)
    probe = re.sub(r"\b\d{1,2}\b", " ", probe)
    probe = re.sub(r"\b\d{4}\b", " ", probe)
    return re.sub(r"[\s\-/.,]+", "", probe)

def _is_segmen_bordero_note(segmen: str) -> bool:
    return segmen.strip() == "" or _strip_bordero_vocab(segmen) == ""

def _handle_bordero_note(original_text: str):
    if not Patterns.BORDERO_KEYWORD_PRESENT.search(original_text):
        return None
    if _strip_bordero_vocab(original_text) == "":
        return [original_text]

    segmen_segmen = [s.strip() for s in re.split(r"[/\-]", original_text) if s.strip() != ""]
    if len(segmen_segmen) < 2:
        return None

    segmen_kode = [s for s in segmen_segmen if not _is_segmen_bordero_note(s)]
    segmen_note = [s for s in segmen_segmen if _is_segmen_bordero_note(s)]

    if not segmen_note or len(segmen_kode) != 1:
        return None

    return [segmen_kode[0].strip().upper()]

def _hapus_catatan_ekor_slip(text: str) -> str:
    hasil = text
    while True:
        dihapus = Patterns.SLIP_CATATAN_EKOR.sub("", hasil).strip()
        if dihapus == hasil:
            dihapus = Patterns.SLIP_MATA_UANG_EKOR.sub("", hasil).strip()
        if dihapus == "" or dihapus == hasil:
            break
        hasil = dihapus
    return hasil

def _breakdown_slip_garis_miring_koma(text: str):
    if "/" not in text or "," not in text:
        return None
    if _mengandung_kata_sd(text) or _sepertinya_pola_range(text):
        return None

    top_parts = [p.strip() for p in text.split("/") if p.strip()]
    if len(top_parts) < 2:
        return None

    hasil = []
    for part in top_parts:
        if Patterns.STANDALONE_BULAN_TAHUN.match(part):
            continue
        if "," in part:
            sub_parts = [s.strip() for s in part.split(",") if s.strip()]
            if len(sub_parts) >= 2 and all(Patterns.SLIP_FIRST_TOKEN_DASH.match(sp) for sp in sub_parts):
                hasil.extend(sp.upper() for sp in sub_parts)
                continue
        hasil.append(part.strip().upper())

    if not hasil or len(hasil) > MAX_HASIL_BREAKDOWN:
        return [text.strip()]

    return hasil

def _breakdown_slip_khusus(raw: str) -> List[str]:
    if Patterns.SLIP_KEEP_AS_IS.search(raw):
        return [_rapikan_teks_label(raw)]

    m_dot = re.match(r"^(\d{11})\.(\d{2,4})$", raw)
    if m_dot:
        base, suffix = m_dot.groups()
        return [base, base[:-len(suffix)] + suffix]

    m_plus = re.match(
        r"^([A-Z]{3,5})(\d{2}[A-Z0-9]{4,5})\s*([+&])\s*(\d{2}[A-Z0-9]{4,5})(?:\s*([-\s]\s*\d{4}))?$",
        raw,
        re.IGNORECASE,
    )
    if m_plus:
        prefix, code1, sep, code2, suffix = m_plus.groups()
        slip1 = f"{prefix}{code1}".upper()
        slip2 = f"{prefix}{code2}".upper()
        if suffix:
            suffix_clean = re.sub(r"\s+", "", suffix).upper()
            slip2 = f"{slip2}{suffix_clean}"
        return [slip1, slip2]

    if ";" in raw:
        return [p.strip().upper() for p in raw.split(";") if p.strip()]

    if "&" in raw:
        parts_mentah = [p.strip() for p in raw.split("&") if p.strip()]
        if parts_mentah:
            digit_pertama = re.sub(r"\D", "", parts_mentah[0])
            if len(digit_pertama) < MINIMAL_DIGIT_NOMOR_POLIS:
                return [raw.strip()]
        return [p.upper() for p in parts_mentah]

    return []

def _breakdown_slip_pengulangan_digit(original_text: str):
    text = original_text.strip()
    if not text or _mengandung_kata_sd(text):
        return None

    if Patterns.SLIP_UPLOADED_FILES.search(text):
        return [original_text]

    if Patterns.SLIP_KEEP_AS_IS.search(text) and not re.match(r'^\d{11}\s+\d{2,3}(?:\s+\d{2,3})+$', text):
        return None

    text = Patterns.MONTH_YEAR_PREFIX.sub("", text).strip().upper()
    text = re.sub(r"\.+$", "", text).strip()

    for connector_re in Patterns.SLIP_PENGULANGAN_CONNECTORS:
        if not connector_re.search(text):
            continue

        raw_parts = [p.strip() for p in connector_re.split(text) if p.strip()]
        if len(raw_parts) < 2:
            continue

        m_first = Patterns.SLIP_FIRST_TOKEN.match(raw_parts[0])
        if not m_first:
            continue

        prefix, first_digits = m_first.groups()
        hasil = [prefix + first_digits]
        gagal = False

        for tok in raw_parts[1:]:
            if Patterns.SLIP_JUNK_TOKEN.match(tok):
                break
            if tok[0].isalpha() and Patterns.SLIP_FIRST_TOKEN.match(tok):
                hasil.append(tok)
                continue
            if not tok.isdigit():
                gagal = True
                break

            if len(tok) == len(first_digits):
                new_digits = tok
            elif len(tok) in SLIP_SUFFIX_ALLOWED_LENGTHS and len(tok) < len(first_digits):
                new_digits = first_digits[: -len(tok)] + tok
            else:
                gagal = True
                break

            hasil.append(prefix + new_digits)

        if gagal or len(hasil) > MAX_HASIL_BREAKDOWN:
            return [original_text]

        return hasil

    return None

def _breakdown_slip_kode_kompleks(original_text: str):
    if Patterns.SLIP_KEEP_AS_IS.search(original_text):
        return None

    text = original_text.strip()
    if not text or _mengandung_kata_sd(text):
        return None

    text = Patterns.SLIP_TAIL_VARIOUS.sub("", text).strip()
    text = Patterns.SLIP_TAIL_END_DOT.sub("", text).strip()
    text = Patterns.SLIP_TAIL_END_PLUS.sub("", text).strip()

    for connector_re in Patterns.SLIP_KOMPLEKS_CONNECTORS:
        if not connector_re.search(text):
            continue

        parts = [p.strip() for p in connector_re.split(text) if p.strip()]
        if len(parts) < 2 or len(parts) > MAX_HASIL_BREAKDOWN:
            continue

        if not Patterns.SLIP_FIRST_SEGMENT_CODE.match(parts[0]):
            continue

        if all(re.match(r'^[A-Za-z]', p) for p in parts):
            return [_clean_code(p) for p in parts]

        return [original_text]

    return None

def _breakdown_slip_titik_group(original_text: str):
    text = original_text.strip()
    if "." not in text:
        return None

    parts = [p.strip() for p in text.split("+") if p.strip()] if "+" in text else [text]
    hasil_total = []

    for part in parts:
        segs = [s.strip() for s in part.split(".")]
        if len(segs) < 2 or not all(s.isdigit() for s in segs):
            return None

        has_long_seg = any(len(s) >= 9 for s in segs)
        total_len = sum(len(s) for s in segs)

        if has_long_seg:
            hasil_part = _rekonstruksi_nomor_dari_pengulangan(".".join(segs))
            if not hasil_part:
                return [original_text]
            hasil_total.extend(hasil_part)
        elif 10 <= total_len <= 13:
            hasil_total.append("".join(segs))
        else:
            return [original_text]

    if not hasil_total or len(hasil_total) > MAX_HASIL_BREAKDOWN:
        return [original_text]

    return hasil_total

def _mengandung_kode_panjang(text: str) -> bool:
    text_clean = _hapus_kata_label(text)
    for m in re.finditer(rf"[A-Za-z0-9]{{{MINIMAL_PANJANG_KODE},}}", text_clean):
        if re.search(r"\d", m.group()):
            return True
    return False

def _mengandung_kata_sd(text: str) -> bool:
    return Patterns.SD.search(text) is not None

def _sepertinya_pola_range(text: str) -> bool:
    kode_kode = Patterns.RANGE_PATTERN.findall(text)
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

def _berisi_banyak_polis_bertag(text: str) -> bool:
    if "," not in text:
        return False
    parts = [p.strip() for p in text.split(",") if p.strip()]
    if len(parts) < 2:
        return False
    return all(Patterns.TAGGED_POLIS.match(p) for p in parts)

def _rekonstruksi_nomor_dari_pengulangan(teks_tanpa_label: str, acuan_awal: str = None) -> List[str]:
    raw_tokens = [t.strip() for t in re.split(r"[+,\-./]", teks_tanpa_label) if t.strip() != ""]
    tokens_expanded = [_clean_code(t) for t in raw_tokens if _clean_code(t)]

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
            prefix_alpha = (
                re.match(r"^[A-Za-z]+", nomor_acuan_terakhir).group()
                if nomor_acuan_terakhir and re.match(r"^[A-Za-z]+", nomor_acuan_terakhir)
                else ""
            )
            digit_acuan = re.sub(r"^[A-Za-z]+", "", nomor_acuan_terakhir) if nomor_acuan_terakhir else ""
            if prefix_alpha and len(tok) == len(digit_acuan):
                tok_lengkap = prefix_alpha + tok
                hasil.append(tok_lengkap)
                nomor_acuan_terakhir = tok_lengkap
            else:
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
                reconstructed = nomor_acuan_terakhir[: -len(tok)] + tok
                hasil.append(reconstructed)
                nomor_acuan_terakhir = reconstructed
            else:
                hasil.append(tok)
                nomor_acuan_terakhir = tok
        else:
            hasil.append(tok)
            nomor_acuan_terakhir = tok

    return hasil

def _proses_satu_segmen(text: str, wajib_11_digit: bool, acuan_awal: str = None) -> List[str]:
    if not _mengandung_kode_panjang(text):
        return [_rapikan_teks_label(text)]

    if _mengandung_kata_sd(text) or _sepertinya_pola_range(text):
        return [re.sub(r"\s+", " ", text).strip()]

    if _berisi_banyak_polis_bertag(text):
        hasil = [p.strip().upper() for p in text.split(",") if p.strip()]
        return hasil if len(hasil) <= MAX_HASIL_BREAKDOWN else [text.strip()]

    if _berisi_banyak_nomor_utuh(text):
        hasil = [_clean_code(t) for t in text.split()]
        return hasil if len(hasil) <= MAX_HASIL_BREAKDOWN else [text.strip()]

    if wajib_11_digit:
        has_valid_polis_format = bool(Patterns.POLIS_VALID_FORMAT.search(text)) or bool(Patterns.POLIS_SUFFIX_FORMAT.search(text))
        if not has_valid_polis_format:
            kode_pertama = re.split(r"[+,\-]", text)[0]
            if len(re.sub(r"\D", "", kode_pertama)) < MINIMAL_DIGIT_NOMOR_POLIS:
                return [text.strip()]

    teks_tanpa_placeholder = Patterns.PLACEHOLDER_TOKENS.sub("", text)
    hasil_breakdown = _rekonstruksi_nomor_dari_pengulangan(teks_tanpa_placeholder, acuan_awal)

    if not hasil_breakdown or len(hasil_breakdown) > MAX_HASIL_BREAKDOWN:
        return [text.strip()]

    return hasil_breakdown

def _chunk_digit_string(digit_str: str, preferred_len: int = None):
    kandidat_panjang = []
    if preferred_len and len(digit_str) % preferred_len == 0:
        kandidat_panjang.append(preferred_len)
    for panjang in VALID_FULL_LENGTHS_POLIS:
        if panjang not in kandidat_panjang and len(digit_str) % panjang == 0:
            kandidat_panjang.append(panjang)

    for panjang in kandidat_panjang:
        jumlah_potongan = len(digit_str) // panjang
        if jumlah_potongan >= 1:
            return [digit_str[i * panjang:(i + 1) * panjang] for i in range(jumlah_potongan)]
    return None

def _breakdown_ampersand_polis(original_text: str):
    if "&" not in original_text:
        return None

    parts = [p.strip() for p in original_text.split("&")]
    if len(parts) < 2:
        return None

    digit_parts = []
    for p in parts:
        if not p or not p.replace(" ", "").isdigit():
            return None
        digit_parts.append(re.sub(r"\D", "", p))

    panjang_panjang = [len(d) for d in digit_parts]
    if any(panjang < MINIMAL_DIGIT_NOMOR_POLIS for panjang in panjang_panjang):
        return [original_text]

    panjang_valid = [p for p in panjang_panjang if p in VALID_FULL_LENGTHS_POLIS]
    preferred_len = panjang_valid[0] if panjang_valid else None

    hasil = []
    for d in digit_parts:
        if len(d) in VALID_FULL_LENGTHS_POLIS:
            hasil.append(d)
            continue
        potongan = _chunk_digit_string(d, preferred_len)
        if not potongan:
            return [original_text]
        hasil.extend(potongan)

    if len(hasil) > MAX_HASIL_BREAKDOWN:
        return [original_text]

    return hasil

def _cek_pisah_spasi_nomor_utuh(text: str):
    groups = [g for g in text.split() if g]
    if len(groups) < 2:
        return None

    while groups and Patterns.TOKEN_SAMPAH_SPASI.match(groups[-1]):
        groups.pop()

    if len(groups) < 2 or not all(re.fullmatch(r"\d+", g) for g in groups):
        return None

    if not all(len(g) >= MINIMAL_DIGIT_NOMOR_POLIS for g in groups) or len(groups) > MAX_HASIL_BREAKDOWN:
        return None

    return groups

def _cek_slash_tanpa_digit_lengkap(text: str, wajib_11_digit: bool):
    if not wajib_11_digit or "/" not in text or _mengandung_kata_sd(text) or _sepertinya_pola_range(text):
        return None

    segmen_segmen = [s.strip() for s in text.split("/") if s.strip() != ""]
    if len(segmen_segmen) < 2:
        return None

    for segmen in segmen_segmen:
        if len(re.sub(r"\D", "", segmen)) >= MINIMAL_DIGIT_NOMOR_POLIS:
            return None

    return [text.strip()]

def _cek_various_tanpa_nomor_asli(original_text: str):
    m = re.match(r"^\s*(?:TBA|VARIOUS|VAR)\b", original_text, re.IGNORECASE)
    if not m:
        return None

    sisa = original_text[m.end():].strip()
    if sisa == "":
        return None

    digit_runs = re.findall(r"\d+", sisa)
    if any(len(d) >= MINIMAL_DIGIT_NOMOR_POLIS for d in digit_runs):
        return None

    return [original_text]

def _cek_awalan_pendek_polis(text: str):
    if not re.search(r"[+,\-]", text):
        return None
    parts = re.split(r"[+,\-]", text)
    if not parts:
        return None
        
    digit_awal = re.sub(r"\D", "", parts[0])
    if 0 < len(digit_awal) < MINIMAL_DIGIT_NOMOR_POLIS:
        return [text.strip()]
    return None

def _breakdown_polis_atau_slip_internal(value: Union[str, float], wajib_11_digit: bool) -> List:
    if pd.isna(value):
        return [value]

    original_text = str(value).strip()
    if original_text == "" or original_text.startswith("*") or Patterns.END_TOKEN.match(original_text):
        return [original_text]

    if not wajib_11_digit:
        hasil_bordero_note = _handle_bordero_note(original_text)
        if hasil_bordero_note is not None:
            return hasil_bordero_note

    original_text = _hapus_end_trailing(original_text).strip()
    if original_text == "":
        return [str(value).strip()]

    _teks_tanpa_bulan_ekor = Patterns.MONTH_YEAR_SUFFIX.sub("", original_text).strip()
    if _teks_tanpa_bulan_ekor != "":
        original_text = _teks_tanpa_bulan_ekor

    if not wajib_11_digit:
        _teks_tanpa_catatan_ekor = _hapus_catatan_ekor_slip(original_text)
        if _teks_tanpa_catatan_ekor != "":
            original_text = _teks_tanpa_catatan_ekor

        _teks_tanpa_prefix_bulan = Patterns.MONTH_YEAR_PREFIX.sub("", original_text).strip()
        if _teks_tanpa_prefix_bulan != "" and _mengandung_kode_panjang(_teks_tanpa_prefix_bulan):
            original_text = _teks_tanpa_prefix_bulan

    if re.fullmatch(r'\d{3}(?:-\d{3}){2,}', original_text.replace(" ", "")):
        return [original_text]

    hasil_various_tanpa_nomor = _cek_various_tanpa_nomor_asli(original_text)
    if hasil_various_tanpa_nomor is not None:
        return hasil_various_tanpa_nomor

    if wajib_11_digit:
        hasil_awalan_pendek = _cek_awalan_pendek_polis(original_text)
        if hasil_awalan_pendek is not None:
            return hasil_awalan_pendek

        hasil_slash_pendek = _cek_slash_tanpa_digit_lengkap(original_text, wajib_11_digit)
        if hasil_slash_pendek is not None:
            return hasil_slash_pendek

        hasil_ampersand = _breakdown_ampersand_polis(original_text)
        if hasil_ampersand is not None:
            return hasil_ampersand

        hasil_spasi_utuh = _cek_pisah_spasi_nomor_utuh(original_text)
        if hasil_spasi_utuh is not None:
            return hasil_spasi_utuh

    if not wajib_11_digit:
        _probe = Patterns.TBA_VAR.sub("", original_text.upper())
        _probe = Patterns.PLACEHOLDER_TOKENS.sub("", _probe)
        _probe = re.sub(r'[+\-/,.\s]', '', _probe)
        if _probe == "":
            return [original_text]

        _tanpa_bulan_probe = Patterns.MONTH_YEAR_PREFIX.sub("", original_text.upper()).strip()
        if _tanpa_bulan_probe != original_text.upper().strip() and not _mengandung_kode_panjang(_tanpa_bulan_probe):
            return [original_text]

        hasil_garis_koma = _breakdown_slip_garis_miring_koma(original_text)
        if hasil_garis_koma is not None:
            return hasil_garis_koma

        hasil_pengulangan = _breakdown_slip_pengulangan_digit(original_text)
        if hasil_pengulangan is not None:
            return hasil_pengulangan

        hasil_kompleks = _breakdown_slip_kode_kompleks(original_text)
        if hasil_kompleks is not None:
            return hasil_kompleks

        hasil_titik = _breakdown_slip_titik_group(original_text)
        if hasil_titik is not None:
            return hasil_titik

    seps = set(re.findall(r'(?<=\d)\s*([.,/+\-])\s*(?=\d)', original_text))
    if len(seps) > 1:
        return [original_text]

    raw = original_text.upper()
    if re.fullmatch(r"(TBA|VARIOUS|VAR|P1|P2|P3)", raw.strip()):
        return [raw.strip()]

    raw = re.sub(r'\s*END\s*$', '', raw).strip()
    raw = re.sub(r'\s*[+\-]\s*$', '', raw).strip()
    raw = Patterns.TBA_VAR.sub("", raw).strip()

    parts_awal = [p for p in re.split(r'[+,\-./]', raw) if p.strip()]
    if len(parts_awal) > MAX_HASIL_BREAKDOWN:
        return [original_text]

    if wajib_11_digit:
        m_dot_tunggal = re.match(r"^(\d{11,16})\.(\d{1,2})$", raw)
        if m_dot_tunggal:
            return [m_dot_tunggal.group(1)]

    is_slip_dot_suffix_pattern = (not wajib_11_digit) and bool(re.match(r"^\d{11}\.\d{2,4}$", raw))
    if Patterns.DOT_GROUPED_CODE.fullmatch(raw) and not is_slip_dot_suffix_pattern:
        return [raw.replace(".", "")]

    if re.fullmatch(r'\d{2,3}-\d{2,3}-\d{2,3}-\d{4,5}', raw):
        return [raw.replace('-', '')]

    raw = _pre_clean_polis_or_slip(raw)
    parts = [p for p in re.split(r'[+,\-./]', raw) if p.strip()]
    if len(parts) > MAX_HASIL_BREAKDOWN:
        return [raw]

    if not wajib_11_digit:
        hasil_khusus = _breakdown_slip_khusus(raw)
        if hasil_khusus:
            return [raw] if len(hasil_khusus) > MAX_HASIL_BREAKDOWN else hasil_khusus

    if Patterns.MASTER_POLICY_ANNOTATION.search(raw) or not _mengandung_kode_panjang(raw):
        return [_rapikan_teks_label(raw)]

    text = _hapus_kata_label(raw)
    if text.strip() == "":
        return [_rapikan_teks_label(raw)]

    segmen_segmen = _pisah_jadi_segmen(text)
    if not segmen_segmen:
        return [_rapikan_teks_label(raw)]

    ada_segmen_berkode = any(_mengandung_kode_panjang(s) for s in segmen_segmen)
    semua_hasil = []
    acuan_terakhir = None

    for segmen in segmen_segmen:
        berkode = _mengandung_kode_panjang(segmen)
        if ada_segmen_berkode and not berkode:
            if re.fullmatch(r"\d+", segmen) and acuan_terakhir and len(acuan_terakhir) > len(segmen):
                hasil_segmen = [acuan_terakhir[: -len(segmen)] + segmen]
            else:
                continue
        else:
            hasil_segmen = _proses_satu_segmen(segmen, wajib_11_digit, acuan_terakhir)

        semua_hasil.extend(hasil_segmen)
        if hasil_segmen and re.fullmatch(r"[A-Za-z0-9]+", hasil_segmen[-1]):
            acuan_terakhir = hasil_segmen[-1]

    hasil_final = []
    for item in semua_hasil:
        item_clean = item.rstrip("&").strip()
        if item_clean:
            hasil_final.append(item_clean)

    if len(hasil_final) > MAX_HASIL_BREAKDOWN:
        return [raw]

    return hasil_final if hasil_final else [_rapikan_teks_label(raw)]

def breakdown_polis_atau_slip(value: Union[str, float], wajib_11_digit: bool) -> List:
    hasil_internal = _breakdown_polis_atau_slip_internal(value, wajib_11_digit)
    if not isinstance(hasil_internal, list):
        hasil_internal = [hasil_internal]
        
    hasil_final = []
    seen_normalized = set() 
    for item in hasil_internal:
        if isinstance(item, str):
            item = item.replace("XXXXPLUSXXXX", "+").replace("XXXXMINUSXXXX", "-")
            item_clean = Patterns.MULTIPLE_SPACES.sub(" ", item.strip()).upper()
            if item_clean and item_clean not in seen_normalized:
                hasil_final.append(item_clean)
                seen_normalized.add(item_clean)
        else:
            if item not in hasil_final:
                hasil_final.append(item)
                
    return hasil_final

def clean_polis(value: Union[str, float]):
    return value if pd.isna(value) else _rapikan_teks_label(str(value))

def breakdown_polis(value: Union[str, float]) -> List:
    return breakdown_polis_atau_slip(value, wajib_11_digit=True)

def clean_slip(value: Union[str, float]):
    if pd.isna(value):
        return value
    original_text = str(value).strip()
    hasil_bordero_note = _handle_bordero_note(original_text)
    if hasil_bordero_note is not None:
        return hasil_bordero_note[0]
    return _rapikan_teks_label(original_text)

def breakdown_slip(value: Union[str, float]) -> List:
    return breakdown_polis_atau_slip(value, wajib_11_digit=False)

def _insert_breakdown_columns(
    df: pd.DataFrame, source_col: str, fn_breakdown, prefix: str, mask=None
) -> None:
    if mask is None:
        hasil_breakdown = df[source_col].apply(fn_breakdown)
    else:
        hasil_breakdown = pd.Series([[] for _ in range(len(df))], index=df.index, dtype=object)
        hasil_breakdown.loc[mask] = df.loc[mask, source_col].apply(fn_breakdown)

    panjang_terpanjang = int(hasil_breakdown.apply(len).max() or 0)
    max_cols = min(panjang_terpanjang, MAX_HASIL_BREAKDOWN) if panjang_terpanjang > 0 else 1

    hasil_breakdown_trimmed = hasil_breakdown.apply(
        lambda x: x[:max_cols] if isinstance(x, list) else x
    )

    df_breakdown = pd.DataFrame(
        hasil_breakdown_trimmed.tolist(),
        columns=[f"{prefix}_{i + 1}" for i in range(max_cols)],
        index=df.index,
    )

    posisi = df.columns.get_loc(source_col)
    for i, col in enumerate(df_breakdown.columns):
        df.insert(posisi + 1 + i, col, df_breakdown[col])

def process_data(input_file: str, sheet_name: str, output_file: str) -> None:
    if not os.path.exists(input_file):
        print(f"[ERROR] File input tidak ditemukan: {input_file}")
        return

    print(f"Membaca data dari: {input_file} (Sheet: {sheet_name}) ...")
    df = pd.read_excel(input_file, sheet_name=sheet_name, header=0)
    df.columns = df.columns.str.strip()

    if CEDANT_COL not in df.columns:
        print(f"[ERROR] Kolom '{CEDANT_COL}' tidak ditemukan di file!")
        return

    df = df[df[CEDANT_COL].astype(str).str.strip() == CEDANT_VALUE].copy()
    if df.empty:
        print("[WARN] Data kosong setelah filter. Proses dihentikan.")
        return

    print("Menjalankan pembersihan dan breakdown kolom...")
    column_mappings = [
        ("FAC_INSURED", breakdown_insured, "FAC_INSURED_CLN"),
        ("FAC_POLICY_NO", breakdown_polis, "FAC_POLICY_CLEAN"),
        ("FAC_SLIP", breakdown_slip, "FAC_SLIP_CLEAN"),
    ]

    for source_col, fn, prefix in column_mappings:
        if source_col in df.columns:
            _insert_breakdown_columns(df, source_col, fn, prefix)
        else:
            print(f"[WARN] Kolom {source_col} tidak ditemukan, dilewati.")

    output_dir = os.path.dirname(output_file)
    if output_dir and not os.path.exists(output_dir):
        os.makedirs(output_dir)

    print(f"Menyimpan hasil ke Excel: {output_file} ...")
    try:
        df.to_excel(output_file, index=False)
        print("Pemrosesan data selesai!")
        print(f"File disimpan di : {output_file}")
    except PermissionError:
        print(f"[ERROR] Tidak dapat menyimpan file. Pastikan file {output_file} tidak sedang dibuka di Excel!")

if __name__ == "__main__":
    process_data(INPUT_FILE, SHEET_NAME, OUTPUT_FILE)