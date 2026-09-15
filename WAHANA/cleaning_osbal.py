import os
import re
import pandas as pd

INPUT_FILE = os.path.join("input", "2b. transaksi Osbal 01.01.23 - 17.08.26.xlsx")
OUTPUT_FILE = os.path.join("output", "wahana_output_osbalnew.xlsx")

CEDANT_COL   = "CCOS_COMP_NAME"
CEDANT_VALUE = "PT.ASURANSI WAHANA TATA"

POLIS_COL   = "FAC_POLICY_NO"
SLIP_COL    = "FAC_SLIP"
INSURED_COL = "FAC_INSURED"
CLASS_COL   = "CLASS_NAME" # Added for Marine Cargo check

CLSDT_POLIS_COL = "CLSDT_POLICY_NO"
CLSDT_SLIP_COL  = "CLSDT_SLIP_NO"
CLSDT_SERTF_COL = "CLSDT_SERTF_NO"

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

# Hapus token PT / QQ dimanapun posisinya (bukan hanya suffix)
INSURED_PT_QQ_RE = re.compile(
    r"""
    \bPT\.?\b
  | \bQQ\b
    """,
    re.IGNORECASE | re.VERBOSE,
)

# Noise token dalam slip (bulan, tahun, mata uang, dsb.)
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

# Kandidat token nomor slip: alfanumerik ≥ 7 karakter
_SLIP_TOKEN_RE = re.compile(r"[A-Z0-9][A-Z0-9\-]{6,}", re.IGNORECASE)


# ─────────────────────────────────────────────────────────────────────────────
# SELEKSI SUMBER DATA (CLSDT vs FAC FALLBACK)
# ─────────────────────────────────────────────────────────────────────────────

def resolve_source_value(clsdt_val, fac_val):
    """
    Prioritaskan CLSDT jika nilainya valid.
    CLSDT tidak valid jika: kosong/NaN atau bernilai TBA, VAR, VARIOUS.
    Jika tidak valid, gunakan FAC sebagai fallback.
    """
    if pd.isna(clsdt_val):
        return fac_val

    clsdt_str = str(clsdt_val).strip()
    if not clsdt_str:
        return fac_val

    if clsdt_str.upper() in {"TBA", "VAR", "VARIOUS"}:
        return fac_val

    return clsdt_val


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
    """Hapus suffix badan usaha (TBK, PT, CV) secara iteratif, lalu hapus
    token PT/QQ dimanapun posisinya dan tanda kurung (isi di dalamnya
    dipertahankan)."""
    name = _normalize_spaces(name)
    for _ in range(3):
        cleaned = _normalize_spaces(INSURED_SUFFIX_RE.sub("", name).strip().strip(","))
        if cleaned == name:
            break
        name = cleaned

    name = INSURED_PT_QQ_RE.sub(" ", name)
    name = name.replace("(", " ").replace(")", " ")
    name = _normalize_spaces(name)
    name = re.sub(r"^[\s.,\-]+|[\s.,\-]+$", "", name)

    return name


def _cap_or_join(items: list) -> list:
    """Jika jumlah item melebihi MAX_SPLIT_COLS, gabungkan dengan koma."""
    if len(items) > MAX_SPLIT_COLS:
        return [",".join(items)]
    return items

def parse_dash_chain(val):
    """
    Memisahkan nomor polis berdasarkan '-'.

    Hanya suffix yang merupakan breakdown polis yang
    akan dibuat menjadi clean polis.

    Contoh:
    09840502012020000115-110-114-109
    -> [
        09840502012020000115,
        09840502012020000110,
        09840502012020000114,
        09840502012020000109
    ]

    Sedangkan:
    02440502012023000394-000061-0003
    -> hanya base karena suffix bukan breakdown.
    """
    m = re.fullmatch(r"(\d+)((?:-\d+)+)", str(val).strip())

    if not m:
        return [], []

    base = m.group(1)
    suffixes = [x for x in m.group(2).split("-") if x]

    polis_list = [base]
    certificate_suffixes = []

    for suffix in suffixes:
        if _is_polis_breakdown_suffix(base, suffix):
            suffix_len = len(suffix)

            # Bentuk nomor polis lengkap berdasarkan suffix
            full_polis = (
                base[:-suffix_len] + suffix
                if len(base) > suffix_len
                else suffix
            )

            polis_list.append(full_polis)
        else:
            # Bukan breakdown -> biarkan certificate menangani
            certificate_suffixes.append(suffix)

    return polis_list, certificate_suffixes

# ─────────────────────────────────────────────────────────────────────────────
# CLEAN POLIS (WAHANA RULES - TIDAK DIUBAH)
# ─────────────────────────────────────────────────────────────────────────────
def _clean_polis_core(val) -> list:
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

    if "/" in val:
        val = val.split("/", 1)[0].strip()

    if re.fullmatch(r"VARIOUS", val_upper):
        return [val.strip()]

    if re.match(r"^VARIOUS\s*-\s*SEE\s+ATTACH", val_upper):
        return [val.strip()]

    if re.fullmatch(r"TBA", val_upper):
        return [val.strip()]

    if re.match(r"^\s*P1\s*/", val, re.IGNORECASE):
        after = re.sub(r"^\s*P1\s*/\s*", "", val, flags=re.IGNORECASE).strip()
        tokens = _extract_polis_tokens(after)
        if tokens:
            return [_normalize_spaces(tokens[0])]

    if re.fullmatch(r"\s*P1\s*(?:\+|\()\s*P2\s*\)?\s*", val, re.IGNORECASE):
        return [_normalize_spaces(val)]

    if re.search(r"[A-Za-z]", val):
        tokens = _extract_polis_tokens(val)
        if tokens:
            m = re.search(r"\d", val)
            if m:
                prefix = val[:m.start()].strip()
                if prefix:
                    return [_normalize_spaces(tokens[0])]

    if re.search(r"S/D", val, re.IGNORECASE):
        sd_parts = re.split(r"\s{2,}", val.strip())
        left = sd_parts[0] if sd_parts else val

        base_m = (re.match(r"^(\d[\d\-]+?)-\d+S/D\d+", left, re.IGNORECASE) or
                  re.match(r"^(\d[\d\-]+)S/D\d+", left, re.IGNORECASE))
        results = ([base_m.group(1).strip()] if base_m else [])

        for part in sd_parts[1:]:
            part = re.sub(r"\+?\s*\bTBA\b\s*", "", part, flags=re.IGNORECASE).strip().strip("+")
            results += [t for t in part.split() if _is_valid_polis_token(t.strip())]

        if results:
            return _cap_or_join(results)

    if re.match(r"^\s*TBA\s{2,}", val, re.IGNORECASE):
        rest = re.sub(r"^\s*TBA\s+", "", val, flags=re.IGNORECASE).strip()
        tokens = _extract_polis_tokens(rest)
        if tokens:
            return _cap_or_join(tokens)


    dash_match = re.fullmatch(r"(\d+)((?:-\d+)+)", val)

    if dash_match:
        base = dash_match.group(1)
        suffixes = [
            x for x in dash_match.group(2).split("-")
            if x
        ]

        results = [base]

        for suffix in suffixes:

            # =========================================================
            # 1. SUFFIX <= 6 DIGIT
            #    Cek apakah ini breakdown polis.
            # =========================================================
            if len(suffix) <= 6:
                if _is_polis_breakdown_suffix(base, suffix):
                    suffix_len = len(suffix)

                    if len(base) > suffix_len:
                        results.append(
                            base[:-suffix_len] + suffix
                        )

            # =========================================================
            # 2. SUFFIX > 6 DIGIT
            #    Tidak mungkin certificate.
            #    Untuk CLEAN POLIS dianggap breakdown / bagian polis.
            # =========================================================
            else:
                # Suffix > 6 digit = breakdown polis langsung
                # Tidak digabung dengan base
                results.append(suffix)

        return _cap_or_join(results)

        # ============================================================
    # DASH CHAIN:
    # BEDAKAN BREAKDOWN POLIS VS CERTIFICATE
    # ============================================================
    

    if re.search(r"\s{2,}", val):
        blocks = re.split(r"\s{2,}", val.strip())
        first_block = blocks[0].strip()

        dash_first = re.fullmatch(r"(\d+)((?:-\d+)+)", first_block)
        if dash_first:
            base = dash_first.group(1)
            suffixes = [x for x in dash_first.group(2).split("-") if x]
            if suffixes:
                lengths = {len(s) for s in suffixes}
                if len(lengths) == 1:
                    suffix_len = len(suffixes[0])
                    if 2 <= suffix_len <= 6:
                        if len(base) > suffix_len:
                            results = [base]
                            for suffix in suffixes:
                                results.append(base[:-suffix_len] + suffix)
                            for block in blocks[1:]:
                                tokens = _extract_polis_tokens(block)
                                for token in tokens:
                                    if _is_valid_polis_token(token.strip()):
                                        results.append(token.strip())
                            return _cap_or_join(results)

        results = []
        for block in blocks:
            tokens = _extract_polis_tokens(block)
            if tokens:
                results.append(_normalize_spaces(tokens[0]))
            if len(blocks) > 1:
                first = blocks[0].strip()
                m = re.fullmatch(r"(.+?)(\d{6})\.00", first)
                if m:
                    prefix = m.group(1)
                    results = [first]
                    ok = True
                    for b in blocks[1:]:
                        b = b.strip()
                        if re.fullmatch(r"\d{6}\.00", b):
                            results.append(prefix + b)
                        else:
                            ok = False
                            break
                    if ok:
                        return _cap_or_join(results)

        if results:
            return _cap_or_join(results)

    if "/" in val:
        slash_parts = re.split(r"\s*/\s*", val, maxsplit=1)
        left = slash_parts[0].strip()
        if left:
            if _is_valid_polis_token(left):
                return [_normalize_spaces(left)]
            tokens = _extract_polis_tokens(left)
            if tokens:
                return [_normalize_spaces(tokens[0])]

    dash_plus_match = re.fullmatch(r"([\d.]+)-(\d+(?:\+\d+)+)", val)

    if dash_plus_match:
        base = dash_plus_match.group(1)
        suffixes = [
            s for s in dash_plus_match.group(2).split("+")
            if s
        ]

        results = [base]

        for suffix in suffixes:
            if _is_polis_breakdown_suffix(base, suffix):
                suffix_len = len(suffix)

                if len(base) > suffix_len:
                    results.append(
                        base[:-suffix_len] + suffix
                    )

        if len(results) > 1:
            return _cap_or_join(results)

        return [_normalize_spaces(base)]

    val = re.sub(r"\s*\+\s*P[3-9]\d*\b", "", val, flags=re.IGNORECASE)
    val = _normalize_spaces(val)

    m = re.fullmatch(r"(.+?\.)(\d{6})((?:\+\d{3})+)", val.replace(" ", ""))
    if m:
        prefix = m.group(1)
        first = m.group(2)
        suffixes = [s for s in m.group(3).split("+") if s]
        results = [prefix + first + ".00"]
        for s in suffixes:
            results.append(prefix + first[:-3] + s + ".00")
        return _cap_or_join(results)

    val = re.sub(r"\b[^\s+]*[A-Za-z][^\s+]*\b", "", val)
    val = re.sub(r"\s*\+\s*", " + ", val)
    val = re.sub(r"^\s*\+\s*|\s*\+\s*$", "", val)
    val = re.sub(r"(?:\+\s*){2,}", "+ ", val)
    val = _normalize_spaces(val)

    if "+" in val:
        return [_normalize_spaces(val)]

    cleaned = _normalize_spaces(val)
    cleaned = cleaned.strip(" -/")

    if not cleaned:
        return []
    return [cleaned]



def clean_polis(val) -> list:
    return [r.replace(".", "") for r in _clean_polis_core(val)]


# ─────────────────────────────────────────────────────────────────────────────
# CLEAN SLIP
# ─────────────────────────────────────────────────────────────────────────────
def _clean_slip_core(val) -> list:
    if pd.isna(val):
        return []
    val = str(val).strip()
    if not val:
        return []
    val = _normalize_spaces(val)
    
    if re.fullmatch(r"(P\d+(?:\s*\+\s*P\d+)*(\s+CANCEL)?|(JANUARI|FEBRUARI|MARET|APRIL|MEI|JUNI|JULI|AGUSTUS|SEPTEMBER|OKTOBER|NOVEMBER|DESEMBER)\s+\d{4}(?:\s*-\s*(IDR|USD|EUR|GBP|SGD|JPY|AUD|CNY))?)", val, flags=re.IGNORECASE):
        return []

    val = re.sub(r"\s*/\s*END\b", "", val, flags=re.IGNORECASE).strip()
    val = re.sub(r"\+VAR\b", "", val, flags=re.IGNORECASE)
    val = _normalize_spaces(val)

    parts = re.split(r"\s{2,}", val)
    parts = [p.strip() for p in parts if p.strip()]
    if len(parts) >= 2:
        results = []
        for p in parts:
            results.append(_normalize_spaces(p))
        return _cap_or_join(results)
   
    if SLIP_EXCEPTION_RE.search(val):
        return [_normalize_spaces(val)]

    if re.fullmatch(r"\d+(?:\.\d+)+(?:-\d+){3,}", val):
        return [_normalize_spaces(val)]

    if "+" in val:
        parts = [p.strip() for p in re.split(r"\s*\+\s*", val) if p.strip()]
        parts = [p for p in parts if not re.fullmatch(r"P\d+", p, flags=re.IGNORECASE)]
        if len(parts) > 1:
            first = parts[0]
            rest = parts[1:]
            rest_is_digit = all(re.fullmatch(r"\d+", r) for r in rest)
            suffix_lengths = ({len(r) for r in rest} if rest_is_digit else set())
            if rest_is_digit and len(suffix_lengths) == 1:
                suffix_len = suffix_lengths.pop()
                m = re.match(rf"^(.*?)(\d{{{suffix_len}}})$", first)
                if m:
                    prefix = m.group(1)
                    results = [first]
                    for s in rest:
                        s = s.zfill(suffix_len)
                        results.append(prefix + s)
                    return _cap_or_join(results)
            if all(re.match(r"^\d+(?:\.\d+)+$", p) for p in parts):
                return _cap_or_join(parts)
            return [_normalize_spaces(val)]

    m = re.match(r"^(.*?)(\d+)-(\d+)$", val)
    if m:
        prefix = m.group(1)
        first = m.group(2)
        second = m.group(3)
        results = [prefix + first]
        if len(second) < len(first):
            second = second.zfill(len(first))
            results.append(prefix + second)
        else:
            results.append(second)
        return _cap_or_join(results)

    slips = re.findall(r"\d+(?:\.\d+)+", val)
    if len(slips) >= 2:
        return _cap_or_join(slips)

    m = re.search(r"\d+(?:\.\d+)+", val)
    if m:
        return [m.group(0)]

    return [_normalize_spaces(val)]     
    
def clean_slip(val) -> list:
    results = _clean_slip_core(val)
    cleaned = []
    for r in results:
        if pd.isna(r):
            continue
        r = str(r)
        r = _normalize_spaces(r)
        r = r.replace(".", "")
        r = r.strip()
        if not r:
            continue
        if r not in cleaned:
            cleaned.append(r)
    return cleaned


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
# NEW FORMAT CERT RANGE HELPER
# ─────────────────────────────────────────────────────────────────────────────
def format_cert_range(start_str, end_str):
    try:
        start_int = int(start_str)
        end_int = int(end_str)
        
        # JIKA TERBALIK (misal 61 lalu 3), JANGAN PAKAI SD! Pecah jadi koma.
        if end_int < start_int:
            return f"{str(start_int).zfill(6)}, {str(end_int).zfill(6)}"
            
        count = end_int - start_int + 1
        if count > 3:
            return f"{str(start_int).zfill(6)} SD {str(end_int).zfill(6)}"
        else:
            return ", ".join(str(i).zfill(6) for i in range(start_int, end_int + 1))
    except ValueError:
        return ""

# ─────────────────────────────────────────────────────────────────────────────
# CLEAN CERTIFICATE
# ─────────────────────────────────────────────────────────────────────────────
def clean_certificate(polis_ori, clsdt_sertf=None, class_name=None):
    """
    Membersihkan certificate.

    PRIORITAS:
    1. Jika CLSDT_SERTF_NO valid -> gunakan CLSDT_SERTF_NO
    2. Jika CLSDT_SERTF_NO tidak valid -> fallback ke POLIS ORI

    CLSDT tidak valid jika:
    - kosong / NaN
    - TBA
    - VAR
    - VARIOUS

    Certificate:
    - maksimal 3
    - numeric
    - maksimal 6 digit
    - jika < 6 digit -> zfill(6)
    - range menggunakan format SD jika > 3 nomor
    """

    # ============================================================
    # HELPER VALIDASI CLSDT CERTIFICATE
    # ============================================================
    def _valid_clsdt_certificate(value):
        if pd.isna(value):
            return False

        value = str(value).strip()

        if not value:
            return False

        if value.upper() in {"TBA", "VAR", "VARIOUS"}:
            return False

        return True

    # ============================================================
    # 1. PRIORITAS CLSDT_SERTF_NO
    # ============================================================
    if _valid_clsdt_certificate(clsdt_sertf):

        clsdt = str(clsdt_sertf).strip()

        # Normalisasi S/D
        clsdt = re.sub(
            r"\s*S\s*/?\s*D\s*",
            " SD ",
            clsdt,
            flags=re.IGNORECASE
        )

        clsdt = _normalize_spaces(clsdt)

        # --------------------------------------------------------
        # CLSDT berbentuk range:
        # 000001 S/D 000005
        # --------------------------------------------------------
        m_sd = re.fullmatch(
            r"(\d{1,6})(?:\.00)?\s*SD\s*(\d{1,6})(?:\.00)?",
            clsdt,
            re.IGNORECASE
        )

        if m_sd:
            result = format_cert_range(
                m_sd.group(1),
                m_sd.group(2)
            )

            return [result] if result else []

        # --------------------------------------------------------
        # CLSDT certificate tunggal
        # --------------------------------------------------------
        if re.fullmatch(r"\d{1,6}(?:\.00)?", clsdt):
            number = re.sub(r"\.00$", "", clsdt)

            return [number.zfill(6)]

        # --------------------------------------------------------
        # CLSDT beberapa certificate dengan koma
        # --------------------------------------------------------
        if "," in clsdt:
            certificates = []

            for part in clsdt.split(","):
                part = part.strip()

                if re.fullmatch(r"\d{1,6}(?:\.00)?", part):
                    part = re.sub(r"\.00$", "", part)
                    certificates.append(part.zfill(6))

                    if len(certificates) >= 3:
                        break

            if certificates:
                return certificates

    # ============================================================
    # 2. FALLBACK KE POLIS ORI
    # ============================================================
    if pd.isna(polis_ori):
        return []

    polis = str(polis_ori).strip()

    if not polis:
        return []

    # ============================================================
    # EXCEPTION
    # ============================================================
    if polis == "02210502012023000128-000129":
        return []

    if polis == "00940502012023001276-136-135-1275":
        return []

    # ============================================================
    # 3. HANDLE S/D DARI POLIS ORI
    # ============================================================
    m_sd = re.search(
        r"(\d{1,6})(?:\.00)?\s*S\s*/?\s*D\s*(\d{1,6})(?:\.00)?\b",
        polis,
        re.IGNORECASE
    )

    if m_sd:
        first = m_sd.group(1)
        last = m_sd.group(2)

        result = format_cert_range(first, last)

        return [result] if result else []

    # ============================================================
    # 4. HANDLE DASH DARI POLIS ORI
    # ============================================================
    if "-" not in polis:
        return []

    parts = [p.strip() for p in polis.split("-")]

    if len(parts) < 2:
        return []

    base = parts[0]
    suffixes = parts[1:]

    certificates = []

    for suffix in suffixes:

        # Harus angka murni
        if not re.fullmatch(r"\d+", suffix):
            continue

        # Lebih dari 6 digit bukan certificate
        if len(suffix) > 6:
            continue

        # ========================================================
        # CEK BREAKDOWN POLIS
        # ========================================================
        if _is_polis_breakdown_suffix(base, suffix):
            continue

        # ========================================================
        # BUKAN BREAKDOWN -> CERTIFICATE
        # ========================================================
        certificates.append(suffix.zfill(6))

        if len(certificates) >= 3:
            break

    return certificates

def _is_polis_breakdown_suffix(base, suffix):
    """
    Menentukan apakah suffix adalah lanjutan nomor polis,
    bukan certificate.

    Contoh:
    09810502012021000074-00073
    -> 00073 = breakdown polis

    09840502012020000115-110
    -> 110 = breakdown polis

    00940502012020001202-1201
    -> 1201 = breakdown polis

    Sedangkan:
    00940502012020001202-0074
    -> 0074 = certificate
    """

    if not re.fullmatch(r"\d+", base):
        return False

    if not re.fullmatch(r"\d+", suffix):
        return False

    suffix_len = len(suffix)

    # Suffix terlalu panjang → bukan breakdown polis
    if suffix_len > 6:
        return False

    # Minimal 2 digit untuk pola breakdown
    if suffix_len < 2:
        return False

    if len(base) <= suffix_len:
        return False

    base_end = base[-suffix_len:]

    try:
        base_num = int(base_end)
        suffix_num = int(suffix)

        diff = suffix_num - base_num

        # Breakdown polis biasanya nomor berdekatan.
        # Mendukung kasus turun seperti:
        # 74 -> 73
        # 1202 -> 1201
        #
        # dan naik seperti:
        # 128 -> 129
        return -10 <= diff <= 10

    except ValueError:
        return False

# ─────────────────────────────────────────────────────────────────────────────
# PROCESS DATA
# ─────────────────────────────────────────────────────────────────────────────
def _insert_clean_columns(df: pd.DataFrame, all_lists: list, prefix: str, max_cols: int) -> list:
    """Buat kolom clean_{prefix}_N dan kembalikan nama-nama kolom baru."""
    added = []
    for i in range(1, max_cols + 1):
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

    required_cols = [
        POLIS_COL, SLIP_COL, INSURED_COL,
        CLSDT_POLIS_COL, CLSDT_SLIP_COL, CLSDT_SERTF_COL,
    ]
    
    # Amankan class column (jika tidak ada, buat dummy agar tidak error)
    if CLASS_COL not in df.columns:
        df[CLASS_COL] = ""
        
    for col in required_cols:
        if col not in df.columns:
            print(f"\n[ERROR] Kolom '{col}' tidak ditemukan di file!")
            print(f"        Kolom tersedia: {list(df.columns)}")
            return

    df[CEDANT_COL] = df[CEDANT_COL].astype(str).str.strip()
    is_wahana = df[CEDANT_COL] == CEDANT_VALUE
    print(f"[2/5] Filter cedant '{CEDANT_VALUE}': {is_wahana.sum():,} baris WAHANA dari total {len(df):,} baris.")

    df.rename(columns={POLIS_COL: "polis_ori", SLIP_COL: "slip_ori", INSURED_COL: "insured_ori"}, inplace=True)

    print("[3/5] Menjalankan proses selection sumber (CLSDT vs FAC) dan cleaning ...")

    all_clean_polis = []
    all_clean_slip = []
    all_clean_ins = []
    all_certificates = []

    max_polis = max_slip = max_ins = 1

    for idx, (_, row) in enumerate(df.iterrows(), 1):
        if idx % 50_000 == 0:
            print(f"      Progress: {idx:,} / {len(df):,} baris diproses...")

        if is_wahana.iloc[idx - 1]:
            raw_polis = resolve_source_value(row.get(CLSDT_POLIS_COL), row.get("polis_ori"))
            raw_slip  = resolve_source_value(row.get(CLSDT_SLIP_COL), row.get("slip_ori"))
            raw_ins   = row.get("insured_ori", "")
            class_nm  = row.get(CLASS_COL, "")

            c_polis = clean_polis(raw_polis)
            c_slip  = clean_slip(raw_slip)
            c_ins   = clean_insured(raw_ins)

            # Untuk Wahana saat ini logic Certificate menganggap kalau dipecah koma (C1, C2)
            # Karena logic clean_polis Wahana outputnya flat list (bukan list of dict), 
            # maka certificate digenerate 1 string. Nanti di step selanjutnya dipecah per kolom
            c_certificate = clean_certificate(
                row.get("polis_ori"),
                row.get(CLSDT_SERTF_COL),
                class_name=class_nm
            )
        else:
            c_polis, c_slip, c_ins = [], [], []
            c_certificate = ""

        max_polis = max(max_polis, len(c_polis))
        max_slip  = max(max_slip,  len(c_slip))
        max_ins   = max(max_ins,   len(c_ins))

        all_clean_polis.append(c_polis)
        all_clean_slip.append(c_slip)
        all_clean_ins.append(c_ins)
        
        # clean_certificate() SUDAH mengembalikan list
        # Contoh:
        # ["000156"]
        # ["000074", "000073"]
        # ["000383", "000045", "000031"]

        if isinstance(c_certificate, list):
            all_certificates.append(c_certificate)
        elif c_certificate:
            all_certificates.append([str(c_certificate).strip()])
        else:
            all_certificates.append([])

    print(f"      Selesai diproses!")

    print("[4/5] Menyusun kolom output ...")

    new_columns = []
    for col in df.columns:
        new_columns.append(col)
        
        if col == "polis_ori":
            # ============================================================
            # FIX: SUSUNAN KOLOM BERPASANGAN
            # (clean polis 1, Certificate 1, clean polis 2, Certificate 2...)
            # ============================================================
            for i in range(1, max_polis + 1):
                polis_col = f"clean polis {i}"
                cert_col = f"Certificate {i}"
                
                df[polis_col] = [lst[i - 1] if i - 1 < len(lst) else None for lst in all_clean_polis]
                new_columns.append(polis_col)
                
                if i <= 3:
                    df[cert_col] = [
                        cert_lst[i - 1] if i - 1 < len(cert_lst) else None
                        for cert_lst in all_certificates
                    ]
                    new_columns.append(cert_col)

        elif col == "slip_ori":
            new_columns += _insert_clean_columns(df, all_clean_slip,  "slip",    max_slip)
        elif col == "insured_ori":
            new_columns += _insert_clean_columns(df, all_clean_ins,   "insured", max_ins)

    # Hanya simpan kolom yang kita butuhkan/sudah tersusun
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
# ENTRY POINT (TEST POLA STRING)
# ─────────────────────────────────────────────────────────────────────────────
# ═════════════════════════════════════════════════════════════════════════════
# ENTRY POINT - TEST CERTIFICATE
# ═════════════════════════════════════════════════════════════════════════════

# if __name__ == "__main__":

#     TEST_DATA = [
#         # Test sebelumnya
#         "02210502012023000128-000129",
#         "02440502012023000394-000061-0003",
#         "00940502012023001276-136-135-1275",
#         "02210502012023000494-70320230003-002",

#         # Test tambahan
#         "0041050201202000183-0045",
#         "09810502012021000074-00073",
#         "01810502012020000259-0008",
#         "09840502012020000115-110-114-109",
#         "00940502012020001202-0074-1201-0073",
#         "02240502012020001376-000320",
#         "01540502012021000442-51481443442445452453457595861",
#         "09810502012020000566-00383-00045-00031",
#     ]

#     print("=" * 100)
#     print("TEST CLEAN POLIS + CERTIFICATE")
#     print("=" * 100)

#     for no, polis_ori in enumerate(TEST_DATA, 1):

#         clean_polis_result = clean_polis(polis_ori)
#         certificate_result = clean_certificate(polis_ori)

#         print(f"\n[{no}]")
#         print(f"POLIS ORI : {polis_ori}")

#         print("\nCLEAN POLIS:")
#         for i in range(5):
#             value = (
#                 clean_polis_result[i]
#                 if i < len(clean_polis_result)
#                 else ""
#             )
#             print(f"  clean polis {i + 1} : {value}")

#         print("\nCERTIFICATE:")
#         for i in range(3):
#             value = (
#                 certificate_result[i]
#                 if i < len(certificate_result)
#                 else ""
#             )
#             print(f"  Certificate {i + 1} : {value}")

#         print("-" * 100)

#     print("\n" + "=" * 100)
#     print("TEST SELESAI")
#     print("=" * 100)

if __name__ == "__main__":
    process_data(INPUT_FILE, OUTPUT_FILE)