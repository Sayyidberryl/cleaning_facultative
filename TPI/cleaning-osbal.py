"""
Cleaning Data 2 - Osbal | Cedant: TPI (PT Asuransi Tugu Pratama Indonesia)
"""

import os
import re

import pandas as pd


# ─────────────────────────────────────────────────────────────────────────────
# KONFIGURASI — ubah bagian ini kalau nama file/kolom berbeda
# ─────────────────────────────────────────────────────────────────────────────

INPUT_FILE  = os.path.join("input", "2a. Transaksi Osbal 01.01.23 - 17.07.26 Rev.xlsx")
OUTPUT_FILE = os.path.join("output", "tpi_output_osbal.xlsx")

CEDANT_COL   = "CCOS_COMP_NAME"
CEDANT_VALUE = "PT ASURANSI TUGU PRATAMA INDONESIA"

POLIS_COL   = "FAC_POLICY_NO"
SLIP_COL    = "FAC_SLIP"
INSURED_COL = "FAC_INSURED"

MAX_SPLIT_COLS = 5   # batas breakdown maksimal, sisanya digabung koma

# Sesuai catatan TPI: polis diawali/diikuti P1 atau P2 dibiarkan apa adanya
_POLIS_KEEP_AS_IS_RE = re.compile(r"(^\s*P[123]\s*/|/\s*P[123]\s*$|\+\s*P[123]\s*$)", re.IGNORECASE)


# ─────────────────────────────────────────────────────────────────────────────
# REGEX & KAMUS PEMBANTU (diadaptasi dari referensi)
# ─────────────────────────────────────────────────────────────────────────────

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

INSURED_JUNK_WORDS = frozenset({
    "PT", "CV", "TBK", "PERSERO", "LTD", "INC", "LLC",
    "AND", "OR", "THE", "OF", "AS",
    "NON FOOD", "DIV",
})

INSURED_SUFFIX_RE = re.compile(
    r"""
    ,?\s*\bTBK\b\s*(?:,?\s*PT\.?)?
  | ,?\s*\(\s*PERSERO\s*\)\s*
  | ,?\s*\bPERSERO\b\s*
  | ,?\s*\bLTD\.?\b\s*
  | ,?\s*\bPT\.?\s*$
  | ,?\s*\bCV\.?\s*$
    """,
    re.IGNORECASE | re.VERBOSE,
)

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
    tok = tok.strip()
    return len(tok) >= 5 and not re.search(r"TBA$", tok, re.IGNORECASE) and bool(re.search(r"\d", tok))


def _extract_polis_tokens(text: str) -> list:
    tokens = []
    for block in re.split(r"\s{2,}", text.strip()):
        for tok in re.split(r"\+|\s+", block.strip()):
            tok = tok.strip().strip("-/")
            if _is_valid_polis_token(tok):
                tokens.append(tok)
    return tokens


def _is_valid_slip_token(tok: str) -> bool:
    tok = tok.strip()
    return len(tok) >= 7 and bool(re.search(r"\d", tok)) and not re.match(r"^\d{4}$", tok)


def _strip_slip_suffix(tok: str) -> str:
    m = re.match(r"^(.+?)-(\d{1,6})$", tok)
    if m and len(m.group(1)) > len(m.group(2)):
        return m.group(1)
    return tok


def _extract_slip_tokens(text: str) -> list:
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


_INSURED_PAREN_JUNK = {"PERSERO", "TBK", "SMALL SHIP FLEET"}


def _handle_insured_paren(match: "re.Match") -> str:
    content_ = match.group(1).strip()
    if content_.upper() in _INSURED_PAREN_JUNK or not content_:
        return ""
    # ambil bagian sebelum koma pertama saja (mis. "44 VESSEL, LIHAT PAGE 3"
    # -> "44 VESSEL"), sisanya biasanya cuma catatan tambahan
    kept = content_.split(",")[0].strip()
    return " " + kept + " " if kept else ""

_REGION_PREFIX_WORDS = {"KALTIM", "KALSEL", "KALTENG", "KALBAR", "KALUT",
                         "SULSEL", "SULUT", "SULTENG", "SULBAR", "SULTRA",
                         "JATIM", "JATENG", "JABAR", "SUMUT", "SUMSEL",
                         "SUMBAR", "NTB", "NTT", "DKI"}


def _merge_region_prefix_parts(parts: list) -> list:
    merged = []
    skip_next_prefix = False
    for i, p in enumerate(parts):
        p_stripped = p.strip()
        if p_stripped.upper() in _REGION_PREFIX_WORDS and i + 1 < len(parts):
            merged.append(p_stripped + " " + parts[i + 1].strip())
            skip_next_prefix = True
        elif skip_next_prefix:
            skip_next_prefix = False
            continue
        else:
            merged.append(p_stripped)
    return merged

def _merge_digit_start_parts(parts: list) -> list:
    """Segmen yang diawali angka (mis. "7 SUPPLY VESSEL", "44 VESSEL") itu
    nama unit/armada, bukan perusahaan baru -> digabung ke segmen
    sebelumnya, bukan dianggap entitas terpisah."""
    merged = []
    for p in parts:
        p_stripped = p.strip()
        if merged and re.match(r"^\d", p_stripped):
            merged[-1] = merged[-1] + " " + p_stripped
        else:
            merged.append(p_stripped)
    return merged


def _name_initials_variants(name: str) -> set:
    words = [w for w in re.split(r"[\s\-]+", name.upper()) if w]
    if not words:
        return set()
    variants = {"".join(w[0] for w in words if w[0].isalpha())}
    if len(words) >= 2:
        variants.add("".join(w[0] for w in words[:-1] if w[0].isalpha()) + words[-1])
    return variants


def _is_abbreviation_of(short: str, long_name: str) -> bool:
    """Cek apakah `short` adalah singkatan dari `long_name`, mis. "TPPI"
    dari "TRANS PACIFIC PETROCHEMICAL INDOTAMA"."""
    short_clean = re.sub(r"[^A-Z0-9]", "", short.upper())
    if len(short_clean) < 2 or len(short_clean) > 8:
        return False
    return short_clean in _name_initials_variants(long_name)


def _strip_trailing_abbrev_word(name: str, priors: list) -> str:
    """Kalau kata TERAKHIR di segmen gabungan ternyata singkatan dari segmen
    sebelumnya (mis. "PDBI DSLNG" -> "DSLNG" singkatan "DONGGI-SENORO LNG"),
    buang kata itu saja, sisanya tetap."""
    words = name.split()
    if len(words) >= 2 and any(_is_abbreviation_of(words[-1], p) for p in priors):
        return " ".join(words[:-1])
    return name


def _clean_insured_name(name: str) -> str:
    name = _normalize_spaces(name)
    name = _INSURED_PREFIX_RE.sub("", name)   # PT/CV di awal
    name = _INSURED_GELAR_RE.sub("", name)    # gelar/sapaan
    name = re.sub(r"\bPTE\.?\b", "", name, flags=re.IGNORECASE)  # kata "PTE" dihapus
    name = re.sub(r"\(([^()]*)\)", _handle_insured_paren, name)  # kurung seimbang
    name = name.replace("(", " ").replace(")", " ")  # sisa kurung yang tidak seimbang
    name = re.sub(r"/", " ", name)            # garis miring nyisa -> spasi (strip "-" DIBIARKAN, bagian nama resmi mis. "PERTA-SAMTAN")
    for _ in range(3):
        cleaned = _normalize_spaces(INSURED_SUFFIX_RE.sub("", name).strip().strip(","))
        if cleaned == name:
            break
        name = cleaned
    name = name.replace(".", "")  # semua karakter titik dihapus
    name = _normalize_spaces(name).strip()
    name = name.upper()  # insured selalu CAPS LOCK
    return name


def _cap_or_join(items: list) -> list:
    if len(items) > MAX_SPLIT_COLS:
        return [",".join(items)]
    return items


# ─────────────────────────────────────────────────────────────────────────────
# KAMUS POLA STANDAR TPI (dari Karakteristik_Polis_dan_Slip.xlsx sheet "TPI"
# + Bismillah_TPI_-_Eksplorasi_Cedant.docx)
# Dipakai sebagai lapisan penyempurnaan SETELAH cleaning utama -- bukan filter,
# hanya membantu merapikan spasi/simpul liar kalau hasil cleaning belum persis
# cocok pola standar.
# ─────────────────────────────────────────────────────────────────────────────

KNOWN_POLIS_PATTERNS = [
    re.compile(r"^\d{14}\s*/\s*\d{4}$"),          # 03250000007295 / 2025
    re.compile(r"^\d{14}$"),                       # 03250000009340
    re.compile(r"^[A-Za-z]{3}\d{7}$"),             # PVE1900115
    re.compile(r"^[A-Za-z]{3}\d{7}\s*/\s*\d{4}$"), # PVV2000087 / 2020
    re.compile(r"^\d{19,20}$"),                    # 1913102022120000372 (Heavy Equipment)
    re.compile(r"^\d{14}-\d{6}\s*/\s*\d{4}$"),     # 04250000000099-000174 / 2026 (Marine Cargo)
    re.compile(r"^\d{6}$"),                        # 140057
]

KNOWN_SLIP_PATTERNS = [
    re.compile(r"^\d{10}$"),                       # 5100486892
    re.compile(r"^[A-Za-z]{3}\d{7}$"),              # FVV2000148
    re.compile(r"^\d{17,19}$"),                     # 29133060324000000 / 891310202212000008
    re.compile(r"^\d{10}\s*/\s*PAYMENT NUMBER\s*\d{10}$", re.IGNORECASE),
    re.compile(r"^TIDAK ADA$", re.IGNORECASE),
    re.compile(r"^\d{18,22}[A-Za-z]{3}[A-Za-z0-9]{3}\d{10}[A-Za-z]\d{2}$"),
]


def _matches_known_pattern(value: str, patterns) -> bool:
    return any(p.match(value) for p in patterns)


def refine_with_known_pattern(value: str, patterns) -> str:
    """Kalau hasil cleaning belum cocok kamus pola standar TPI, coba beberapa
    perbaikan ringan (spasi, simbol liar) sebelum menyerah dan mengembalikan
    hasil cleaning apa adanya."""
    if not value:
        return value
    if _matches_known_pattern(value, patterns):
        return value

    candidates = [
        re.sub(r"\s+", "", value),
        re.sub(r"\s*/\s*", " / ", value).strip(),
        value.strip(" .-/"),
        re.sub(r"\.0+$", "", value),
        "0" + value if value.isdigit() else value,        # angka polos kehilangan 0 di depan
        re.sub(r"^(\d+)(\s*/\s*(?:19|20)\d{2})$", r"0\1\2", value),  # "angka / tahun" kehilangan 0 di depan
    ]
    for c in candidates:
        if _matches_known_pattern(c, patterns):
            return c

    return value


def _refine_list(items: list, patterns) -> list:
    return [refine_with_known_pattern(v, patterns) for v in items]


# ─────────────────────────────────────────────────────────────────────────────
# MESIN "REPETITION CHAIN" — untuk pola berulang seperti:
#   PFV2100285 + PFV2100284 + PFV2100286 + 287 + 283
#   5100376775+729+826 + P2
#   03240000001429 / 2023 + 1408 + 1419 + 1429
#   PUF2200041-045-046-038-042-044-043
#   FVF2200491-EVF2200120-FVF2200550
#   PVF2300768 / 2023        PVF2300771 / 2023   (spasi panjang)
# Aturan (per arahan terbaru):
#   - Buang tahun (semua gaya: " / 2020", "/2020", ".2020", "-2020" (bukan
#     bagian dari SUSPENSE...CEDANT))
#   - Buang P1/P2/P3/TBA yang MENEMPEL ke kode/angka; kalau BERDIRI SENDIRI
#     (utuh cuma "P1", "P1 + P2", atau "TBA" saja) dibiarkan apa adanya
#   - Buang narasi seperti "SLIP FOR NEXT 12 MONTHS", "TANPA INDUK SLIP"
#   - Kode pendek yang cuma angka & lebih pendek dari basenya -> direkonstruksi
#     pakai basenya (ganti N digit terakhir)
#   - Duplikat (hasil akhir sama persis) -> ditampilkan sekali saja
#   - Kalau hasil pecahannya > 5 -> JANGAN dipecah, tampilkan sebagai satu
#     string gabungan bentuk pendek (join pakai "-"), bukan direkonstruksi
# ─────────────────────────────────────────────────────────────────────────────

_MONTH_NAMES_RE = (
    r"JANUARI|FEBRUARI|SEPTEMBER|NOVEMBER|DESEMBER|OKTOBER|AGUSTUS|MARET"
    r"|JANUARY|FEBRUARY|SEPTEMBER|NOVEMBER|DECEMBER|OCTOBER|AUGUST|MARCH"
    r"|APRIL|JUNI|JULI|JUNE|JULY|MEI|MAY"
    r"|SEPT|AUGUS|AGUS|OKT|DES|JAN|FEB|MAR|APR|JUN|JUL|AUG|SEP|OCT|NOV|DEC"
)

_JUNK_PHRASE_RE = re.compile(
    r"""
    SLIP\s+FOR\s+NEXT\s+\d+\s+MONTHS
  | TANPA\s+INDUK\s+SLIP
  | TANPA\s+SLIP\s+INDUK
  | SLIP\s+MASIH\s+PAKAI\s+SLIP\s+SEMENTARA
  | PAKAI\s+SLIP\s+SEMENTARA
  | SUPEN\s+SLIP
  | SUSPEN\s+SLIP
  | PAYMENT\s+NUMBER\s*:?
  | BINDER\s+NO\.?
  | REN\.?\s*POL\.?\s*NO\.?
  | REPLACEMENT
  | FINAL\s+ADJUSTMENT
  | ACTUAL\s+PREMIUM\s+FOR
  | MINDEPPREM\.?\s*FOR
  | MINDEP\s+PREMI(?:UM)?
  | MINDEP\s+FOR
  | \bMINDEP\b
  | \bMIN\.?
  | \bPERIOD\b
  | \bFOR\b
  | DATE\s+\d{1,2}\.\d{1,2}\.\d{2,4}
  | (?:19|20)\d{2}\s*-\s*\d{2}\b
  | \.?\s*\+?\s*SLIP\s+END\s*\d*(?:\s*\+\s*\d+\.?)*
  | \.?\s*\+?\s*END\.?\s*\d*(?:\s*\+\s*\d+\.?)*
  | DE(?:KLARASI|CLARASI)\s*:?
  | DECL\.?\s*:?
  | INTER\s+ISLAND
  | \bCOL\.?\s*NUMBER\b
  | \bCREDIT\s+NOTE\b
  | \bNPWN\b
  | BIND\.?\s*NO\.?
  | \bPN\.?\s+(?=\S)
  | \bR\.?\s*PREMIUM\b
  | \bMENJADI\b
  | \bS\s*/\s*D\b
  | \bDEPOSIT\b.*?\bFOR\b
  | AND\s+AC\.?\s*PREMI\.?\s*FOR
  | \bFROM\b
  | \bLIHAT\s+PAGE\s*\d*\b
  | \bEX\s*:?
  | TANPA\s+ADA\s+SLIP\s+INDUK
  | \bIDR\b
  | \bUSD\b
  | \bCN\b\s*:?
  | \bSLIP\b
  | \bID\s*:
  | \bAND\b
  | \bPR\.
  | \bACTUAL\s+PREMIUM\b
  | \bEQ\b
  | \bNO\b\s*:?
    """
    + r"| (?:" + _MONTH_NAMES_RE + r")\.?\s*\d{0,4}",
    re.IGNORECASE | re.VERBOSE,
)

# Narasi yang sengaja DIBIARKAN apa adanya (bukan pola berulang/tahun)
_STANDALONE_NARRATIVE_RE = re.compile(r"SUSPENSE.*CEDANT", re.IGNORECASE)

# P1/P2/P3/TBA/VAR BERDIRI SENDIRI (boleh gabungan "P1 + P2" / "P1,P2") -> dibiarkan
_STANDALONE_P_RE = re.compile(
    r"^\s*(?:P[123]|TBA|VAR)(?:\s*[+,]\s*(?:P[123]|TBA|VAR))*\s*$", re.IGNORECASE
)

_YEAR_RE = re.compile(r"/\s*(?:19|20)\d{2}\b")
_YEAR_2DIGIT_RE = re.compile(r"/\s*\d{2}$")
# Pola "KODE.TAHUN" / ".TAHUN" (titik lalu tahun) di akhir -- HANYA kalau
# titik cuma muncul sekali di seluruh nilai (kode dengan banyak titik seperti
# "017.PAR.EQ.6.2025" atau "009.1050.201.2014.001670.00" TIDAK kena ini)
_DOT_YEAR_RE = re.compile(r"^(.*)\.(?:19|20)?\d{2,4}$")

# Suffix pendek angka murni cuma direkonstruksi kalau <= ini banyaknya digit
# (di atas ini dianggap nomor lengkap yang beda, bukan potongan/suffix)
_MAX_SUFFIX_LEN = 7

# Nilai literal yang eksplisit "biarkan apa adanya" (format nyeleneh yang
# tidak masuk pola umum manapun, tidak aman digeneralisasi lewat regex)
_LITERAL_WHITELIST = {
    "017.PAR.EQ.6.2025",
}


def _is_standalone_exception(val: str) -> bool:
    if val.strip() in _LITERAL_WHITELIST:
        return True
    return bool(_STANDALONE_P_RE.match(val)) or bool(_STANDALONE_NARRATIVE_RE.search(val))


def _strip_junk_and_year(text: str) -> str:
    text = _JUNK_PHRASE_RE.sub("", text)
    text = _YEAR_RE.sub("", text)
    return text


def _strip_attached_noise(text: str) -> str:
    # P1/P2/P3/TBA/VAR yang menempel ke kode/angka lain (bukan berdiri sendiri)
    text = re.sub(r"\bP[123]\b\s*/?\s*", "", text, flags=re.IGNORECASE)
    text = re.sub(r"\bTBA\b\s*/?\s*", "", text, flags=re.IGNORECASE)
    text = re.sub(r"\bVAR\b\s*/?\s*", "", text, flags=re.IGNORECASE)
    text = re.sub(r"^P-(?=\d)", "", text)  # "P-96102" -> "96102" (P tunggal di depan)
    return text


def _clean_token(tok: str) -> str:
    tok = tok.strip(" :;/")
    if not tok:
        return tok
    # Titik dihapus kalau: (a) tokennya angka murni, atau (b) tokennya kode
    # sederhana "HURUF...ANGKA.ANGKA" (mis. "FVF23.00087" -> titik cuma typo).
    # TAPI kalau hurufnya nyebar di beberapa tempat (mis. "017.PAR.EQ.6.2025"
    # -- ada blok huruf setelah angka juga), titik dipertahankan -> biarkan
    # apa adanya, itu format kode terstruktur, bukan typo.
    if not re.search(r"[A-Za-z]", tok):
        tok = tok.replace(".", "")
    elif re.match(r"^[A-Za-z]+[\d.]+$", tok):
        tok = tok.replace(".", "")
    return tok


def _classify_and_reconstruct(tokens: list) -> list:
    results = []
    current_base = None
    for tok in tokens:
        tok = _clean_token(tok)
        if not tok:
            continue
        digits_only = re.match(r"^\d+$", tok) is not None
        base_has_letter = bool(current_base and re.search(r"[A-Za-z]", current_base))
        if (
            digits_only
            and current_base is not None
            and len(tok) < len(current_base)
            and len(tok) <= _MAX_SUFFIX_LEN
            # kalau base-nya angka polos (tanpa huruf), cuma direkonstruksi
            # kalau base-nya cukup panjang (>=7 digit) -- base pendek/angka
            # polos terlalu ambigu buat diasumsikan pola perulangan
            and (base_has_letter or len(current_base) >= 7)
        ):
            reconstructed = current_base[: -len(tok)] + tok
            results.append(reconstructed)
        else:
            results.append(tok)
            current_base = tok
    return results


def _dedup_preserve_order(items: list) -> list:
    seen = set()
    deduped = []
    for r in items:
        if r not in seen:
            seen.add(r)
            deduped.append(r)
    return deduped


def expand_repetition_chain(val, patterns=None) -> list:
    if pd.isna(val):
        return []
    val = str(val).strip()
    if not val:
        return []

    if _is_standalone_exception(val):
        return [_normalize_spaces(val)]

    _had_narrative = bool(_JUNK_PHRASE_RE.search(val))

    cleaned_text = _strip_junk_and_year(val)
    cleaned_text = _strip_attached_noise(cleaned_text)
    # "&" sekarang jadi pemisah asli (bukan dihapus buta) -- dibiarkan, nanti
    # ikut split di raw_tokens. Sisa "&" dari narasi (OCT & NOV) otomatis
    # hilang sendiri karena jadi pemisah kosong setelah nama bulan dihapus.
    # tahun 2 digit di ujung mis. "PNF980025+26/97" -> "/97" dibuang
    cleaned_text = _YEAR_2DIGIT_RE.sub("", cleaned_text)
    # "KODE.TAHUN" (titik tunggal + tahun di akhir) -> tahun dibuang, HANYA
    # kalau titik cuma muncul sekali (kode dot-chain seperti "017.PAR.EQ..."
    # tetap dibiarkan apa adanya)
    if cleaned_text.count(".") == 1:
        m = _DOT_YEAR_RE.match(cleaned_text)
        if m:
            cleaned_text = m.group(1)

    # Dot-chain angka murni mis. "1400834.832.826.824" -- kalau semua
    # bagian setelah yang pertama sama panjang & lebih pendek (pola
    # base+suffix jelas), anggap titik sebagai pemisah repetisi. Kalau
    # tidak (mis. "009.1050.201.2014.001670.00", panjang tidak konsisten),
    # titik cuma dihapus & digabung jadi satu kode (ditangani _clean_token).
    if "." in cleaned_text and not re.search(r"[A-Za-z]", cleaned_text):
        _dot_parts = cleaned_text.split(".")
        if len(_dot_parts) > 1 and all(p.isdigit() for p in _dot_parts):
            _rest_lens = {len(p) for p in _dot_parts[1:]}
            if len(_rest_lens) == 1 and next(iter(_rest_lens)) < len(_dot_parts[0]):
                cleaned_text = "+".join(_dot_parts)
    # huruf tunggal nyempil di ujung (mis. ".../V") -> dibuang
    cleaned_text = re.sub(r"[/\-+,]\s*[A-Za-z]\s*$", "", cleaned_text)
    cleaned_text = cleaned_text.replace(":", "").replace(";", "").replace("(", "").replace(")", "")
    cleaned_text = cleaned_text.strip()
    if not cleaned_text:
        # semuanya cuma narasi/junk (mis. "END.1" berdiri sendiri) -> jangan
        # dihapus semua, biarkan nilai aslinya apa adanya
        return [_normalize_spaces(val)]

    # pecah jadi token mentah: spasi panjang, +, -, koma, garis miring, &
    # kode yang terpisah spasi dari angkanya (mis. "EUC 080030", "FUH 0800003")
    # disatukan dulu sebelum di-split, berlaku untuk semua prefix kode
    cleaned_text = re.sub(r"\b([A-Za-z]{2,5})\s+(\d{5,})\b", r"\1\2", cleaned_text)
    cleaned_text = cleaned_text.replace("*", "")  # tanda bintang liar dibuang

    raw_tokens = [t for t in re.split(r"\s+|[+\-,/&]", cleaned_text) if t.strip()]
    raw_tokens = [_clean_token(t) for t in raw_tokens]
    raw_tokens = [t for t in raw_tokens if t]
    if not raw_tokens:
        return [_normalize_spaces(val)]

    # Kalau SEMUA token cuma angka pendek tanpa huruf sama sekali (mis.
    # "105-106"), tidak ada "kode dasar" yang jelas untuk dijadikan acuan
    # breakdown -> terlalu ambigu, biarkan apa adanya (jangan dipecah)
    if len(raw_tokens) > 1 and all(
        re.match(r"^\d+$", t) and len(t) <= 4 for t in raw_tokens
    ):
        return [_normalize_spaces(val)]

    reconstructed = _classify_and_reconstruct(raw_tokens)

    # Samakan format (tambah 0 di depan dst) SEBELUM dedup, supaya duplikat
    # yang cuma beda karena kehilangan 0 di depan ikut kedeteksi sama
    if patterns:
        reconstructed = [refine_with_known_pattern(r, patterns) for r in reconstructed]

    reconstructed = _dedup_preserve_order(reconstructed)

    if len(reconstructed) == 1:
        return [_normalize_spaces(reconstructed[0])]

    if len(reconstructed) <= MAX_SPLIT_COLS:
        return [_normalize_spaces(r) for r in reconstructed]

    # > 5 hasil -> jangan dipecah, gabung bentuk pendek pakai "-"
    # > 5 hasil:
    # - kalau tadinya ada narasi yang dibuang -> ini kasus "kode + cerita
    #   panjang", cukup sisakan kode paling awal
    # - kalau TIDAK ada narasi (murni rantai angka/kode panjang) -> biarkan
    #   apa adanya, jangan dipecah
    if _had_narrative:
        return [_normalize_spaces(reconstructed[0])]
    return [_normalize_spaces(cleaned_text)]


# ─────────────────────────────────────────────────────────────────────────────
# CLEAN POLIS  (+ pengecualian khusus TPI untuk P1/P2)
# ─────────────────────────────────────────────────────────────────────────────

def clean_polis(val) -> list:
    if pd.isna(val):
        return []
    val = str(val).strip()
    if not val:
        return []

    # Pemicu mesin repetition-chain: ada tanda "+", spasi panjang, tahun,
    # P1/P2/P3/TBA/VAR yang menempel, atau ada "-" / "/" yang menyambungkan
    # lebih dari 1 kode -> proses lewat mesin baru
    _trigger = (
        "+" in val
        or re.search(r"\s{2,}", val)
        or _YEAR_RE.search(val)
        or re.search(r"\bP[123]\b", val, re.IGNORECASE)
        or re.search(r"\bTBA\b", val, re.IGNORECASE)
        or re.search(r"\bVAR\b", val, re.IGNORECASE)
        or _JUNK_PHRASE_RE.search(val)
        or "/" in val
        or "-" in val
        or "&" in val
        or "." in val
    )
    if _trigger:
        return expand_repetition_chain(val, KNOWN_POLIS_PATTERNS)

    if POLIS_EXCEPTION_RE.search(val):
        return [_normalize_spaces(val)]

    val_upper = val.upper().strip()
    if re.match(r"^VARIOUS\s*$", val_upper):
        return [val.strip()]
    if re.match(r"^TBA\s*$", val_upper):
        return [val.strip()]

    if "+" in val:
        blocks = re.split(r"\s{2,}", val.strip())
        left_parts = [p.strip() for p in blocks[0].split("+")
                      if p.strip() and not re.match(r"^TBA$", p.strip(), re.IGNORECASE)]
        right_blocks = blocks[1:]
        if left_parts:
            first = left_parts[0]
            rest_parts = left_parts[1:]
            any_long = any(len(r) >= 10 for r in rest_parts)
            all_suffix = rest_parts and all(re.match(r"^\d{1,6}$", r) for r in rest_parts)
            if all_suffix and len(first) >= 10 and not any_long:
                results = [first[:-len(s)] + s if len(first) > len(s) else first + s for s in rest_parts]
                for rb in right_blocks:
                    results += [t for t in rb.split() if _is_valid_polis_token(t.strip())]
                return _cap_or_join(results) if results else [first]
            all_tokens = _extract_polis_tokens(val)
            if all_tokens:
                return _cap_or_join(all_tokens)

    if re.search(r"\s{2,}", val):
        tokens = _extract_polis_tokens(val)
        if tokens:
            return _cap_or_join(tokens)

    val = re.sub(r"\bVARIOUS\b\s*", "", val, flags=re.IGNORECASE).strip()
    val = re.sub(r"\bVAR\b\s*", "", val, flags=re.IGNORECASE).strip()

    parts = re.split(r"\s*/\s*", val)
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

    _trigger = (
        "+" in val
        or re.search(r"\s{2,}", val)
        or _YEAR_RE.search(val)
        or re.search(r"\bP[123]\b", val, re.IGNORECASE)
        or re.search(r"\bTBA\b", val, re.IGNORECASE)
        or re.search(r"\bVAR\b", val, re.IGNORECASE)
        or _JUNK_PHRASE_RE.search(val)
        or "/" in val
        or "-" in val
        or "&" in val
        or "." in val
    )
    if _trigger:
        return expand_repetition_chain(val, KNOWN_SLIP_PATTERNS)

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

    cleaned = re.sub(r"\bVARIOUS\b\s*", "", val, flags=re.IGNORECASE)
    cleaned = re.sub(r"^\s*-\s*", "", cleaned).strip()
    cleaned = _normalize_spaces(cleaned)
    return [cleaned] if cleaned else []


# ─────────────────────────────────────────────────────────────────────────────
# CLEAN INSURED
# ─────────────────────────────────────────────────────────────────────────────

# PT/CV di AWAL nama (paling umum: "PT ANGKASA PURA")
_INSURED_PREFIX_RE = re.compile(r"^\s*(?:PT|CV)\.?\s+", re.IGNORECASE)

# Gelar / sapaan (Bapak, Ibu, Ny, Mr, Mrs, Ms, Dr, Ir, dst)
_INSURED_GELAR_RE = re.compile(
    r"""
    \b(?:BAPAK|IBU|SDRI?|NYONYA|NY|TUAN|MR|MRS|MS|
        DR|IR|PROF|HJ|H)\b\.?\s*
    """,
    re.IGNORECASE | re.VERBOSE,
)

# Pemisah antar-insured: koma, garis miring, QQ, atau "and/or"
_INSURED_SPLIT_RE = r"\bQQ\b|\bAND\s*/\s*OR\b|/|,|–|¿"


def _remove_polis_slip_from_insured(text: str, polis_ori, slip_ori) -> str:
    """Buang nomor polis/slip yang nyempil di teks insured (pakai nilai asli
    sebelum cleaning, sesuai contoh di dokumen eksplorasi TPI)."""
    for token in (polis_ori, slip_ori):
        if pd.notna(token):
            t = str(token).strip()
            if t and t != "-":
                text = text.replace(t, "")
    return text


def clean_insured(val, polis_ori=None, slip_ori=None) -> list:
    if pd.isna(val):
        return []
    val = str(val).strip()
    if not val:
        return []

    val = _remove_polis_slip_from_insured(val, polis_ori, slip_ori)

    # "NAMA,PT ..." / "NAMA, PT ..." adalah PT ditulis terbalik (bukan
    # pemisah ke perusahaan baru) -> gabung, jangan dipecah di titik itu
    val = re.sub(r",\s*(?:PT|CV)\.?\s+", " ", val, flags=re.IGNORECASE)
    val = re.sub(r",\s*KSO\.?\s*$", " KSO", val, flags=re.IGNORECASE)

    raw_parts = re.split(_INSURED_SPLIT_RE, val, flags=re.IGNORECASE)
    raw_parts = _merge_digit_start_parts(raw_parts)  # nama unit/armada digabung, bukan dipisah
    raw_parts = _merge_region_prefix_parts(raw_parts)

    cleaned = []
    for p in raw_parts:
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

    # buang segmen yang ternyata cuma singkatan dari segmen sebelumnya
    # (mis. "TPPI" setelah "TRANS PACIFIC PETROCHEMICAL INDOTAMA")
    final = []
    for name in cleaned:
        name = _strip_trailing_abbrev_word(name, final)
        if any(_is_abbreviation_of(name, prior) for prior in final):
            continue
        final.append(name)

    return _cap_or_join(final)


# ─────────────────────────────────────────────────────────────────────────────
# PROSES UTAMA
# ─────────────────────────────────────────────────────────────────────────────

def _insert_clean_columns(df, all_lists, prefix, max_cols):
    added = []
    for i in range(1, max_cols + 1):
        col_name = f"{prefix}_{i}"
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

    for col in [POLIS_COL, SLIP_COL, INSURED_COL]:
        if col not in df.columns:
            print(f"\n[ERROR] Kolom '{col}' tidak ditemukan di file!")
            print(f"        Kolom tersedia: {list(df.columns)}")
            return

    df.rename(columns={
        POLIS_COL:   "FAC_POLICY_NO_ORI",
        SLIP_COL:    "FAC_SLIP_ORI",
        INSURED_COL: "FAC_INSURED_ORI",
    }, inplace=True)

    print("[3/5] Menjalankan proses cleaning ...")

    all_clean_polis, all_clean_slip, all_clean_ins = [], [], []
    max_polis = max_slip = max_ins = 1

    for _, row in df.iterrows():
        c_polis = _refine_list(clean_polis(row.get("FAC_POLICY_NO_ORI", "")), KNOWN_POLIS_PATTERNS)
        c_slip  = _refine_list(clean_slip(row.get("FAC_SLIP_ORI", "")), KNOWN_SLIP_PATTERNS)
        c_ins   = clean_insured(row.get("FAC_INSURED_ORI", ""), row.get("FAC_POLICY_NO_ORI", ""), row.get("FAC_SLIP_ORI", ""))

        max_polis = max(max_polis, len(c_polis))
        max_slip  = max(max_slip, len(c_slip))
        max_ins   = max(max_ins, len(c_ins))

        all_clean_polis.append(c_polis)
        all_clean_slip.append(c_slip)
        all_clean_ins.append(c_ins)

    print("[4/5] Menyusun kolom output ...")

    new_columns = []
    for col in df.columns:
        new_columns.append(col)
        if col == "FAC_POLICY_NO_ORI":
            new_columns += _insert_clean_columns(df, all_clean_polis, "FAC_POLICY_NO", max_polis)
        elif col == "FAC_SLIP_ORI":
            new_columns += _insert_clean_columns(df, all_clean_slip, "FAC_SLIP", max_slip)
        elif col == "FAC_INSURED_ORI":
            new_columns += _insert_clean_columns(df, all_clean_ins, "FAC_INSURED", max_ins)

    df = df[new_columns]

    os.makedirs(os.path.dirname(output_file), exist_ok=True)
    print(f"[5/5] Menyimpan hasil ke: {output_file} ...")
    df.to_excel(output_file, index=False)

    print(f"\n{'=' * 55}")
    print("  [OK] Selesai!")
    print(f"  Total baris output  : {len(df):,}")
    print(f"  Total kolom output  : {len(df.columns)}")
    print(f"  File disimpan di    : {output_file}")
    print(f"{'=' * 55}")


if __name__ == "__main__":
    process_data(INPUT_FILE, OUTPUT_FILE)