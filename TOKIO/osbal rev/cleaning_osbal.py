import os
import re

import pandas as pd


INPUT_FILE = os.path.join("input", "inosbal_facul_rev1.xlsx")
OUTPUT_FILE = os.path.join("output", "tokio_output_inosbalfacul.xlsx")

CEDANT_COL   = "CCOS_COMP_NAME"
CEDANT_VALUE = "PT.ASURANSI TOKIO MARINE INDONESIA"

POLIS_COL   = "FAC_POLICY_NO"
SLIP_COL    = "FAC_SLIP"
INSURED_COL = "FAC_INSURED"

MAX_SPLIT_COLS = 5

# Polis: kata/prefix yang menyebabkan nilai dibiarkan apa adanya
POLIS_EXCEPTION_RE = re.compile(
    r"""
    MOP\s*MARINE
  | (?:LINE\s*SLIP|LINESLIP)
  | \b(?:P1|P2|P3|P73)\s*CANCEL
  | \b(?:P1|P2|P3|P73)\b
  | \bCANCEL\b
  | PENYELESAIAN(?:\s+SUSPENSE)?
  | HUTANG
  | UTANG
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
  | \bP1\b
  | \bP2\b
  | \bP3\b
  | PENYELESAIAN(?:\s+SUSPENSE)?
  | HUTANG
  | UTANG
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
    "MR", "MR.", "MRS", "MRS.", "MS", "MS.",
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
# HELPERS
# ─────────────────────────────────────────────────────────────────────────────


def _normalize_spaces(text: str) -> str:
    return re.sub(r"\s{2,}", " ", text).strip()


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

    # Exception: jangan di-clean
    if re.fullmatch(r"AEON\s+MALL\s*\(TENANTS\)", name, re.IGNORECASE):
        return name.upper()

    for _ in range(3):
        cleaned = _normalize_spaces(INSURED_SUFFIX_RE.sub("", name).strip().strip(","))
        if cleaned == name:
            break
        name = cleaned

    # Hapus token PT / QQ dimanapun posisinya (prefix, tengah, dsb.)
    name = INSURED_PT_QQ_RE.sub(" ", name)

    # ==========================
    # EXCEPTION
    # AEON MALL (TENANTS)
    # biarkan apa adanya
    # ==========================
    if re.fullmatch(r"\s*AEON\s+MALL\s+\(TENANTS\)\s*", name, flags=re.IGNORECASE):
        return "AEON MALL (TENANTS)"

    # Khusus (PERSERO) dihapus
    name = re.sub(
        r"\(\s*PERSERO\s*\)",
        "",
        name,
        flags=re.IGNORECASE,
    )

    # Kurung selain (PERSERO) hanya dihilangkan tandanya,
    # isinya tetap dipertahankan
    name = re.sub(r"[()]", " ", name)

    # Hapus tanda kurung saja, isi tetap dipertahankan
    name = name.replace("(", " ")
    name = name.replace(")", " ")

    name = _normalize_spaces(name)
    # Bersihkan sisa tanda baca (titik/koma/strip) di ujung
    name = re.sub(r"^[\s.,\-]+|[\s.,\-]+$", "", name)

    return name


def _cap_or_join(items: list) -> list:
    """Jika jumlah item melebihi MAX_SPLIT_COLS, gabungkan dengan koma."""
    if len(items) > MAX_SPLIT_COLS:
        return [",".join(items)]
    return items

# =============================================================================
# HELPER POLIS TOKIO
# =============================================================================

TOKIO_PREFIX_RE = re.compile(
    r"""
    (?:[A-Z]{2,5}/[A-Z]{2,5}/\d{2}-)
    |
    (?:[A-Z]{2,10}\d{2})
    """,
    re.IGNORECASE | re.VERBOSE,
)

POLIS_CODE_RE = re.compile(
    r"[A-Z]\d{7,8}",
    re.IGNORECASE,
)


def _remove_tokio_prefix(text: str) -> str:
    """
    Contoh:
    TMD/FIAR/24-F0035890
        -> F0035890

    TMIFIAR23F5004951
        -> F5004951

    BTDFIAR24F0039468
        -> F0039468
    """

    text = TOKIO_PREFIX_RE.sub("", text)
    return text.strip()


def _expand_suffix(base: str, suffix: str) -> str:
    """
    Contoh

    base   = F5004951
    suffix = 4959
    hasil  = F5004959

    base   = F5004951
    suffix = 957
    hasil  = F5004957

    base   = F0046054
    suffix = 5089
    hasil  = F0045089
    """

    m = re.match(r"([A-Z])(\d+)$", base)

    if not m:
        return suffix

    head = m.group(1)
    num = m.group(2)

    if len(suffix) >= len(num):
        return head + suffix

    return head + num[:-len(suffix)] + suffix


def _extract_tokio_codes(text: str):
    """
    Ambil semua kode polis format:

    F0035890
    M5569642
    A0944430
    E0005714
    """

    return POLIS_CODE_RE.findall(text.upper())


def _unique(items):
    out = []

    for x in items:
        if x and x not in out:
            out.append(x)

    return out


# ─────────────────────────────────────────────────────────────────────────────
# CLEAN POLIS
# ─────────────────────────────────────────────────────────────────────────────
# ─────────────────────────────────────────────────────────────────────────────
# CLEAN POLIS TOKIO
# ─────────────────────────────────────────────────────────────────────────────

def clean_polis_raw(val) -> list:
    if pd.isna(val):
        return []

    val = str(val).strip().upper()

    if not val:
        return []

    # ============================================================
    # NORMALISASI
    # ============================================================

    val = re.sub(r"\bVARIOUS\b", "", val)

    # hapus VAR hanya kalau di belakang
    val = re.sub(r"\s*/\s*VAR$", "", val)

    # TBA tetap kalau berdiri sendiri
    if val != "TBA":
        val = re.sub(r"\bTBA\b", "", val)

    val = val.replace("&", "+")
    val = _normalize_spaces(val)

    # Hapus prefix tahun "24." sebelum kode polis
    # Contoh:
    # 24.F5023454 + F5023455
    # -> F5023454 + F5023455
    val = re.sub(
        r"^24\.(?=[A-Z]\d{7,8})",
        "",
        val,
        flags=re.IGNORECASE
    )

    # hapus titik
    val = val.replace(".", "")

    # hapus suffix polis belakang
    # contoh:
    # TMD/FEAQ/24-F5040057-005-02
    # menjadi:
    # TMD/FEAQ/24-F5040057
    val = re.sub(r"-\d{3}-\d{2}$", "", val)
    # ============================================================
    # PENYELESAIAN / HUTANG PIUTANG
    # ============================================================

    if POLIS_EXCEPTION_RE.search(val):
        return [_normalize_spaces(val)]

    # ============================================================
    # OPEN COVER
    # ============================================================

    if re.search(
        r"(JANUARI|FEBRUARI|MARET|APRIL|MEI|JUNI|JULI|AGUSTUS|SEPTEMBER|OKTOBER|NOVEMBER|DESEMBER)\s+\d{4}",
        val,
        flags=re.IGNORECASE,
    ):
        return [_normalize_spaces(val)]

    # ============================================================
    # PREFIX TOKIO
    #
    # TMD/AMBL/25-A0952369-000-00
    # -> TMD/AMBL/25-A0952369
    #
    # TMI/FEAQ/23-F5041279-000-01
    # -> TMI/FEAQ/23-F5041279
    # ============================================================

    m = re.fullmatch(
        r"([A-Z]{3}/[A-Z]{4}/\d{2}-[A-Z]\d{7,8})(?:-000-\d+)?",
        val,
        flags=re.IGNORECASE,
    )

    if m:
        return [m.group(1).upper()]

    # ============================================================
    # TANPA PREFIX
    #
    # A0944430-000-03
    # -> A0944430
    #
    # F5041279-000-01
    # -> F5041279
    # ============================================================

    m = re.fullmatch(
        r"([A-Z]\d{7,8})(?:-000-\d+)?",
        val,
        flags=re.IGNORECASE,
    )

    if m:
        return [m.group(1).upper()]

    # ============================================================
    # POLA S/D
    # ============================================================

    if "S/D" in val:
        hasil = re.findall(r"[A-Z]\d{7,8}", val)
        if hasil:
            return _cap_or_join(_unique(hasil))

    # ============================================================
    # HAPUS SUFFIX -001 / -002 / dst
    #
    # TMD/FIAR/23-F5021295-001
    # -> TMD/FIAR/23-F5021295
    #
    # TMD/FEAQ/23-F5021296-001
    # -> TMD/FEAQ/23-F5021296
    # ============================================================

    val = re.sub(
        r"([A-Z]\d{7,8})-\d{3}(?=\s|\+|$)",
        r"\1",
        val,
        flags=re.IGNORECASE,
    )

    # ============================================================
    # PISAHKAN POLIS YANG KETEMPEL SETELAH SPASI
    #
    # F5021296 TMD/FEAQ/23-F5006567
    # -> F5021296+TMD/FEAQ/23-F5006567
    # ============================================================

    val = re.sub(
        r"([A-Z]\d{7,8})\s+(TMD/|TMI/)",
        r"\1+\2",
        val,
        flags=re.IGNORECASE,
    )
    # ============================================================
    # POLA +
    # ============================================================
    if "+" in val:

        parts = [p.strip() for p in val.split("+") if p.strip()]

        results = []

        tokio_prefix = None
        tokio_company = None

        for p in parts:

            # ====================================================
            # SIMPAN PREFIX COMPANY TOKIO
            #
            # TMD/FIAR/23F0043653
            # -> simpan TMD/
            #
            # supaya:
            # FEAQ/23F0043656
            # -> TMD/FEAQ/23F0043656
            # ====================================================

            m_company = re.match(
                r"^(TMD|TMI)/",
                p,
                flags=re.IGNORECASE,
            )

            if m_company:
                tokio_company = m_company.group(1).upper() + "/"

            # ====================================================
            # PREFIX TOKIO TERPOTONG
            #
            # TMD/FIAR/23F0043653+FEAQ/23F0043656
            #
            # ->
            # TMD/FIAR/23F0043653
            # TMD/FEAQ/23F0043656
            # ====================================================

            m = re.fullmatch(
                r"([A-Z]{4}/\d{2}[A-Z]\d{7,8})",
                p,
                flags=re.IGNORECASE,
            )

            if m and tokio_company:
                results.append(
                    tokio_company + m.group(1).upper()
                )
                continue

            # ====================================================
            # PREFIX + BEBERAPA KODE
            #
            # TMD/FEAQ/23-F0000263-F0003746
            # ->
            # TMD/FEAQ/23-F0000263
            # TMD/FEAQ/23-F0003746
            # ====================================================

            m = re.fullmatch(
                r"^([A-Z]{3}/[A-Z]{4}/\d{2}-)"
                r"([A-Z]\d{7,8})"
                r"((?:-[A-Z]\d{7,8})+)$",
                p,
                flags=re.IGNORECASE,
            )

            if m:
                tokio_prefix = m.group(1).upper()

                results.append(
                    tokio_prefix + m.group(2).upper()
                )

                suffixes = re.findall(
                    r"[A-Z]\d{7,8}",
                    m.group(3),
                    flags=re.IGNORECASE,
                )

                for suffix in suffixes:
                    results.append(
                        tokio_prefix + suffix.upper()
                    )

                continue

            # ====================================================
            # PREFIX TOKIO TERPOTONG
            #
            # TMD/FIAR/23F0043653+FEAQ/23F0043656
            #
            # ->
            # TMD/FIAR/23F0043653
            # TMD/FEAQ/23F0043656
            #
            # Prefix TMD diwariskan ke polis berikutnya
            # ====================================================

            m = re.fullmatch(
                r"([A-Z]{4}/\d{2}[A-Z]\d{7,8})",
                p,
                flags=re.IGNORECASE,
            )

            if m and tokio_prefix:
                results.append(
                    "TMD/" + m.group(1).upper()
                )
                continue
            
            # bagian lengkap
            m = re.match(
                r"^([A-Z]{3}/[A-Z]{4}/\d{2}-)([A-Z]\d{7,8})$",
                p,
                flags=re.IGNORECASE,
            )

            if m:
                tokio_prefix = m.group(1).upper()
                results.append(tokio_prefix + m.group(2).upper())
                continue

                        # format titik
            # MDD.FPAR.23.F5031581 + F5031582

            m = re.match(
                r"^([A-Z]{7,8}\d{2})([A-Z]\d{7,8})$",
                p,
                flags=re.IGNORECASE,
            )

            if m:
                tokio_prefix = m.group(1).upper()
                results.append(tokio_prefix + m.group(2).upper())
                continue
                
            # hanya kode
            m = re.fullmatch(
                r"[A-Z]\d{7,8}",
                p,
                flags=re.IGNORECASE,
            )

            if m:

                code = m.group(0).upper()

                if tokio_prefix:
                    results.append(tokio_prefix + code)
                else:
                    results.append(code)

                continue

            # fallback
            if p.upper() not in ["VAR", "VARIOUS", "TBA"]:
                results.append(p.upper())

        return _cap_or_join(_unique(results))

    # ============================================================
    # POLA PREFIX TOKIO DENGAN - ANTAR POLIS
    #
    # TMD/FEAQ/23-F0000263-F0003746
    # ->
    # TMD/FEAQ/23-F0000263
    # TMD/FEAQ/23-F0003746
    #
    # Prefix tetap diwariskan ke polis berikutnya
    # ============================================================

    m = re.fullmatch(
        r"([A-Z]{3}/[A-Z]{4}/\d{2}-)"
        r"([A-Z]\d{7,8})"
        r"((?:-[A-Z]\d{7,8})+)",
        val,
        flags=re.IGNORECASE,
    )

    if m:
        tokio_prefix = m.group(1).upper()
        first_code = m.group(2).upper()

        results = [tokio_prefix + first_code]

        # Ambil kode setelah tanda -
        suffixes = re.findall(
            r"[A-Z]\d{7,8}",
            m.group(3),
            flags=re.IGNORECASE,
        )

        for suffix in suffixes:
            results.append(
                tokio_prefix + suffix.upper()
            )

        return _cap_or_join(_unique(results))

    # ============================================================
    # POLA TOKIO COMPACT:
    #
    # TMDFIAR21F5017883-7884-TMDEMCB21E0005602
    #
    # ->
    # TMDFIAR21F5017883
    # TMDFIAR21F5017884
    # TMDEMCB21E0005602
    #
    # Kode 4 digit setelah "-" mewarisi prefix
    # dari kode sebelumnya.
    # ============================================================

    m = re.fullmatch(
        r"((?:TMD|TMI)[A-Z]{4}\d{2})"
        r"([A-Z]\d{7,8})"
        r"-(\d{4})"
        r"-((?:TMD|TMI)[A-Z]{4}\d{2}[A-Z]\d{7,8})",
        val,
        flags=re.IGNORECASE,
    )

    if m:
        tokio_prefix = m.group(1).upper()
        first_code = m.group(2).upper()
        second_suffix = m.group(3)
        third_policy = m.group(4).upper()

        # F5017883 -> prefix F501
        code_prefix = first_code[:-4]

        return _cap_or_join(_unique([
            tokio_prefix + first_code,
            tokio_prefix + code_prefix + second_suffix,
            third_policy,
        ]))

    # ============================================================
    # POLA TOKIO DENGAN -
    # ============================================================

    # ------------------------------------------------------------
    # 1. TMDFIAR21F5026865-F5026866
    # -> TMDFIAR21F5026865
    # -> TMDFIAR21F5026866
    # ------------------------------------------------------------

    m = re.fullmatch(
        r"((?:TMD|TMI)[A-Z]{4}\d{2})([A-Z]\d{7,8})-([A-Z]\d{7,8})",
        val,
        flags=re.IGNORECASE
    )

    if m:
        tokio_prefix = m.group(1).upper()

        return [
            tokio_prefix + m.group(2).upper(),
            tokio_prefix + m.group(3).upper()
        ]


    # ------------------------------------------------------------
    # 2. PREFIX TOKIO DENGAN SUFFIX 3 DIGIT / 4 DIGIT
    #
    # Contoh 4 digit:
    # TMDFIAR21F5023320-3324-3318
    # ->
    # TMDFIAR21F5023320
    # TMDFIAR21F5023324
    # TMDFIAR21F5023318
    #
    # Contoh 3 digit:
    # TMDFIAR22F5028544-545-757-758
    # ->
    # TMDFIAR22F5028544
    # TMDFIAR22F5028545
    # TMDFIAR22F5028757
    # TMDFIAR22F5028758
    #
    # Kalau suffix 3 digit:
    # ambil prefix sampai sebelum 3 digit terakhir
    #
    # Kalau suffix 4 digit:
    # ambil prefix sampai sebelum 4 digit terakhir
    # ------------------------------------------------------------

    m = re.fullmatch(
        r"((?:TMD|TMI)[A-Z]{4}\d{2})"
        r"([A-Z]\d{7,8})"
        r"((?:-\d{3,4})+)",
        val,
        flags=re.IGNORECASE
    )

    if m:
        tokio_prefix = m.group(1).upper()
        first_code = m.group(2).upper()

        results = [
            tokio_prefix + first_code
        ]

        # Ambil semua suffix 3 atau 4 digit
        suffixes = re.findall(
            r"\d{3,4}",
            m.group(3)
        )

        for suffix in suffixes:

            if len(suffix) == 3:
                # Contoh:
                # F5028544-545
                # F50285 + 545
                # = F5028545
                code_prefix = first_code[:-3]

            else:
                # Contoh:
                # F5023320-3324
                # F502 + 3324
                # = F5023324
                code_prefix = first_code[:-4]

            results.append(
                tokio_prefix + code_prefix + suffix
            )

        return _cap_or_join(_unique(results))

    # ------------------------------------------------------------
    # 3. POLA CAMPURAN / TIDAK KONSISTEN
    #
    # TMDFIAR21F0039255-688-312-784-726-888-552295888566
    # -> biarkan apa adanya
    # ------------------------------------------------------------

    if re.fullmatch(
        r"(?:TMD|TMI)[A-Z]{4}\d{2}[A-Z]\d{7,8}(?:-\d+)+",
        val,
        flags=re.IGNORECASE
    ):
        return [val]

    # ============================================================
    # MULTI POLIS SPASI / -
    # ============================================================

    hasil = re.findall(
        r"(?:TMD|TMI)?[A-Z]{3,5}\d{2}[A-Z]\d{7,8}",
        val,
        flags=re.IGNORECASE
    )

    if len(hasil) > 1:
        return _cap_or_join(_unique([x.upper() for x in hasil]))

    # ============================================================
    # MULTI POLIS DENGAN /
    # A0672872 / A0781730 / A0896105 / A0897232
    # ============================================================

    hasil = re.findall(
        r"[A-Z]\d{7,8}",
        val,
        flags=re.IGNORECASE
    )

    if len(hasil) > 1:
        return _cap_or_join(_unique([x.upper() for x in hasil]))
    
    # ============================================================
    # DEFAULT
    # ============================================================

    # ============================================================
    # TOKIO FULL CODE
    # TMDFIAR23SF0044266
    # ============================================================

    if re.match(r"^(TMD|TMI)[A-Z0-9]+$", val):
        return [val.upper()]


    m = re.search(
        r"[A-Z]\d{7,8}",
        val
    )

    if m:
        return [m.group(0).upper()]

    if m:
        return [m.group(0).upper()]

    m = re.search(r"[A-Z]\d{7,8}", val)

    if m:
        return [m.group(0).upper()]

    return []

# ============================================================
# FINAL CLEAN POLIS
# HAPUS SEMUA PEMISAH
# ============================================================

def clean_polis(val) -> list:
    # TAMBAHAN KHUSUS:
    # titik dihapus, tanda + tetap dipertahankan
    if pd.notna(val):
        val_str = str(val).strip().upper()

        if re.fullmatch(
            r"TMD\.FIAR\.22\.F\d{7,8}(?:\+\d{3})+",
            val_str
        ):
            return [val_str.replace(".", "")]

    # KODE ASLI - JANGAN DIUBAH
    hasil = clean_polis_raw(val)

    hasil_clean = []

    for x in hasil:
        x = re.sub(r"[^A-Z0-9]", "", str(x).upper())

        if x:
            hasil_clean.append(x)

    return hasil_clean

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

    # --------------------------------------------------
    # Hapus "/ AC NO : xxxxx"
    # --------------------------------------------------
    val = re.sub(
        r"([A-Z]{3}/[A-Z]{4}/\d{2}-[A-Z]\d{7,8})-\d{3}-\d{2}",
        r"\1",
        val,
        flags=re.IGNORECASE,
    )

    val = re.sub(
        r"([A-Z]\d{7,8})-\d{3}-\d{2}",
        r"\1",
        val,
        flags=re.IGNORECASE,
    )

    # --------------------------------------------------
    # 10017359 S/D 10017359 ...
    # ambil nomor pertama
    # --------------------------------------------------
    m = re.match(
        r"^(\d+)\s+S/D\s+\d+",
        val,
        flags=re.IGNORECASE,
    )

    if m:
        return [m.group(1)]

    # --------------------------------------------------
    # 10019492/00187246/TMD/AMBL/23-A0944430
    # ambil kode Tokio
    # --------------------------------------------------
    m = re.search(
        r"(TM[DI]/[A-Z]{4}/\d{2}-[A-Z]\d{7,8})",
        val,
        flags=re.IGNORECASE,
    )

    if m:
        return [m.group(1)]

    # Open cover FACILITY → biarkan apa adanya
    if "FACILITY" in val:
        return [val]

    # ============================================================
    # PENYELESAIAN SUSPENSE / HUTANG PIUTANG
    # Dibiarkan apa adanya
    # ============================================================

    if POLIS_EXCEPTION_RE.search(val):
        return [_normalize_spaces(val)]

    # ============================================================
    # 1. HAPUS BARIS YANG HANYA BERISI KETERANGAN
    #
    # Contoh:
    # P1 + P2
    # P1 CANCEL
    # NOVEMBER 2024
    # JANUARI 2024 - USD
    # ============================================================

    # ============================================================
    # 1. HANYA BUANG KETERANGAN YANG MEMANG TIDAK DIPAKAI
    #
    # P1 / P2 tetap dipertahankan
    # Open cover seperti:
    # FEBRUARI 2025 - AUD
    # DESEMBER 2024 - IDR
    # juga tetap dipertahankan
    # ============================================================

    if re.fullmatch(
        r"P\d+(?:\s*\+\s*P\d+)*(\s+CANCEL)?",
        val,
        flags=re.IGNORECASE,
    ):
        return [_normalize_spaces(val)]

    # ============================================================
    # 2. HAPUS "/ END"
    # ============================================================

    val = re.sub(
        r"\s*/\s*END\b",
        "",
        val,
        flags=re.IGNORECASE,
    ).strip()

    # ============================================================
    # 2A. HAPUS KETERANGAN VAR ACCOUNT / FACILITY
    #
    # JULI 2022 - AUD/VAR ACCOUNT
    # -> JULI 2022 - AUD
    #
    # JUNI 2024 - IDR / FACILITY AUTOMOBILE
    # -> JUNI 2024 - IDR
    # ============================================================

    val = re.sub(
        r"\s*/\s*(VAR\s+ACCOUNT|FACILITY\s+AUTOMOBILE|FACILITY)\b",
        "",
        val,
        flags=re.IGNORECASE,
    )

    val = re.sub(
        r"\bVARIOUS\b",
        "",
        val,
        flags=re.IGNORECASE
    )

    val = re.sub(
        r"\bVAR\b",
        "",
        val,
        flags=re.IGNORECASE
    )

    val = _normalize_spaces(val)

    # ============================================================
    # 2B. HAPUS KETERANGAN OPEN COVER
    #
    # Contoh:
    # JULI 2022 - AUD/VAR ACCOUNT
    # -> JULI 2022 - AUD
    #
    # JUNI 2024 - IDR / FACILITY AUTOMOBILE
    # -> JUNI 2024 - IDR
    #
    # OKTOBER 2024 - EUR / FACILITY
    # -> OKTOBER 2024 - EUR
    # ============================================================

    val = re.sub(
        r"/\s*VAR\s+ACCOUNT\b",
        "",
        val,
        flags=re.IGNORECASE,
    )

    val = re.sub(
        r"/\s*FACILITY\s+AUTOMOBILE\b",
        "",
        val,
        flags=re.IGNORECASE,
    )

    val = re.sub(
        r"/\s*FACILITY\b",
        "",
        val,
        flags=re.IGNORECASE,
    )

    val = _normalize_spaces(val)

    # print("DEBUG SLIP:", val)

    # ============================================================
    # DUA SLIP DALAM SATU CELL (dipisahkan spasi panjang)
    # ============================================================

    # ============================================================
    # DUA SLIP DALAM SATU CELL (dipisahkan spasi panjang)
    #
    # Contoh:
    # 000.6005.202.2023.001021 + 001587 + 315
    #                                     000.6005.202.2023.001021.00
    #
    # atau
    #
    # 000.6005.201.2024.003281.00+VAR
    #                                 0006005202202400230200
    # ============================================================

    parts = re.split(r"\s{2,}", val)
    parts = [p.strip() for p in parts if p.strip()]

    if len(parts) >= 2:

        results = []

        for p in parts:
            results.append(_normalize_spaces(p))

        return _cap_or_join(results)
   
    # ============================================================
    # 3. EXCEPTION
    # ============================================================

    if SLIP_EXCEPTION_RE.search(val):
        return [_normalize_spaces(val)]

    # ============================================================
    # 3A. EXCEPTION KHUSUS
    # Jangan breakdown pola seperti:
    # 6005.201.2020.002029-1469-2021-1461
    # ============================================================

    if re.fullmatch(
        r"\d+(?:\.\d+)+(?:-\d+){3,}",
        val
    ):
        return [_normalize_spaces(val)]

    # ============================================================
    # 4. POLA +
    # ============================================================
    if "+" in val:

        parts = [
            p.strip()
            for p in re.split(r"\s*\+\s*", val)
            if p.strip()
        ]

        # Buang P1, P2, dst
        parts = [
            p
            for p in parts
            if not re.fullmatch(r"P\d+", p, flags=re.IGNORECASE)
        ]
        

        if len(parts) > 1:

            first = parts[0]
            rest = parts[1:]

            # ----------------------------------------------------
            # Semua bagian nomor slip lengkap
            #
            # Contoh:
            # 4001092401819 + 4501032400015
            #
            # hasil:
            # 4001092401819
            # 4501032400015
            # ----------------------------------------------------

            if all(
                re.fullmatch(r"\d+", p)
                for p in parts
            ):
                return _cap_or_join(parts)

            # ----------------------------------------------------
            # Breakdown suffix
            # Contoh:
            # 000715 + 000440 + 000114
            # ----------------------------------------------------

            rest_is_digit = all(
                re.fullmatch(r"\d+", r)
                for r in rest
            )

            suffix_lengths = (
                {len(r) for r in rest}
                if rest_is_digit
                else set()
            )

            if rest_is_digit and len(suffix_lengths) == 1:

                suffix_len = suffix_lengths.pop()

                m = re.match(
                    rf"^(.*?)(\d{{{suffix_len}}})$",
                    first,
                )

                if m:

                    prefix = m.group(1)

                    results = [first]

                    for s in rest:

                        # ubah panjang suffix menjadi sama dengan nomor pertama
                        s = s.zfill(suffix_len)

                        results.append(prefix + s)

                    return _cap_or_join(results)
            # ----------------------------------------------------
            # Semua bagian sudah berupa nomor slip lengkap
            # ----------------------------------------------------

            if all(
                re.match(r"^\d+(?:\.\d+)+$", p)
                for p in parts
            ):
                return _cap_or_join(parts)

            return [_normalize_spaces(val)]

    # ============================================================
    # 5. BREAKDOWN DENGAN "-"
    #
    # Contoh 1:
    # 6005.201.2020.00741-549
    # ->
    # 6005.201.2020.00741
    # 6005.201.2020.00549
    #
    # Contoh 2:
    # 6005.202.2020.00139490898578-0018605642397146
    # ->
    # 6005.202.2020.00139490898578
    # 0018605642397146
    # ============================================================

    m = re.match(
        r"^(.*?)(\d+)-(\d+)$",
        val
    )

    if m:

        prefix = m.group(1)
        first = m.group(2)
        second = m.group(3)

        results = [prefix + first]

        # Jika nomor kedua lebih pendek,
        # anggap masih satu kelompok nomor pertama
        if len(second) < len(first):

            second = second.zfill(len(first))
            results.append(prefix + second)

        # Jika panjang sama / lebih panjang,
        # simpan apa adanya tanpa prefix
        else:
            results.append(second)

        return _cap_or_join(results)

    # ============================================================
    # 5A. DUA NOMOR SLIP DALAM SATU CELL
    #
    # Contoh:
    # 000.6005.201.2024.003281.00      0006005202202400230200
    #
    # Hasil:
    # 000.6005.201.2024.003281.00
    # 0006005202202400230200
    # ============================================================

    # slips = re.findall(
    #     r"(?:\d+(?:\.\d+)+|\d{12,})",
    #     val
    # )

    # if len(slips) >= 2:
    #     return _cap_or_join(slips)

    # ============================================================
    # 5A. DUA NOMOR SLIP DALAM SATU CELL
    # DIPISAHKAN SPASI PANJANG
    #
    # Contoh:
    # 6005.201.2021.000519                  000.6005.202.2021.000773.00
    #
    # Hasil:
    # 6005.201.2021.000519
    # 000.6005.202.2021.000773.00
    # ============================================================

    slips = re.findall(
        r"\d+(?:\.\d+)+",
        val
    )

    if len(slips) >= 2:
        return _cap_or_join(slips)

    # ============================================================
    # 6. AMBIL NOMOR SLIP UTAMA
    #
    # Contoh:
    # 000.6005.101.2023.002208.00 / MARET 2023 - EUR
    # ->
    # 000.6005.101.2023.002208.00
    #
    # 000.6005.101.2026.000457.00 - VAR /
    # REALISASI JANUARI 2026 - CNY
    # ->
    # 000.6005.101.2026.000457.00
    # ============================================================

    # ============================================================
    # OPEN COVER / KETERANGAN BISNIS
    #
    # Contoh:
    # FEBRUARI 2025 - AUD
    # JUNI 2024 - IDR / FACILITY AUTOMOBILE
    # OKTOBER 2024 - EUR / FACILITY
    #
    # Biarkan apa adanya
    # ============================================================

    if re.search(
        r"(JANUARI|FEBRUARI|MARET|APRIL|MEI|JUNI|JULI|AGUSTUS|SEPTEMBER|OKTOBER|NOVEMBER|DESEMBER)"
        r"\s+\d{4}",
        val,
        flags=re.IGNORECASE,
    ):
        return [_normalize_spaces(val)]

    m = re.search(
        r"\d+(?:\.\d+)+",
        val
    )

    if m:
        return [m.group(0)]

    # ============================================================
    # 7. DEFAULT
    # ============================================================

    return [_normalize_spaces(val)]     
    
def clean_slip_raw(val) -> list:
    """
    Cleaning Slip Osbal

    Rules:
    - Menjalankan _clean_slip_core()
    - Menghapus titik (.)
    - Merapikan spasi
    - Menghapus duplikasi
    - Menghapus nilai kosong
    """

    results = _clean_slip_core(val)

    cleaned = []

    for r in results:

        if pd.isna(r):
            continue

        r = str(r).upper()

        r = _normalize_spaces(r)

        r = r.replace(".", "")

        # hapus TBA kalau ada setelah nomor slip
        r = re.sub(r'(?<=\d)\s*TBA\b', '', r)

        r = r.strip()

        if not r:
            continue

        if r not in cleaned:
            cleaned.append(r)

    # ============================================================
    # SLIP 4 DIGIT
    #
    # 6056
    # -> 00006056
    #
    # Hanya berlaku jika tepat 4 digit
    # ============================================================

    val_str = "" if pd.isna(val) else str(val).strip()

    if re.fullmatch(r"\d{4}", val_str):
        return ["0000" + val_str]

    return cleaned


# ============================================================
# FINAL CLEAN SLIP
# HAPUS SEMUA PEMISAH
# / - + . & spasi dll
# ============================================================

def clean_slip(val) -> list:
    hasil = clean_slip_raw(val)

    hasil_clean = []

    for x in hasil:
        x = str(x).upper()

        # ============================================================
        # PERTAHANKAN SPASI UNTUK:
        # - Nama bulan
        # - PENYELESAIAN
        # - HUTANG / UTANG
        # ============================================================
        if re.search(
            r"\b(JANUARI|FEBRUARI|MARET|APRIL|MEI|JUNI|JULI|"
            r"AGUSTUS|SEPTEMBER|OKTOBER|NOVEMBER|DESEMBER|"
            r"JANUARY|FEBRUARY|MARCH|APRIL|MAY|JUNE|JULY|"
            r"AUGUST|SEPTEMBER|OCTOBER|DECEMBER)\b"
            r"|\bPENYELESAIAN\b"
            r"|\bHUTANG\b"
            r"|\bUTANG\b",
            x,
            flags=re.IGNORECASE
        ):
            x = _normalize_spaces(x)

        else:
            # Selain itu → hapus semua pemisah termasuk spasi
            x = re.sub(r"[^A-Z0-9]", "", x)

        if x:
            hasil_clean.append(x)

    return hasil_clean
# ─────────────────────────────────────────────────────────────────────────────
# CLEAN INSURED
# ─────────────────────────────────────────────────────────────────────────────
def clean_insured(val) -> list:
    if pd.isna(val):
        return []

    val = str(val).strip()

    if not val:
        return []

    # Ambil isi dalam kurung
    bracket_parts = re.findall(r"\(([^)]*)\)", val)

    # Hapus seluruh isi kurung dari string utama
    val = re.sub(r"\([^)]*\)", "", val)

    # Breakdown berdasarkan delimiter
    parts = re.split(r"/|,|QQ", val, flags=re.IGNORECASE)

    # Tambahkan isi kurung sebagai insured baru
    parts.extend(bracket_parts)

    cleaned = []

    for p in parts:

        p = _normalize_spaces(p.strip())

        p = re.sub(
        r"^(MR|MRS|MS|DR|IR)\.?\s+",
        "",
        p,
        flags=re.IGNORECASE,
    )

        if len(p) <= 2:
            continue

        if p.upper() in INSURED_JUNK_WORDS:
            continue

        if re.match(r"^[^A-Za-z0-9]+$", p):
            continue

        p_clean = _clean_insured_name(p)

        if not p_clean:
            continue

        if p_clean.upper() in INSURED_JUNK_WORDS:
            continue

        if p_clean not in cleaned:
            cleaned.append(p_clean.upper())

    if not cleaned:
        fallback = _clean_insured_name(_normalize_spaces(val))
        return [fallback.upper()] if fallback else []

    return _cap_or_join(cleaned)
# ─────────────────────────────────────────────────────────────────────────────
# PROCESS DATA
# ─────────────────────────────────────────────────────────────────────────────

def _insert_clean_columns(
    df: pd.DataFrame,
    all_lists: list,
    prefix: str,
    max_cols: int,
) -> list:
    """Buat kolom clean_{prefix}_N dan kembalikan nama-nama kolom baru."""
    added = []
    for i in range(1, max_cols + 1):
        col_name = f"clean {prefix} {i}"
        df[col_name] = [lst[i - 1] if i - 1 < len(lst) else None for lst in all_lists]
        added.append(col_name)
    return added

def is_valid_clstd_polis(val):
    if pd.isna(val):
        return False

    val = str(val).strip().upper()

    if not val:
        return False

    if val in {"TBA", "VAR", "VARIOUS"}:
        return False

    return True

def _is_valid_clstd_polis(val) -> bool:
    if pd.isna(val):
        return False

    val = str(val).strip().upper()

    if not val:
        return False

    if val in {"TBA", "VAR", "VARIOUS"}:
        return False

    return True

def process_data(input_file: str, output_file: str) -> None:
    print(f"[1/5] Membaca data dari: {input_file} ...")
    df = pd.read_excel(input_file, header=0)

    print(f"      Total baris keseluruhan: {len(df):,}")

    # ============================================================
    # CEK KOLOM
    # ============================================================

    required_cols = [
        CEDANT_COL,
        POLIS_COL,
        SLIP_COL,
        INSURED_COL,
        "CLSDT_POLICY_NO",
        "CLSDT_SLIP_NO",
    ]

    for col in required_cols:
        if col not in df.columns:
            print(f"\n[ERROR] Kolom '{col}' tidak ditemukan!")
            print(f"        Kolom tersedia: {list(df.columns)}")
            return

    # ============================================================
    # FILTER CEDANT
    # ============================================================

    df[CEDANT_COL] = df[CEDANT_COL].astype(str).str.strip()

    is_tokio = df[CEDANT_COL] == CEDANT_VALUE

    print(
        f"[2/5] Filter cedant '{CEDANT_VALUE}': "
        f"{is_tokio.sum():,} baris TOKIO dari total {len(df):,} baris."
    )

    # ============================================================
    # RENAME DATA ORIGINAL
    # ============================================================

    df.rename(
        columns={
            POLIS_COL: "polis_ori",
            SLIP_COL: "slip_ori",
            INSURED_COL: "insured_ori",
        },
        inplace=True,
    )
    # ============================================================
    # CLEANING
    # ============================================================

    print("[3/5] Menjalankan proses cleaning hanya untuk TOKIO ...")

    all_clean_polis = []
    all_clean_slip = []
    all_clean_ins = []

    max_polis = 1
    max_slip = 1
    max_ins = 1

    total_clstd_used = 0
    total_fallback_fac = 0

    for idx, (_, row) in enumerate(df.iterrows(), 1):

        if idx % 50_000 == 0:
            print(
                f"      Progress: {idx:,} / "
                f"{len(df):,} baris diproses..."
            )

        # ========================================================
        # HANYA CLEANING TOKIO
        # ========================================================

        if is_tokio.iloc[idx - 1]:

            polis_ori = row.get("polis_ori", "")
            clstd_polis = row.get("CLSDT_POLICY_NO", "")

            # ====================================================
            # TENTUKAN SUMBER POLIS CLEAN
            # ====================================================

            if _is_valid_clstd_polis(clstd_polis):

                # CLSDT valid → gunakan CLSDT
                polis_source = clstd_polis
                total_clstd_used += 1

            else:

                # CLSDT kosong / TBA / VAR / VARIOUS
                # → kembali ke POLIS ORI
                polis_source = polis_ori
                total_fallback_fac += 1

            # ====================================================
            # CLEAN POLIS
            #
            # PENTING:
            # clean_polis() TIDAK DIUBAH.
            # Yang berubah hanya sumber inputnya.
            # ====================================================

            c_polis = clean_polis(polis_source)

            # ========================================================
            # CLEAN SLIP
            # ========================================================

            clstd_slip = row.get("CLSDT_SLIP_NO", "")
            slip_ori = row.get("slip_ori", "")

            if _is_valid_clstd_polis(clstd_slip):

                slip_source = clstd_slip
                slip_source_type = "CLSDT"

            else:

                slip_source = slip_ori
                slip_source_type = "FAC"

            c_slip = clean_slip(slip_source)

            # DEBUG SLIP
            if str(clstd_slip).strip() == "00006056" and str(slip_ori).strip() == "P1":
                print("=" * 80)
                print("DEBUG SLIP")
                print(f"SLIP ORI       : {slip_ori}")
                print(f"CLSDT SLIP     : {clstd_slip}")
                print(f"SUMBER         : {slip_source_type}")
                print(f"SLIP SOURCE    : {slip_source}")
                print(f"CLEAN SLIP     : {c_slip}")
                print("=" * 80)
            # ====================================================
            # CLEAN INSURED
            # ====================================================

            c_ins = clean_insured(row.get("insured_ori", ""))

        else:

            c_polis = []
            c_slip = []
            c_ins = []

        # ========================================================
        # UPDATE MAX COLUMN
        # ========================================================

        max_polis = max(max_polis, len(c_polis))
        max_slip = max(max_slip, len(c_slip))
        max_ins = max(max_ins, len(c_ins))

        all_clean_polis.append(c_polis)
        all_clean_slip.append(c_slip)
        all_clean_ins.append(c_ins)

    print("      Selesai diproses!")

    print(f"      -> Jumlah kolom clean polis  : {max_polis}")
    print(f"      -> Jumlah kolom clean slip   : {max_slip}")
    print(f"      -> Jumlah kolom clean insured: {max_ins}")

    print()
    print("      SUMBER POLIS CLEAN:")
    print(f"      -> Menggunakan CLSDT : {total_clstd_used:,} baris")
    print(f"      -> Fallback ke FAC   : {total_fallback_fac:,} baris")

    # ============================================================
    # SUSUN KOLOM OUTPUT
    #
    # CLEAN TETAP LANGSUNG DI SEBELAH ORI
    # ============================================================

    print("[4/5] Menyusun kolom output ...")

    new_columns = []

    for col in df.columns:

        new_columns.append(col)

        if col == "polis_ori":

            new_columns += _insert_clean_columns(
                df,
                all_clean_polis,
                "polis",
                max_polis,
            )

        elif col == "slip_ori":

            new_columns += _insert_clean_columns(
                df,
                all_clean_slip,
                "slip",
                max_slip,
            )

        elif col == "insured_ori":

            new_columns += _insert_clean_columns(
                df,
                all_clean_ins,
                "insured",
                max_ins,
            )

    df = df[new_columns]

    # ============================================================
    # SAVE
    # ============================================================

    print(f"[5/5] Menyimpan hasil ke: {output_file} ...")

    df.to_excel(output_file, index=False)

    print()
    print("=" * 55)
    print("  [OK] Selesai!")
    print(f"  Total baris output  : {len(df):,}")
    print(f"  Total kolom output  : {len(df.columns)}")
    print(f"  File disimpan di    : {output_file}")
    print("=" * 55)
# ─────────────────────────────────────────────────────────────────────────────
# ENTRY POINT
# ─────────────────────────────────────────────────────────────────────────────
# ─────────────────────────────────────────────────────────────────────────────
# TEST RUN
# ─────────────────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    print("MASUK MAIN")
    process_data(INPUT_FILE, OUTPUT_FILE)
 
