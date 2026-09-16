import os
import re
from pathlib import Path

import pandas as pd


# ============================================================
# CONFIG
# ============================================================

BASE_DIR = Path(__file__).resolve().parent

INPUT_FILE = os.path.join(BASE_DIR, "input", "2b. transaksi Osbal 01.01.23 - 17.08.26.xlsx")
OUTPUT_FILE = os.path.join(BASE_DIR, "output", "mag_output_osbal.xlsx")

CEDANT_COL = "CCOS_COMP_NAME"

# HANYA CEDANT INI YANG DICLEANING
CEDANT_VALUE = "PT.ASURANSI MULTI ARTHA GUNA"

POLIS_COL = "FAC_POLICY_NO"
SLIP_COL = "FAC_SLIP"
INSURED_COL = "FAC_INSURED"
CLSDT_SERTF_COL = "CLSDT_SERTF_NO"

MAX_SPLIT_COLS = 5



# ============================================================
# EXCEPTION POLIS
# ============================================================

POLIS_EXCEPTION_RE = re.compile(
    r"""
    MOP\s*/\s*MARINE
    |
    (?:LINE\s*/\s*SLIP|LINESLIP)
    |
    \b(?:P1|P2|P3|P73)\s*/\s*CANCEL
    |
    \b(?:P1|P2|P3|P73)\b
    |
    \bCANCEL\b
    |
    PENYELESAIAN(?:\s+SUSPENSE)?
    |
    HUTANG
    |
    UTANG
    """,
    re.IGNORECASE | re.VERBOSE,
)


# ============================================================
# EXCEPTION SLIP
# ============================================================

SLIP_EXCEPTION_RE = re.compile(
    r"""
    \bSUMMARY\b
    |
    \bBORDER[OA]\b
    |
    \bBORDRO\b
    |
    \bSINGGLESHIPMENT\b
    |
    \bP1\b
    |
    \bP2\b
    |
    \bP3\b
    |
    PENYELESAIAN(?:\s+SUSPENSE)?
    |
    HUTANG
    |
    UTANG
    """,
    re.IGNORECASE | re.VERBOSE,
)


MONTH_NAMES = frozenset({
    "JANUARI",
    "FEBRUARI",
    "MARET",
    "APRIL",
    "MEI",
    "JUNI",
    "JULI",
    "AGUSTUS",
    "SEPTEMBER",
    "OKTOBER",
    "NOVEMBER",
    "DESEMBER",
    "JANUARY",
    "FEBRUARY",
    "MARCH",
    "MAY",
    "JUNE",
    "JULY",
    "AUGUST",
    "SEPTEMBER",
    "OCTOBER",
    "NOVEMBER",
    "DECEMBER",
})


# ============================================================
# INSURED
# ============================================================

INSURED_JUNK_WORDS = frozenset({
    "PT",
    "CV",
    "TBK",
    "PERSERO",
    "LTD",
    "INC",
    "LLC",
    "AND",
    "OR",
    "THE",
    "OF",
    "AS",
    "NON FOOD",
    "DIV",
    "MR",
    "MR.",
    "MRS",
    "MRS.",
    "MS",
    "MS.",
})


INSURED_SUFFIX_RE = re.compile(
    r"""
    ,?\s*\bTBK\b\s*(?:,?\s*PT\.?)?
    |
    ,?\s*\bPT\.?\s*$
    |
    ,?\s*\bCV\.?\s*$
    """,
    re.IGNORECASE | re.VERBOSE,
)


INSURED_PT_QQ_RE = re.compile(
    r"""
    \bPT\.?\b
    |
    \bQQ\b
    """,
    re.IGNORECASE | re.VERBOSE,
)


# ============================================================
# SLIP NOISE
# ============================================================

_SLIP_NOISE_RE = re.compile(
    r"""
    \b(?:JANUARI|FEBRUARI|MARET|APRIL|MEI|JUNI|JULI|AGUSTUS
        |SEPTEMBER|OKTOBER|NOVEMBER|DESEMBER)\b
    |
    \b(?:JANUARY|FEBRUARY|MARCH|APRIL|MAY|JUNE|JULY
        |AUGUST|SEPTEMBER|OCTOBER|NOVEMBER|DECEMBER)\b
    |
    \b(?:IDR|USD|ENG)\b
    |
    \b(?:P1|P2|P3|P73)\b
    |
    \bNEW\b
    |
    \bVARIOUS\b
    |
    \b20[0-9]{2}\b
    """,
    re.IGNORECASE | re.VERBOSE,
)


_SLIP_TOKEN_RE = re.compile(
    r"[A-Z0-9][A-Z0-9\-]{6,}",
    re.IGNORECASE,
)


# ============================================================
# HELPERS
# ============================================================

def _normalize_spaces(text: str) -> str:
    return re.sub(r"\s{2,}", " ", text).strip()


def _is_valid_slip_token(tok: str) -> bool:
    tok = tok.strip()

    return (
        len(tok) >= 7
        and bool(re.search(r"\d", tok))
        and not re.match(r"^\d{4}$", tok)
    )


def _strip_slip_suffix(tok: str) -> str:

    m = re.match(
        r"^(.+?)-(\d{1,6})$",
        tok,
    )

    if m and len(m.group(1)) > len(m.group(2)):
        return m.group(1)

    return tok


def _extract_slip_tokens(text: str) -> list:

    results = []

    for block in re.split(
        r"\s{2,}",
        text.strip(),
    ):

        clean = _SLIP_NOISE_RE.sub(
            " ",
            block,
        )

        clean = re.sub(
            r"^[\s\-\/+,]+|[\s\-\/+,]+$",
            "",
            clean,
        ).strip()

        for cand in _SLIP_TOKEN_RE.findall(clean):

            cand = cand.strip("-")

            parts = cand.split("-")

            if len(parts) == 2:

                a, b = parts

                if (
                    _is_valid_slip_token(a)
                    and _is_valid_slip_token(b)
                    and abs(len(a) - len(b)) <= 2
                ):
                    results.extend([a, b])
                    continue

            if _is_valid_slip_token(cand):

                results.append(
                    _strip_slip_suffix(cand)
                )

    return results


# ============================================================
# CLEAN INSURED NAME
# ============================================================

def _clean_insured_name(name: str) -> str:

    name = _normalize_spaces(name)

    if re.fullmatch(
        r"AEON\s+MALL\s*\(TENANTS\)",
        name,
        re.IGNORECASE,
    ):
        return name.upper()

    for _ in range(3):

        cleaned = _normalize_spaces(
            INSURED_SUFFIX_RE.sub(
                "",
                name,
            ).strip().strip(",")
        )

        if cleaned == name:
            break

        name = cleaned

    name = INSURED_PT_QQ_RE.sub(
        " ",
        name,
    )

    if re.fullmatch(
        r"\s*AEON\s+MALL\s+\(TENANTS\)\s*",
        name,
        flags=re.IGNORECASE,
    ):
        return "AEON MALL (TENANTS)"

    name = re.sub(
        r"\(\s*PERSERO\s*\)",
        "",
        name,
        flags=re.IGNORECASE,
    )

    name = re.sub(
        r"[()]",
        " ",
        name,
    )

    name = _normalize_spaces(name)

    name = re.sub(
        r"^[\s.,\-]+|[\s.,\-]+$",
        "",
        name,
    )

    return name


def _cap_or_join(items: list) -> list:

    if len(items) > MAX_SPLIT_COLS:
        return [",".join(items)]

    return items


# ============================================================
# BUILD REPEATED POLIS
# ============================================================

def _build_repeated_polis(
    base: str,
    suffix: str,
) -> str:

    base = re.sub(
        r"[^A-Z0-9]",
        "",
        str(base).upper(),
    )

    suffix = re.sub(
        r"[^A-Z0-9]",
        "",
        str(suffix).upper(),
    )

    if not base or not suffix:
        return ""

    if len(suffix) >= len(base):
        return suffix

    return (
        base[:-len(suffix)]
        + suffix
    )


def _is_short_suffix(value: str) -> bool:

    value = value.strip().upper()

    return bool(
        re.fullmatch(
            r"\d{1,6}",
            value,
        )
    )


# ============================================================
# STRIP REFERENSI "EX" (POLIS SEBELUMNYA)
# ============================================================
# REVISI: referensi ke polis sebelumnya (EX / EX. POLICY NO /
# EX.<nomor>) hanya catatan, BUKAN bagian dari polis atau
# certificate. Buang seluruhnya.
#
# Contoh:
# 11020123000087 EX. POLICY NO : 11020122000132
# -> 11020123000087
#
# 45100422000467 - 000001 EX.45100421000501
# -> 45100422000467 - 000001
# ============================================================

def _strip_ex_reference(text: str) -> str:

    return re.sub(
        r"\s*EX\.?\s*(?:POLICY\s*NO\.?\s*)?:?\s*\d{4,}\s*$",
        "",
        text,
        flags=re.IGNORECASE,
    ).strip()


# ============================================================
# CLEAN POLIS RAW
# ============================================================

def clean_polis_raw(val) -> list:

    if pd.isna(val):
        return []

    original = str(val).strip().upper()

    if not original:
        return []

    original = _strip_ex_reference(original)

    if not original:
        return []

    # ========================================================
    # RULE BARU MULTI ARTHA GUNA (DATA 2 OSBAL):
    # RANGKAIAN P-TOKEN SAJA (P1 + P2 + P3 + P4 ...)
    #
    # Tidak ada digit polis sama sekali, hanya penanda P1/P2/dst.
    # Ini contoh pola CLSDT yang tidak valid (mis. "P1 CANCEL
    # P2 CANCEL P3 CANCEL"). BIARKAN APA ADANYA, jangan dipecah.
    #
    # Contoh:
    # P1 + P2 + P3 + P4
    # -> P1 + P2 + P3 + P4
    # ========================================================

    if re.fullmatch(
        r"\s*P\d+\s*(?:\+\s*P\d+\s*)*",
        original,
        flags=re.IGNORECASE,
    ):
        return [original]

        # ========================================================
    # RULE BARU MULTI ARTHA GUNA:
    # BORDERO + TBA
    #
    # TBA hanya dibuang.
    # Selain TBA, format bordero dibiarkan apa adanya.
    #
    # Contoh:
    # 40010922042857.000752 S/D 001152+507.001 - 004 TBA
    # ->
    # 40010922042857.000752 S/D 001152+507.001 - 004
    # ========================================================
    bordero_tba_match = re.fullmatch(
        r"(.*?\S)\s+TBA\s*",
        original,
        flags=re.IGNORECASE,
    )

    if bordero_tba_match and "S/D" in original.upper():
        return [bordero_tba_match.group(1).strip()]

    # ========================================================
    # RULE BARU MULTI ARTHA GUNA (DATA 2 OSBAL):
    # DAFTAR KODE PENDEK (1-3 digit) DIPISAH "+" + TBA DI AKHIR
    #
    # TBA hanya penanda, cukup dibuang. Format "+" lainnya
    # dibiarkan apa adanya.
    #
    # Contoh:
    # 026 + 015 + 037 + 048 + 195 + 162 + 173 + 184 TBA
    # ->
    # 026 + 015 + 037 + 048 + 195 + 162 + 173 + 184
    # ========================================================

    short_code_tba_match = re.fullmatch(
        r"""
        \s*
        (\d{1,3}(?:\s*\+\s*\d{1,3})+)
        \s+
        TBA
        \s*
        """,
        original,
        flags=re.IGNORECASE | re.VERBOSE,
    )

    if short_code_tba_match:
        return [
            _normalize_spaces(
                short_code_tba_match.group(1)
            )
        ]

    # ========================================================
    # RULE BARU MULTI ARTHA GUNA (DATA 2 OSBAL):
    # POLIS TBA (+ P2 + P3 + ...)
    #
    # TBA dan P2/P3/dst hanya penanda, semua dibuang.
    # Clean polis = POLIS saja.
    #
    # Contoh:
    # 45013025001791 TBA
    # -> 45013025001791
    #
    # 50010925000078 TBA + P2 + P3
    # -> 50010925000078
    # ========================================================

    plain_polis_tba_match = re.fullmatch(
        r"""
        \s*
        (\d{10,})
        \s+
        TBA
        (?:\s*\+\s*P\d+)*
        \s*
        """,
        original,
        flags=re.IGNORECASE | re.VERBOSE,
    )

    if plain_polis_tba_match:
        return [
            plain_polis_tba_match.group(1)
        ]

    # ========================================================
    # RULE BARU MULTI ARTHA GUNA (DATA 2 OSBAL):
    # BASE[.-]CERT1 S/D [BASE[.-]]CERT2
    #
    # Satu polis, dua certificate. Clean polis = BASE saja.
    # (Certificate-nya diproses terpisah lewat
    # _extract_certificate_list().)
    #
    # Contoh:
    # 02031121000065-000366 S/D 02031121000065-000385
    # -> 02031121000065
    #
    # 40010922042857.001214 S/D 001224
    # -> 40010922042857
    #
    # 40010924046082-001199 S/D 40010924046082.001210
    # -> 40010924046082
    #
    # 40010925046051.000001 s/d 40010925046051.001359
    # -> 40010925046051
    # ========================================================

    two_cert_polis_match = re.fullmatch(
        r"""
        \s*
        (\d{10,})
        [.\-]
        \d{1,6}
        \s*S\s*/?\s*D\s*
        (?:\1[.\-])?
        \d{1,7}
        \s*
        """,
        original,
        flags=re.IGNORECASE | re.VERBOSE,
    )

    if two_cert_polis_match:
        return [
            two_cert_polis_match.group(1)
        ]

    # ========================================================
    # CERTIFICATE FORMAT
    #
    # Certificate hanya mengambil POLIS INDUK.
    #
    # Contoh:
    # 1010100823000179 - CERTIF : 000001 , 000002
    # -> 1010100823000179
    #
    # 1010100822000097 - 000001 , 000002 , 000003
    # -> 1010100822000097
    #
    # P1 / 1071031124000024 - 000251 SD 0000131
    # -> 1071031124000024
    # ========================================================

    certif_polis_match = re.fullmatch(
        r"""
        \s*
        (?:P1\s*/\s*)?
        (\d{10,})
        \s*-\s*
        (?:
            CERTIF\s*:\s*
        )?
        \d{1,6}
        \s*,\s*
        \d{1,6}
        (?:
            \s*,\s*
            \d{1,6}
        )?
        \s*
        """,
        original,
        flags=re.IGNORECASE | re.VERBOSE,
    )

    if certif_polis_match:
        return [
            certif_polis_match.group(1)
        ]

    certif_range_match = re.fullmatch(
        r"""
        \s*
        P1\s*/\s*
        (\d{10,})
        \s*-\s*
        \d{1,6}
        \s*S\s*/?\s*D\s*
        \d{1,7}
        \s*
        """,
        original,
        flags=re.IGNORECASE | re.VERBOSE,
    )

    if certif_range_match:
        return [
            certif_range_match.group(1)
        ]
    
    # ========================================================
    # POLIS / P2
    #
    # Contoh:
    # 2011090125000012 / P2
    #
    # P2 dihapus dari clean polis.
    # ========================================================

    p2_match = re.fullmatch(
        r"""
        \s*
        (\d{10,})
        \s*/\s*
        P2
        \s*
        """,
        original,
        flags=re.IGNORECASE | re.VERBOSE,
    )

    if p2_match:
        return [
            p2_match.group(1)
        ]

    # ========================================================
    # P1 + P2 / TBA + P2
    #
    # HARUS DIBIARKAN APA ADANYA.
    # JANGAN BREAKDOWN.
    # ========================================================

    if re.fullmatch(
        r"""
        (?:P1|TBA)
        \s*\+\s*
        P2
        """,
        original,
        flags=re.IGNORECASE | re.VERBOSE,
    ):
        return [original]

    # yang ditambahin
    # ========================================================
    # CERTIFICATE FORMAT KHUSUS
    #
    # Certificate TIDAK BOLEH ikut menjadi clean polis.
    # Ambil hanya polis induknya.
    # ========================================================

    cert_polis_match = re.fullmatch(
        r"""
        \s*
        (?:P1\s*/\s*)?
        (\d{10,})
        \s*-\s*
        (?:
            CERTIF\s*:\s*
        )?
        \d{1,6}
        (?:
            \s*,\s*\d{1,6}
        )+
        \s*
        """,
        original,
        flags=re.IGNORECASE | re.VERBOSE,
    )

    if cert_polis_match:

        return [
            cert_polis_match.group(1)
        ]

    # ========================================================
    # CERTIFICATE RANGE
    #
    # P1 / POLIS - 000251 SD 0000131
    #
    # Clean polis = POLIS SAJA
    # ========================================================

    cert_range_match = re.fullmatch(
        r"""
        \s*
        P1\s*/\s*
        (\d{10,})
        \s*-\s*
        \d{1,6}
        \s*S\s*/?\s*D\s*
        \d{1,7}
        \s*
        """,
        original,
        flags=re.IGNORECASE | re.VERBOSE,
    )

    if cert_range_match:

        return [
            cert_range_match.group(1)
        ]

    repeat_match = re.fullmatch(
        r"(\d+)\s*-\s*(\d{1,7})",
        original,
    )

    # ========================================================
    # KHUSUS MULTI ARTHA GUNA:
    # BASE - NNNNNN S/D NNNNNN (TANPA prefix "P1/")
    # Pemisah "-" boleh ada atau tidak (kadang cuma spasi).
    #
    # Contoh:
    # 40080521000122 - 000001 S/D 000010
    # 03120622000013 - 000001 s/d 000006
    # 45031122001162 000064 S/D 000071
    #
    # Clean polis = BASE saja.
    # Certificate diproses TERPISAH lewat _extract_certificate_list().
    # ========================================================

    sd_no_prefix_match = re.fullmatch(
        r"""
        \s*
        (\d{10,})
        [\s-]+
        \d{1,6}
        \s*S\s*/?\s*D\s*
        \d{1,7}
        \s*
        """,
        original,
        flags=re.IGNORECASE | re.VERBOSE,
    )

    if sd_no_prefix_match:
        return [
            sd_no_prefix_match.group(1)
        ]

    # ========================================================
    # KHUSUS MULTI ARTHA GUNA (REVISI):
    # BASE - CERT1 - CERT2 (dua-duanya nomor certificate,
    # dipisah "-" biasa, TANPA teks S/D).
    #
    # Nomor certificate dibatasi 5-7 digit supaya tidak
    # tertukar dengan suffix pendek (3-4 digit) pada rantai
    # perulangan polis (chain_match di bawah).
    #
    # Contoh:
    # 02031123000014-000006 - 000023
    # -> 02031123000014
    # ========================================================

    dash_range_no_prefix_match = re.fullmatch(
        r"""
        \s*
        (\d{10,})
        [\s-]+
        \d{5,7}
        \s*-\s*
        \d{5,7}
        \s*
        """,
        original,
        flags=re.IGNORECASE | re.VERBOSE,
    )

    if dash_range_no_prefix_match:
        return [
            dash_range_no_prefix_match.group(1)
        ]

    # ========================================================
    # KHUSUS MULTI ARTHA GUNA (REVISI):
    # BASE RANGE,EXTRA(/VARIOUS)
    #
    # Contoh:
    # 02031125000018 000455-000536,000467/VARIOUS
    # -> 02031125000018
    # ========================================================

    cert_range_extra_no_prefix_match = re.fullmatch(
        r"""
        \s*
        (\d{10,})
        [\s-]+
        \d{5,7}
        \s*-\s*
        \d{5,7}
        \s*,\s*
        \d{5,7}
        (?:\s*/\s*(?:VAR|VARIOUS))?
        \s*
        """,
        original,
        flags=re.IGNORECASE | re.VERBOSE,
    )

    if cert_range_extra_no_prefix_match:
        return [
            cert_range_extra_no_prefix_match.group(1)
        ]

    # ========================================================
    # KHUSUS MULTI ARTHA GUNA:
    # PERULANGAN POLIS RANTAI (BASE-suf1-suf2-...-sufN)
    #
    # Contoh:
    # 40012121031606-26085-5849-5862   -> 4 polis
    # 40013022002828-2817-1396-1385    -> 4 polis
    # 02010922005704-1073-186          -> 3 polis
    #
    # Jika total polis (base + jumlah suffix) melebihi
    # MAX_SPLIT_COLS, BIARKAN APA ADANYA (jangan dipecah).
    # ========================================================

    chain_match = re.fullmatch(
        r"(\d{10,})((?:-\d{1,20}){2,})",
        original,
    )

    if chain_match:

        base = chain_match.group(1)

        suffix_groups = re.findall(
            r"\d{1,20}",
            chain_match.group(2),
        )

        results = [base]

        for suffix in suffix_groups:

            # ------------------------------------------------
            # KHUSUS MULTI ARTHA GUNA:
            # Suffix >= 10 digit dalam rantai dianggap sudah
            # jadi nomor polis independen (bukan suffix untuk
            # di-build), sisanya (< 10 digit) tetap dibentuk
            # dengan replace trailing digit dari base.
            #
            # Contoh:
            # 40010922026607-6618-6664-6675-67446777678868026813
            # -> 40010922026607, 40010922026618, 40010922026664,
            #    40010922026675, 67446777678868026813
            # ------------------------------------------------

            if len(suffix) >= 10:
                candidate = suffix
            else:
                candidate = _build_repeated_polis(
                    base,
                    suffix,
                )

            if candidate and candidate not in results:
                results.append(candidate)

        if len(results) > MAX_SPLIT_COLS:
            cleaned_original = re.sub(
                r"\s*\+\s*TBA\b",
                "",
                original,
                flags=re.IGNORECASE,
            )

            return [_normalize_spaces(cleaned_original)]

        return results

    if repeat_match:

        base = repeat_match.group(1)
        suffix = repeat_match.group(2)

        # ------------------------------------------------------
        # KHUSUS MULTI ARTHA GUNA:
        # Suffix 6 digit -> BUKAN perulangan polis, tapi
        # certificate. Clean polis = BASE saja.
        #
        # Contoh:
        # 40030121085615 - 000108   -> clean polis: 40030121085615
        # 45013022003422 - 000001   -> clean polis: 45013022003422
        # ------------------------------------------------------

        if len(suffix) == 6:
            return [base]

        if len(suffix) < len(base):

            candidate = _build_repeated_polis(
                base,
                suffix,
            )

            results = [base]

            if candidate and candidate not in results:
                results.append(candidate)

            return results

    # ========================================================
    # KHUSUS MULTI ARTHA GUNA:
    # BASE - SUFFIX_PANJANG (suffix >= 10 digit)
    #
    # Suffix sepanjang ini adalah nomor polis independen kedua,
    # BUKAN suffix perulangan yang perlu di-build.
    #
    # Contoh:
    # 05010922012345-05010522002309   -> 2 polis
    # 45013022020627-45013222013849   -> 2 polis
    # ========================================================

    long_pair_match = re.fullmatch(
        r"(\d{10,})\s*-\s*(\d{10,})",
        original,
    )

    if long_pair_match:

        base = long_pair_match.group(1)
        second = long_pair_match.group(2)

        results = [base]

        if second not in results:
            results.append(second)

        return results

    pr_pattern = (
        re.search(
            r"^\s*P1\s*/",
            original,
            flags=re.IGNORECASE,
        )
        or
        re.search(
            r"^\s*\d+\s*-\s*\d{4,7}(?:\s*/|$)",
            original,
            flags=re.IGNORECASE,
        )
        or
        re.search(
            r"^\s*\d+\s*-\s*\d{4,7}\s*/\s*\d+",
            original,
            flags=re.IGNORECASE,
        )
    )

    if pr_pattern:

        pr_val = original

        pr_val = re.sub(
            r"^\s*P1\s*/\s*",
            "",
            pr_val,
            flags=re.IGNORECASE,
        ).strip()

        pr_val = re.sub(
            r"\s*/\s*(?:VAR|VARIOUS)\s*$",
            "",
            pr_val,
            flags=re.IGNORECASE,
        ).strip()

        parts = re.split(
            r"/|,|QQ|&",
            val,
            flags=re.IGNORECASE,
        )

        results = []

        for part in parts:

            part = part.strip()

            if not part:
                continue

            if part.upper() in {
                "P1",
                "VAR",
                "VARIOUS",
            }:
                continue

            if re.fullmatch(
                r"\d{1,7}",
                part,
            ):
                continue

            part = re.sub(
                r"^(\d+)\s*-\s*\d{4,7}$",
                r"\1",
                part,
            ).strip()

            part = re.sub(
                r"\s*/\s*\d{1,7}$",
                "",
                part,
            ).strip()

            cleaned = re.sub(
                r"[^A-Z0-9]",
                "",
                part.upper(),
            )

            if (
                cleaned
                and cleaned not in {
                    "P1",
                    "VAR",
                    "VARIOUS",
                }
            ):
                results.append(cleaned)

        if results:
            return results

    if original == "207010423000-08+09+207010223000-08+09":
        return [original]

    if original == "SUSPENSE VARIOUS CEDANT 2018-2019":
        return [original]

    if re.search(
        r"\bSUSPENSE\b",
        original,
        flags=re.IGNORECASE,
    ):
        return [original]

    # ========================================================
    # TBA + P... HARUS DIBIARKAN APA ADANYA
    # ========================================================

    if re.fullmatch(
        r"TBA\s*\+\s*P\d+",
        original,
        flags=re.IGNORECASE,
    ):
        return [original]

    if re.search(
        r"\b(PENYELESAIAN|HUTANG|UTANG|PIUTANG|ENDORSEMENT)\b",
        original,
        flags=re.IGNORECASE,
    ):
        return [original]

    sd_match = re.fullmatch(
        r"(.+?)\s+S\s*/?\s*D\s+(.+)",
        str(val),
        flags=re.IGNORECASE,
    )

    if sd_match:

        start_raw = sd_match.group(1).strip()
        end_raw = sd_match.group(2).strip()

        start_clean = re.sub(
            r"[^A-Z0-9]",
            "",
            start_raw.upper(),
        )

        end_clean = re.sub(
            r"[^A-Z0-9]",
            "",
            end_raw.upper(),
        )

        if (
            re.fullmatch(r"\d+", start_clean)
            and re.fullmatch(r"\d+", end_clean)
            and len(start_clean) == len(end_clean)
        ):

            start_num = int(start_clean)
            end_num = int(end_clean)

            step = (
                1
                if end_num >= start_num
                else -1
            )

            jumlah = abs(
                end_num - start_num
            ) + 1

            if jumlah > MAX_SPLIT_COLS:
                return [original]

            return [
                str(num).zfill(len(start_clean))
                for num in range(
                    start_num,
                    end_num + step,
                    step,
                )
            ]

        return [original]

    val_str = re.sub(
        r"\s*/\s*(?:VARIOUS|VAR)\s*$",
        "",
        original,
        flags=re.IGNORECASE,
    ).strip()

    if re.match(
        r"^TBA\s*/\s*",
        val_str,
        flags=re.IGNORECASE,
    ):
        val_str = re.sub(
            r"^TBA\s*/\s*",
            "",
            val_str,
            flags=re.IGNORECASE,
        ).strip()

    val_str = re.sub(
        r"-0$",
        "",
        val_str,
    ).strip()

    if re.search(
        r"\bS\s*/\s*D\b",
        val_str,
        flags=re.IGNORECASE,
    ):

        parts_sd = re.split(
            r"\s*S\s*/\s*D\s*",
            val_str,
            flags=re.IGNORECASE,
        )

        if len(parts_sd) != 2:
            return [original]

        start_raw = re.sub(
            r"[^0-9]",
            "",
            parts_sd[0],
        )

        end_raw = re.sub(
            r"[^0-9]",
            "",
            parts_sd[1],
        )

        if not start_raw or not end_raw:
            return [original]

        if len(start_raw) != len(end_raw):
            return [original]

        start_num = int(start_raw)
        end_num = int(end_raw)

        step = (
            1
            if end_num >= start_num
            else -1
        )

        jumlah = abs(
            end_num - start_num
        ) + 1

        if jumlah > MAX_SPLIT_COLS:
            return [original]

        return [
            str(num).zfill(len(start_raw))
            for num in range(
                start_num,
                end_num + step,
                step,
            )
        ]

    if val_str in {
        "TBA",
        "VAR",
        "VARIOUS",
    }:
        return []

    m = re.match(
        r"^(\d+)-(\d{1,7})$",
        val_str,
    )

    if m:

        base = m.group(1)
        suffix = m.group(2)

        if len(suffix) < len(base):

            candidate = _build_repeated_polis(
                base,
                suffix,
            )

            results = [base]

            if candidate and candidate not in results:
                results.append(candidate)

            return results

    m = re.match(
        r"^(\d+)-(\d{3}(?:-\d{3})+)$",
        val_str,
    )

    if m:

        base = m.group(1)

        suffixes = re.findall(
            r"\d{3}",
            m.group(2),
        )

        results = [base]

        prefix = base[:-3]

        for suffix in suffixes:

            candidate = prefix + suffix

            if candidate not in results:
                results.append(candidate)

        if len(results) > MAX_SPLIT_COLS:
            return [original]

        return results

    val_str = re.sub(
        r"^P1\s*/\s*",
        "",
        val_str,
        flags=re.IGNORECASE,
    ).strip()

    if "/" in val_str:

        slash_parts = [
            p.strip()
            for p in re.split(
                r"\s*/\s*",
                val_str,
            )
            if p.strip()
        ]

        results_slash = []

        for p in slash_parts:

            if p in {
                "VAR",
                "VARIOUS",
            }:
                continue

            if p:
                results_slash.append(p)

        if len(results_slash) > 1:
            return results_slash

    parts = re.split(
        r"\s*(?:\+|&|,)\s*",
        val_str,
    )

    parts = [
        p.strip()
        for p in parts
        if p.strip()
    ]

    results = []
    current_base = None

    for p in parts:

        # ----------------------------------------------------------
        # RULE BARU MULTI ARTHA GUNA (DATA 2 OSBAL):
        # TOKEN P1/P2/P3/... HANYA PENANDA, BUANG SAJA.
        # Jangan pernah dianggap sebagai nomor polis tersendiri.
        # ----------------------------------------------------------

        if re.fullmatch(
            r"P\d+",
            p,
            flags=re.IGNORECASE,
        ):
            continue

        m = re.match(
            r"^([A-Z0-9]+)-(\d{1,6})$",
            p,
        )

        if m:

            base = m.group(1)
            suffix = m.group(2)

            current_base = base

            # --------------------------------------------------
            # KHUSUS MULTI ARTHA GUNA:
            # Suffix 5-6 digit dianggap CERTIFICATE, bukan
            # perulangan polis. Ambil BASE saja, buang suffix
            # (certificate diproses terpisah).
            # --------------------------------------------------

            if len(suffix) >= 5:

                if base not in results:
                    results.append(base)

                continue

            if len(suffix) < len(base):

                candidate = (
                    base[:-len(suffix)]
                    + suffix
                )

                if base not in results:
                    results.append(base)

                if candidate not in results:
                    results.append(candidate)

                continue

            cleaned = re.sub(
                r"[^A-Z0-9]",
                "",
                p,
            )

            if cleaned:
                results.append(cleaned)

            continue

        if (
            current_base
            and re.fullmatch(
                r"\d{1,6}",
                p,
            )
            and len(p) < len(current_base)
        ):

            # --------------------------------------------------
            # KHUSUS MULTI ARTHA GUNA:
            # Suffix pendek 5-6 digit = CERTIFICATE, bukan
            # perulangan polis. Skip, jangan ubah current_base.
            # --------------------------------------------------

            if len(p) >= 5:
                continue

            candidate = (
                current_base[:-len(p)]
                + p
            )

            if candidate not in results:
                results.append(candidate)

            continue

        cleaned = re.sub(
            r"[^A-Z0-9]",
            "",
            p,
        )

        if (
            cleaned
            and cleaned not in {
                "TBA",
                "VAR",
                "VARIOUS",
            }
            and not re.fullmatch(
                r"P\d+",
                cleaned,
                flags=re.IGNORECASE,
            )
        ):

            results.append(cleaned)

            if re.search(
                r"\d",
                cleaned,
            ):
                current_base = cleaned

    if len(parts) == 1:

        space_parts = re.split(
            r"\s+",
            val_str.strip(),
        )

        if len(space_parts) > 1:

            temp = []

            for p in space_parts:

                cleaned = re.sub(
                    r"[^A-Z0-9]",
                    "",
                    p,
                )

                if (
                    cleaned
                    and cleaned not in {
                        "TBA",
                        "VAR",
                        "VARIOUS",
                    }
                    and not re.fullmatch(
                        r"P\d+",
                        cleaned,
                        flags=re.IGNORECASE,
                    )
                ):
                    temp.append(cleaned)

            if len(temp) > 1:
                results = temp

    results = [
        re.sub(
            r"^VARIOUS",
            "",
            x,
            flags=re.IGNORECASE,
        ).strip()
        for x in results
    ]

    results = [
        x
        for x in results
        if x
    ]

    unique_results = []

    for item in results:

        if item not in unique_results:
            unique_results.append(item)

    if len(unique_results) > MAX_SPLIT_COLS:

        # ----------------------------------------------------
        # KHUSUS MULTI ARTHA GUNA:
        # Saat dibiarkan apa adanya (melebihi MAX_SPLIT_COLS),
        # tetap hapus token "TBA" dan "P1/P2/P3/..." yang hanya
        # penanda.
        #
        # Contoh:
        # 05012121043301+106+341+117+128+139+141+TBA
        # -> 05012121043301+106+341+117+128+139+141
        #
        # 01013024000704 + 715 + 131 + 726 + 737 + 583 + P2
        # -> 01013024000704 + 715 + 131 + 726 + 737 + 583
        # ----------------------------------------------------

        cleaned_original = re.sub(
            r"\s*\+\s*TBA\b",
            "",
            original,
            flags=re.IGNORECASE,
        )

        cleaned_original = re.sub(
            r"\bTBA\s*\+\s*",
            "",
            cleaned_original,
            flags=re.IGNORECASE,
        ).strip()

        cleaned_original = re.sub(
            r"\s*\+\s*P\d+\b",
            "",
            cleaned_original,
            flags=re.IGNORECASE,
        )

        cleaned_original = re.sub(
            r"\bP\d+\s*\+\s*",
            "",
            cleaned_original,
            flags=re.IGNORECASE,
        ).strip()

        return [cleaned_original]

    return unique_results


# ============================================================
# CLEAN POLIS
# ============================================================

def clean_polis(val) -> list:

    if pd.isna(val):
        return []

    original = str(val).strip().upper()

    # ========================================================
    # POLIS + P2
    #
    # Contoh:
    # 1010010925002545 + P2
    #
    # P2 dihapus, polis tetap
    # ========================================================
    p2_plus_match = re.fullmatch(
        r"(\d{10,})\s*\+\s*P2",
        original,
        flags=re.IGNORECASE,
    )

    if p2_plus_match:
        return [
            p2_plus_match.group(1)
        ]

    # ========================================================
    # P1 + P2
    #
    # HARUS DIBIARKAN APA ADANYA
    # ========================================================
    if re.fullmatch(
        r"P1\s*\+\s*P2",
        original,
        flags=re.IGNORECASE,
    ):
        return [original]

    # ========================================================
    # KHUSUS ETIQA:
    # 3039010925001162 + 138 + 151 + 149 + P2
    # P2 hanya penanda, jangan masuk hasil polis
    # ========================================================
    m = re.fullmatch(
        r"(\d{10,})\s*\+\s*(\d+)\s*\+\s*(\d+)\s*\+\s*(\d+)\s*\+\s*P2",
        original
    )

    if m:
        base = m.group(1)
        prefix = base[:-4]

        return [
            base,
            prefix + m.group(2).zfill(4),
            prefix + m.group(3).zfill(4),
            prefix + m.group(4).zfill(4),
        ]

    if not original:
        return []

    if original in {
        "207010423000-08+09+207010223000-08+09",
        "SUSPENSE VARIOUS CEDANT 2018-2019",
    }:
        return [original]

    original = re.sub(
        r"\s*/\s*0\s*/\s*(?:VAR|VARIOUS)\s*$",
        "",
        original,
        flags=re.IGNORECASE,
    ).strip()

    original = re.sub(
        r"\s*-\s*\d+\s*/\s*(?:VAR|VARIOUS)\s*$",
        "",
        original,
        flags=re.IGNORECASE,
    ).strip()

    original = re.sub(
        r"\s*-\s*(?:VAR|VARIOUS)\s*$",
        "",
        original,
        flags=re.IGNORECASE,
    ).strip()

    hasil = clean_polis_raw(original)

    hasil_clean = []

    for x in hasil:

        x = str(x).strip().upper()

        if (
            x in {
                "SWADHARMA QQ SECTOOR",
                "VARIOUS",
                "TBA",
                "ENDORSEMENT 2024",
            }
            or re.search(
                r"\b(PENYELESAIAN|HUTANG|UTANG|PIUTANG|ENDORSEMENT)\b",
                x,
                flags=re.IGNORECASE,
            )
            or re.search(
                r"\bS\s*/\s*D\b",
                x,
                flags=re.IGNORECASE,
            )
            or re.search(
                r"EX\.?\s*POLICY\s*NO",
                x,
                flags=re.IGNORECASE,
            )
            or "+" in x
            or x.count("-") >= 2
        ):
            cleaned = x

        else:

            if re.fullmatch(
                r"\d+\s*-\s*\d+",
                x,
            ):
                cleaned = x

            else:
                cleaned = re.sub(
                    r"[^A-Z0-9]",
                    "",
                    x,
                )

        if (
            cleaned
            and cleaned not in hasil_clean
        ):
            hasil_clean.append(cleaned)

    return hasil_clean


# ============================================================
# CLEAN SLIP
# ============================================================

def clean_slip(val):

    if pd.isna(val):
        return []

    original = str(val).strip()

    # ========================================================
    # KHUSUS ETIQA:
    # Format CN/... harus dibiarkan sebagai 1 slip utuh
    # P1 dan VAR hanya penanda/deskripsi
    # ========================================================
    m = re.fullmatch(
        r"(?:P1\s*/\s*)?(CN/\d+/\d+/\d+/\d+)(?:\s*-\s*VAR)?",
        original
    )

    if m:
        return [m.group(1)]

    if not original:
        return []

    s = original.upper().strip()
    
    # ========================================================
    # MULTI ARTHA GUNA - HAPUS P2 / P3 DI AKHIR
    # ========================================================
    s = re.sub(
        r"(?:\s*\+\s*P\d+)+\s*$",
        "",
        s,
        flags=re.IGNORECASE,
    ).strip()

    original = s

    # KHUSUS MULTI ARTHA GUNA:
    # TBA di akhir setelah tanda + hanya sebagai penanda,
    # jadi hapus TBA tanpa mengubah pola slip lainnya.
    if re.search(r"\+\s*TBA\s*$", s, flags=re.IGNORECASE):
        s = re.sub(
            r"\s*\+\s*TBA\s*$",
            "",
            s,
            flags=re.IGNORECASE,
        ).strip()

        original = s

    if re.search(r"\+\s*TBA\s*$", s, flags=re.IGNORECASE):
        s = re.sub(
            r"\s*\+\s*TBA\s*$",
            "",
            s,
            flags=re.IGNORECASE,
        ).strip()

    if ";" in s:

        semicolon_parts = [
            p.strip()
            for p in re.split(
                r"\s*;\s*",
                s,
            )
            if p.strip()
        ]

        if len(semicolon_parts) <= 5:
            return semicolon_parts

        return [
            _normalize_spaces(s)
        ]

    special_slip_match = re.match(
        r"^\s*(\d{10,})\s*/\s*0\s*/\s*VAR\b",
        s,
        flags=re.IGNORECASE,
    )

    if special_slip_match:

        return [
            special_slip_match.group(1)
        ]

    def fallback_original():

        return [
            _normalize_spaces(
                original.upper()
            )
        ]

    def build_repeated_number(
        base,
        suffix,
    ):

        base = re.sub(
            r"[^A-Z0-9]",
            "",
            str(base).upper(),
        )

        suffix = re.sub(
            r"[^0-9]",
            "",
            str(suffix).upper(),
        )

        if not base or not suffix:
            return ""

        if len(suffix) >= len(base):
            return suffix

        return (
            base[:-len(suffix)]
            + suffix.zfill(len(suffix))
        )
    

    # ============================================================
    # RULE BARU MULTI ARTHA GUNA - REPEATED SLIP
    # ============================================================

    # 1. FORMAT: BASE-SUFFIX-SUFFIX-SUFFIX
    # Maksimal 4 slip → expand
    # Lebih dari 4 slip → tetap original
    hyphen_parts = [
        p.strip()
        for p in re.split(r"\s*-\s*", s)
        if p.strip()
    ]

    if (
        len(hyphen_parts) >= 2
        and re.fullmatch(r"\d{10,}", hyphen_parts[0])
        and all(re.fullmatch(r"\d{1,6}", p) for p in hyphen_parts[1:])
    ):
        total_slip = len(hyphen_parts)

        if total_slip <= 4:
            base = hyphen_parts[0]
            result = [base]

            for suffix in hyphen_parts[1:]:
                candidate = build_repeated_number(base, suffix)

                if candidate and candidate not in result:
                    result.append(candidate)

            return result

        # > 4 slip → jangan expand, tapi hapus TBA
        cleaned_original = re.sub(
            r"\s*\+\s*TBA\b",
            "",
            original,
            flags=re.IGNORECASE,
        )

        return [_normalize_spaces(cleaned_original)]

        # ========================================================
    # RULE BARU MULTI ARTHA GUNA (REVISI):
    # "( EX : ... )" HANYA REFERENSI SLIP SEBELUMNYA, BUKAN
    # BAGIAN SLIP. BUANG SELURUHNYA (termasuk END kalau ada).
    #
    # Contoh:
    # 4609012200005 / ( EX : 4609011900002 ) / END
    # -> 4609012200005
    # ========================================================
    ex_paren_match = re.match(
        r"^\s*(\d{10,})\s*/\s*\(\s*EX\s*:\s*\d{10,}\s*\)",
        s,
        flags=re.IGNORECASE,
    )

    if ex_paren_match:
        return [ex_paren_match.group(1)]

    cn_slip_match = re.match(
        r"^\s*(\d{7,})\s*-\s*\d+\s*/\s*CN\b",
        s,
        flags=re.IGNORECASE,
    )

    if cn_slip_match:

        return [
            cn_slip_match.group(1)
        ]

    if s == "TBA":
        return ["TBA"]

    if s == "VARIOUS":
        return ["VARIOUS"]

    if s == "P1":
        return ["P1"]

    if re.search(
        r"\bPENYELESAIAN\b|\bSUSPENSE\b|\bHUTANG PIUTANG\b|\bUTANG PIUTANG\b",
        s,
        flags=re.IGNORECASE,
    ):

        return [
            _normalize_spaces(s)
        ]

    if "SUSPENSE VARIOUS CEDANT" in s:

        return [
            _normalize_spaces(s)
        ]

    if "PENYELESAIANSUSPEN" in s:

        return [
            _normalize_spaces(s)
        ]

    if re.search(
        r"\bENDORSEMENT\b",
        s,
        flags=re.IGNORECASE,
    ):

        return [
            _normalize_spaces(s)
        ]

    s = re.sub(
        r"^\s*TBA\s*/\s*",
        "",
        s,
        flags=re.IGNORECASE,
    ).strip()

    month_pattern = (
        r"JANUARI|FEBRUARI|MARET|APRIL|MEI|JUNI|JULI|"
        r"AGUSTUS|SEPTEMBER|OKTOBER|NOVEMBER|DESEMBER|"
        r"JANUARY|FEBRUARY|MARCH|APRIL|MAY|JUNE|JULY|"
        r"AUGUST|SEPTEMBER|OCTOBER|NOVEMBER|DECEMBER"
    )

    month_sd_match = re.search(
        rf"(\d{{7,}})\s+S\s*/?\s*D\s+(\d{{7,}})",
        s,
        flags=re.IGNORECASE,
    )

    if month_sd_match:

        slip_start = month_sd_match.group(1)
        slip_end = month_sd_match.group(2)

        start_num = int(slip_start)
        end_num = int(slip_end)

        jumlah = abs(
            end_num - start_num
        ) + 1

        if jumlah > MAX_SPLIT_COLS:

            return [
                f"{slip_start} S/D {slip_end}"
            ]

        step = (
            1
            if end_num >= start_num
            else -1
        )

        return [
            str(num).zfill(len(slip_start))
            for num in range(
                start_num,
                end_num + step,
                step,
            )
        ]

    if re.search(
        rf"\b(?:{month_pattern})\b\s+\d{{4}}",
        s,
        flags=re.IGNORECASE,
    ):

        month_numbers = re.findall(
            r"(?<!\d)\d{7,}(?!\d)",
            s,
        )

        month_numbers = list(
            dict.fromkeys(month_numbers)
        )

        if month_numbers:

            if len(month_numbers) > MAX_SPLIT_COLS:

                return [
                    _normalize_spaces(s)
                ]

            return month_numbers

    if "BORDERO" in s:

        return [
            _normalize_spaces(s)
        ]

    sd_match = re.search(
        r"(\d{7,})\s+S\s*/?\s*D\s+(\d{7,})",
        s,
        flags=re.IGNORECASE,
    )

    if sd_match:

        slip_start = sd_match.group(1)
        slip_end = sd_match.group(2)

        start_num = int(slip_start)
        end_num = int(slip_end)

        step = (
            1
            if end_num >= start_num
            else -1
        )

        jumlah = abs(
            end_num - start_num
        ) + 1

        if jumlah > MAX_SPLIT_COLS:

            return [
                _normalize_spaces(s)
            ]

        return [
            str(num).zfill(len(slip_start))
            for num in range(
                start_num,
                end_num + step,
                step,
            )
        ]

    if re.fullmatch(
        r"(JANUARI|FEBRUARI|MARET|APRIL|MEI|JUNI|JULI|"
        r"AGUSTUS|SEPTEMBER|OKTOBER|NOVEMBER|DESEMBER)"
        r"\s+\d{4}",
        s,
        re.IGNORECASE,
    ):

        return [
            _normalize_spaces(s)
        ]

    sd_match = re.search(
        r"^"
        r"(\d{10,})"
        r"\s*-\s*"
        r"(\d{1,6})"
        r"\s*S\s*/?\s*D\s*"
        r"(\d{1,6})"
        r"$",
        s,
        flags=re.IGNORECASE,
    )

    if sd_match:

        base = sd_match.group(1)

        suffix_start = sd_match.group(2)
        suffix_end = sd_match.group(3)

        result = [
            base,
            build_repeated_number(
                base,
                suffix_start,
            ),
            build_repeated_number(
                base,
                suffix_end,
            ),
        ]

        result = [
            x
            for x in result
            if x
        ]

        result = list(
            dict.fromkeys(result)
        )

        if len(result) > 5:
            return fallback_original()

        return result

    s = re.sub(
        r"^\s*P1\s*/\s*",
        "",
        s,
        flags=re.IGNORECASE,
    )

    s = re.sub(
        r"\s*/\s*P1\s*$",
        "",
        s,
        flags=re.IGNORECASE,
    )

    cn_numbers = re.findall(
        r"(?<!\d)(\d{4,})\s*/\s*CN\b",
        s,
        flags=re.IGNORECASE,
    )

    if cn_numbers:

        result = list(
            dict.fromkeys(cn_numbers)
        )

        if len(result) > 5:
            return fallback_original()

        return result

    s = re.sub(
        r"(?<!\d)(\d{10,})\s*-\s*(\d)(?!\d)",
        r"\1\2",
        s,
    )

    plus_repeat_match = re.fullmatch(
        r"(\d{10,})-(\d{1,6})(?:\+(\d{1,6}))+",
        s,
    )

    if plus_repeat_match:

        base = plus_repeat_match.group(1)

        suffix_part = s.split(
            "-",
            1,
        )[1]

        suffixes = re.findall(
            r"\d{1,6}",
            suffix_part,
        )

        result = [base]

        for suffix in suffixes:

            candidate = build_repeated_number(
                base,
                suffix,
            )

            if (
                candidate
                and candidate not in result
            ):
                result.append(candidate)

        if len(result) > 5:
            return fallback_original()

        return result

    repeated_parts = re.split(
        r"\s*[;/]\s*",
        s,
    )

    if len(repeated_parts) > 1:

        repeated_result = []
        current_base = None

        for part in repeated_parts:

            part = part.strip()

            if not part:
                continue

            if part in {
                "VARIOUS",
                "VAR",
                "TBA",
            }:
                continue

            m = re.fullmatch(
                r"(\d{10,})\s*-\s*(\d{1,6})",
                part,
            )

            if m:

                base = m.group(1)
                suffix = m.group(2)

                current_base = base

                if base not in repeated_result:
                    repeated_result.append(base)

                candidate = build_repeated_number(
                    base,
                    suffix,
                )

                if (
                    candidate
                    and candidate not in repeated_result
                ):
                    repeated_result.append(candidate)

                continue

            if re.fullmatch(
                r"\d{10,}",
                part,
            ):

                if part not in repeated_result:
                    repeated_result.append(part)

                current_base = part

                continue

            if (
                current_base
                and re.fullmatch(
                    r"\d{1,6}",
                    part,
                )
            ):

                candidate = build_repeated_number(
                    current_base,
                    part,
                )

                if (
                    candidate
                    and candidate not in repeated_result
                ):
                    repeated_result.append(candidate)

                continue

        if repeated_result:

            if len(repeated_result) > 5:
                return fallback_original()

            return repeated_result

    if "+" in s:

        parts = [
            x.strip()
            for x in re.split(
                r"\s*\+\s*",
                s,
            )
            if x.strip()
        ]

        if (
            parts
            and re.fullmatch(
                r"\d{10,}",
                parts[0],
            )
        ):

            first = parts[0]
            prefix = first[:-4]

            result = [first]

            for part in parts[1:]:

                if part in {
                    "VAR",
                    "VARIOUS",
                    "TBA",
                    "P1",
                }:
                    continue

                if re.fullmatch(
                    r"\d{1,4}",
                    part,
                ):

                    result.append(
                        prefix
                        + part.zfill(4)
                    )

                    continue

                if re.fullmatch(
                    r"\d{10,}",
                    part,
                ):

                    result.append(part)

                    continue

            result = list(
                dict.fromkeys(result)
            )

            if len(result) > 5:
                return fallback_original()

            return result

    s = re.sub(
        r"\s*/\s*307\s*$",
        "",
        s,
        flags=re.IGNORECASE,
    )

    s = re.sub(
        r"\bVARIOUS\b",
        "",
        s,
        flags=re.IGNORECASE,
    )

    s = re.sub(
        r"\bVAR\b",
        "",
        s,
        flags=re.IGNORECASE,
    )

    s = re.sub(
        r"\bTBA\b",
        "",
        s,
        flags=re.IGNORECASE,
    )

    s = s.strip(
        " /+-;"
    )

    if re.search(
        r"SUB\s*CLASS\s+OF\s+BUSINESS",
        original,
        flags=re.IGNORECASE,
    ):

        return [
            _normalize_spaces(
                original.upper()
            )
        ]

    long_numbers = re.findall(
        r"(?<!\d)\d{10,}(?!\d)",
        s,
    )

    if long_numbers:

        long_numbers = list(
            dict.fromkeys(long_numbers)
        )

        if len(long_numbers) > 5:
            return fallback_original()

        return long_numbers

    if re.fullmatch(
        r"\d{4}\.[A-Z]+\.\d+\.\d{4}",
        original.upper(),
    ):

        return [
            _normalize_spaces(
                original.upper()
            )
        ]

    sd_match = re.search(
        r"(?<!\d)"
        r"(\d{7,})"
        r"\s+S\s*/?\s*D\s+"
        r"(\d{1,})"
        r"(?!\d)",
        s,
        flags=re.IGNORECASE,
    )

    if sd_match:

        start_raw = sd_match.group(1)
        end_raw = sd_match.group(2)

        if len(start_raw) == len(end_raw):

            start_num = int(start_raw)
            end_num = int(end_raw)

            step = (
                1
                if end_num >= start_num
                else -1
            )

            jumlah = abs(
                end_num - start_num
            ) + 1

            if jumlah > 5:
                return fallback_original()

            return [
                str(num).zfill(
                    len(start_raw)
                )
                for num in range(
                    start_num,
                    end_num + step,
                    step,
                )
            ]

        return fallback_original()

    parts = re.split(
        r"\s*[;/]\s*",
        s,
    )

    result = []

    for part in parts:

        part = part.strip()

        if not part:
            continue

        if part in {
            "VAR",
            "VARIOUS",
            "TBA",
            "P1",
        }:
            continue

        cn_match = re.search(
            r"(?<!\d)(\d{4,})\s*CN\b",
            part,
            flags=re.IGNORECASE,
        )

        if cn_match:

            number = cn_match.group(1)

            if number not in result:
                result.append(number)

            continue

        if re.fullmatch(
            r"\d+",
            part,
        ):

            if part not in result:
                result.append(part)

            continue

        part = re.sub(
            r"\b(?:VAR|VARIOUS|TBA)\b",
            "",
            part,
            flags=re.IGNORECASE,
        ).strip()

        if not part:
            continue

        cleaned = _normalize_spaces(part)

        if not cleaned:
            continue

        if cleaned.upper() in {
            "VAR",
            "VARIOUS",
            "TBA",
        }:
            continue

        if cleaned not in result:
            result.append(cleaned)

    if len(result) > 5:
        return fallback_original()

    if not result:

        return [
            _normalize_spaces(
                original.upper()
            )
        ]

    return result


# ============================================================
# CLEAN INSURED
# ============================================================

def clean_insured(val) -> list:
    if pd.isna(val):
        return []

    val = str(val).strip()

    if not val:
        return []

    cleaned = []

    # Breakdown berdasarkan:
    # /  -> pemisah perusahaan
    # ,  -> pemisah perusahaan / suffix PT
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

# ============================================================
# CERTIFICATE
# ============================================================
def _extract_certificate_list(polis_ori):
    """
    CERTIFICATE MENEMPEL KE 1 POLIS.

    Jadi:
        1010100822000097 - 000001 , 000002 , 000003

    HASIL:
        ["000001, 000002, 000003"]

    BUKAN:
        ["000001", "000002", "000003"]

    Contoh:
        1010100823000179 - CERTIF : 000001 , 000002
        -> ["000001, 000002"]

        P1 / 1071031124000024 - 000251 SD 0000131
        -> ["000251 SD 0000131"]

    Certificate selalu dianggap milik polis induknya.
    """

    if pd.isna(polis_ori):
        return []

    polis = str(polis_ori).strip().upper()

    if not polis:
        return []

    polis = _strip_ex_reference(polis)

    if not polis:
        return []

    # Jangan ambil certificate dari data suspense / endorsement
    if re.search(
        r"\bSUSPENSE\b|\bENDORSEMENT\b",
        polis,
        flags=re.IGNORECASE,
    ):
        return []

    # ========================================================
    # FORMAT 0 (KHUSUS MULTI ARTHA GUNA - DATA 2 OSBAL)
    #
    # BASE[.-]CERT1 S/D [BASE[.-]]CERT2
    #
    # Satu polis, DUA certificate saja (TIDAK di-expand jadi
    # range penuh seperti FORMAT 4 di bawah).
    #
    # Contoh:
    # 02031121000065-000366 S/D 02031121000065-000385
    # -> certificate 1 = 000366, 000385
    #
    # 40010922042857.001214 S/D 001224
    # -> certificate 1 = 001214, 001224
    #
    # 40010924046082-001199 S/D 40010924046082.001210
    # -> certificate 1 = 001199, 001210
    #
    # 40010925046051.000001 s/d 40010925046051.001359
    # -> certificate 1 = 000001, 001359
    # ========================================================

    two_cert_match = re.fullmatch(
        r"""
        \s*
        (\d{10,})
        [.\-]
        (\d{1,6})
        \s*S\s*/?\s*D\s*
        (?:\1[.\-])?
        (\d{1,7})
        \s*
        """,
        polis,
        flags=re.IGNORECASE | re.VERBOSE,
    )

    if two_cert_match:

        cert1 = two_cert_match.group(2).zfill(6)
        cert2 = two_cert_match.group(3).zfill(6)

        return [
            f"{cert1}, {cert2}"
        ]

    # ========================================================
    # FORMAT 1
    #
    # 1010100823000179 - CERTIF : 000001 , 000002
    #
    # HASIL:
    # certificate 1 = 000001, 000002
    # ========================================================

    certif_match = re.fullmatch(
        r"""
        \s*
        \d{10,}
        \s*-\s*
        CERTIF\s*:\s*
        (\d{1,6})
        \s*,\s*
        (\d{1,6})
        (?:
            \s*,\s*
            (\d{1,6})
        )?
        \s*
        """,
        polis,
        flags=re.IGNORECASE | re.VERBOSE,
    )

    if certif_match:

        certificates = [
            value.zfill(6)
            for value in certif_match.groups()
            if value
        ]

        return [
            ", ".join(certificates)
        ]

    # ========================================================
    # FORMAT 2
    #
    # 1010100822000097 - 000001 , 000002 , 000003
    #
    # HASIL:
    # certificate 1 = 000001, 000002, 000003
    # ========================================================

    normal_match = re.fullmatch(
        r"""
        \s*
        \d{10,}
        \s*-\s*
        (\d{1,6})
        \s*,\s*
        (\d{1,6})
        (?:
            \s*,\s*
            (\d{1,6})
        )?
        \s*
        """,
        polis,
        flags=re.IGNORECASE | re.VERBOSE,
    )

    if normal_match:

        certificates = [
            value.zfill(6)
            for value in normal_match.groups()
            if value
        ]

        return [
            ", ".join(certificates)
        ]

    # ========================================================
    # FORMAT 3
    #
    # P1 / 1071031124000024 - 000251 SD 0000131
    #
    # HASIL:
    # certificate 1 = 000251 SD 0000131
    #
    # BUKAN:
    # certificate 1 = 000251
    # certificate 2 = 0000131
    # ========================================================

    range_match = re.fullmatch(
        r"""
        \s*
        P1\s*/\s*
        \d{10,}
        \s*-\s*
        (\d{1,6})
        \s*S\s*/?\s*D\s*
        (\d{1,7})
        \s*
        """,
        polis,
        flags=re.IGNORECASE | re.VERBOSE,
    )

    if range_match:

        start = range_match.group(1)
        end = range_match.group(2)

        if len(start) > 6:
            return []

        return [
            f"{start.zfill(6)} SD {end.zfill(7)}"
        ]

    # ========================================================
    # FORMAT 4 (KHUSUS MULTI ARTHA GUNA)
    #
    # BASE - NNNNNN S/D NNNNNN (TANPA prefix "P1/")
    # Pemisah "-" boleh ada atau tidak (kadang cuma spasi).
    #
    # REVISI: pemisah antar 2 nomor certificate boleh juga
    # tanda "-" biasa, tanpa teks "S/D" eksplisit.
    #
    # Nomor certificate dibatasi 5-7 digit supaya TIDAK
    # tertukar dengan suffix pendek (1-4 digit) pada rantai
    # perulangan polis (mis. "02010922005704-1073-186" itu
    # BUKAN certificate, itu 3 perulangan polis).
    #
    # Contoh:
    # 40080521000122 - 000001 S/D 000010
    # 03120622000013 - 000001 s/d 000006
    # 45031122001162 000064 S/D 000071
    # 02031125000018 000637-000745
    #
    # Rule sama seperti certificate dari CLSDT_SERTF_NO:
    # - jumlah <= 3 -> expand pakai koma
    # - jumlah > 3  -> tetap format "AWAL SD AKHIR"
    # ========================================================

    range_no_prefix_match = re.fullmatch(
        r"""
        \s*
        \d{10,}
        [\s-]+
        (\d{5,7})
        \s*(?:S\s*/?\s*D\s*|-\s*)
        (\d{5,7})
        (?:\s*/\s*(?:VAR|VARIOUS))?
        \s*
        """,
        polis,
        flags=re.IGNORECASE | re.VERBOSE,
    )


    if range_no_prefix_match:

        start_raw, end_raw = range_no_prefix_match.groups()

        start_num = int(start_raw)
        end_num = int(end_raw)

        step = (
            1
            if end_num >= start_num
            else -1
        )

        jumlah = abs(
            end_num - start_num
        ) + 1

        if jumlah <= 3:

            return [
                ", ".join(
                    str(num).zfill(6)
                    for num in range(
                        start_num,
                        end_num + step,
                        step,
                    )
                )
            ]

        return [
            f"{start_raw.zfill(6)} SD {end_raw.zfill(6)}"
        ]

    # ========================================================
    # FORMAT 5 (KHUSUS MULTI ARTHA GUNA)
    #
    # BASE - NNNNNN (suffix persis 6 digit, tanpa S/D)
    # Boleh diikuti penanda "/ VAR" atau "/ VARIOUS" (dibuang).
    #
    # Contoh:
    # 40030121085615 - 000108   -> certificate 1 = 000108
    # 45013022003422 - 000001   -> certificate 1 = 000001
    # 02031126000011 - 000001 / VAR -> certificate 1 = 000001
    # ========================================================

    single_sertf_match = re.fullmatch(
        r"""
        \s*
        \d{10,}
        \s*-\s*
        (\d{6})
        (?:\s*/\s*(?:VAR|VARIOUS))?
        \s*
        """,
        polis,
        flags=re.IGNORECASE | re.VERBOSE,
    )

    if single_sertf_match:

        return [
            single_sertf_match.group(1)
        ]

    # ========================================================
    # FORMAT 6 (KHUSUS MULTI ARTHA GUNA - DATA 2 OSBAL)
    #
    # BASE RANGE,EXTRA(/VARIOUS)
    #
    # Contoh:
    # 02031125000018 000455-000536,000467/VARIOUS
    # -> certificate 1 = 000455 SD 000536, 000467
    # ========================================================

    cert_range_extra_match = re.fullmatch(
        r"""
        \s*
        \d{10,}
        [\s-]+
        (\d{5,7})
        \s*-\s*
        (\d{5,7})
        \s*,\s*
        (\d{5,7})
        (?:\s*/\s*(?:VAR|VARIOUS))?
        \s*
        """,
        polis,
        flags=re.IGNORECASE | re.VERBOSE,
    )

    if cert_range_extra_match:

        start_raw, end_raw, extra_raw = (
            cert_range_extra_match.groups()
        )

        return [
            f"{start_raw.zfill(6)} SD {end_raw.zfill(6)}, {extra_raw.zfill(6)}"
        ]

    # ========================================================
    # FORMAT 7 (KHUSUS MULTI ARTHA GUNA - DATA 2 OSBAL)
    #
    # P1 / BASE1 - CERT & BASE2
    # Certificate hanya menempel ke BASE1 (polis pertama).
    #
    # Contoh:
    # P1 / 45100426000023 - 000001 & 45100426000034
    # -> certificate 1 = 000001
    # ========================================================

    p1_amp_match = re.fullmatch(
        r"""
        \s*
        P1\s*/\s*
        \d{10,}
        \s*-\s*
        (\d{1,6})
        \s*&\s*
        \d{10,}
        \s*
        """,
        polis,
        flags=re.IGNORECASE | re.VERBOSE,
    )

    if p1_amp_match:

        return [
            p1_amp_match.group(1).zfill(6)
        ]

    # ========================================================
    # FORMAT 8 (KHUSUS MULTI ARTHA GUNA - DATA 2 OSBAL)
    #
    # BASE-CERT1+CERT2+CERT3+...
    # (semua angka certificate, 5-7 digit tiap angka, supaya
    # tidak tertukar dengan rantai perulangan polis yang
    # angkanya pendek, mis. "500130230005-12+23+34+45" itu
    # BUKAN certificate, itu 5 perulangan polis.)
    #
    # Contoh:
    # 03010924000363-000001+000002+000003+000004
    # -> certificate 1 = 000001 SD 000004
    #
    # Rule sama seperti format lain:
    # - jumlah <= 3 -> expand pakai koma
    # - jumlah > 3  -> tetap format "AWAL SD AKHIR"
    # ========================================================

    plus_chain_match = re.fullmatch(
        r"""
        \s*
        \d{10,}
        \s*-\s*
        (\d{5,7}(?:\s*\+\s*\d{5,7})+)
        \s*
        """,
        polis,
        flags=re.IGNORECASE | re.VERBOSE,
    )

    if plus_chain_match:

        raw_certs = re.findall(
            r"\d{5,7}",
            plus_chain_match.group(1),
        )

        if len(raw_certs) <= 3:

            return [
                ", ".join(
                    c.zfill(6)
                    for c in raw_certs
                )
            ]

        return [
            f"{raw_certs[0].zfill(6)} SD {raw_certs[-1].zfill(6)}"
        ]

    # ========================================================
    # FORMAT 9 (KHUSUS MULTI ARTHA GUNA - DATA 2 OSBAL)
    #
    # BASE-suf1+suf2+...+CERT1+CERT2 (CAMPURAN)
    #
    # Suffix pendek (1-4 digit) di depan = perulangan polis,
    # BUKAN certificate. Angka panjang (5-7 digit) di BELAKANG
    # = certificate beneran.
    #
    # Contoh:
    # 46013023011-836+847+858+869+00039+000041
    # -> 836/847/858/869 = perulangan polis (diabaikan di sini)
    # -> certificate 1 = 000039, 000041
    # ========================================================

    mixed_chain_match = re.fullmatch(
        r"""
        \s*
        \d{10,}
        (?:[-+]\d{1,4})*
        ((?:\+\d{5,7})+)
        \s*
        """,
        polis,
        flags=re.IGNORECASE | re.VERBOSE,
    )

    if mixed_chain_match:

        raw_certs = re.findall(
            r"\d{5,7}",
            mixed_chain_match.group(1),
        )

        if len(raw_certs) <= 3:

            return [
                ", ".join(
                    c.zfill(6)
                    for c in raw_certs
                )
            ]

        return [
            f"{raw_certs[0].zfill(6)} SD {raw_certs[-1].zfill(6)}"
        ]

    # ========================================================
    # SELAIN FORMAT DI ATAS
    # BUKAN CERTIFICATE
    # ========================================================

    return []


# ============================================================
# CERTIFICATE DARI CLSDT_SERTF_NO
#
# KHUSUS MULTI ARTHA GUNA:
# Certificate diprioritaskan dari CLSDT_SERTF_NO.
#
# Rule valid certificate:
# - Angka 1-6 digit -> valid, zero padding jadi 6 digit.
# - Lebih dari 6 digit -> BUKAN certificate, diabaikan.
# - TBA / VAR / VARIOUS -> tidak valid.
# - Support format "S/D" atau "SD" (range).
#   Jika jumlah anggota range <= 3 -> expand pakai koma.
#   Jika jumlah anggota range > 3  -> tetap format "AWAL SD AKHIR".
# - Beberapa certificate dipisah koma -> masing-masing dinormalize.
#
# Certificate selalu ditempel ke polis pertama (certificate 1),
# konsisten dengan _extract_certificate_list() di atas.
# ============================================================

def _normalize_single_certificate_number(tok: str):

    tok = tok.strip()

    if not tok:
        return None

    digits = re.sub(
        r"[^0-9]",
        "",
        tok,
    )

    if not digits:
        return None

    # Lebih dari 6 digit -> bukan certificate
    if len(digits) > 6:
        return None

    return digits.zfill(6)


def _extract_certificate_from_clsdt_sertf(value) -> list:

    if pd.isna(value):
        return []

    text = str(value).strip().upper()

    if not text:
        return []

    if text in {
        "TBA",
        "VAR",
        "VARIOUS",
    }:
        return []

    # ========================================================
    # FORMAT RANGE: 000001 SD 000005 / 000001 S/D 000005
    # ========================================================

    sd_match = re.fullmatch(
        r"""
        \s*
        (\d{1,6})
        \s*S\s*/?\s*D\s*
        (\d{1,6})
        \s*
        """,
        text,
        flags=re.IGNORECASE | re.VERBOSE,
    )

    if sd_match:

        start_raw, end_raw = sd_match.groups()

        start_num = int(start_raw)
        end_num = int(end_raw)

        step = (
            1
            if end_num >= start_num
            else -1
        )

        jumlah = abs(
            end_num - start_num
        ) + 1

        # Range <= 3 certificate -> expand pakai koma
        if jumlah <= 3:

            return [
                ", ".join(
                    str(num).zfill(6)
                    for num in range(
                        start_num,
                        end_num + step,
                        step,
                    )
                )
            ]

        # Range > 3 certificate -> tetap format "AWAL SD AKHIR"
        return [
            f"{start_raw.zfill(6)} SD {end_raw.zfill(6)}"
        ]

    # ========================================================
    # BEBERAPA CERTIFICATE DIPISAH KOMA
    # (atau hanya satu certificate tunggal)
    # ========================================================

    parts = [
        p.strip()
        for p in text.split(",")
        if p.strip()
    ]

    if not parts:
        return []

    normalized = []

    for part in parts:

        if part in {
            "TBA",
            "VAR",
            "VARIOUS",
        }:
            continue

        num = _normalize_single_certificate_number(
            part
        )

        if num:
            normalized.append(num)

    if not normalized:
        return []

    return [
        ", ".join(normalized)
    ]


# ============================================================
# VALIDASI CLSDT
# ============================================================

def _is_valid_clstd(val) -> bool:

    if pd.isna(val):
        return False

    val = str(
        val
    ).strip().upper()

    if not val:
        return False

    # ========================================================
    # INVALID CLSDT
    # ========================================================
    # Jika CLSDT berisi penanda P<n> CANCEL / P<n>/CANCEL --
    # baik satu maupun berulang seperti
    # "P1 CANCEL P2 CANCEL P3 CANCEL" -- jangan gunakan CLSDT.
    # Harus fallback ke POLIS ORI / SLIP ORI.
    # ========================================================

    if re.search(
        r"P\d+\s*/?\s*CANCEL",
        val,
        flags=re.IGNORECASE,
    ):
        return False

    if val in {
        "TBA",
        "VAR",
        "VARIOUS",
    }:
        return False

    return True

# ============================================================
# FIX LEADING ZERO CLSDT POLICY NO / SLIP NO
# ============================================================
# REVISI: sebagian CLSDT_POLICY_NO / CLSDT_SLIP_NO kehilangan
# digit 0 di depan (biasa karena Excel membaca sebagai angka),
# sehingga "0701....." terbaca jadi "701.....".
#
# Kalau nilainya (setelah dibuang ".0" sisa float) diawali
# "701", tambahkan kembali "0" di depan.
#
# Contoh:
# 701912200001 -> 0701912200001
# ============================================================

def _fix_clstd_leading_zero(val):

    if pd.isna(val):
        return val

    text = str(val).strip()

    if not text:
        return val

    # buang sisa ".0" dari nilai numerik yang dibaca sebagai float
    text = re.sub(r"\.0+$", "", text)

    if text.upper().startswith("701"):
        return "0" + text

    return val

def _get_clstd_source(
    clstd_value,
    fac_value,
):

    clstd_value = _fix_clstd_leading_zero(
        clstd_value
    )

    if _is_valid_clstd(
        clstd_value
    ):
        return (
            clstd_value,
            "CLSDT",
        )

    return (
        fac_value,
        "FAC",
    )


# ============================================================
# NORMALIZE CEDANT
# ============================================================

def _normalize_cedant(val) -> str:

    if pd.isna(val):
        return ""

    return re.sub(
        r"[^A-Z0-9]",
        "",
        str(val).upper(),
    )


# ============================================================
# INSERT CLEAN COLUMNS
# ============================================================

def _insert_clean_columns(
    df: pd.DataFrame,
    all_lists: list,
    prefix: str,
    max_cols: int,
) -> list:

    added = []

    # SELALU BUAT 5 KOLOM
    for i in range(
        1,
        MAX_SPLIT_COLS + 1,
    ):

        col_name = (
            f"clean {prefix} {i}"
        )

        df[col_name] = [
            (
                lst[i - 1]
                if i - 1 < len(lst)
                else None
            )
            for lst in all_lists
        ]

        added.append(col_name)

    return added


# ============================================================
# PROCESS DATA
# ============================================================

def process_data(
    input_file: Path,
    output_file: Path,
) -> None:

    # ========================================================
    # 1. BACA SELURUH DATA
    # ========================================================

    print(
        f"[1/5] Membaca data dari: "
        f"{input_file}"
    )

    from io import StringIO

    try:
        from xlsx2csv import Xlsx2csv
    except ImportError:
        raise ImportError(
            "Library xlsx2csv belum terinstall. "
            "Install dengan: pip install xlsx2csv"
        )

    print(
        "[1/5] Mode pembacaan Excel cepat..."
    )

    buffer = StringIO()

    Xlsx2csv(
        str(input_file),
        skip_empty_lines=False,
    ).convert(buffer)

    buffer.seek(0)

    df = pd.read_csv(
        buffer,
        dtype=str,
        low_memory=False,
    )

    original_row_count = len(df)

    print(
        f"      Total baris keseluruhan: "
        f"{original_row_count:,}"
    )

    # ========================================================
    # 2. CEK KOLOM
    # ========================================================

    required_cols = [
        CEDANT_COL,
        POLIS_COL,
        SLIP_COL,
        INSURED_COL,
        "CLSDT_POLICY_NO",
        "CLSDT_SLIP_NO",
        CLSDT_SERTF_COL,
    ]

    for col in required_cols:

        if col not in df.columns:

            print(
                f"\n[ERROR] Kolom '{col}' "
                f"tidak ditemukan!"
            )

            print(
                "Kolom tersedia:"
            )

            print(
                list(df.columns)
            )

            raise RuntimeError(
                f"Kolom wajib tidak ditemukan: {col}"
            )

    # ========================================================
    # 3. IDENTIFIKASI CEDANT
    #
    # PENTING:
    # is_etiqa HANYA MASK UNTUK MENENTUKAN BARIS
    # YANG DICLEANING.
    #
    # TIDAK ADA FILTER DATAFRAME.
    # ========================================================

    cedant_target = _normalize_cedant(
        CEDANT_VALUE
    )

    cedant_normalized = (
        df[CEDANT_COL]
        .map(_normalize_cedant)
    )

    is_etiqa = (
        cedant_normalized
        .eq(cedant_target)
    )

    jumlah_etiqa = int(
        is_etiqa.sum()
    )

    jumlah_non_etiqa = (
        len(df)
        - jumlah_etiqa
    )

    print(
        "[2/5] Identifikasi cedant:"
    )

    print(
        f"      Target cleansing : "
        f"{CEDANT_VALUE}"
    )

    print(
        f"      Baris target     : "
        f"{jumlah_etiqa:,}"
    )

    print(
        f"      Baris cedant lain: "
        f"{jumlah_non_etiqa:,}"
    )

    print(
        "      Semua baris tetap dipertahankan."
    )

    if jumlah_etiqa == 0:

        print(
            "\n[WARNING] Cedant target tidak ditemukan."
        )

        print(
            "          Tidak ada baris yang menjalankan cleansing."
        )

        print(
            "          Semua data tetap dipertahankan."
        )

    # ========================================================
    # RENAME KOLOM
    # ========================================================

    df.rename(
        columns={
            POLIS_COL: "polis_ori",
            SLIP_COL: "slip_ori",
            INSURED_COL: "insured_ori",
        },
        inplace=True,
    )

    # ========================================================
    # 4. CLEANING
    # ========================================================

    print(
        "[3/5] Menjalankan cleansing "
        "hanya untuk cedant target..."
    )

    polis_ori_list = (
        df["polis_ori"].tolist()
    )

    slip_ori_list = (
        df["slip_ori"].tolist()
    )

    insured_ori_list = (
        df["insured_ori"].tolist()
    )

    clsdt_polis_list = (
        df["CLSDT_POLICY_NO"].tolist()
    )

    clsdt_slip_list = (
        df["CLSDT_SLIP_NO"].tolist()
    )

    clsdt_sertf_list = (
        df[CLSDT_SERTF_COL].tolist()
    )

    is_etiqa_list = (
        is_etiqa.tolist()
    )

    # ========================================================
    # HASIL CLEANING
    # ========================================================

    all_clean_polis = []
    all_clean_slip = []
    all_clean_ins = []

    # CERTIFICATE PER CLEAN POLIS
    all_certificates_per_polis = []

    total_etiqa_processed = 0

    total_clstd_polis_used = 0
    total_fallback_polis_fac = 0

    total_clstd_slip_used = 0
    total_fallback_slip_fac = 0

    total_clstd_sertf_used = 0
    total_fallback_sertf_polis = 0

    total_rows = len(df)

    # ========================================================
    # LOOP SEMUA BARIS
    # ========================================================

    for idx in range(total_rows):

        if (
            idx + 1
        ) % 50_000 == 0:

            print(
                f"      Progress: "
                f"{idx + 1:,} / "
                f"{total_rows:,} "
                f"baris diproses..."
            )

        # ====================================================
        # HANYA ETIQA
        # ====================================================

        if is_etiqa_list[idx]:

            total_etiqa_processed += 1

            # ------------------------------------------------
            # POLIS
            # ------------------------------------------------

            p_ori = (
                polis_ori_list[idx]
            )

            c_p_no = (
                clsdt_polis_list[idx]
            )

            polis_source, polis_source_name = (
                _get_clstd_source(
                    c_p_no,
                    p_ori,
                )
            )

            if (
                polis_source_name
                == "CLSDT"
            ):

                total_clstd_polis_used += 1

            else:

                total_fallback_polis_fac += 1

            c_polis = clean_polis(
                polis_source
            )

            # ------------------------------------------------
            # CERTIFICATE
            #
            # Prioritas:
            # 1. CLSDT_SERTF_NO (jika valid).
            # 2. Fallback ke extraction dari polis_ori
            #    (rule existing _extract_certificate_list),
            #    hanya jika CLSDT_SERTF_NO kosong/NaN/
            #    TBA/VAR/VARIOUS atau gagal dinormalize.
            #
            # Certificate TIDAK diambil sembarangan dari polis;
            # breakdown polis dan certificate diproses terpisah.
            # ------------------------------------------------

            c_sertf_no = (
                clsdt_sertf_list[idx]
            )

            if _is_valid_clstd(
                c_sertf_no
            ):

                c_certificates = (
                    _extract_certificate_from_clsdt_sertf(
                        c_sertf_no
                    )
                )

                if c_certificates:
                    total_clstd_sertf_used += 1
                else:
                    c_certificates = (
                        _extract_certificate_list(
                            polis_source
                        )
                    )
                    total_fallback_sertf_polis += 1

            else:

                c_certificates = (
                    _extract_certificate_list(
                        polis_source
                    )
                )

                total_fallback_sertf_polis += 1

            # Maksimal 3 certificate
            c_certificates = c_certificates[:3]
            # ------------------------------------------------
            # SLIP
            # ------------------------------------------------

            s_ori = (
                slip_ori_list[idx]
            )

            c_s_no = (
                clsdt_slip_list[idx]
            )

            slip_source, slip_source_name = (
                _get_clstd_source(
                    c_s_no,
                    s_ori,
                )
            )

            if (
                slip_source_name
                == "CLSDT"
            ):

                total_clstd_slip_used += 1

            else:

                total_fallback_slip_fac += 1

            c_slip = clean_slip(
                slip_source
            )

            # ------------------------------------------------
            # INSURED
            # ------------------------------------------------

            c_ins = clean_insured(
                insured_ori_list[idx]
            )

        # ====================================================
        # CEDANT LAIN
        #
        # JANGAN DICLEAN.
        # ====================================================

        else:

            c_polis = []
            c_slip = []
            c_ins = []
            c_certificates = []

        # ====================================================
        # SIMPAN HASIL
        # ====================================================

        all_clean_polis.append(
            c_polis
        )

        all_clean_slip.append(
            c_slip
        )

        all_clean_ins.append(
            c_ins
        )

        all_certificates_per_polis.append(
            c_certificates
        )

    # ========================================================
    # SUMMARY CLEANING
    # ========================================================

    print(
        "      Selesai diproses!"
    )

    print(
        f"      -> Baris target yang dicleansing : "
        f"{total_etiqa_processed:,}"
    )

    print(
        f"      -> Baris cedant lain            : "
        f"{jumlah_non_etiqa:,}"
    )

    print(
        f"      -> CLSDT POLICY digunakan       : "
        f"{total_clstd_polis_used:,}"
    )

    print(
        f"      -> POLICY fallback ke FAC       : "
        f"{total_fallback_polis_fac:,}"
    )

    print(
        f"      -> CLSDT SLIP digunakan         : "
        f"{total_clstd_slip_used:,}"
    )

    print(
        f"      -> SLIP fallback ke FAC         : "
        f"{total_fallback_slip_fac:,}"
    )

    print(
        f"      -> CLSDT_SERTF_NO digunakan     : "
        f"{total_clstd_sertf_used:,}"
    )

    print(
        f"      -> Certificate fallback ke polis: "
        f"{total_fallback_sertf_polis:,}"
    )

    # ========================================================
    # 5. SUSUN KOLOM OUTPUT
    # ========================================================

    print(
        "[4/5] Menyusun kolom output..."
    )

    new_columns = []

    # ========================================================
    # LOOP KOLOM EXISTING
    # ========================================================

    for col in df.columns:

        # ----------------------------------------------------
        # POLIS
        # ----------------------------------------------------

        if col == "polis_ori":

            # =================================================
            # WAJIB:
            #
            # polis_ori
            # clean polis 1
            # certificate 1
            # clean polis 2
            # certificate 2
            # clean polis 3
            # certificate 3
            # clean polis 4
            # clean polis 5
            # =================================================

            new_columns.append(
                "polis_ori"
            )

            # -----------------------------------------------
            # CLEAN POLIS 1
            # -----------------------------------------------

            df["clean polis 1"] = [
                (
                    lst[0]
                    if len(lst) > 0
                    else None
                )
                for lst in all_clean_polis
            ]

            new_columns.append(
                "clean polis 1"
            )

            # -----------------------------------------------
            # CERTIFICATE 1
            # -----------------------------------------------

            df["certificate 1"] = [
                (
                    certs[0]
                    if len(certs) > 0
                    else ""
                )
                for certs in all_certificates_per_polis
            ]

            new_columns.append(
                "certificate 1"
            )

            # -----------------------------------------------
            # CLEAN POLIS 2
            # -----------------------------------------------

            df["clean polis 2"] = [
                (
                    lst[1]
                    if len(lst) > 1
                    else None
                )
                for lst in all_clean_polis
            ]

            new_columns.append(
                "clean polis 2"
            )

            # -----------------------------------------------
            # CERTIFICATE 2
            # -----------------------------------------------

            df["certificate 2"] = [
                (
                    certs[1]
                    if len(certs) > 1
                    else ""
                )
                for certs in all_certificates_per_polis
            ]

            new_columns.append(
                "certificate 2"
            )

            # -----------------------------------------------
            # CLEAN POLIS 3
            # -----------------------------------------------

            df["clean polis 3"] = [
                (
                    lst[2]
                    if len(lst) > 2
                    else None
                )
                for lst in all_clean_polis
            ]

            new_columns.append(
                "clean polis 3"
            )

            # -----------------------------------------------
            # CERTIFICATE 3
            # -----------------------------------------------

            df["certificate 3"] = [
                (
                    certs[2]
                    if len(certs) > 2
                    else ""
                )
                for certs in all_certificates_per_polis
            ]

            new_columns.append(
                "certificate 3"
            )

            # -----------------------------------------------
            # CLEAN POLIS 4
            # -----------------------------------------------

            df["clean polis 4"] = [
                (
                    lst[3]
                    if len(lst) > 3
                    else None
                )
                for lst in all_clean_polis
            ]

            new_columns.append(
                "clean polis 4"
            )

            # -----------------------------------------------
            # CLEAN POLIS 5
            # -----------------------------------------------

            df["clean polis 5"] = [
                (
                    lst[4]
                    if len(lst) > 4
                    else None
                )
                for lst in all_clean_polis
            ]

            new_columns.append(
                "clean polis 5"
            )

        # ----------------------------------------------------
        # SLIP
        # ----------------------------------------------------

        elif col == "slip_ori":

            new_columns.append(
                "slip_ori"
            )

            slip_clean_columns = (
                _insert_clean_columns(
                    df,
                    all_clean_slip,
                    "slip",
                    MAX_SPLIT_COLS,
                )
            )

            new_columns += (
                slip_clean_columns
            )

        # ----------------------------------------------------
        # INSURED
        # ----------------------------------------------------

        elif col == "insured_ori":

            new_columns.append(
                "insured_ori"
            )

            insured_clean_columns = (
                _insert_clean_columns(
                    df,
                    all_clean_ins,
                    "insured",
                    MAX_SPLIT_COLS,
                )
            )

            new_columns += (
                insured_clean_columns
            )

        # ----------------------------------------------------
        # KOLOM EXISTING LAIN
        # ----------------------------------------------------

        else:

            new_columns.append(
                col
            )

    # ========================================================
    # REORDER
    # ========================================================

    df = df[
        new_columns
    ]

    # ========================================================
    # VALIDASI JUMLAH BARIS
    # ========================================================

    print(
        "\n"
        + "=" * 60
    )

    print(
        "VALIDASI AKHIR"
    )

    print(
        "=" * 60
    )

    print(
        f"Total baris input  : "
        f"{original_row_count:,}"
    )

    print(
        f"Total baris output : "
        f"{len(df):,}"
    )

    print(
        f"Baris target       : "
        f"{jumlah_etiqa:,}"
    )

    print(
        f"Baris cedant lain  : "
        f"{jumlah_non_etiqa:,}"
    )

    print(
        f"CLSDT POLICY       : "
        f"{total_clstd_polis_used:,}"
    )

    print(
        f"Fallback POLICY    : "
        f"{total_fallback_polis_fac:,}"
    )

    print(
        f"CLSDT SLIP         : "
        f"{total_clstd_slip_used:,}"
    )

    print(
        f"Fallback SLIP      : "
        f"{total_fallback_slip_fac:,}"
    )

    print(
        f"CLSDT SERTF        : "
        f"{total_clstd_sertf_used:,}"
    )

    print(
        f"Fallback SERTF     : "
        f"{total_fallback_sertf_polis:,}"
    )

    # ========================================================
    # VALIDASI FATAL
    # ========================================================

    if len(df) != original_row_count:

        print(
            "\n[ERROR FATAL]"
        )

        print(
            f"Baris input  : "
            f"{original_row_count:,}"
        )

        print(
            f"Baris output : "
            f"{len(df):,}"
        )

        print(
            "File TIDAK disimpan."
        )

        raise RuntimeError(
            "Jumlah baris output berbeda "
            "dengan jumlah baris input!"
        )

    print(
        "\n[OK] VALIDASI JUMLAH BARIS BERHASIL"
    )

    print(
        f"{original_row_count:,} = "
        f"{len(df):,}"
    )

    print(
        "Semua baris input tetap dipertahankan."
    )

    # ========================================================
    # SIMPAN
    # ========================================================

    print(
        f"\n[5/5] Menyimpan hasil ke:"
    )

    print(
        f"{output_file}"
    )

    Path(
        output_file
    ).parent.mkdir(
        parents=True,
        exist_ok=True,
    )

    df.to_excel(
        output_file,
        index=False,
        engine="xlsxwriter",
    )

    # ========================================================
    # FINAL SUMMARY
    # ========================================================

    print(
        "\n"
        + "=" * 60
    )

    print(
        "[OK] SELESAI!"
    )

    print(
        "=" * 60
    )

    print(
        f"Total baris input  : "
        f"{original_row_count:,}"
    )

    print(
        f"Total baris output : "
        f"{len(df):,}"
    )

    print(
        f"Baris target       : "
        f"{jumlah_etiqa:,}"
    )

    print(
        f"Cedant lain        : "
        f"{jumlah_non_etiqa:,}"
    )

    print(
        f"CLSDT POLICY       : "
        f"{total_clstd_polis_used:,}"
    )

    print(
        f"Fallback POLICY    : "
        f"{total_fallback_polis_fac:,}"
    )

    print(
        f"CLSDT SLIP         : "
        f"{total_clstd_slip_used:,}"
    )

    print(
        f"Fallback SLIP      : "
        f"{total_fallback_slip_fac:,}"
    )

    print(
        f"CLSDT SERTF        : "
        f"{total_clstd_sertf_used:,}"
    )

    print(
        f"Fallback SERTF     : "
        f"{total_fallback_sertf_polis:,}"
    )

    print(
        f"Total kolom output : "
        f"{len(df.columns)}"
    )

    print(
        f"File disimpan di   : "
        f"{output_file}"
    )

    print(
        "=" * 60
    )


# ============================================================
# ENTRY POINT
# ============================================================


if __name__ == "__main__":

    print("MASUK MAIN")

    process_data(
        INPUT_FILE,
        OUTPUT_FILE,
    )


# ============================================================
# ENTRY POINT
# ============================================================
# if __name__ == "__main__":
#     print("=" * 80)
#     print("TEST CLEANING DATA 2 OSBAL")
#     print("=" * 80)

#     # =========================
#     # TEST POLIS
#     # =========================
#     test_polis = [
#         "40010922042857.000752 S/D 001152+507.001 - 004 TBA",
#         "11020123000087 Ex. Policy No : 11020122000132",
#     ]

#     print("\nTEST POLIS")
#     for val in test_polis:
#         result = clean_polis(val)
#         print(f"\nORI          : {val}")
#         print(f"CLEAN POLIS  : {result}")

#     # =========================
#     # TEST SLIP
#     # =========================
#     test_slip = [
#         "4609012200005 / ( EX : 4609011900002 ) / END",
#     ]

#     print("\n" + "=" * 80)
#     print("TEST SLIP")
#     print("=" * 80)

#     for val in test_slip:
#         result = clean_slip(val)
#         print(f"\nORI          : {val}")
#         print(f"CLEAN SLIP   : {result}")