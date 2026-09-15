import os
import re
import numpy as np
import pandas as pd

# ============================================================
# CONFIG
# ============================================================
INPUT_FILE = os.path.join("input", "1b. Transaksi Facul 01.01.23 - 17.08.2026.xlsx")
OUTPUT_FILE = os.path.join("output", "etiqa_output_facul.xlsx")

CEDANT_COL = "COMP_NAME"
CEDANT_VALUE = "PT. AS. ETIQA INT. IND.(EX. AS. ASOKA MAS)"
POLIS_COL = "FAC_POLICY_NO"
SLIP_COL = "FAC_SLIP"
INSURED_COL = "FAC_INSURED"
BROKER_NAME_COL = "COMP_NAME_1"
BROKER_CODE_COL = "FAC_BROKER"
MAX_SPLIT_COLS = 5

POLIS_EXCEPTION_RE = re.compile(
    r"MOP\s*MARINE|(?:LINE\s*SLIP|LINESLIP)|\b(?:P1|P2|P3|P73)\s*CANCEL\b|\b(?:P1|P2|P3|P73)\b|\bCANCEL\b|PENYELESAIAN(?:\s+SUSPENSE)?|HUTANG|UTANG",
    re.I,
)

SLIP_EXCEPTION_RE = re.compile(
    r"\bSUMMARY\b|\bBORDER[OA]\b|\bBORDRO\b|\bSINGGLESHIPMENT\b|PENYELESAIAN(?:\s+SUSPENSE)?|HUTANG|UTANG",
    re.I,
)

MONTH_NAMES = frozenset({
    "JANUARI", "FEBRUARI", "MARET", "APRIL", "MEI", "JUNI", "JULI", "AGUSTUS",
    "SEPTEMBER", "OKTOBER", "NOVEMBER", "DESEMBER", "JANUARY", "FEBRUARY", "MARCH",
    "APRIL", "MAY", "JUNE", "JULY", "AUGUST", "SEPTEMBER", "OCTOBER", "NOVEMBER", "DECEMBER"
})

INSURED_REMOVE_RE = re.compile(
    r"\b(PT|CV|TBK|PERSERO|LTD|PTE|INC|LLC|MR|MRS|MS|BAPAK|BPK|IBU|NY|DR|DRS|DRA|IR|H|HJ)\b",
    re.I,
)
INSURED_POLIS_RE = re.compile(r"(POLIS\s*NO\.? .*)|(POLICY\s*NO\.? .*)|(SLIP\s*NO\.? .*)", re.I | re.X)
INSURED_SPLIT_RE = re.compile(r"\s*,\s*|\s*/\s*|\s+QQ\s+|\s+AND/OR\s+|\s*&\s*|\s*\+\s*", re.I)
INSURED_JUNK_WORDS = {"", "AND", "OR", "THE", "OF", "AS"}


def _normalize_spaces(text):
    return re.sub(r"\s{2,}", " ", str(text)).strip()


def _cap_or_join(items):
    items = list(dict.fromkeys([str(x) for x in items if str(x).strip()]))
    return [",".join(items)] if len(items) > MAX_SPLIT_COLS else items

def _cap_or_original(items, original):
    items = list(dict.fromkeys([str(x) for x in items if str(x).strip()]))
    return [original] if len(items) > MAX_SPLIT_COLS else items

def _clean_insured_name(name):
    if pd.isna(name):
        return ""
    name = str(name).upper()
    name = re.sub(r"\s*/\s*BORD(?:ER|ERO)?\b.*$", "", name, flags=re.I)
    name = re.sub(r"POLIS\s*NO\.?\s*.*$|POLICY\s*NO\.?\s*.*$|SLIP\s*NO\.?\s*.*$", "", name, flags=re.I)
    name = INSURED_REMOVE_RE.sub(" ", name)
    name = re.sub(r"[()]", " ", name)
    name = name.replace("/", " ").replace("-", " ")
    return _normalize_spaces(name)


def split_insured(name):
    if pd.isna(name):
        return []
    hasil = []
    for p in INSURED_SPLIT_RE.split(str(name)):
        p = _clean_insured_name(p)
        if p and p not in INSURED_JUNK_WORDS:
            hasil.append(p)
    return hasil


def clean_insured(val):
    if pd.isna(val):
        return []
    val = str(val).strip()
    if not val:
        return []
    upper = val.upper()
    if re.match(r"^BIO\s+FARMA\b", upper):
        return [_normalize_spaces(re.sub(r"\s*[/&]\s*", " ", upper))]
    cleaned = []
    for ins in split_insured(val):
        ins = _normalize_spaces(ins).upper()
        if len(ins) <= 2 or ins in INSURED_JUNK_WORDS:
            continue
        if re.fullmatch(
            r"BORD(?:ER|ERO)?\s+(?:JANUARI|FEBRUARI|MARET|APRIL|MEI|JUNI|JULI|AGUSTUS|SEPTEMBER|OKTOBER|NOVEMBER|DESEMBER|JANUARY|FEBRUARY|MARCH|MAY|JUNE|JULY|AUGUST|SEPTEMBER|OCTOBER|NOVEMBER|DECEMBER)(?:\s+20\d{2})?",
            ins,
            re.I,
        ):
            continue
        if ins not in cleaned:
            cleaned.append(ins)
    if not cleaned:
        fallback = _clean_insured_name(val)
        if fallback:
            cleaned = [fallback.upper()]
    return _cap_or_join(cleaned)


# ============================================================
# CERTIFICATE
# ============================================================

def _normalize_cert_num(value):
    value = value.strip()
    if not value.isdigit():
        return value
    # Certificate umumnya 6 digit; jangan memotong angka yang lebih panjang.
    return value.zfill(6) if len(value) < 6 else value

# ============================================================
# CERTIFICATE
# ============================================================

def _format_certificate_numbers(values):
    """
    Format certificate menjadi 6 digit.

    Rule:
    - <= 3 nomor  -> dipisahkan koma
    - > 3 nomor   -> gunakan SD
    - angka 0 / 000 bukan certificate
    """
    cleaned = []

    for value in values:
        if value is None:
            continue

        value = str(value).strip()

        if not value.isdigit():
            continue

        # 000 / 00 / 0 bukan certificate
        if int(value) == 0:
            continue

        cleaned.append(value.zfill(6))

    # Hilangkan duplicate tetapi tetap pertahankan urutan
    cleaned = list(dict.fromkeys(cleaned))

    if not cleaned:
        return ""

    # <= 3 certificate -> breakdown
    if len(cleaned) <= 3:
        return ", ".join(cleaned)

    # > 3 certificate -> SD
    return f"{cleaned[0]} SD {cleaned[-1]}"


def _parse_certificate_tail(tail):
    """
    Membaca bagian setelah policy.

    Contoh:
        000176
        000162-000170
        000001 S/D 000004
        000001,000002,000003

    Return:
        string certificate atau ""
    """

    if not tail:
        return ""

    tail = _normalize_spaces(str(tail).upper())

    # Standardisasi S/D -> SD
    tail = re.sub(r"\bS\s*/\s*D\b", "SD", tail, flags=re.I)

    # --------------------------------------------------------
    # RANGE SD
    # --------------------------------------------------------
    m = re.fullmatch(
        r"(\d{1,6})\s+SD\s+(\d{1,6})",
        tail,
        re.I,
    )

    if m:
        start = m.group(1)
        end = m.group(2)

        # Kalau jumlah digit berbeda, jangan tebak
        if len(start) != len(end):
            return ""

        a = int(start)
        b = int(end)

        # Jangan generate range terlalu besar
        count = abs(b - a) + 1

        if count <= 0:
            return ""

        # > 3 -> tetap SD
        if count > 3:
            return f"{start.zfill(6)} SD {end.zfill(6)}"

        # <= 3 -> breakdown
        step = 1 if a <= b else -1

        values = [
            str(x).zfill(6)
            for x in range(a, b + step, step)
            if x != 0
        ]

        return ", ".join(values)

    # --------------------------------------------------------
    # COMMA
    # --------------------------------------------------------
    if "," in tail:
        values = [
            x.strip()
            for x in tail.split(",")
        ]

        # Semua harus certificate numeric 1-6 digit
        if not all(re.fullmatch(r"\d{1,6}", x) for x in values):
            return ""

        return _format_certificate_numbers(values)

    # --------------------------------------------------------
    # DASH CHAIN
    #
    # 000162-000170
    # 000046-048
    # 000046-000048
    # --------------------------------------------------------
    if "-" in tail:
        values = [
            x.strip()
            for x in tail.split("-")
        ]

        if not all(re.fullmatch(r"\d{1,6}", x) for x in values):
            return ""

        # 000 dianggap separator, bukan certificate
        values = [
            x for x in values
            if int(x) != 0
        ]

        if not values:
            return ""

        # Hanya 1 certificate
        if len(values) == 1:
            return values[0].zfill(6)

        # Dua/lebih angka:
        # Jangan otomatis membuat range kalau bukan format SD.
        # Perlakukan sebagai daftar certificate.
        return _format_certificate_numbers(values)

    # --------------------------------------------------------
    # SINGLE CERTIFICATE
    # --------------------------------------------------------
    if re.fullmatch(r"\d{1,6}", tail):
        if int(tail) == 0:
            return ""

        return tail.zfill(6)

    # --------------------------------------------------------
    # POLA ACAK / TIDAK VALID
    # --------------------------------------------------------
    return ""

# ============================================================
# CERTIFICATE
# ============================================================

def _format_certificate_list(values):
    """
    Format certificate:

    <= 3 certificate
        000001, 000002, 000003

    > 3 certificate
        000001 SD 000005

    Semua certificate wajib 6 digit.
    """

    cleaned = []

    for value in values:
        value = str(value).strip()

        if not value:
            continue

        # Harus angka 1-6 digit
        if not re.fullmatch(r"\d{1,6}", value):
            continue

        # 0 / 00 / 000 bukan certificate
        if int(value) == 0:
            continue

        value = value.zfill(6)

        if value not in cleaned:
            cleaned.append(value)

    if not cleaned:
        return ""

    # Maksimal 3 -> breakdown koma
    if len(cleaned) <= 3:
        return ", ".join(cleaned)

    # Lebih dari 3 -> SD
    return f"{cleaned[0]} SD {cleaned[-1]}"


def _expand_certificate_range(start, end):
    """
    Expand range certificate.

    Contoh:
        000001 SD 000003
        ->
        000001, 000002, 000003

    Kalau > 3:
        000001 SD 000004
    """

    if not (
        re.fullmatch(r"\d{1,6}", start)
        and re.fullmatch(r"\d{1,6}", end)
    ):
        return ""

    a = int(start)
    b = int(end)

    if a == 0 or b == 0:
        return ""

    count = abs(b - a) + 1

    # Range > 3 -> SD
    if count > 3:
        return f"{start.zfill(6)} SD {end.zfill(6)}"

    # Range <= 3 -> breakdown
    step = 1 if a <= b else -1

    values = []

    for number in range(a, b + step, step):
        if number == 0:
            continue

        values.append(
            str(number).zfill(6)
        )

    return ", ".join(values)


def _parse_certificate_part(tail):
    """
    Parse bagian certificate setelah policy.

    Support:

        000176
        000162-000170
        000180-188
        000251 S/D 0000131
        000379/380/381/382/VARIOUS
        000-487,489,494,500,502
    """

    if not tail:
        return ""

    tail = _normalize_spaces(
        str(tail).upper()
    )

    # --------------------------------------------------------
    # Buang VAR / VARIOUS / TBA hanya sebagai descriptor
    # --------------------------------------------------------
    tail = re.sub(
        r"/\s*(?:VAR|VARIOUS|TBA)\b.*$",
        "",
        tail,
        flags=re.I,
    )

    tail = re.sub(
        r"\b(?:VAR|VARIOUS|TBA)\b.*$",
        "",
        tail,
        flags=re.I,
    )

    tail = _normalize_spaces(tail).strip()

    if not tail:
        return ""

    # --------------------------------------------------------
    # STANDARDISASI S/D
    # --------------------------------------------------------
    tail = re.sub(
        r"\bS\s*/\s*D\b",
        "SD",
        tail,
        flags=re.I,
    )

    # --------------------------------------------------------
    # RANGE SD
    #
    # 000251 SD 000131
    # --------------------------------------------------------
    m = re.fullmatch(
        r"(\d{1,6})\s+SD\s+(\d{1,6})",
        tail,
        re.I,
    )

    if m:
        return _expand_certificate_range(
            m.group(1),
            m.group(2),
        )

    # --------------------------------------------------------
    # COMMA
    #
    # 000487,489,494,500,502
    # --------------------------------------------------------
    if "," in tail:

        values = []

        # Bisa terdapat dash separator sebelum angka pertama.
        for part in tail.split(","):

            part = part.strip()

            # Ambil angka terakhir kalau diawali "-"
            m = re.search(
                r"(\d{1,6})$",
                part,
            )

            if not m:
                return ""

            values.append(
                m.group(1)
            )

        return _format_certificate_list(
            values
        )

    # --------------------------------------------------------
    # SLASH
    #
    # 000379/380/381/382
    # --------------------------------------------------------
    if "/" in tail:

        values = []

        parts = [
            x.strip()
            for x in tail.split("/")
        ]

        for part in parts:

            # Ambil angka saja
            m = re.fullmatch(
                r"\d{1,6}",
                part,
            )

            if not m:
                return ""

            values.append(
                m.group(0)
            )

        return _format_certificate_list(
            values
        )

    # --------------------------------------------------------
    # DASH CHAIN
    #
    # 000162-000170
    # 000180-188
    # 000-046-048
    # --------------------------------------------------------
    if "-" in tail:

        parts = [
            x.strip()
            for x in re.split(
                r"\s*-\s*",
                tail
            )
        ]

        values = []

        for part in parts:

            if not re.fullmatch(
                r"\d{1,6}",
                part,
            ):
                return ""

            # 000 bukan certificate
            if int(part) == 0:
                continue

            values.append(part)

        return _format_certificate_list(
            values
        )

    # --------------------------------------------------------
    # SINGLE CERTIFICATE
    # --------------------------------------------------------
    if re.fullmatch(
        r"\d{1,6}",
        tail,
    ):

        if int(tail) == 0:
            return ""

        return tail.zfill(6)

    # --------------------------------------------------------
    # POLA ACAK / TIDAK VALID
    # --------------------------------------------------------
    return ""

# ============================================================
# CERTIFICATE
# ============================================================

def _normalize_certificate_number(value):
    """
    Normalisasi certificate menjadi 6 digit.

    Contoh:
        176     -> 000176
        000176  -> 000176
        0000131 -> 000131
    """
    value = str(value).strip()

    if not value.isdigit():
        return None

    # Certificate maksimal 6 digit.
    # Jika lebih dari 6 digit, ambil 6 digit terakhir.
    if len(value) > 6:
        value = value[-6:]

    # 000 / 00 / 0 bukan certificate valid
    if int(value) == 0:
        return None

    return value.zfill(6)


def clean_certificate(polis_ori):
    if pd.isna(polis_ori):
        return []

    raw = _normalize_spaces(str(polis_ori).upper())

    if not raw:
        return []

    # ========================================================
    # 1. POLICY - CERTIFICATE S/D CERTIFICATE
    #
    # Contoh:
    # 1071031124000024 - 000251 S/D 0000131
    #
    # Hasil:
    # 000251 SD 000131
    # ========================================================

    m = re.fullmatch(
        r"\s*(\d{8,})\s*-\s*(\d+)\s*S\s*/?\s*D\s*(\d+)\s*",
        raw,
        re.I,
    )

    if m:
        cert_start = _normalize_certificate_number(m.group(2))
        cert_end = _normalize_certificate_number(m.group(3))

        if cert_start and cert_end:
            return [f"{cert_start} SD {cert_end}"]

        return []

    # ========================================================
    # 2. POLICY - CERTIFICATE - CERTIFICATE
    #
    # Contoh:
    # 1071031117000018 - 000162-000170
    #
    # Hasil:
    # 000162, 000170
    # ========================================================

    m = re.fullmatch(
        r"\s*(\d{8,})\s*-\s*(\d+)\s*-\s*(\d+)\s*",
        raw,
        re.I,
    )

    if m:
        certificates = []

        for value in (m.group(2), m.group(3)):
            cert = _normalize_certificate_number(value)

            if cert:
                certificates.append(cert)

        return certificates[:3]

    # ========================================================
    # 3. POLICY - CERTIFICATE
    #
    # Contoh:
    # 1071031117000018 - 000176
    #
    # Hasil:
    # 000176
    # ========================================================

    m = re.fullmatch(
        r"\s*(\d{8,})\s*-\s*(\d+)\s*",
        raw,
        re.I,
    )

    if m:
        cert = _normalize_certificate_number(m.group(2))

        if cert:
            return [cert]

        return []

    # ========================================================
    # 4. POLICY - CERTIFICATE/CERTIFICATE/CERTIFICATE
    #
    # Contoh:
    # 1010031116000187 - 000379/380/381/382/VARIOUS
    #
    # Hasil:
    # 000379 SD 000382
    # ========================================================

    m = re.fullmatch(
        r"\s*(\d{8,})\s*-\s*(.+?)\s*",
        raw,
        re.I,
    )

    if m:
        suffix = m.group(2).strip()

        # Buang descriptor VARIOUS / VAR / TBA
        suffix = re.sub(
            r"\b(?:VARIOUS|VAR|TBA)\b",
            "",
            suffix,
            flags=re.I,
        ).strip(" ,/-")

        # ----------------------------------------------------
        # Pecah berdasarkan / atau ,
        # ----------------------------------------------------
        parts = re.split(r"\s*[/,]\s*", suffix)

        certificates = []

        for part in parts:
            part = part.strip()

            if not part:
                continue

            # Kalau ada "-"
            # contoh: 000379-000382
            range_match = re.fullmatch(
                r"(\d+)\s*-\s*(\d+)",
                part,
            )

                        # ----------------------------------------------------
            # KHUSUS FORMAT 000-487
            #
            # 000 dianggap prefix/dummy, bukan certificate.
            # Yang dipakai adalah angka setelah "-".
            #
            # Contoh:
            # 000-487
            # -> 000487
            # ----------------------------------------------------
            dummy_prefix_match = re.fullmatch(
                r"0+\s*-\s*(\d+)",
                part,
            )

            if dummy_prefix_match:
                cert = _normalize_certificate_number(
                    dummy_prefix_match.group(1)
                )

                if cert:
                    certificates.append(cert)

                continue

            if range_match:
                start = _normalize_certificate_number(range_match.group(1))
                end = _normalize_certificate_number(range_match.group(2))

                if start and end:
                    certificates.extend([start, end])

                continue

            cert = _normalize_certificate_number(part)

            if cert:
                certificates.append(cert)

        # Hilangkan duplikat
        unique_certificates = []

        for cert in certificates:
            if cert not in unique_certificates:
                unique_certificates.append(cert)

        certificates = unique_certificates

        if not certificates:
            return []

        # Jika lebih dari 3 certificate,
        # gunakan SD dari pertama sampai terakhir.
        if len(certificates) > 3:
            return [
                f"{certificates[0]} SD {certificates[-1]}"
            ]

        return certificates[:3]

    # ========================================================
    # 5. FORMAT LAIN / RANDOM
    # ========================================================

    return []

def clean_polis_and_certificate(val):
    """
    Menghasilkan pasangan:

        [
            (clean_polis_1, certificate_1),
            (clean_polis_2, certificate_2),
            ...
        ]

    Tujuan:
        Polis dan certificate tidak tertukar.

    Contoh:

        1801281700019 - 000602 /
        1801051700398 - 000596

    menjadi:

        [
            ("1801281700019", "000602"),
            ("1801051700398", "000596")
        ]
    """

    if pd.isna(val):
        return []

    raw = _normalize_spaces(str(val).upper())

    if not raw:
        return []

        # ========================================================
    # TAMBAHAN RULE BARU
    # POLICY - CERTIFICATE CHAIN
    #
    # Rule ini HANYA menangani certificate yang berada
    # setelah "-" pada satu policy.
    #
    # Contoh:
    # 1010031116000176 - 000118/1119/1120
    # 1010031116000187 - 000379/380/381/382/VARIOUS
    # 1071031124000024 - 000251 S/D 0000131
    #
    # Tidak mengganggu rule policy + policy yang sudah ada.
    # ========================================================

    m = re.fullmatch(
        r"(\d{8,})\s*-\s*(.+)",
        raw,
        re.I,
    )

    if m:
        policy = m.group(1)
        tail = m.group(2).strip()

        # ----------------------------------------------------
        # BUANG DESCRIPTOR DI AKHIR
        # ----------------------------------------------------
        tail_clean = re.sub(
            r"(?:/|\s+)?\b(?:VAR|VARIOUS|TBA)\b\s*$",
            "",
            tail,
            flags=re.I,
        ).strip(" ,/-")

        # ----------------------------------------------------
        # RULE A
        # CERTIFICATE S/D CERTIFICATE
        #
        # 000251 S/D 0000131
        # -> 000251 SD 000131
        # ----------------------------------------------------
        sd_match = re.fullmatch(
            r"(\d{1,7})\s*S\s*/?\s*D\s*(\d{1,7})",
            tail_clean,
            re.I,
        )

        if sd_match:
            start = sd_match.group(1)
            end = sd_match.group(2)

            start = _normalize_certificate_number(start)
            end = _normalize_certificate_number(end)

            if start and end:
                return [
                    (policy, f"{start} SD {end}")
                ]

        # ----------------------------------------------------
        # RULE B
        # CERTIFICATE DENGAN /
        #
        # 000118/1119/1120
        # -> 000118, 001119, 001120
        #
        # > 3:
        # 000379/380/381/382
        # -> 000379 SD 000382
        # ----------------------------------------------------
        if "/" in tail_clean:
            parts = [
                x.strip()
                for x in tail_clean.split("/")
                if x.strip()
            ]

            certificates = []

            for part in parts:
                if not re.fullmatch(r"\d{1,6}", part):
                    certificates = []
                    break

                cert = _normalize_certificate_number(part)

                if cert:
                    certificates.append(cert)

            if certificates:
                certificates = list(
                    dict.fromkeys(certificates)
                )

                if len(certificates) <= 3:
                    return [
                        (
                            policy,
                            ", ".join(certificates)
                        )
                    ]

                return [
                    (
                        policy,
                        f"{certificates[0]} SD {certificates[-1]}"
                    )
                ]

        # ----------------------------------------------------
        # RULE C
        # CERTIFICATE DENGAN DASH
        #
        # 000162-000170
        # -> 000162, 000170
        #
        # Tetap menggunakan rule certificate,
        # bukan policy expansion.
        # ----------------------------------------------------
        dash_match = re.fullmatch(
            r"(\d{1,6})\s*-\s*(\d{1,6})",
            tail_clean,
            re.I,
        )

        if dash_match:
            certificates = []

            for value in dash_match.groups():
                cert = _normalize_certificate_number(value)

                if cert:
                    certificates.append(cert)

            if certificates:
                return [
                    (
                        policy,
                        ", ".join(certificates)
                    )
                ]

        # ----------------------------------------------------
        # RULE D
        # SINGLE CERTIFICATE
        #
        # 000176
        # -> 000176
        # ----------------------------------------------------
        if re.fullmatch(r"\d{1,6}", tail_clean):
            cert = _normalize_certificate_number(
                tail_clean
            )

            if cert:
                return [
                    (policy, cert)
                ]

    # ========================================================
    # RULE LAMA DILANJUTKAN DI BAWAH SINI
    # ========================================================

    # --------------------------------------------------------
    # TAMBAHAN RULE:
    # POLICY - CERTIFICATE CHAIN
    #
    # Jangan pecah "/" jika "/" tersebut adalah separator
    # antar certificate.
    #
    # Contoh:
    # 1010031116000176 - 000118/1119/1120
    #
    # 1010031116000187 - 000379/380/381/382/VARIOUS
    # --------------------------------------------------------
    certificate_chain_match = re.fullmatch(
        r"(\d{8,})\s*-\s*(.+)",
        raw,
        re.I,
    )

    if certificate_chain_match:
        policy = certificate_chain_match.group(1)
        tail = certificate_chain_match.group(2).strip()

        # Hanya jalankan rule ini jika tail memang
        # terlihat seperti certificate chain.
        #
        # Contoh valid:
        # 000118/1119/1120
        # 000379/380/381/382/VARIOUS
        # 000251 S/D 0000131
        #
        # Descriptor VAR/VARIOUS/TBA boleh ada di belakang.
        if (
            "/" in tail
            or re.search(r"\bS\s*/?\s*D\b", tail, re.I)
        ):
            certificate = _parse_certificate_tail(tail)

            if certificate:
                return [
                    (policy, certificate)
                ]

    # --------------------------------------------------------
    # TAMBAHAN RULE:
    # POLICY - MULTIPLE CERTIFICATE DENGAN SLASH
    #
    # Contoh:
    # 1010031116000176 - 000118/1119/1120
    #
    # Jangan split "/" sebagai pemisah policy.
    # "/" di sini adalah pemisah certificate.
    #
    # Hasil:
    # policy      = 1010031116000176
    # certificate = 000118, 001119, 001120
    # --------------------------------------------------------
    m = re.fullmatch(
        r"(\d{8,})\s*-\s*(\d+(?:\s*/\s*\d+)+)",
        raw,
        re.I,
    )

    if m:
        policy = m.group(1)
        certificate_tail = m.group(2).strip()

        certificate = _parse_certificate_tail(
            certificate_tail
        )

        return [
            (policy, certificate)
        ]

    # --------------------------------------------------------
    # CASE 1
    # Multiple policy menggunakan /
    #
    # POLICY-CERT / POLICY-CERT
    # --------------------------------------------------------
    slash_parts = [
        p.strip()
        for p in re.split(r"\s*/\s*", raw)
        if p.strip()
    ]

    pairs = []

    for part in slash_parts:

        # POLICY - CERT
        m = re.fullmatch(
            r"(\d{8,})\s*-\s*(.+)",
            part,
        )

        if m:
            policy = m.group(1)
            tail = m.group(2).strip()

            certificate = _parse_certificate_tail(tail)

            if certificate:
                pairs.append(
                    (policy, certificate)
                )
                continue

            # Tidak ada certificate valid
            pairs.append(
                (policy, "")
            )
            continue

        # Policy tanpa certificate
        m = re.fullmatch(
            r"(\d{8,})",
            part,
        )

        if m:
            pairs.append(
                (m.group(1), "")
            )

    # Kalau berhasil parsing pasangan
    if pairs:
        return pairs

    # --------------------------------------------------------
    # CASE 2
    # Satu policy - certificate
    # --------------------------------------------------------
    m = re.fullmatch(
        r"(\d{8,})\s*-\s*(.+)",
        raw,
    )

    if m:
        policy = m.group(1)
        tail = m.group(2)

        certificate = _parse_certificate_tail(tail)

        return [
            (policy, certificate)
        ]

    # --------------------------------------------------------
    # CASE 3
    # Policy biasa
    # --------------------------------------------------------
    m = re.fullmatch(
        r"\d{8,}",
        raw,
    )

    if m:
        return [
            (raw, "")
        ]

    return []

# ============================================================
# POLIS
# ============================================================
def _extract_main_policy(val):
    m = re.search(r"\b\d{8,}\b", val)
    return m.group(0) if m else ""


def clean_polis(val):
    if pd.isna(val):
        return []
    original = _normalize_spaces(str(val).upper())
    if not original:
        return []

    # Exception wajib dipertahankan.
    if POLIS_EXCEPTION_RE.search(original):
        # P1 / nomor polis -> nomor polis saja; P1 murni tetap P1.
        m = re.search(r"\bP[123P]?\s*/\s*(\d{8,})", original, re.I)
        if m:
            return [m.group(1)]
        if re.fullmatch(r"P[123](?:\s*\+\s*P[123])?", original, re.I):
            return [original]
        if re.fullmatch(r"TBA\s*/.*", original, re.I):
            return [original]
        if re.fullmatch(r"(?:P1|P2|P3|P73)", original, re.I):
            return [original]

    if re.fullmatch(r"VARIOUS", original, re.I):
        return ["VARIOUS"]

    # Hapus descriptor VARIOUS/VAR/TBA setelah policy valid.
    val2 = re.sub(r"\s*/\s*(?:VAR|VARIOUS|TBA)\b.*$", "", original, flags=re.I)
    val2 = re.sub(r"\s+(?:VAR|VARIOUS)\b.*$", "", val2, flags=re.I)
    val2 = _normalize_spaces(val2)

    # P1 / policy atau P2 / policy.
    m = re.fullmatch(r"P[123]\s*/\s*(\d{8,})(?:\s+.*)?", val2, re.I)
    if m:
        return [m.group(1)]

    # Untuk format policy - certificate, suffix certificate TIDAK dimasukkan ke clean polis.
    # Ini penting untuk pola Etiqa seperti:
    # 1071031117000018 - 000162-000170
    # 301001091900866-797-809
    # 1010031116000017-000-046-048
    main = _extract_main_policy(val2)
    if main and re.search(r"\bS\s*\.?\s*/?\s*D\b", val2, re.I):
        return [main]

    # perulangan salah
    # if main and re.match(r"^\s*\d{8,}\s*-", val2):
    #     # Jika setelah policy hanya numeric suffix 1-6 digit, anggap certificate.
    #     tail = val2[val2.find(main) + len(main):]
    #     if re.search(r"(?:-|,|/)", tail):
    #         if re.fullmatch(r"(?:\s*[-/]\s*\d{1,6}|\s*[-/]\s*\d{1,6}(?:\s*[-/,]\s*\d{1,6})*)(?:\s*)", tail):
    #             return [main]
    #     if re.fullmatch(r"\s*-\s*\d{1,6}\s*", tail):
    #         return [main]

    # --------------------------------------------------------
    # DASH CHAIN
    # Contoh:
    # 00940502012023001276-136-135-1275
    #
    # Ini adalah perulangan polis, BUKAN certificate.
    # Suffix pendek dibentuk dari base policy.
    # --------------------------------------------------------
    if re.fullmatch(r"\d{8,}(?:-\d{1,6})+", val2):
        parts = [p for p in val2.split("-") if p]

        base = parts[0]
        hasil = [base]

        for suffix in parts[1:]:
            if len(suffix) < len(base):
                full = base[:len(base) - len(suffix)] + suffix
            else:
                full = suffix

            if full not in hasil:
                hasil.append(full)

        return _cap_or_original(hasil, original)

    # Numeric S/D range tanpa policy prefix: expand jika <=5, selain itu pertahankan.
    m = re.fullmatch(r"(\d+)\s*S\s*/?\s*D\s*(\d+)", val2, re.I)
    if m and len(m.group(1)) == len(m.group(2)):
        a, b = int(m.group(1)), int(m.group(2))
        count = abs(b - a) + 1
        if count > MAX_SPLIT_COLS:
            return [original]
        step = 1 if a <= b else -1
        return [str(n).zfill(len(m.group(1))) for n in range(a, b + step, step)]

    # Ampersand = separate polis.
    if "&" in val2:
        parts = [re.sub(r"[^A-Z0-9]", "", p.strip()) for p in re.split(r"\s*&\s*", val2)]
        parts = [p for p in parts if p]
        return _cap_or_original(parts, original)

    # Plus = separate policy. Suffix pendek dibentuk dari policy sebelumnya.
    if "+" in val2:
        parts = [p.strip() for p in re.split(r"\s*\+\s*", val2) if p.strip()]
        hasil = []
        base = None
        for p in parts:
            if re.fullmatch(r"P[123]", p, re.I) or p.upper() in {"VAR", "VARIOUS", "TBA"}:
                continue
            clean = re.sub(r"[^A-Z0-9]", "", p)
            if len(clean) >= 8:
                hasil.append(clean)
                base = clean
                continue
            if base and re.fullmatch(r"\d{1,6}", p):
                suffix = p
                full = base[: len(base) - len(suffix)] + suffix if len(suffix) < len(base) else suffix
                if full not in hasil:
                    hasil.append(full)
                base = full
        if hasil:
            if hasil:
                return _cap_or_original(hasil, original)

    # Slash yang hanya memisahkan policy valid.
    if "/" in val2:
        parts = [p.strip() for p in re.split(r"\s*/\s*", val2)]
        candidates = []
        for p in parts:
            if re.fullmatch(r"(?:VAR|VARIOUS|TBA|P[123])", p, re.I):
                continue
            clean = re.sub(r"[^A-Z0-9]", "", p)
            if len(clean) >= 8 and re.search(r"\d", clean):
                candidates.append(clean)
        if candidates:
            return _cap_or_original(candidates, original)

    # Policy numerik murni / dotted.
    if re.fullmatch(r"[0-9.]+", val2):
        return [val2.replace(".", "")]

    # Dash yang bukan certificate: fallback bersih tanpa separator.
    if re.fullmatch(r"[0-9-]+", val2):
        parts = [p for p in val2.split("-") if p]
        if len(parts) == 1:
            return [parts[0]]
        # Jika ada policy panjang di awal, ambil policy utama saja.
        if len(parts[0]) >= 8:
            return [parts[0]]
        return [val2.replace("-", "")]

    # Ambil policy panjang jika descriptor mengikutinya.
    candidates = re.findall(r"\b\d{8,}\b", val2)
    if candidates:
        return _cap_or_original(list(dict.fromkeys(candidates)), original)

    return [val2] if val2 else []


# ============================================================
# SLIP
# ============================================================
def clean_slip(val):
    if pd.isna(val):
        return []
    original = _normalize_spaces(str(val).upper())
    if not original:
        return []

    val2 = original.replace(".", "")

    # Exception/descriptive slip tetap satu nilai.
    if SLIP_EXCEPTION_RE.search(val2):
        return [_normalize_spaces(val2)]

    # Month/year + various -> month/year saja.
    m = re.fullmatch(r"([A-Z]+)\s+(20\d{2})\s*/\s*VARIOUS", val2, re.I)
    if m and m.group(1) in MONTH_NAMES:
        return [f"{m.group(1)} {m.group(2)}"]

    # P1 / slip-number.
    m = re.fullmatch(r"P[123]\s*/\s*(\d{7,})", val2, re.I)
    if m:
        return [m.group(1)]

    # slash antar 2 slip panjang.
    m = re.fullmatch(r"(\d{7,})\s*/\s*(\d{7,})", val2)
    if m:
        return [m.group(1), m.group(2)]

    # 123456789 / 0 => nomor utama.
    m = re.fullmatch(r"(\d{7,})\s*/\s*0", val2)
    if m:
        return [m.group(1)]

    # CN descriptor setelah slip utama.
    m = re.match(r"^(\d{7,})\s*-\s*\d+/CN/", val2, re.I)
    if m:
        return [m.group(1)]

    # S/D slip range. <=5 di-expand, >5 tetap S/D.
    m = re.search(r"(\d{7,})\s*S\s*\.?\s*/?\s*D\s*(\d{7,})", val2, re.I)
    if m and len(m.group(1)) == len(m.group(2)):
        a, b = int(m.group(1)), int(m.group(2))
        count = abs(b - a) + 1
        if count > MAX_SPLIT_COLS:
            return [f"{m.group(1)} S/D {m.group(2)}"]
        step = 1 if a <= b else -1
        return [str(n).zfill(len(m.group(1))) for n in range(a, b + step, step)]

    # Slip panjang / suffix pendek -> expand base + reconstructed suffix.
    m = re.fullmatch(r"(\d{7,})\s*/\s*(\d{1,6})", val2)
    if m:
        base, suffix = m.groups()
        if len(suffix) < len(base):
            return [base, base[:len(base) - len(suffix)] + suffix]

    # Plus: full slip + suffix pendek.
    if "+" in val2:
        parts = [p.strip() for p in val2.split("+") if p.strip()]
        if parts and re.fullmatch(r"\d{7,}", parts[0]):
            base = parts[0]
            hasil = [base]
            for suffix in parts[1:]:
                if re.fullmatch(r"\d{3,6}", suffix):
                    full = base[:len(base) - len(suffix)] + suffix if len(suffix) < len(base) else suffix
                    if full not in hasil:
                        hasil.append(full)
                elif suffix.upper() not in {"VAR", "VARIOUS", "TBA", "P1", "P2", "P3"}:
                    # Full slip berikutnya tetap dipertahankan.
                    if len(suffix) >= 7 and suffix not in hasil:
                        hasil.append(suffix)
            return _cap_or_join(hasil)

    # Numeric dash chain: jangan menganggap certificate; untuk slip, expand suffix.
    if re.fullmatch(r"\d{7,}(?:\s*-\s*\d{1,6})+", val2):
        parts = [p.strip() for p in re.split(r"\s*-\s*", val2)]
        base = parts[0]
        hasil = [base]
        for suffix in parts[1:]:
            if re.fullmatch(r"\d{1,6}", suffix):
                full = base[:len(base) - len(suffix)] + suffix if len(suffix) < len(base) else suffix
                if full not in hasil:
                    hasil.append(full)
        return _cap_or_join(hasil)

    # Hapus noise umum dari slip deskriptif.
    val3 = re.sub(r"^SEE\s+ATTACHMENT\s*/\s*", "", val2, flags=re.I).strip()
    val3 = re.sub(r"\bEND(?:\.?1)?\b", "", val3, flags=re.I)
    val3 = re.sub(r"\b(?:USD|IDR|SGD|EUR|JPY|AUD|GBP|ORI|ORIGINAL|COPY|REALISASI|REALIZATION|CANCEL(?:LED)?|ENDORSEMENT|SA|P1|P2|P3|VAR|REVISI|REV)\b.*$", "", val3, flags=re.I)
    val3 = _normalize_spaces(val3).strip("-_,.; ")

    # Ambil nomor slip panjang dari string.
    candidates = re.findall(r"\b\d{7,}\b", val3)
    if candidates:
        candidates = list(dict.fromkeys(candidates))
        return candidates if len(candidates) <= MAX_SPLIT_COLS else [val3]

    # Slash CN / format non-numeric harus dipertahankan.
    if "/CN/" in val3 or re.search(r"[A-Z]{2,}/\d", val3):
        return [val3]

    if "," in val3:
        parts = [p.strip() for p in val3.split(",") if p.strip()]
        return _cap_or_original(parts, original)

    if re.fullmatch(r"\d{1,6}(?:\s*[&+-]\s*\d{1,6})+", val3):
        return [val3]

    return [val3] if val3 else []


# ============================================================
# BUSINESS PARTNERS / EXCEL HELPERS
# ============================================================
def clean_business_partners(val):
    if pd.isna(val):
        return ""
    val = str(val).strip().upper()
    val = re.sub(r"\bPT\.\s*", "PT ", val, flags=re.I)
    return _normalize_spaces(val)


def get_mitra_bisnis(broker_name, broker_code, cedant):
    broker_name = "" if pd.isna(broker_name) else str(broker_name).strip()
    broker_code = "" if pd.isna(broker_code) else str(broker_code).strip()
    cedant = "" if pd.isna(cedant) else str(cedant).strip()
    if broker_name and broker_name.upper() != "DIRECT":
        return clean_business_partners(broker_name)
    if broker_code and broker_code.upper() != "DIRECT":
        return clean_business_partners(broker_code)
    return clean_business_partners(cedant)


def _insert_clean_columns(df, all_lists, prefix, max_cols):
    added = []
    for i in range(1, max_cols + 1):
        col_name = f"clean {prefix} {i}"
        df[col_name] = [lst[i - 1] if i - 1 < len(lst) else None for lst in all_lists]
        added.append(col_name)
    return added


def _fast_read_excel(path, sheet_name=None, header=0):
    import openpyxl
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb[sheet_name] if sheet_name else wb.worksheets[0]
    rows = ws.iter_rows(values_only=True)
    for _ in range(header):
        next(rows)
    cols = list(next(rows))
    seen = {}
    new_cols = []
    for col in cols:
        col = "" if col is None else str(col).strip()
        if col not in seen:
            seen[col] = 0
            new_cols.append(col)
        else:
            seen[col] += 1
            new_cols.append(f"{col}_{seen[col]}")
    data = list(rows)
    wb.close()
    return pd.DataFrame(data, columns=new_cols)


def _fast_write_excel(df, path):
    import openpyxl
    from openpyxl.cell import WriteOnlyCell
    output_dir = os.path.dirname(path)
    if output_dir:
        os.makedirs(output_dir, exist_ok=True)
    wb = openpyxl.Workbook(write_only=True)
    ws = wb.create_sheet("Sheet1")
    ws.append(list(df.columns))
    for row in df.itertuples(index=False, name=None):
        out = []
        for value in row:
            cell = WriteOnlyCell(ws, value=None if pd.isna(value) else value)
            if isinstance(value, str):
                cell.number_format = "@"
            out.append(cell)
        ws.append(out)
    wb.save(path)


# ============================================================
# PROCESS DATA
# ============================================================
def process_data(input_file, output_file):
    print(f"[1/5] Membaca data dari: {input_file} ...")
    df = _fast_read_excel(input_file, header=0)
    print(f"      Total baris keseluruhan: {len(df):,}")

    if CEDANT_COL not in df.columns:
        print(f"[ERROR] Kolom '{CEDANT_COL}' tidak ditemukan!")
        return

    required_cols = [POLIS_COL, SLIP_COL, INSURED_COL, BROKER_NAME_COL, BROKER_CODE_COL]
    for col in required_cols:
        if col not in df.columns:
            print(f"[ERROR] Kolom '{col}' tidak ditemukan di file!")
            return

    df[CEDANT_COL] = (
        df[CEDANT_COL].fillna("").astype(str).str.upper().str.strip()
        .str.replace(r"\.", "", regex=True).str.replace(r"\s+", " ", regex=True)
    )
    cedant_value_normalized = CEDANT_VALUE.upper().replace(".", "")
    is_etiqa = df[CEDANT_COL].eq(cedant_value_normalized)
    total_etiqa = int(is_etiqa.sum())
    print(f"[2/5] Filter cedant '{CEDANT_VALUE}': {total_etiqa:,} baris ETIQA dari total {len(df):,} baris.")

    broker_name_s = df[BROKER_NAME_COL].fillna("").astype(str).str.strip()
    broker_code_s = df[BROKER_CODE_COL].fillna("").astype(str).str.strip()
    cedant_s = df[CEDANT_COL].fillna("").astype(str).str.strip()
    mitra_values = [get_mitra_bisnis(a, b, c) for a, b, c in zip(broker_name_s, broker_code_s, cedant_s)]

    df.rename(columns={POLIS_COL: "polis_ori", SLIP_COL: "slip_ori", INSURED_COL: "insured_ori"}, inplace=True)

    insert_pos = list(df.columns).index(BROKER_NAME_COL) + 1 if BROKER_NAME_COL in df.columns else len(df.columns)
    df.insert(insert_pos, "BUSINESS PARTNERS", mitra_values)

    print("[3/5] Menjalankan proses cleaning hanya untuk baris ETIQA ...")
    n = len(df)
    all_clean_polis = [[] for _ in range(n)]
    all_clean_slip = [[] for _ in range(n)]
    all_clean_ins = [[] for _ in range(n)]

    # Certificate per polis
    all_certificates = [[] for _ in range(n)]
    max_polis = max_slip = max_ins = 1

    etiqa_idx = np.flatnonzero(is_etiqa.to_numpy())
    polis_vals = df["polis_ori"].to_numpy()
    slip_vals = df["slip_ori"].to_numpy()
    insured_vals = df["insured_ori"].to_numpy()

    for n_done, pos in enumerate(etiqa_idx, 1):
        if n_done % 5000 == 0:
            print(f"      Progress: {n_done:,} / {total_etiqa:,} baris ETIQA diproses...")
        # ========================================================
        # POLIS + CERTIFICATE HARUS DIPASANGKAN
        # ========================================================

        pairs = clean_polis_and_certificate(
            polis_vals[pos]
        )

        if pairs:
            c_polis = [p for p, c in pairs]
            c_cert = [c for p, c in pairs]
        else:
            c_polis = clean_polis(
                polis_vals[pos]
            )

            c_cert = [
                clean_certificate(
                    polis_vals[pos]
                )
            ]

        # Hapus certificate kosong
        c_cert = [
            c for c in c_cert
            if c
        ]

        c_slip = clean_slip(
            slip_vals[pos]
        )

        c_ins = clean_insured(
            insured_vals[pos]
        )

        max_polis = max(
            max_polis,
            len(c_polis)
        )

        max_slip = max(
            max_slip,
            len(c_slip)
        )

        max_ins = max(
            max_ins,
            len(c_ins)
        )

        all_clean_polis[pos] = c_polis
        all_clean_slip[pos] = c_slip
        all_clean_ins[pos] = c_ins
        all_certificates[pos] = c_cert
        max_slip = max(max_slip, len(c_slip))
        max_ins = max(max_ins, len(c_ins))
        all_clean_polis[pos] = c_polis
        all_clean_slip[pos] = c_slip
        all_clean_ins[pos] = c_ins
        all_certificates[pos] = c_cert

    print(f"      Selesai diproses! polis={max_polis}, slip={max_slip}, insured={max_ins}")

    # ============================================================
    # CERTIFICATE PER POLIS
    # ============================================================

    max_cert = max(
        [len(x) for x in all_certificates] + [1]
    )

    for i in range(1, max_cert + 1):
        col_name = f"Certificate {i}"

        df[col_name] = [
            certs[i - 1]
            if i - 1 < len(certs) and certs[i - 1]
            else None
            for certs in all_certificates
        ]
    new_columns = []
    for col in df.columns:
        if col == "CERTIFICATE":
            continue
        new_columns.append(col)
        if col == "polis_ori":

            polis_cols = _insert_clean_columns(
                df,
                all_clean_polis,
                "polis",
                max_polis
            )

            # ========================================================
            # POLIS 1 -> CERTIFICATE 1
            # POLIS 2 -> CERTIFICATE 2
            # POLIS 3 -> CERTIFICATE 3
            # ========================================================

            for i, polis_col in enumerate(
                polis_cols,
                start=1
            ):

                new_columns.append(
                    polis_col
                )

                cert_col = f"Certificate {i}"

                if cert_col in df.columns:
                    new_columns.append(
                        cert_col
                    )
        elif col == "slip_ori":
            new_columns += _insert_clean_columns(df, all_clean_slip, "slip", max_slip)
        elif col == "insured_ori":
            new_columns += _insert_clean_columns(df, all_clean_ins, "insured", max_ins)

    df = df[new_columns]
    print(f"[5/5] Menyimpan hasil ke: {output_file} ...")
    _fast_write_excel(df, output_file)
    print(f"[OK] Selesai. Total baris={len(df):,}, kolom={len(df.columns)}, ETIQA={total_etiqa:,}")


# ============================================================
# ENTRY POINT
# ============================================================
# if __name__ == "__main__":

#     print("=" * 80)
#     print("TEST ENTRY POINT - CLEAN POLIS + CERTIFICATE")
#     print("=" * 80)

#     test_cases = [
#         "1010031116000176 - 000118/1119/1120",
#         "1010031116000187 - 000379/380/381/382/VARIOUS",
#         "1010031116000187 - 000385/386/388/389/390/VARIOUS",
#         "1071031117000018 - 000176",
#         "1071031117000018 - 000162-000170",
#         "1071031124000024 - 000251 S/D 0000131",
#         "1010031116000017-000-046-048",

#         # Bukan certificate
#         "+ 3010010519000236",
#         "+ 3010010520005142 + 051",
#         "+ 017",
#         "/ P2",
#     ]

#     for i, test in enumerate(test_cases, 1):

#         result = clean_polis_and_certificate(test)

#         print(f"\nTEST {i}")
#         print(f"ORI       : {test}")
#         print(f"HASIL     : {result}")

#         if result:
#             for j, (polis, certificate) in enumerate(result, 1):
#                 print(f"  Polis {j}      : {polis}")
#                 print(f"  Certificate {j}: {certificate}")
#         else:
#             print("  Tidak ada hasil")

#     print("\n" + "=" * 80)
#     print("TEST ENTRY POINT SELESAI")
#     print("DATA BESAR BELUM DIJALANKAN")
#     print("=" * 80)

if __name__ == "__main__":

    print("MASUK MAIN")

    process_data(
        INPUT_FILE,
        OUTPUT_FILE,
    )
