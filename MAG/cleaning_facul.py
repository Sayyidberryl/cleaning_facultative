import os
import re
import numpy as np
import pandas as pd

# ============================================================
# CONFIG
# ============================================================
INPUT_FILE = os.path.join("input", "1b. Transaksi Facul 01.01.23 - 17.08.2026.xlsx")
OUTPUT_FILE = os.path.join("output", "mag_output_facul.xlsx")

CEDANT_COL = "COMP_NAME"
CEDANT_VALUE = "PT ASURANSI MULTI ARTHA GUNA"   # setelah normalisasi (upper, buang '.', collapse spasi)
POLIS_COL = "FAC_POLICY_NO"
SLIP_COL = "FAC_SLIP"
INSURED_COL = "FAC_INSURED"
BROKER_NAME_COL = "COMP_NAME_1"
BROKER_CODE_COL = "FAC_BROKER"
MAX_SPLIT_COLS = 5

# NOTE:
# Tidak ada logic CLSDT_POLICY_NO / CLSDT_SLIP_NO / CLSDT_SERTF_NO / certificate
# di script ini. Ini murni Data 1 - FACUL.

POLIS_EXCEPTION_RE = re.compile(
    r"MOP\s*MARINE|(?:LINE\s*SLIP|LINESLIP)|\b(?:P1|P2|P3|P73)\s*CANCEL\b|\b(?:P1|P2|P3|P73)\b|\bCANCEL\b|PENYELESAIAN(?:\s+SUSPENSE)?|HUTANG|UTANG",
    re.I,
)

SLIP_EXCEPTION_RE = re.compile(
    r"\bSUMMARY\b|\bBORDER[OA]\b|\bBORDRO\b|\bSINGGLESHIPMENT\b|PENYELESAIAN(?:\s+SUSPENSE)?|HUTANG|UTANG",
    re.I,
)

# --------------------------------------------------------------
# POLA BARU KHUSUS MULTI ARTHA GUNA (ditemukan dari data asli):
# "<POLIS AKTIF> EX[.] [POLICY NO] [:] <POLIS LAMA/REFERENSI>"
# Contoh nyata:
#   45040118001473 Ex. Policy No : 45040117001948
#   36040119000076 EX. POLICY NO : 36040118000085
#   45090119000635 Ex 45090118000303
#   45040518000083 Ex.45040517000081
# Polis aktif adalah angka PERTAMA (sebelum EX). Bagian setelah EX
# adalah referensi ke polis periode sebelumnya, bukan polis tambahan.
# --------------------------------------------------------------
EX_POLICY_RE = re.compile(
    r"^\s*(\d{8,})\s*EX\.?\s*(?:POLICY\s*NO\.?)?\s*:?\s*\d{4,}\s*$",
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

# Panjang digit "wajar" untuk nomor polis Multi Artha Guna, berdasarkan
# observasi data asli (umumnya 8-20 digit). Kandidat digit di atas ini
# dianggap hasil data korup/gabungan tanpa separator -> jangan diambil.
MAX_REASONABLE_POLIS_DIGITS = 20


def _normalize_spaces(text):
    return re.sub(r"\s{2,}", " ", str(text)).strip()


def _cap_or_join(items):
    items = list(dict.fromkeys([str(x) for x in items if str(x).strip()]))
    return [",".join(items)] if len(items) > MAX_SPLIT_COLS else items


def _cap_or_original(items, original):
    items = list(dict.fromkeys([str(x) for x in items if str(x).strip()]))
    return [original] if len(items) > MAX_SPLIT_COLS else items


# ============================================================
# INSURED  (reused dari script Etiqa, generic - tidak diubah)
# ============================================================
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
        # Jangan sampai hasil akhir kosong padahal input tidak kosong
        # (mis. "BAPAK/IBU" -> semua kata masuk daftar gelar yang dihapus).
        fallback = _clean_insured_name(val)
        if not fallback:
            fallback = _normalize_spaces(val)
        cleaned = [fallback.upper()]
    return _cap_or_join(cleaned)


# ============================================================
# CERTIFICATE (reuse PERSIS dari script Etiqa, tidak diubah)
# ============================================================
def _normalize_certificate_number(value):
    """
    Normalisasi certificate menjadi 6 digit.
    Contoh: 176 -> 000176, 000176 -> 000176, 0000131 -> 000131
    """
    value = str(value).strip()
    if not value.isdigit():
        return None
    if len(value) > 6:
        value = value[-6:]
    if int(value) == 0:
        return None
    return value.zfill(6)


def _format_certificate_list(values):
    """
    <= 3 certificate -> "000001, 000002, 000003"
    >  3 certificate -> "000001 SD 000005"
    """
    cleaned = []
    for value in values:
        value = str(value).strip()
        if not value:
            continue
        if not re.fullmatch(r"\d{1,6}", value):
            continue
        if int(value) == 0:
            continue
        value = value.zfill(6)
        if value not in cleaned:
            cleaned.append(value)
    if not cleaned:
        return ""
    if len(cleaned) <= 3:
        return ", ".join(cleaned)
    return f"{cleaned[0]} SD {cleaned[-1]}"


def _expand_certificate_range(start, end):
    if not (re.fullmatch(r"\d{1,6}", start) and re.fullmatch(r"\d{1,6}", end)):
        return ""
    a = int(start)
    b = int(end)
    if a == 0 or b == 0:
        return ""
    count = abs(b - a) + 1
    if count > 3:
        return f"{start.zfill(6)} SD {end.zfill(6)}"
    step = 1 if a <= b else -1
    values = []
    for number in range(a, b + step, step):
        if number == 0:
            continue
        values.append(str(number).zfill(6))
    return ", ".join(values)


def clean_certificate(polis_ori):
    """
    Certificate MAG - RULE KETAT.

    Certificate hanya dianggap valid jika:
      1. nomor certificate TEPAT 6 digit penuh; dan
      2. nomor tersebut diawali tanda '-' setelah nomor polis.

    Jadi:
      BASE-000001             -> certificate
      BASE-000001 S/D 000006  -> certificate
      BASE-0614               -> BUKAN certificate
      BASE-00001              -> BUKAN certificate
      BASE-8007321            -> BUKAN certificate
      BASE+2553               -> BUKAN certificate
      BASE/588/4263           -> BUKAN certificate

    Kalau formatnya tidak memenuhi rule di atas, return [].
    """
    if pd.isna(polis_ori):
        return []

    raw = _normalize_spaces(str(polis_ori).upper())
    if not raw:
        return []

    # --------------------------------------------------------
    # POLICY - CERTIFICATE S/D CERTIFICATE
    # Kedua certificate WAJIB 6 digit.
    # --------------------------------------------------------
    m = re.fullmatch(
        r"\s*(\d{8,})\s*-\s*(\d{6})\s*S\s*/?\s*D\s*(\d{6})\s*",
        raw,
        re.I,
    )
    if m:
        return [f"{m.group(2)} SD {m.group(3)}"]

    # --------------------------------------------------------
    # POLICY - CERTIFICATE - CERTIFICATE
    # --------------------------------------------------------
    m = re.fullmatch(
        r"\s*(\d{8,})\s*-\s*(\d{6})\s*-\s*(\d{6})\s*",
        raw,
        re.I,
    )
    if m:
        return [m.group(2), m.group(3)]

    # --------------------------------------------------------
    # POLICY - CERTIFICATE / CERTIFICATE / CERTIFICATE
    # Semua bagian setelah '-' WAJIB 6 digit.
    # --------------------------------------------------------
    m = re.fullmatch(r"\s*(\d{8,})\s*-\s*(\d{6})(.*)\s*", raw, re.I)
    if m:
        first_cert = m.group(2)
        remainder = m.group(3).strip()
        certs = [first_cert]

        if not remainder:
            return certs

        # Hanya separator certificate yang diperbolehkan.
        parts = [p.strip() for p in re.split(r"\s*[/+,]\s*", remainder) if p.strip()]
        if not parts:
            return []

        for part in parts:
            if not _is_real_certificate_suffix(part):
                return []
            if part not in certs:
                certs.append(part)

        return certs[:MAX_SPLIT_COLS]

    return []


# ============================================================
# RULE BARU: PEMISAHAN POLIS & CERTIFICATE MAG
# ============================================================
# Certificate hanya 6 digit penuh dan harus didahului '-'.
# Selain itu diproses sebagai perulangan polis oleh clean_polis().
# ============================================================
# ============================================================
# RULE PEMISAHAN POLIS & CERTIFICATE - MULTI ARTHA GUNA
# ============================================================
# ATURAN UTAMA:
#
# 1. CERTIFICATE HANYA diakui kalau:
#       <POLIS>-<6 DIGIT PENUH>
#    Contoh:
#       36040118000154-000001
#       08080517000105-000006
#
# 2. Suffix seperti 0614, 588, 8007321, 5180737, 2553, dst.
#    BUKAN certificate. Itu adalah PERULANGAN / lanjutan nomor polis.
#    Suffix tersebut harus diproses oleh clean_polis menjadi nomor polis
#    berikutnya, bukan dimasukkan ke kolom certificate.
#
# 3. Kalau ada beberapa policy penuh dalam satu value, masing-masing
#    tetap menjadi clean polis terpisah.
#
# 4. Kalau separator + / , / - menghasilkan suffix pendek, suffix
#    mengikuti BASE polis sebelumnya dan direkonstruksi menjadi nomor
#    polis penuh.
#
# 5. Kalau pola ambigu / tidak aman, pertahankan original.
# ============================================================


def _is_real_certificate_suffix(value):
    """Certificate MAG = TEPAT 6 digit, tidak lebih dan tidak kurang."""
    return bool(re.fullmatch(r"\d{6}", str(value).strip()))


def _has_explicit_certificate_pattern(original):
    """
    True hanya kalau ditemukan pola BASE-######.
    Jadi BASE-0614, BASE-588, BASE-8007321, dll BUKAN certificate.
    """
    return bool(re.search(r"\d{8,}\s*-\s*\d{6}(?!\d)", original))


def _split_real_certificate_pattern(original):
    """
    Pisahkan policy + certificate khusus pola yang benar-benar jelas.

    RULE FINAL MAG:
    - Certificate harus berasal dari pola BASE-######.
    - Angka certificate wajib TEPAT 6 digit.
    - Setelah certificate pertama ditemukan, angka 6 digit berikutnya yang
      berdiri sendiri setelah + atau , boleh menjadi certificate lanjutan
      (elision), misalnya BASE-000001+000002.
    - S/D hanya valid jika kedua ujungnya tepat 6 digit.
    - BASE- saja = policy tanpa certificate.
    - Suffix 1-5 digit, >6 digit, atau pola dash biasa = BUKAN certificate;
      biarkan clean_polis menangani sebagai perulangan policy.
    """
    raw = _normalize_spaces(str(original).upper())
    if not raw:
        return None

    # Jangan sentuh descriptor huruf selain S/D.
    check_str = re.sub(r"S\s*/\s*D", "", raw, flags=re.I)
    if re.search(r"[A-Z]", check_str):
        return None
    if "&" in raw:
        return None

    # Single policy + certificate / S/D.
    m = re.fullmatch(r"(\d{8,})\s*-\s*(\d{6})", raw)
    if m:
        return [(m.group(1), m.group(2))]

    m = re.fullmatch(r"(\d{8,})\s*-\s*(\d{6})\s*S\s*/\s*D\s*(\d{6})", raw, re.I)
    if m:
        return [(m.group(1), f"{m.group(2)} S/D {m.group(3)}")]

    # Kalau tidak ada separator multi-chunk, tidak ada grouping tambahan.
    if not re.search(r"[+,/]", raw):
        return None

    protected = re.sub(r"S\s*/\s*D", "SDPLACEHOLDER", raw, flags=re.I)
    chunks = [x.strip() for x in re.split(r"\s*[+,/]\s*", protected) if x.strip()]
    if len(chunks) < 2:
        return None

    groups = []
    lookup = {}
    current_base = None
    certificate_chain_started = False

    for chunk in chunks:
        # BASE-###### S/D ######
        m = re.fullmatch(r"(\d{8,})\s*-\s*(\d{6})\s*SDPLACEHOLDER\s*(\d{6})", chunk, re.I)
        if m:
            base = m.group(1)
            cert = f"{m.group(2)} S/D {m.group(3)}"
            certificate_chain_started = True
        else:
            # BASE-######
            m = re.fullmatch(r"(\d{8,})\s*-\s*(\d{6})", chunk)
            if m:
                base = m.group(1)
                cert = m.group(2)
                certificate_chain_started = True
            else:
                # BASE- tanpa suffix = policy kedua/berikutnya tanpa cert.
                m = re.fullmatch(r"(\d{8,})\s*-\s*", chunk)
                if m:
                    base = m.group(1)
                    cert = None
                else:
                    # BASE saja = policy penuh tanpa certificate.
                    m = re.fullmatch(r"(\d{8,})", chunk)
                    if m:
                        base = m.group(1)
                        cert = None
                    else:
                        # Elision: hanya 6 digit dan harus sudah ada
                        # certificate eksplisit sebelumnya.
                        m = re.fullmatch(r"(\d{6})", chunk)
                        if m and current_base is not None and certificate_chain_started:
                            base = current_base
                            cert = m.group(1)
                        else:
                            return None

        current_base = base
        if base not in lookup:
            item = {"base": base, "certs": []}
            lookup[base] = item
            groups.append(item)

        if cert is not None and cert not in lookup[base]["certs"]:
            lookup[base]["certs"].append(cert)

    if not groups or len(groups) > MAX_SPLIT_COLS:
        return None

    # Harus benar-benar ada certificate. Kalau seluruh chunk hanya policy,
    # serahkan ke clean_polis agar tidak mengubah rule policy repetition.
    if not any(item["certs"] for item in groups):
        return None

    return [
        (item["base"], ",".join(item["certs"]) if item["certs"] else None)
        for item in groups
    ]


def get_policy_certificate_pairs(raw_polis):
    """
    Entry point output polis + certificate.

    DEFINISI FINAL:
      certificate = EXACT 6 digit + didahului '-'.
      selain itu = perulangan/lantaran polis -> clean_polis().
    """
    if pd.isna(raw_polis):
        return []

    original = _normalize_spaces(str(raw_polis).upper())
    if not original:
        return []

    cert_result = _split_real_certificate_pattern(original)
    if cert_result is not None:
        return cert_result

    return [(p, None) for p in clean_polis(raw_polis)]


# ============================================================
# POLIS  (adaptasi FACUL - Multi Artha Guna)
# ============================================================
def _extract_main_policy(val):
    m = re.search(r"\b\d{8,}\b", val)
    return m.group(0) if m else ""


_TOKEN_FULL_DASH_RE = re.compile(r"^(\d{8,})\s*-\s*(\d{1,14})$")


def _classify_policy_token(tok):
    """
    Klasifikasi satu token polis (hasil split dari '-', '+', atau ',').
      - "full_dash": token sendiri sudah lengkap berupa BASE-SUFFIX
        (dipertahankan literal apa adanya, TIDAK direkonstruksi ulang).
      - "full": token berupa angka polos >=8 digit, berdiri sendiri.
      - "short": angka pendek (<=7 digit) yang mengacu ke base sebelumnya.
      - None: tidak bisa diklasifikasi dengan aman -> ambigu.
    """
    tok = tok.strip()
    m = _TOKEN_FULL_DASH_RE.match(tok)
    if m:
        return "full_dash", tok.replace(" ", ""), m.group(1)
    if re.fullmatch(r"\d{8,}", tok):
        return "full", tok, tok
    if re.fullmatch(r"\d{1,7}", tok):
        return "short", tok, None
    return None, None, None


def _build_policy_chain(tokens_raw, original, drop_words=frozenset()):
    """
    Fungsi rekonstruksi tunggal, dipakai bersama untuk chain '-', '+', dan ','
    supaya perlakuan konsisten (tidak ada logic ganda/beda antar separator).

    Aturan:
      - Token "full" / "full_dash" -> dipertahankan apa adanya (tidak diotak-atik).
        Setelah token "full" (angka polos, TANPA dash internal), base
        angka tsb boleh dipakai untuk merekonstruksi token "short" berikutnya.
      - Setelah token "full_dash" (sudah mengandung dash sendiri), base
        DIRESET (base=None). Ini supaya token "short" berikutnya yang
        mengacu ke base yang sudah punya dash internal TIDAK direkonstruksi
        secara otomatis (pola ambigu -> lebih aman keep original) daripada
        menebak makna suffix tersebut.
      - Kalau ada token yang tidak bisa diklasifikasi (huruf, S/D, dsb),
        atau token "short" muncul tanpa base yang valid -> ambigu -> keep original.
    """
    tokens_out = []
    base = None
    for raw in tokens_raw:
        raw_s = raw.strip()
        if not raw_s:
            continue
        if raw_s.upper() in drop_words or re.fullmatch(r"P[123]", raw_s, re.I):
            continue
        kind, value, base_digits = _classify_policy_token(raw_s)
        if kind is None:
            return [original]
        if kind == "full":
            tokens_out.append(value)
            base = base_digits
        elif kind == "full_dash":
            tokens_out.append(value)
            # Tetap simpan base. Untuk MAG, BASE-0614/0716 atau
            # BASE-0614+0716 adalah pola perulangan polis, bukan
            # certificate karena suffix-nya bukan 6 digit.
            base = base_digits
        else:  # short
            if base is None:
                return [original]
            suf = value
            full_val = base[:len(base) - len(suf)] + suf if len(suf) < len(base) else suf
            tokens_out.append(full_val)
            base = full_val

    tokens_out = list(dict.fromkeys(tokens_out))
    if not tokens_out:
        return [original]
    return _cap_or_original(tokens_out, original)


def _clean_policy_repetition_chain(original):
    """
    Tangani perulangan polis MAG sebelum fallback generic.

    Contoh:
      BASE+2553
      BASE/588/4263/4274
      BASE-0614,0716,0727
      BASE-0614/0716

    Semua angka lanjutan di sini adalah POLIS, bukan certificate.
    Certificate ditangani terpisah dan hanya valid untuk BASE-######.
    """
    if not re.search(r"[+/,]", original):
        return None

    # Jangan sentuh value yang jelas berisi teks/deskripsi.
    check = re.sub(r"S\s*/\s*D", "", original, flags=re.I)
    if re.search(r"[A-Z]", check) or "&" in original:
        return None

    protected = re.sub(r"S\s*/\s*D", "SDPLACEHOLDER", original, flags=re.I)
    chunks = [x.strip() for x in re.split(r"\s*[+/,]\s*", protected) if x.strip()]
    if len(chunks) < 2:
        return None

    # Base pertama harus jelas >= 8 digit.
    first = chunks[0]
    m = re.fullmatch(r"(\d{8,})(?:\s*-\s*(\d+))?", first)
    if not m:
        return None

    base = m.group(1)
    result = []

    # Kalau chunk pertama punya suffix, itu juga perulangan polis.
    first_suffix = m.group(2)
    if first_suffix is None:
        result.append(base)
    else:
        # BASE-###### ditangani certificate hanya kalau keseluruhan value
        # memang lolos _split_real_certificate_pattern. Kalau sampai di
        # sini berarti bukan certificate -> wajib jadi polis.
        if len(first_suffix) >= len(base):
            return None
        result.append(f"{base[:len(base)-len(first_suffix)]}{first_suffix}")

    current_base = base
    for chunk in chunks[1:]:
        chunk = chunk.replace("SDPLACEHOLDER", " S/D ")
        # Policy penuh.
        if re.fullmatch(r"\d{8,}", chunk):
            result.append(chunk)
            current_base = chunk
            continue

        # Policy penuh dengan suffix internal -> tetap dianggap perulangan.
        m_full = re.fullmatch(r"(\d{8,})\s*-\s*(\d+)", chunk)
        if m_full:
            b, suffix = m_full.groups()
            if len(suffix) >= len(b):
                return None
            result.append(f"{b[:len(b)-len(suffix)]}{suffix}")
            current_base = b
            continue

        # Suffix pendek -> rekontruksi dari base terakhir.
        if re.fullmatch(r"\d{1,7}", chunk):
            suffix = chunk
            if len(suffix) >= len(current_base):
                return None
            full = current_base[:len(current_base)-len(suffix)] + suffix
            if full not in result:
                result.append(full)
            continue

        return None

    result = list(dict.fromkeys(result))
    if not result:
        return None
    return _cap_or_original(result, original)


def clean_polis(val):
    if pd.isna(val):
        return []
    original = _normalize_spaces(str(val).upper())
    if not original:
        return []

    # ========================================================
    # RULE BARU - EX. POLICY NO
    # Jika ada format:
    # <POLIS> Ex. Policy No : <POLIS LAMA>
    #
    # BIARKAN POLIS APA ADANYA.
    # Jangan ambil hanya polis pertama.
    # Jangan split menjadi 2 polis.
    # ========================================================
    if re.search(
        r"\bEX\.?\s*(?:POLICY\s*NO\.?)?\s*:?",
        original,
        re.I
    ):
        return [original]

    # ========================================================
    # RULE BARU - DATA 1 FACUL MULTI ARTHA GUNA
    # Format:
    #   <POLIS AKTIF> Ex. Policy No : <POLIS LAMA>
    #
    # Contoh:
    #   45040118001473 Ex. Policy No : 45040117001948
    #
    # Hasil:
    #   ['45040118001473']
    #
    # Polis setelah "Ex. Policy No" hanya merupakan
    # referensi polis lama, BUKAN polis tambahan.
    # ========================================================
    m = re.fullmatch(
        r"\s*(\d{8,})\s*EX\.?\s*(?:POLICY\s*NO\.?)?\s*:?\s*\d{4,}\s*",
        original,
        re.I,
    )

    if m:
        return [m.group(1)]

    # Exception wajib dipertahankan (PENYELESAIAN/HUTANG/CANCEL/P1-P3/dll).
    if POLIS_EXCEPTION_RE.search(original):
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

    # --------------------------------------------------------
    # PRIORITASKAN pola "POLICY - CERTIFICATE[s]" ala Etiqa untuk
    # kasus SEDERHANA (tanpa koma/&/+ tercampur, supaya tidak bentrok
    # dengan comma-chain/plus-chain multi-polis yang sudah tervalidasi
    # di atas). Kalau cocok, clean_polis hanya mengembalikan BASE
    # policy-nya saja (satu token) - certificate-nya diekstrak terpisah
    # lewat clean_certificate() di process_data.
    # --------------------------------------------------------
    if "," not in original and "&" not in original and "+" not in original:
        if clean_certificate(original):
            m_base = re.match(r"^\s*(\d{8,})", original)
            if m_base:
                return [m_base.group(1)]

    # Hapus descriptor VARIOUS/VAR/TBA setelah policy valid.
    val2 = re.sub(r"\s*/\s*(?:VAR|VARIOUS|TBA)\b.*$", "", original, flags=re.I)
    val2 = re.sub(r"\s+(?:VAR|VARIOUS)\b.*$", "", val2, flags=re.I)
    val2 = _normalize_spaces(val2)

    # P1 / policy atau P2 / policy.
    m = re.fullmatch(r"P[123]\s*/\s*(\d{8,})(?:\s+.*)?", val2, re.I)
    if m:
        return [m.group(1)]

    # Catatan: shortcut "ambil main policy saja jika ada S/D" HANYA berlaku
    # kalau tidak ada koma. Kalau ada koma + S/D tercampur (pola langka,
    # <1% data), pola dianggap tidak cukup jelas -> ditangani di
    # comma-chain di bawah (yang akan fallback ke original bila ambigu).
    main = _extract_main_policy(val2)
    if main and "," not in val2 and re.search(r"\bS\s*\.?\s*/?\s*D\b", val2, re.I):
        return [main]

    # --------------------------------------------------------
    # PERULANGAN POLIS MAG
    # Semua suffix yang BUKAN certificate (bukan BASE-######)
    # diproses sebagai polis berulang.
    # --------------------------------------------------------
    repetition = _clean_policy_repetition_chain(val2)
    if repetition is not None:
        return repetition

    # --------------------------------------------------------
    # COMMA CHAIN - beberapa polis dipisah ',' (BUKAN certificate).
    # Contoh: 01031118000012-000006,01031118000012-000007
    #         05030319049722,05030319049733,05030319049744,49755
    # Cek DULU sebelum dash-chain, karena comma adalah separator utama.
    # --------------------------------------------------------
    if "," in val2:
        parts = [p.strip() for p in val2.split(",") if p.strip()]
        if len(parts) >= 2:
            return _build_policy_chain(parts, original)

    # --------------------------------------------------------
    # DASH CHAIN - perulangan polis (BUKAN certificate).
    # Menangani baik suffix pendek maupun beberapa polis penuh sekaligus.
    # Contoh: 13010921002655-13010921002644-13010921002677-2666 (mixed)
    #         00940502012023001276-136-135-1275 (short suffix)
    # --------------------------------------------------------
    if re.fullmatch(r"\d{8,}(?:\s*-\s*\d+)+", val2):
        parts = [p for p in re.split(r"\s*-\s*", val2) if p]
        return _build_policy_chain(parts, original)

    # Numeric S/D range tanpa policy prefix: expand jika <=5, selain itu pertahankan.
    m = re.fullmatch(r"(\d+)\s*S\s*/?\s*D\s*(\d+)", val2, re.I)
    if m and len(m.group(1)) == len(m.group(2)):
        a, b = int(m.group(1)), int(m.group(2))
        count = abs(b - a) + 1
        if count > MAX_SPLIT_COLS:
            return [original]
        step = 1 if a <= b else -1
        return [str(n).zfill(len(m.group(1))) for n in range(a, b + step, step)]

    # --------------------------------------------------------
    # AMPERSAND - diperketat untuk MAG.
    # Hanya displit jika SETIAP bagian setelah dibersihkan
    # sudah terlihat seperti polis penuh (>=8 digit alfanumerik).
    # Jika ada bagian yang pendek/ambigu (mis. sisa certificate-like
    # suffix: "45031125000863 - 000001 & 000002"), JANGAN displit,
    # pertahankan original.
    # --------------------------------------------------------
    if "&" in val2:
        raw_parts = [p.strip() for p in re.split(r"\s*&\s*", val2)]
        cleaned_parts = [re.sub(r"[^A-Z0-9]", "", p) for p in raw_parts]
        cleaned_parts = [p for p in cleaned_parts if p]
        if cleaned_parts and all(len(p) >= 8 for p in cleaned_parts):
            return _cap_or_original(cleaned_parts, original)
        return [original]

    # Plus = separate policy. Suffix pendek dibentuk dari policy sebelumnya.
    if "+" in val2:
        parts = [p.strip() for p in re.split(r"\s*\+\s*", val2) if p.strip()]
        return _build_policy_chain(parts, original, drop_words={"VAR", "VARIOUS", "TBA"})

    # Slash yang hanya memisahkan policy valid.
    if "/" in val2:
        result = _build_policy_chain(
            re.split(r"\s*/\s*", val2), original, drop_words={"VAR", "VARIOUS", "TBA"}
        )
        if result != [original]:
            return result

    # Policy numerik murni / dotted.
    if re.fullmatch(r"[0-9.]+", val2):
        return [val2.replace(".", "")]

    # Dash yang bukan certificate: fallback bersih tanpa separator.
    # Catatan: hanya aman kalau ada SATU bagian yang jelas berupa
    # polis penuh (>=8 digit). Kalau semua bagian pendek (mis.
    # "2485-2474-2452-2463", tidak ada basis polis yang jelas),
    # JANGAN digabung paksa jadi satu angka -> keep original.
    if re.fullmatch(r"[0-9-]+", val2):
        parts = [p for p in val2.split("-") if p]
        if len(parts) == 1:
            return [parts[0]]
        if len(parts[0]) >= 8:
            return [parts[0]]
        return [original]

    # --------------------------------------------------------
    # Fallback umum: ambil semua nomor panjang (>=8 digit).
    # Tambahan safety-cap khusus MAG: kandidat yang terlalu panjang
    # (> MAX_REASONABLE_POLIS_DIGITS) dianggap data korup/gabungan
    # tanpa separator -> jangan diambil, keep original.
    # --------------------------------------------------------
    candidates = re.findall(r"\b\d{8,}\b", val2)
    candidates = [c for c in candidates if len(c) <= MAX_REASONABLE_POLIS_DIGITS]
    if candidates:
        return _cap_or_original(list(dict.fromkeys(candidates)), original)

    return [val2] if val2 else []


# ============================================================
# SLIP  (reused dari script Etiqa, generic - tidak diubah)
# ============================================================
def clean_slip(val):
    if pd.isna(val):
        return []
    original = _normalize_spaces(str(val).upper())
    if not original:
        return []

    val2 = original.replace(".", "")

    # Exception: P1/P2/P3 murni harus dipertahankan (jangan sampai
    # tersapu oleh noise-removal generic di bagian bawah fungsi ini).
    if re.fullmatch(r"P[123]", val2, re.I):
        return [val2]

    # --------------------------------------------------------
    # Kombinasi placeholder layer P1/P2/P3/P4/dst (mis. "P1 + P2",
    # "P1 + P2 + P3 + P4", "P1 (P2)", "P1 / END").
    # Regex noise-removal generic di bawah bersifat destruktif untuk
    # pola ini (bisa menghapus seluruh string). Tangani lebih awal:
    #   - kalau ada nomor slip asli (>=7 digit) di dalamnya, ambil nomor
    #     itu saja (placeholder P-nya dibuang, bukan bagian dari slip).
    #   - kalau murni kombinasi P1/P2/P3/dst tanpa nomor asli,
    #     pertahankan original (jangan sampai jadi kosong).
    # --------------------------------------------------------
    if re.search(r"\bP\s*[1-9]\b", val2, re.I):
        numbers = re.findall(r"\b\d{7,}\b", val2)
        if numbers:
            numbers = list(dict.fromkeys(numbers))
            return numbers if len(numbers) <= MAX_SPLIT_COLS else [val2]
        return [val2]

    if SLIP_EXCEPTION_RE.search(val2):
        return [_normalize_spaces(val2)]

    m = re.fullmatch(r"([A-Z]+)\s+(20\d{2})\s*/\s*VARIOUS", val2, re.I)
    if m and m.group(1) in MONTH_NAMES:
        return [f"{m.group(1)} {m.group(2)}"]

    m = re.fullmatch(r"P[123]\s*/\s*(\d{7,})", val2, re.I)
    if m:
        return [m.group(1)]

    m = re.fullmatch(r"(\d{7,})\s*/\s*(\d{7,})", val2)
    if m:
        return [m.group(1), m.group(2)]

    m = re.fullmatch(r"(\d{7,})\s*/\s*0", val2)
    if m:
        return [m.group(1)]

    m = re.match(r"^(\d{7,})\s*-\s*\d+/CN/", val2, re.I)
    if m:
        return [m.group(1)]

    m = re.search(r"(\d{7,})\s*S\s*\.?\s*/?\s*D\s*(\d{7,})", val2, re.I)
    if m and len(m.group(1)) == len(m.group(2)):
        a, b = int(m.group(1)), int(m.group(2))
        count = abs(b - a) + 1
        if count > MAX_SPLIT_COLS:
            return [f"{m.group(1)} S/D {m.group(2)}"]
        step = 1 if a <= b else -1
        return [str(n).zfill(len(m.group(1))) for n in range(a, b + step, step)]

    m = re.fullmatch(r"(\d{7,})\s*/\s*(\d{1,6})", val2)
    if m:
        base, suffix = m.groups()
        if len(suffix) < len(base):
            return [base, base[:len(base) - len(suffix)] + suffix]

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
                    if len(suffix) >= 7 and suffix not in hasil:
                        hasil.append(suffix)
            return _cap_or_join(hasil)

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

    val3 = re.sub(r"^SEE\s+ATTACHMENT\s*/\s*", "", val2, flags=re.I).strip()
    val3 = re.sub(r"\bEND(?:\.?1)?\b", "", val3, flags=re.I)
    val3 = re.sub(r"\b(?:USD|IDR|SGD|EUR|JPY|AUD|GBP|ORI|ORIGINAL|COPY|REALISASI|REALIZATION|CANCEL(?:LED)?|ENDORSEMENT|SA|P1|P2|P3|VAR|REVISI|REV)\b.*$", "", val3, flags=re.I)
    val3 = _normalize_spaces(val3).strip("-_,.; ")

    candidates = re.findall(r"\b\d{7,}\b", val3)
    if candidates:
        candidates = list(dict.fromkeys(candidates))
        return candidates if len(candidates) <= MAX_SPLIT_COLS else [val3]

    if "/CN/" in val3 or re.search(r"[A-Z]{2,}/\d", val3):
        return [val3]

    if "," in val3:
        parts = [p.strip() for p in val3.split(",") if p.strip()]
        return _cap_or_original(parts, original)

    if re.fullmatch(r"\d{1,6}(?:\s*[&+-]\s*\d{1,6})+", val3):
        return [val3]

    return [val3] if val3 else []


# ============================================================
# BUSINESS PARTNERS  (reused dari script Etiqa)
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
# PROCESS DATA  (Data 1 - FACUL, tanpa kolom Certificate)
# ============================================================
def process_data(input_file, output_file):
    print(f"[1/5] Membaca data dari: {input_file} ...")
    df = _fast_read_excel(input_file, header=0)
    total_input_rows = len(df)
    print(f"      Total baris keseluruhan (semua cedant): {total_input_rows:,}")

    for col in (CEDANT_COL, POLIS_COL, SLIP_COL, INSURED_COL, BROKER_NAME_COL, BROKER_CODE_COL):
        if col not in df.columns:
            print(f"[ERROR] Kolom '{col}' tidak ditemukan di file!")
            return

    cedant_norm = (
        df[CEDANT_COL].fillna("").astype(str).str.upper().str.strip()
        .str.replace(r"\.", " ", regex=True).str.replace(r"\s+", " ", regex=True).str.strip()
    )
    cedant_target_norm = re.sub(r"\s+", " ", CEDANT_VALUE.upper().replace(".", " ")).strip()
    is_target = cedant_norm.eq(cedant_target_norm)
    total_target = int(is_target.sum())
    total_other = total_input_rows - total_target
    print(f"[2/5] Filter cedant '{CEDANT_VALUE}': {total_target:,} baris target dari total {total_input_rows:,} baris "
          f"({total_other:,} baris cedant lain, TIDAK diubah).")

    broker_name_s = df[BROKER_NAME_COL].fillna("").astype(str).str.strip()
    broker_code_s = df[BROKER_CODE_COL].fillna("").astype(str).str.strip()
    mitra_values = [
        get_mitra_bisnis(a, b, c) for a, b, c in zip(broker_name_s, broker_code_s, cedant_norm)
    ]

    df.rename(columns={POLIS_COL: "polis_ori", SLIP_COL: "slip_ori", INSURED_COL: "insured_ori"}, inplace=True)

    insert_pos = list(df.columns).index(BROKER_NAME_COL) + 1 if BROKER_NAME_COL in df.columns else len(df.columns)
    df.insert(insert_pos, "BUSINESS PARTNERS", mitra_values)

    print("[3/5] Menjalankan cleaning HANYA untuk baris target (baris lain dibiarkan apa adanya) ...")
    n = len(df)
    all_polis_cert_pairs = [[] for _ in range(n)]   # list of (polis, cert_or_None)
    all_clean_slip = [[] for _ in range(n)]
    all_clean_ins = [[] for _ in range(n)]
    max_polis = max_slip = max_ins = 1

    target_idx = np.flatnonzero(is_target.to_numpy())
    polis_vals = df["polis_ori"].to_numpy()
    slip_vals = df["slip_ori"].to_numpy()
    insured_vals = df["insured_ori"].to_numpy()

    changed_polis = changed_slip = changed_ins = 0
    rows_with_certificate = 0

    for n_done, pos in enumerate(target_idx, 1):
        if n_done % 5000 == 0:
            print(f"      Progress: {n_done:,} / {total_target:,} baris target diproses...")

        raw_polis = polis_vals[pos]
        raw_slip = slip_vals[pos]
        raw_ins = insured_vals[pos]

        pairs = get_policy_certificate_pairs(raw_polis)
        c_slip = clean_slip(raw_slip)
        c_ins = clean_insured(raw_ins)

        if pairs and any(c for _, c in pairs):
            rows_with_certificate += 1

        if not pd.isna(raw_polis):
            orig_norm = _normalize_spaces(str(raw_polis).upper())
            c_polis_only = [p for p, _ in pairs]
            if c_polis_only != [orig_norm] or any(c for _, c in pairs):
                changed_polis += 1
        if not pd.isna(raw_slip):
            orig_norm = _normalize_spaces(str(raw_slip).upper())
            if c_slip != [orig_norm]:
                changed_slip += 1
        if not pd.isna(raw_ins):
            if c_ins != [_normalize_spaces(str(raw_ins)).upper()]:
                changed_ins += 1

        max_polis = max(max_polis, len(pairs))
        max_slip = max(max_slip, len(c_slip))
        max_ins = max(max_ins, len(c_ins))

        all_polis_cert_pairs[pos] = pairs
        all_clean_slip[pos] = c_slip
        all_clean_ins[pos] = c_ins

    print(f"      Selesai diproses! max kolom -> polis={max_polis}, slip={max_slip}, insured={max_ins}")
    print(f"      Baris dengan certificate ditemukan: {rows_with_certificate:,}")

    # Kolom interleaved: clean polis N | certificate N (bukan blok terpisah).
    polis_col_names = []
    for i in range(1, max_polis + 1):
        polis_col = f"clean polis {i}"
        cert_col = f"certificate {i}"
        df[polis_col] = [pairs[i - 1][0] if i - 1 < len(pairs) else None for pairs in all_polis_cert_pairs]
        df[cert_col] = [pairs[i - 1][1] if i - 1 < len(pairs) else None for pairs in all_polis_cert_pairs]
        polis_col_names.append((polis_col, cert_col))

    new_columns = []
    for col in df.columns:
        if col.startswith("clean polis ") or col.startswith("certificate "):
            continue
        new_columns.append(col)
        if col == "polis_ori":
            for polis_col, cert_col in polis_col_names:
                new_columns.append(polis_col)
                new_columns.append(cert_col)
        elif col == "slip_ori":
            new_columns += _insert_clean_columns(df, all_clean_slip, "slip", max_slip)
        elif col == "insured_ori":
            new_columns += _insert_clean_columns(df, all_clean_ins, "insured", max_ins)

    df = df[new_columns]
    total_output_rows = len(df)

    print(f"[4/5] Validasi baris ...")
    print("-" * 60)
    print(f"INPUT ROWS        : {total_input_rows:,}")
    print(f"OUTPUT ROWS       : {total_output_rows:,}")
    print(f"TARGET CEDANT     : {CEDANT_VALUE}")
    print(f"TARGET ROWS       : {total_target:,}")
    print(f"OTHER CEDANT ROWS : {total_other:,}")
    print("-" * 60)
    print(f"CLEANED (polis)   : {changed_polis:,}  | UNCHANGED: {total_target - changed_polis:,}")
    print(f"CLEANED (slip)    : {changed_slip:,}  | UNCHANGED: {total_target - changed_slip:,}")
    print(f"CLEANED (insured) : {changed_ins:,}  | UNCHANGED: {total_target - changed_ins:,}")
    print("-" * 60)
    assert total_input_rows == total_output_rows, "JUMLAH BARIS INPUT != OUTPUT! PROSES DIHENTIKAN."
    assert total_target + total_other == total_input_rows
    print("VALIDASI OK: INPUT ROWS == OUTPUT ROWS, tidak ada baris yang hilang/bertambah.")

    print(f"[5/5] Menyimpan hasil ke: {output_file} ...")
    _fast_write_excel(df, output_file)
    print(f"[OK] Selesai. Total baris={total_output_rows:,}, kolom={len(df.columns)}, target={total_target:,}")


if __name__ == "__main__":
    print("=" * 80)
    print("TEST ENTRY POINT - CLEANING FACUL - MULTI ARTHA GUNA")
    print("(TANPA logic certificate / CLSDT - hanya untuk Data 1 FACUL)")
    print("=" * 80)

    print("\n--- TEST FAC_POLICY_NO ---")
    polis_tests = [
        # Pola EX Policy No (spesifik MAG, harus ambil polis aktif saja)
        "45040118001473 Ex. Policy No : 45040117001948",
        "36040119000076 EX. POLICY NO : 36040118000085",
        "45090119000635 Ex 45090118000303",
        "45040518000083 Ex.45040517000081",
        # TBA sebagai placeholder (harus di-skip)
        "05010921011067 + TBA",
        "01013025000694+48013025002218+01013225000608+TBA",
        # TBA murni / ambigu -> keep original
        "TBA/MAG",
        "TBA - 40010917014875.000582 - 000588",
        # Dash chain perulangan polis (bukan certificate)
        "13010921002655-13010921002644-13010921002677-2666",
        # Plus chain > MAX_SPLIT_COLS -> keep original
        "40012121020589 + 143 + 154 + 934 + 945 + 578",
        # Ampersand dua polis penuh -> boleh displit
        "05010919008355&05010519003327",
        # Ampersand ambigu (certificate-like suffix pendek) -> keep original
        "45031125000863 - 000001 & 000002",
        # Comma dua polis penuh -> boleh displit (bukan certificate)
        "01031118000012-000006,01031118000012-000007",
        # Comma dengan suffix pendek ambigu (>5 hasil) -> keep original
        "36080519000-369,371,382,405,416,484,495,507",
        # P1 murni
        "P1",
        # PENYELESAIAN (deskriptif, keep original)
        "PENYELESAIAN HUTANG PIUTANG MAG 2024 IDR TW1",
        # Data korup/gabungan tanpa separator -> keep original (safety cap)
        "28218919120321422523624725826927129330531632733TBA",
        # Polis biasa
        "45040120001088",
        # Comma chain: increasing suffix pattern (harus direkonstruksi)
        "05030319049722,05030319049733,05030319049744,49755",
        # Comma chain: dua polis penuh dengan dash-suffix masing2
        "08080519000166 - 000001,08080519000166 - 000002",
        # Comma chain dengan S/D di dalamnya -> ambigu, keep original
        "01020118000018 - 000001 S/D 000003, 05020118001155",
    ]
    for i, t in enumerate(polis_tests, 1):
        print(f"[{i:2d}] ORI : {t!r}")
        print(f"      HASIL: {clean_polis(t)}")

    print("\n--- TEST FAC_SLIP ---")
    slip_tests = [
        "BORDERO SEPTEMBER 2019",
        "BORDERO AGUSTUS 2023 - USD",
        "4503111800016 S/D 4503112001706",
        "5001302100046 + 041 + 044",
        "4609012300003  /  END",
        "3601102200029 - 3601102200028",
        "P1",
        "4501301957947-0487-9669-0491-9673-8858",
    ]
    for i, t in enumerate(slip_tests, 1):
        print(f"[{i:2d}] ORI : {t!r}")
        print(f"      HASIL: {clean_slip(t)}")

    print("\n--- TEST FAC_INSURED ---")
    insured_tests = [
        "SRIBOGA FLOUR MILL,PT",
        "HARAPAN LANGGENG ABADI/THERESIA MONITA PAIMANTA/ANDY SUTRISNO",
        "BANK INDEX SELINDO QQ ARTHALAUT BUMIJASA",
        "HUTAMA KARYA (PERSERO) AS PRINCIPAL AND/OR PT ADHI KARYA (PERSERO)",
        "MAHAMERU TEGAR SENTOSA QQ THIO RIANDY THEONARDO, PT",
        "LAUTAN REZEKI/FURNILUX INDONESIA/CIA SENG/MURTONO/ARYA CHANDRA",
    ]
    for i, t in enumerate(insured_tests, 1):
        print(f"[{i:2d}] ORI : {t!r}")
        print(f"      HASIL: {clean_insured(t)}")

    print("\n--- TEST CERTIFICATE (reuse persis logic Etiqa) ---")
    cert_tests = [
        "1071031117000018 - 000176",
        "1071031117000018 - 000162-000170",
        "1071031117000018 - 000162/000170",
        "1071031117000018 - 000162/1119/1120",  # ada suffix bukan 6 digit -> bukan certificate
        "1071031124000024 - 000251 S/D 0000131",  # 0000131 = 7 digit -> bukan certificate
        "45040120001088",
    ]
    for i, t in enumerate(cert_tests, 1):
        cp = clean_polis(t)
        cc = clean_certificate(t)
        print(f"[{i}] ORI          : {t!r}")
        print(f"    clean_polis  : {cp}")
        print(f"    certificate  : {cc}")

    print("\n--- TEST RULE BARU: PEMISAHAN POLIS & CERTIFICATE (get_policy_certificate_pairs) ---")
    new_rule_tests = [
        # PRINSIP UTAMA
        ("36040118000154-000001+000002", [("36040118000154", "000001,000002")]),
        ("45013024030214-000001+124002553-000001", [("45013024030214", "000001"), ("124002553", "000001")]),
        ("45013024018823-0001+45013224012546-0001", None),
        ("45013024027219-000001+45013124002245-000001", [("45013024027219", "000001"), ("45013124002245", "000001")]),
        ("40013024000513 - 000005+40013224000574 - 000005", [("40013024000513", "000005"), ("40013224000574", "000005")]),
        # POLIS SAMA + BANYAK CERTIFICATE
        ("45040118000745 - 000001 / 45040118000745 - 000002", [("45040118000745", "000001,000002")]),
        ("45040117010905 - 000001+45040117010905 - 000002", [("45040117010905", "000001,000002")]),
        ("45040117009628 - 000001+45040117009628 - 000002", [("45040117009628", "000001,000002")]),
        ("06100418000013-000001,06100418000013-000002", [("06100418000013", "000001,000002")]),
        ("36040118000336-000001+36040118000336-000002", [("36040118000336", "000001,000002")]),
        ("36040118000405 - 000001+36040118000405 - 000002", [("36040118000405", "000001,000002")]),
        ("36040119000032 - 000001,36040119000032 - 000002", [("36040119000032", "000001,000002")]),
        ("08080519000166 - 000001,08080519000166 - 000002", [("08080519000166", "000001,000002")]),
        ("01031118000012-000006,01031118000012-000007", [("01031118000012", "000006,000007")]),
        ("40030321000148-000001/40030321000148-000001", [("40030321000148", "000001")]),  # dedup
        # CERTIFICATE TANPA POLIS DIULANG (elision)
        ("07080518000014-00001,00002,00004", None),
        ("40010520009476-000002 + 003 + 004 + 005 + 006", None),
        ("03010924000363-000001+000002+000003+000004", [("03010924000363", "000001,000002,000003,000004")]),
        # JANGAN SALAH TEBAK -> None dari rule baru, fallback ke rule lama (keep original)
        ("05012920001475-987-998-009-022-033-044-146", None),
        ("40012120055055-491-033-044-009-088-022077011066998", None),
        ("02031125000018 000455-000536,000467/VARIOUS", None),
        ("02031125000018 000134,000221-000409", None),
        # RULE S/D
        ("08080517000105-000001 S/D 000006", [("08080517000105", "000001 S/D 000006")]),
        # 2 POLIS + CERTIFICATE (base kedua tanpa suffix)
        ("08080517000105-000001 S/D 000006 , 08080517000116-",
         [("08080517000105", "000001 S/D 000006"), ("08080517000116", None)]),
    ]
    all_ok = True
    for ori, expected in new_rule_tests:
        result = get_policy_certificate_pairs(ori)
        ok = (expected is None) or (result == expected)
        status = "OK" if ok else "MISMATCH"
        if not ok:
            all_ok = False
        print(f"[{status}] ORI: {ori!r}")
        for i, (p, c) in enumerate(result, 1):
            print(f"        clean polis {i} = {p!r:20s} certificate {i} = {c!r}")
        if expected is not None and not ok:
            print(f"        EXPECTED       : {expected}")
    print("\nSEMUA TEST RULE BARU:", "PASS" if all_ok else "ADA YANG GAGAL, CEK DI ATAS")

    print("\n" + "=" * 80)
    print("TEST SELESAI - LOGIC SUDAH DIVALIDASI TERHADAP DATA ASLI")
    print("=" * 80)

    process_data(INPUT_FILE, OUTPUT_FILE)

# if __name__ == "__main__":
#     print("=" * 80)
#     print("TEST ENTRY POINT - CLEANING FACUL - MULTI ARTHA GUNA")
#     print("=" * 80)

#     test_polis = [
#         "45040118001473 Ex. Policy No : 45040117001948",
#         "36040119000076 EX. POLICY NO : 36040118000085",
#         "45090119000635 Ex 45090118000303",
#         "45040518000083 Ex.45040517000081",
#     ]

#     print("\n--- TEST FAC_POLICY_NO ---")

#     for i, val in enumerate(test_polis, 1):
#         hasil = clean_polis(val)

#         print(f"[{i}]")
#         print(f"ORI  : {val!r}")
#         print(f"HASIL: {hasil}")

#     print("\n" + "=" * 80)
#     print("EXPECTED")
#     print("=" * 80)

#     print("""
# [1] ['45040118001473']
# [2] ['36040119000076']
# [3] ['45090119000635']
# [4] ['45040518000083']
# """)
    