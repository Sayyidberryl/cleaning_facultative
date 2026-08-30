import os
import re

import pandas as pd


INPUT_FILE = os.path.join("input", "inosbal_facul_rev1.xlsx")
OUTPUT_FILE = os.path.join("output", "wahana_output_osbal_rev2.xlsx")

CEDANT_COL   = "CCOS_COMP_NAME"
CEDANT_VALUE = "PT.ASURANSI WAHANA TATA"

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

    # Hapus token PT / QQ dimanapun posisinya (prefix, tengah, dsb.)
    name = INSURED_PT_QQ_RE.sub(" ", name)
    # Hapus tanda kurung, isi di dalamnya dipertahankan
    name = name.replace("(", " ").replace(")", " ")
    name = _normalize_spaces(name)
    # Bersihkan sisa tanda baca (titik/koma/strip) di ujung
    name = re.sub(r"^[\s.,\-]+|[\s.,\-]+$", "", name)

    return name


def _cap_or_join(items: list) -> list:
    """Jika jumlah item melebihi MAX_SPLIT_COLS, gabungkan dengan koma."""
    if len(items) > MAX_SPLIT_COLS:
        return [",".join(items)]
    return items


# ─────────────────────────────────────────────────────────────────────────────
# CLEAN POLIS
# ─────────────────────────────────────────────────────────────────────────────
def _clean_polis_core(val) -> list:
    """
    CLEAN POLIS - WAHANA REV.2

    RULE:
    - POLIS_CLN_1 = base / nomor polis utama.
    - Suffix yang merupakan pola pengulangan dibuat menjadi
      POLIS_CLN_2, POLIS_CLN_3, dst.
    - Panjang suffix boleh 2, 3, 4, 5, atau 6 digit.
    - Untuk beberapa suffix, semua suffix harus memiliki
      panjang digit yang sama.
    - Base tidak ikut dipotong selain mengganti digit terakhir
      sebanyak panjang suffix.
    - Dash biasa yang tidak jelas sebagai pola pengulangan
      tetap dibiarkan sebagai 1 polis.
    """

    if pd.isna(val):
        return []

    val = str(val).strip()

    if not val:
        return []

    # ============================================================
    # EXCEPTION
    # ============================================================

    if POLIS_EXCEPTION_RE.search(val):
        return [_normalize_spaces(val)]

    if re.search(r"\d+TBA\d+", val, re.IGNORECASE):
        return [_normalize_spaces(val)]

    val_upper = val.upper().strip()

    # ============================================================
    # HAPUS KETERANGAN SETELAH "/"
    # Contoh:
    # 098.4050.201.2024.000019.00/CN/FPAR/23/11/132
    # ->
    # 098.4050.201.2024.000019.00
    # ============================================================

    if "/" in val:
        val = val.split("/", 1)[0].strip()

    # ============================================================
    # VARIOUS / TBA
    # ============================================================

    if re.fullmatch(r"VARIOUS", val_upper):
        return [val.strip()]

    if re.match(
        r"^VARIOUS\s*-\s*SEE\s+ATTACH",
        val_upper
    ):
        return [val.strip()]

    if re.fullmatch(r"TBA", val_upper):
        return [val.strip()]

    # ============================================================
    # P1 / NOMOR POLIS
    #
    # P1/023.4050.601.2025.000063.00
    #
    # -> 023.4050.601.2025.000063.00
    # ============================================================

    if re.match(
        r"^\s*P1\s*/",
        val,
        re.IGNORECASE
    ):

        after = re.sub(
            r"^\s*P1\s*/\s*",
            "",
            val,
            flags=re.IGNORECASE
        ).strip()

        tokens = _extract_polis_tokens(after)

        if tokens:
            return [_normalize_spaces(tokens[0])]

    # ============================================================
    # P1 + P2 / P1 (P2)
    #
    # Jangan breakdown.
    # ============================================================

    if re.fullmatch(
        r"\s*P1\s*(?:\+|\()\s*P2\s*\)?\s*",
        val,
        re.IGNORECASE
    ):
        return [_normalize_spaces(val)]

    # ============================================================
    # TEKS DI DEPAN + NOMOR POLIS
    #
    # LINE SLIP ASWATA Q3 IDR    026.4050.201.2025.000357.00
    #
    # -> 026.4050.201.2025.000357.00
    # ============================================================

    if re.search(r"[A-Za-z]", val):

        tokens = _extract_polis_tokens(val)

        if tokens:

            # Cari bagian pertama yang benar-benar dimulai angka
            m = re.search(r"\d", val)

            if m:
                prefix = val[:m.start()].strip()

                if prefix:
                    return [_normalize_spaces(tokens[0])]

    # ============================================================
    # POLA S/D
    # ============================================================

    if re.search(r"S/D", val, re.IGNORECASE):

        sd_parts = re.split(
            r"\s{2,}",
            val.strip()
        )

        left = sd_parts[0] if sd_parts else val

        base_m = (
            re.match(
                r"^(\d[\d\-]+?)-\d+S/D\d+",
                left,
                re.IGNORECASE,
            )
            or
            re.match(
                r"^(\d[\d\-]+)S/D\d+",
                left,
                re.IGNORECASE,
            )
        )

        results = (
            [base_m.group(1).strip()]
            if base_m
            else []
        )

        for part in sd_parts[1:]:

            part = re.sub(
                r"\+?\s*\bTBA\b\s*",
                "",
                part,
                flags=re.IGNORECASE,
            ).strip().strip("+")

            results += [
                t
                for t in part.split()
                if _is_valid_polis_token(t.strip())
            ]

        if results:
            return _cap_or_join(results)

    # ============================================================
    # TBA DI AWAL + SPASI GANDA
    # ============================================================

    if re.match(
        r"^\s*TBA\s{2,}",
        val,
        re.IGNORECASE
    ):

        rest = re.sub(
            r"^\s*TBA\s+",
            "",
            val,
            flags=re.IGNORECASE
        ).strip()

        tokens = _extract_polis_tokens(rest)

        if tokens:
            return _cap_or_join(tokens)

    # ============================================================
    # POLA DASH - REV.2
    #
    # CONTOH:
    #
    # 09840502012020000187-000156
    #
    # base   = 09840502012020000187
    # suffix = 000156
    #
    # hasil:
    # 09840502012020000187
    # 09840502012020000156
    #
    # --------------------------------
    #
    # 0041050201202000183-0045
    #
    # hasil:
    # 0041050201202000183
    # 0041050201202000045
    #
    # --------------------------------
    #
    # 09840502012020000115-110-114-109
    #
    # hasil:
    # base
    # base + 110
    # base + 114
    # base + 109
    # ============================================================

    dash_match = re.fullmatch(
        r"(\d+)((?:-\d+)+)",
        val
    )

    if dash_match:

        base = dash_match.group(1)

        suffix_part = dash_match.group(2)

        suffixes = [
            x
            for x in suffix_part.split("-")
            if x != ""
        ]

        # ============================================================
        # KHUSUS:
        # Jika suffix pertama > 6 digit, ambil BASE saja.
        #
        # Contoh:
        # 02210502012023000494-70320230003-002
        #
        # Hasil:
        # 02210502012023000494
        # ============================================================

        if suffixes and len(suffixes[0]) > 6:
            return [base]

        # KHUSUS:
        # Kalau ada suffix 6 digit, hanya suffix 6 digit yang dipakai.
        # Suffix 4 digit seperti 0006 dan 0005 diabaikan.

        suffix_6_digit = [
            s for s in suffixes
            if len(s) == 6
        ]

        if suffix_6_digit and len(base) > 6:
            results = [base]

            for suffix in suffix_6_digit:
                new_polis = base[:-6] + suffix

                if new_polis not in results:
                    results.append(new_polis)

            return _cap_or_join(results)

        # ============================================================
        # POLA SUFFIX CAMPURAN 4 DAN 6 DIGIT
        #
        # Contoh:
        # 00940502012023001727-1726-000163-000164-001728
        #
        # Base   : 00940502012023001727
        # 1726   : ganti 4 digit terakhir
        # 000163 : ganti 6 digit terakhir
        # 000164 : ganti 6 digit terakhir
        # 001728 : ganti 6 digit terakhir
        #
        # Hasil:
        # 00940502012023001727
        # 00940502012023001726
        # 00940502012023000163
        # 00940502012023000164
        # 00940502012023001728
        # ============================================================

        if (
            len(suffixes) >= 2
            and all(re.fullmatch(r"\d+", s) for s in suffixes)
            and all(len(s) in {4, 6} for s in suffixes)
            and len(base) > 6
        ):
            results = [base]

            for suffix in suffixes:

                suffix_len = len(suffix)

                new_polis = (
                    base[:-suffix_len] + suffix
                )

                if new_polis not in results:
                    results.append(new_polis)

            return _cap_or_join(results)

        # ============================================================
        # POLA TIDAK BERATURAN
        # Contoh:
        # 01540502012021000442-51481443442445452453457595861
        # 01540502012023000758759-000165000164
        #
        # Dibiarkan apa adanya.
        # ============================================================

        if len(suffixes) == 1:

            second = suffixes[0]

            # Jika suffix bukan pola penggantian digit
            # (terlalu panjang atau panjangnya tidak masuk akal)
            if (
                len(second) > 6
                and len(second) < len(base)
            ):
                return [_normalize_spaces(val)]

        # ============================================================
        # Jika hanya ada SATU dash dan panjang suffix > 6 digit,
        # anggap dua nomor polis berbeda.
        #
        # Contoh:
        # 09810502012022000192-09810502022022000141
        # ->
        # 09810502012022000192
        # 09810502022022000141
        # ============================================================

        # ============================================================
        # POLA TIDAK BERATURAN
        #
        # Contoh:
        # 01540502012021000442-51481443442445452453457595861
        # 01540502012023000758759-000165000164
        #
        # Biarkan apa adanya.
        # ============================================================

        if len(suffixes) == 1:

            # ============================================================
            # POLA TIDAK BERATURAN
            #
            # Contoh:
            # 01540502012021000442-51481443442445452453457595861
            # 01540502012023000758759-000165000164
            #
            # Biarkan apa adanya.
            # ============================================================

            if (
                len(second) > 6
                and len(second) < len(base)
                and not (
                    len(second) == len(base)
                    or second.startswith(base[:8])
                )
            ):
                return [_normalize_spaces(val)]

            second = suffixes[0]

            # suffix terlalu panjang tetapi bukan nomor polis penuh
            if (
                len(second) > 6
                and len(second) < len(base)
            ):
                return [_normalize_spaces(val)]

        if len(suffixes) == 1 and len(suffixes[0]) > 6:
            return _cap_or_join([
                base,
                suffixes[0]
            ])

        # --------------------------------------------------------
        # HARUS ADA SUFFIX
        # --------------------------------------------------------

        if suffixes:

            # ========================================================
            # KHUSUS WAHANA:
            # HANYA SUFFIX 6 DIGIT YANG BOLEH DIPROSES
            #
            # Contoh:
            # 02210502012023000128-000129-0006-0005
            #
            # 000129 -> DIPAKAI
            # 0006   -> DIABAIKAN
            # 0005   -> DIABAIKAN
            #
            # Hasil:
            # 02210502012023000128
            # 02210502012023000129
            # ========================================================

            suffix_6_digit = [
                s for s in suffixes
                if len(s) == 6
            ]

            if suffix_6_digit and len(base) > 6:

                results = [base]

                for suffix in suffix_6_digit:
                    new_polis = base[:-6] + suffix

                    if new_polis not in results:
                        results.append(new_polis)

                return _cap_or_join(results)

            # ========================================================
            # Kalau TIDAK ADA suffix 6 digit,
            # jangan proses suffix 4 digit / 3 digit / 2 digit
            # sebagai pengulangan polis.
            # ========================================================

            return [_normalize_spaces(val)]
    # ============================================================
    # POLA SPASI GANDA
    #
    # Contoh:
    #
    # 09810502012021000074-00073
    #     098.1050.201.2021.000146.00
    #
    # Polis utama tetap diproses dari bagian pertama.
    # ============================================================

    if re.search(r"\s{2,}", val):

        blocks = re.split(
            r"\s{2,}",
            val.strip()
        )

        first_block = blocks[0].strip()

        # --------------------------------------------------------
        # Kalau blok pertama punya pola dash,
        # proses ulang menggunakan rule dash di atas.
        # --------------------------------------------------------

        dash_first = re.fullmatch(
            r"(\d+)((?:-\d+)+)",
            first_block
        )

        if dash_first:

            base = dash_first.group(1)

            suffixes = [
                x
                for x in dash_first.group(2).split("-")
                if x
            ]

            if suffixes:

                lengths = {
                    len(s)
                    for s in suffixes
                }

                if len(lengths) == 1:

                    suffix_len = len(suffixes[0])

                    if 2 <= suffix_len <= 6:

                        if len(base) > suffix_len:

                            results = [base]

                            for suffix in suffixes:

                                results.append(
                                    base[:-suffix_len]
                                    + suffix
                                )

                            # ------------------------------------------------
                            # Tambahkan nomor polis dari blok berikutnya
                            # ------------------------------------------------

                            for block in blocks[1:]:

                                tokens = _extract_polis_tokens(
                                    block
                                )

                                for token in tokens:

                                    if _is_valid_polis_token(
                                        token.strip()
                                    ):
                                        results.append(
                                            token.strip()
                                        )

                            return _cap_or_join(results)

        # --------------------------------------------------------
        # Kalau bukan pola dash,
        # ambil nomor polis dari SEMUA blok (dipisah spasi panjang),
        # masing-masing jadi entry breakdown terpisah.
        #
        # Contoh:
        # 015.1050.201.2023.000716.00        000.6005.202.2023.001965.01
        # ->
        # 015.1050.201.2023.000716.00
        # 000.6005.202.2023.001965.01
        # --------------------------------------------------------

        results = []

        for block in blocks:

            tokens = _extract_polis_tokens(block)

            if tokens:
                results.append(_normalize_spaces(tokens[0]))

            # --------------------------------------------------------
            # POLA PENGULANGAN
            #
            # Contoh:
            # 004.1050.101.2023.000256.00  000262.00  000276.00
            # --------------------------------------------------------

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

    # ============================================================
    # POLA SLASH
    #
    # 091.1050.601.2021.000002.00/PUH2100331
    #
    # ->
    # 091.1050.601.2021.000002.00
    # ============================================================

    if "/" in val:

        slash_parts = re.split(
            r"\s*/\s*",
            val,
            maxsplit=1
        )

        left = slash_parts[0].strip()

        if left:

            if _is_valid_polis_token(left):
                return [
                    _normalize_spaces(left)
                ]

            tokens = _extract_polis_tokens(left)

            if tokens:
                return [
                    _normalize_spaces(tokens[0])
                ]

    # ============================================================
    # POLA DASH + PLUS CAMPURAN
    #
    # Contoh:
    # 004.4050.202.2023.0015-68+69+70+79
    #
    # base   = 004.4050.202.2023.0015
    # suffix = 68, 69, 70, 79 (semua 2 digit, konsisten)
    #
    # hasil:
    # 004.4050.202.2023.0015
    # 004.4050.202.2023.0068
    # 004.4050.202.2023.0069
    # 004.4050.202.2023.0070
    # 004.4050.202.2023.0079
    # ============================================================

    dash_plus_match = re.fullmatch(
        r"([\d.]+)-(\d+(?:\+\d+)+)",
        val
    )

    if dash_plus_match:

        base = dash_plus_match.group(1)
        suffixes = [
            s for s in dash_plus_match.group(2).split("+") if s
        ]

        if suffixes and all(re.fullmatch(r"\d+", s) for s in suffixes):

            suffix_lengths = {len(s) for s in suffixes}

            if len(suffix_lengths) == 1:

                suffix_len = suffix_lengths.pop()

                if 2 <= suffix_len <= 6 and len(base) > suffix_len:

                    results = [base] + [
                        base[:-suffix_len] + suffix for suffix in suffixes
                    ]

                    return _cap_or_join(results)

        # Suffix tidak konsisten panjangnya -> jangan mengarang,
        # pertahankan nilai aslinya apa adanya.
        return [_normalize_spaces(val)]

    # ============================================================
    # POLA PLUS - DUA NOMOR POLIS LENGKAP
    #
    # Contoh:
    # 022.4050.201.2023.000394 + 022.4050.202.2023.00077
    #
    # ->
    # 022.4050.201.2023.000394
    # 022.4050.202.2023.00077
    # ============================================================

    if "+" in val:

        plus_parts = [
            p.strip()
            for p in re.split(r"\s*\+\s*", val)
            if p.strip()
        ]

        if len(plus_parts) > 1:

            # Semua bagian harus berupa nomor polis lengkap
            if all(
                re.fullmatch(r"\d+(?:\.\d+)+", p)
                for p in plus_parts
            ):
                return _cap_or_join(plus_parts)

    # ============================================================
    # HAPUS P3, P4, P5, DST
    # ============================================================

    val = re.sub(
        r"\s*\+\s*P[3-9]\d*\b",
        "",
        val,
        flags=re.IGNORECASE
    )

    val = _normalize_spaces(val)

    # ============================================================
    # POLA PENGULANGAN 3 DIGIT
    #
    # Contoh:
    # 098.4050.201.2023.000009+007+013+018+001
    #
    # ->
    # 098.4050.201.2023.000009.00
    # 098.4050.201.2023.000007.00
    # 098.4050.201.2023.000013.00
    # 098.4050.201.2023.000018.00
    # 098.4050.201.2023.000001.00
    # ============================================================

    m = re.fullmatch(
        r"(.+?\.)(\d{6})((?:\+\d{3})+)",
        val.replace(" ", "")
    )

    if m:

        prefix = m.group(1)
        first = m.group(2)

        suffixes = [
            s
            for s in m.group(3).split("+")
            if s
        ]

        results = [
            prefix + first + ".00"
        ]

        for s in suffixes:
            results.append(
                prefix + first[:-3] + s + ".00"
            )

        return _cap_or_join(results)

    # ============================================================
    # POLA "+"
    #
    # Biarkan apa adanya.
    # ============================================================

    # ============================================================
    # HAPUS SEMUA TOKEN YANG MENGANDUNG HURUF
    # Contoh:
    # P1, P2, P3
    # S/D
    # VAR
    # TBA
    # END
    # CANCEL
    # REALISASI
    # Dll.
    # ============================================================

    val = re.sub(
        r"\b[^\s+]*[A-Za-z][^\s+]*\b",
        "",
        val
    )

    # Rapikan spasi dan tanda +
    val = re.sub(r"\s*\+\s*", " + ", val)
    val = re.sub(r"^\s*\+\s*|\s*\+\s*$", "", val)
    val = re.sub(r"(?:\+\s*){2,}", "+ ", val)

    val = _normalize_spaces(val)

    if "+" in val:
        return [_normalize_spaces(val)]

    # ============================================================
    # CLEANUP BIASA
    # ============================================================

    cleaned = _normalize_spaces(val)

    cleaned = cleaned.strip(" -/")

    if not cleaned:
        return []

    return [cleaned]


def clean_polis(val) -> list:
    """Wrapper: jalankan cleaning polis lalu hapus semua titik dari hasilnya."""
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
    

    # ============================================================
    # 1. HAPUS BARIS YANG HANYA BERISI KETERANGAN
    #
    # Contoh:
    # P1 + P2
    # P1 CANCEL
    # NOVEMBER 2024
    # JANUARI 2024 - USD
    # ============================================================

    if re.fullmatch(
        r"(P\d+(?:\s*\+\s*P\d+)*(\s+CANCEL)?|"
        r"(JANUARI|FEBRUARI|MARET|APRIL|MEI|JUNI|JULI|AGUSTUS|SEPTEMBER|OKTOBER|NOVEMBER|DESEMBER)"
        r"\s+\d{4}(?:\s*-\s*(IDR|USD|EUR|GBP|SGD|JPY|AUD|CNY))?)",
        val,
        flags=re.IGNORECASE,
    ):
        return []

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
    # 2A. HAPUS +VAR
    #
    # Contoh:
    # 000.6005.201.2024.003281.00+VAR
    # ->
    # 000.6005.201.2024.003281.00
    # ============================================================

    val = re.sub(
        r"\+VAR\b",
        "",
        val,
        flags=re.IGNORECASE
    )

    val = _normalize_spaces(val)

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
    
def clean_slip(val) -> list:
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
    
def _is_valid_clstd(val) -> bool:
    if pd.isna(val):
        return False

    val = str(val).strip().upper()

    if not val:
        return False

    # Nilai CLSDT yang memang tidak bisa dipakai
    if val in {"TBA", "VAR", "VARIOUS"}:
        return False

    # CLSDT hanya berisi keterangan P1/P2/... CANCEL
    # dianggap tidak jelas → fallback ke ORI
    if re.fullmatch(
        r"(?:P\d+\s*CANCEL\s*)+",
        val
    ):
        return False

    return True
def process_data(input_file: str, output_file: str) -> None:
    print(f"[1/5] Membaca data dari: {input_file} ...")
    df = pd.read_excel(input_file, header=0)
    print(f"      Total baris keseluruhan: {len(df):,}")

    if CEDANT_COL not in df.columns:
        print(f"\n[ERROR] Kolom '{CEDANT_COL}' tidak ditemukan!")
        print(f"        Kolom tersedia: {list(df.columns)}")
        return

    for col in [POLIS_COL, SLIP_COL, INSURED_COL]:
        if col not in df.columns:
            print(f"\n[ERROR] Kolom '{col}' tidak ditemukan di file!")
            print(f"        Kolom tersedia: {list(df.columns)}")
            return

    df[CEDANT_COL] = df[CEDANT_COL].astype(str).str.strip()
    is_wahana = df[CEDANT_COL] == CEDANT_VALUE
    print(f"[2/5] Filter cedant '{CEDANT_VALUE}': {is_wahana.sum():,} baris WAHANA dari total {len(df):,} baris.")

    df.rename(columns={POLIS_COL: "polis_ori", SLIP_COL: "slip_ori", INSURED_COL: "insured_ori"},
              inplace=True)

    print("[3/5] Menjalankan proses cleaning hanya untuk baris WAHANA ...")

    all_clean_polis, all_clean_slip, all_clean_ins = [], [], []
    max_polis = max_slip = max_ins = 1

    for idx, (_, row) in enumerate(df.iterrows(), 1):
        if idx % 50_000 == 0:
            print(f"      Progress: {idx:,} / {len(df):,} baris diproses...")

        if is_wahana.iloc[idx - 1]:

            polis_ori = row.get("polis_ori", "")
            clstd_polis = row.get("CLSDT_POLICY_NO", "")

            # ====================================================
            # TENTUKAN SUMBER POLIS CLEAN
            # ====================================================

            if _is_valid_clstd(clstd_polis):
                polis_source = clstd_polis
            else:
                polis_source = polis_ori

            c_polis = clean_polis(polis_source)

            # ====================================================
            # TENTUKAN SUMBER SLIP CLEAN
            # ====================================================

            slip_ori = row.get("slip_ori", "")
            clstd_slip = row.get("CLSDT_SLIP_NO", "")

            if _is_valid_clstd(clstd_slip):
                slip_source = clstd_slip
            else:
                slip_source = slip_ori

            c_slip = clean_slip(slip_source)

            # ====================================================
            # CLEAN INSURED
            # ====================================================

            c_ins = clean_insured(row.get("insured_ori", ""))
        else:
            c_polis, c_slip, c_ins = [], [], []

        max_polis = max(max_polis, len(c_polis))
        max_slip  = max(max_slip,  len(c_slip))
        max_ins   = max(max_ins,   len(c_ins))

        all_clean_polis.append(c_polis)
        all_clean_slip.append(c_slip)
        all_clean_ins.append(c_ins)

    print(f"      Selesai diproses!")
    print(f"      -> Jumlah kolom clean polis  : {max_polis}")
    print(f"      -> Jumlah kolom clean slip   : {max_slip}")
    print(f"      -> Jumlah kolom clean insured: {max_ins}")

    print("[4/5] Menyusun kolom output ...")

    new_columns = []
    for col in df.columns:
        new_columns.append(col)
        if col == "polis_ori":
            new_columns += _insert_clean_columns(df, all_clean_polis, "polis",   max_polis)
        elif col == "slip_ori":
            new_columns += _insert_clean_columns(df, all_clean_slip,  "slip",    max_slip)
        elif col == "insured_ori":
            new_columns += _insert_clean_columns(df, all_clean_ins,   "insured", max_ins)

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
# ENTRY POINT
# ─────────────────────────────────────────────────────────────────────────────
# ─────────────────────────────────────────────────────────────────────────────
# ENTRY POINT - TEST RULE
# ─────────────────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    print("MASUK MAIN")
    process_data(INPUT_FILE, OUTPUT_FILE)