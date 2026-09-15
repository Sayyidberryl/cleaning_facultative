"""
Cleaning Data 3 - Suspend | Cedant: KB INS (KOOKMIN BEST INS. EX. LIG INS.)

CATATAN VERSI: v2 -- rewrite setelah BUG_KB_INS.md (bug report pertama dari
hasil pengerjaan v1). Semua perubahan di bawah ini dipicu oleh contoh
konkret di bug report tsb, bukan tebakan baru.

CATATAN ARSITEKTUR:
- Base polis/slip KB INS pada umumnya 15-17 digit angka (mayoritas 16
  digit), tergantung COB (Kebakaran/IAR/EQ/Marine Cargo/MBI/CAR) --
  sesuai sheet "KB Ins" di Karakteristik_Polis_dan_Slip.xlsx.
- POLIS & SLIP: keduanya membuang total ekor "-NNNNNN" (sertifikat) atau
  "-N/N" (fraksi/instalment) -- BEDA dari cedant lain (mis. JRP) yang
  cuma memproses kolom POLIS. Versi bersih ini dipakai untuk MATCHING;
  versi ORI tetap disimpan apa adanya.
- [BARU v2] Nilai yang diikuti narasi penjelasan (mis. "23910000 - masuk
  suspend karena masuk produksi Juni 2025") -- narasi setelah " - "
  DIBUANG, base tetap dipakai. Dideteksi lewat heuristik: teks setelah
  " - " dianggap narasi (bukan kode) kalau berisi >=3 kata dan didominasi
  huruf kecil.
- [BARU v2] Nilai "garbled"/format asing (mis. "P0B3-8006F3-2403A-7740",
  yang isinya campuran huruf+angka dalam satu token, BUKAN kode dikenal
  spt TBA/VAR/P<n>/S/D) -- DIBIARKAN APA ADANYA, tidak diproses sama
  sekali (supaya "-7740" di ujungnya tidak salah dianggap ekor
  sertifikat/fraksi lalu ikut terpotong).
- CERTIFICATE: diambil dari POLIS_ORI dulu, fallback ke SLIP_NO_ORI kalau
  POLIS_ORI tidak mengandung pola sertifikat. [BARU v2] build_certificate()
  sekarang mendukung format CAMPURAN (list diskrit + range dalam satu
  nilai, mis. "001 s/d 009, 011" -> "000001 SD 000009, 000011") dan
  varian pemisah "S.D." (dengan titik, bukan cuma "S/D"/"SD"/"S"). Setiap
  potongan yang dipisah koma diformat SENDIRI-SENDIRI (bukan lagi
  dikompres global kalau >3 potongan) -- lihat build_certificate().
- Nomor credit note berpola "NNNNNN/CN/NN/NN/NN" dibiarkan apa adanya.
- Insured -- SEMUA rule di bawah dikonfirmasi lewat BUG_KB_INS.md (contoh
  yang sama juga berlaku di Data 1 & Data 2, jadi konsisten dipakai di
  ketiga file):
    * "&" TIDAK LAGI jadi pemisah antar-insured (supaya "L&B", "E&C" tidak
      pecah) -- pemisah resmi: QQ, AND/OR(.), "/", ",", en-dash "-",
      " - " (spasi-dash-spasi), "+".
    * Nama lokasi/cabang tunggal yang nempel lewat pemisah (mis.
      "...,GROBOGAN", ".../BREBES", ".../PEMALANG FACTORY",
      ".../FACTORY II") DIGABUNG BALIK ke entitas sebelumnya, TIDAK
      dipecah jadi insured terpisah.
    * "PREFIX FACTORY <n1> & <n2>" DIPECAH per cabang ("PREFIX FACTORY
      <n1>", "PREFIX FACTORY <n2>").
    * Isi dalam kurung yang berupa singkatan/alias dari teks sebelumnya
      (mengandung "PT"/"CV", atau cocok pola inisial/awalan kata dari
      teks sebelumnya) DIBUANG; isi kurung yang berupa identitas cabang/
      pabrik/pihak berkepentingan (tidak match sbg singkatan) tetap
      DIPERTAHANKAN.
    * Frasa boilerplate "AS OWNER/PRINCIPAL/..." dibuang sebelum split.
    * Entitas hasil split yang ternyata singkatan dari entitas lain yang
      sudah diambil (baik lewat kamus dikenal spt "SEIN", pola inisial,
      pola awalan kata, atau pola "X&Y" pendek dgn kata pertama sama
      dgn entitas sebelumnya spt "LOTTE E&C" vs "LOTTE ENGINEERING")
      DIBUANG, tidak dihitung sbg insured terpisah.
- Baris berstatus "SUSPENSE" saja yang diproses; baris "ADJUSTED" dibuang
  total sebelum cleaning.

=====================================================================
CATATAN YANG MASIH PERLU DIKONFIRMASI:
1. Threshold "narasi" pada penghapusan teks setelah " - " (heuristik:
   >=3 kata & didominasi huruf kecil) adalah perkiraan dari SATU contoh
   ("23910000 - masuk suspend..."). Kalau ada pola narasi lain yang lolos
   atau kode asli yang malah salah terpotong, mohon dikabari.
2. Daftar lokasi yang dikenali untuk merge-balik (_KNOWN_LOCATION_WORDS:
   BREBES, GROBOGAN, PEMALANG, PASURUAN, JOMBANG) hanya mencakup yang
   muncul di bug report. Nama kota/kabupaten lain yang belum ada di
   daftar ini AKAN tetap dipecah jadi insured terpisah -- perlu
   ditambahkan manual kalau ketemu kasus baru.
=====================================================================
"""

import os
import re

import pandas as pd


# ─────────────────────────────────────────────────────────────────────────────
# KONFIGURASI
# ─────────────────────────────────────────────────────────────────────────────

INPUT_FILE  = os.path.join("input", "3b. Database Suspense 150826.xlsx")
INPUT_SHEET = "Sheet1"
HEADER_ROW  = 2

OUTPUT_FILE = os.path.join("output", "KbIns_output_suspend.xlsx")

CEDANT_FILTER_COL   = "CEDANT SHRT NAME"
CEDANT_FILTER_VALUE = "KB INS"

STATUS_COL          = "STATUS"
STATUS_KEEP_VALUE   = "SUSPENSE"

INSURED_COL = "INSURED"
POLIS_COL   = "POLIS"
SLIP_COL    = "SLIP NO"

MAX_SPLIT_COLS = 5


# ─────────────────────────────────────────────────────────────────────────────
# KAMUS POLA STANDAR KB INS
# ─────────────────────────────────────────────────────────────────────────────

KNOWN_POLIS_PATTERNS = [re.compile(r"^\d{13,19}$")]
KNOWN_SLIP_PATTERNS  = [re.compile(r"^\d{13,19}$")]


def _matches_known_pattern(value: str, patterns) -> bool:
    return any(p.match(value) for p in patterns)


def refine_with_known_pattern(value: str, patterns) -> str:
    if not value:
        return value
    if _matches_known_pattern(value, patterns):
        return value
    candidates = [re.sub(r"\s+", "", value), value.strip(" .-/")]
    for c in candidates:
        if _matches_known_pattern(c, patterns):
            return c
    return value


def _normalize_spaces(text: str) -> str:
    return re.sub(r"\s{2,}", " ", text).strip()


# ─────────────────────────────────────────────────────────────────────────────
# KOLOM CERTIFICATE [DIROMBAK v2] -- mendukung list campuran (diskrit + range)
# dan varian pemisah "S.D." dengan titik.
# ─────────────────────────────────────────────────────────────────────────────

_CERT_PAIR_SEP_RE = re.compile(
    r"\s*(?:S\s*\.\s*D\s*\.?|S\s*/\s*D|SD|S(?!\d)|-)\s*", re.IGNORECASE
)


def _format_cert_chunk(chunk: str):
    """Format SATU potongan (hasil split koma level-atas). Balikan None
    kalau potongan tidak masuk akal (bukan angka murni / bukan pasangan
    range angka murni)."""
    chunk = chunk.strip()
    if not chunk:
        return None
    parts = [p for p in _CERT_PAIR_SEP_RE.split(chunk) if p.strip()]
    if len(parts) == 2 and all(re.match(r"^\d{1,7}$", p.strip()) for p in parts):
        lo, hi = int(parts[0]), int(parts[1])
        if hi < lo:
            lo, hi = hi, lo
        count = hi - lo + 1
        if count > 3:
            return f"{str(lo).zfill(6)} SD {str(hi).zfill(6)}"
        return ", ".join(str(n).zfill(6) for n in range(lo, hi + 1))
    if len(parts) == 1 and re.match(r"^\d{1,7}$", parts[0].strip()):
        return str(int(parts[0])).zfill(6)
    return None


def build_certificate(cert_raw) -> str:
    """Format nomor sertifikat/register KB INS:
    - Dipecah dulu berdasarkan koma level-atas -> tiap potongan diformat
      SENDIRI-SENDIRI (bukan dikompres global berdasar total jumlah item).
    - Potongan berupa pasangan range 2-angka (dipisah "-"/"S/D"/"SD"/"S"/
      "S.D."): >3 anggota -> notasi "AWAL SD AKHIR"; <=3 anggota ->
      breakdown penuh dipisah koma.
    - Potongan berupa angka tunggal -> tetap apa adanya (di-pad 6 digit).
    - Semua angka di-pad ke 6 digit; semua varian S/D distandarisasi "SD".
    - Kalau ADA SATU SAJA potongan yang tidak masuk akal -> seluruh
      CERTIFICATE dikosongkan (supaya tidak menampilkan hasil parsial
      yang salah).
    """
    if cert_raw is None:
        return ""
    text = str(cert_raw).strip()
    if not text or text.upper() in {"-", "NAN", "NONE", "NULL"}:
        return ""
    top_chunks = [c.strip() for c in text.split(",") if c.strip()]
    if not top_chunks:
        return ""
    formatted = []
    for c in top_chunks:
        f = _format_cert_chunk(c)
        if f is None:
            return ""
        formatted.append(f)
    return ", ".join(formatted)


# ─────────────────────────────────────────────────────────────────────────────
# DETEKSI BLOK SERTIFIKAT/REGISTER YANG NEMPEL DI POLIS/SLIP
# ─────────────────────────────────────────────────────────────────────────────

_CERT_BLOCK_SLASH_RE = re.compile(r"^\s*(\d{10,})\s*/\s*(.+)$")
_CERT_BLOCK_SPACE_RE = re.compile(r"^\s*(\d{10,})\s+-?\s*(.+)$")
_CERT_BLOCK_TIGHT_DASH_RE = re.compile(r"^\s*(\d{10,})\s*[-.]\s*(\d{5,7})\s*$")


def try_extract_cert_block(primary: str):
    """Coba deteksi pola 'BASIS + pemisah + blok sertifikat/register'.
    Balikan (basis, certificate_terformat) kalau cocok, else (None, None).
    """
    for rx in (_CERT_BLOCK_SLASH_RE, _CERT_BLOCK_SPACE_RE):
        m = rx.match(primary)
        if m:
            base, rest = m.groups()
            rest = rest.strip()
            if rest and not re.match(r"^\d{10,}", rest):
                cert = build_certificate(rest)
                if cert:
                    return base, cert
    m = _CERT_BLOCK_TIGHT_DASH_RE.match(primary)
    if m:
        base, rest = m.groups()
        cert = build_certificate(rest)
        if cert:
            return base, cert
    return None, None


def extract_certificate(polis_raw, slip_raw) -> str:
    """Ambil sertifikat dari POLIS_ORI dulu; fallback ke SLIP_NO_ORI."""
    for raw in (polis_raw, slip_raw):
        if pd.isna(raw):
            continue
        val = str(raw).strip()
        if not val:
            continue
        _, cert = try_extract_cert_block(val)
        if cert:
            return cert
    return ""


# ─────────────────────────────────────────────────────────────────────────────
# CLEANING POLIS & SLIP
# ─────────────────────────────────────────────────────────────────────────────

_POLIS_KEEP_AS_IS_RE = re.compile(r"^\s*P\d+\s*/", re.IGNORECASE)
_CN_PATTERN_RE = re.compile(r"^\s*\d{3,6}\s*/\s*CN\s*/\s*\d{1,4}\s*/\s*\d{2}\s*/\s*\d{2}\s*$", re.IGNORECASE)

# [BARU v2] narasi penjelasan setelah " - " (mis. "23910000 - masuk suspend
# karena masuk produksi Juni 2025") -- dibuang, basis tetap dipakai.
_NARRATIVE_SPLIT_RE = re.compile(r"^(.*?)\s+-\s+(.+)$")


def _strip_narrative_tail(val: str) -> str:
    m = _NARRATIVE_SPLIT_RE.match(val)
    if not m:
        return val
    head, tail = m.groups()
    if re.search(r"[a-z]{3,}", tail) and len(tail.split()) >= 3:
        return head.strip()
    return val


# [BARU v2] nilai "garbled"/format asing (huruf+angka campur dalam satu
# token, bukan kode dikenal) -- dibiarkan apa adanya, tidak diproses.
_GARBLED_TOKEN_RE = re.compile(r"^(?=.*[A-Za-z])(?=.*\d).+$")
_JUNK_TOKEN_RE = re.compile(r"^(?:P\d+|VAR|VARIOUS|TBA)$", re.IGNORECASE)


def _looks_garbled(val: str) -> bool:
    tokens = re.split(r"[\s+\-/,]", val)
    for t in tokens:
        t = t.strip()
        if not t:
            continue
        if _JUNK_TOKEN_RE.match(t):
            continue
        if re.fullmatch(r"S/?D\.?", t, re.IGNORECASE):
            continue
        if _GARBLED_TOKEN_RE.match(t):
            return True
    return False


def _strip_tail(val: str) -> str:
    """Buang total ekor '-NNNNNN' (sertifikat) atau '-N/N' (fraksi /
    instalment) secara berulang sampai tidak ada lagi yang cocok."""
    p = val
    while True:
        stripped = re.sub(r"-\d+(?:/\d+)?$", "", p)
        if stripped == p:
            break
        p = stripped
    return p if p else val


def _clean_polis_or_slip(val) -> str:
    if pd.isna(val):
        return ""
    val = str(val).strip()
    if not val:
        return ""
    if _CN_PATTERN_RE.match(val) or _POLIS_KEEP_AS_IS_RE.match(val):
        return val

    val = _strip_narrative_tail(val)
    if _looks_garbled(val):
        return val

    base, cert = try_extract_cert_block(val)
    if base:
        return base

    return _strip_tail(val)


def clean_polis(val) -> str:
    return _clean_polis_or_slip(val)


def clean_slip(val) -> str:
    return _clean_polis_or_slip(val)


def _remove_polis_slip_from_text(text: str, polis_ori, slip_ori) -> str:
    if pd.notna(polis_ori):
        token = str(polis_ori).strip()
        if token and token != "-":
            text = text.replace(token, "")
    if pd.notna(slip_ori):
        token = str(slip_ori).strip()
        if token and token != "-":
            text = text.replace(token, "")
    return text


# ─────────────────────────────────────────────────────────────────────────────
# CLEAN INSURED [DIROMBAK v2] -- lihat CATATAN ARSITEKTUR di atas
# ─────────────────────────────────────────────────────────────────────────────

_GELAR_PREFIX_RE = re.compile(
    r"^\s*(?:BAPAK|IBU|SDRI?|NYONYA|NY|TUAN|TN|MR|MRS|MS|DR|IR|PROF|HJ|H)\b\.?\s*",
    re.IGNORECASE,
)
_GELAR_SUFFIX_RE = re.compile(
    r"\s*,?\s*\b(?:S\.?E|S\.?H|S\.?SI|S\.?T|S\.?SOS|S\.?KOM|S\.?PD|"
    r"M\.?SI|M\.?M|M\.?H|M\.?KOM|SH|SE)\b\.?\s*$",
    re.IGNORECASE,
)
_SUSPENSE_KEEP_RE = re.compile(r"SUSPEN|\(?H\)?\s*UTANG\s*PIUTANG", re.IGNORECASE)

# pemisah: QQ, AND/OR(.), "/", ",", en-dash "-", " - " (spasi-dash-spasi), "+".
# "&" SENGAJA TIDAK termasuk (supaya "L&B", "E&C" tidak ikut pecah).
_SPLIT_RE = re.compile(r"\bQQ\b|\bAND\s*/\s*OR\.?\b|/{1,2}|,|–|\s-\s|\+|¿", re.IGNORECASE)

INSURED_JUNK_WORDS = frozenset({
    "PT", "CV", "TBK", "PERSERO", "LTD", "INC", "LLC",
    "AND", "OR", "THE", "OF", "AS", "VARIOUS", "VAR", "QQ", "TBA",
})

# lokasi/kabupaten yang harus TETAP MENYATU dgn entitas sebelumnya
_KNOWN_LOCATION_WORDS = {"BREBES", "GROBOGAN", "PEMALANG", "PASURUAN", "JOMBANG"}
_BRANCH_MERGE_RE = re.compile(
    r"^(?:FACTORY|PABRIK|CABANG)\s+\S+$|^\S+\s+(?:FACTORY|PABRIK|CABANG)$",
    re.IGNORECASE,
)


def _is_location_merge_part(p: str) -> bool:
    up = p.strip().upper()
    if up in _KNOWN_LOCATION_WORDS:
        return True
    if _BRANCH_MERGE_RE.match(up):
        return True
    return False


# "PREFIX FACTORY <n1> & <n2>" -> pecah per cabang
_FACTORY_MULTI_RE = re.compile(
    r"^(.*?)\bFACTORY\s+([IVXLCDM0-9]+(?:\s*(?:,|&|DAN)\s*[IVXLCDM0-9]+)+)\s*$",
    re.IGNORECASE,
)


def _expand_factory_multi(val: str):
    m = _FACTORY_MULTI_RE.match(val.strip())
    if not m:
        return None
    prefix, nums_part = m.groups()
    prefix = prefix.strip()
    nums = [n.strip() for n in re.split(r"\s*(?:,|&|DAN)\s*", nums_part, flags=re.IGNORECASE) if n.strip()]
    if len(nums) < 2:
        return None
    return [f"{prefix} FACTORY {n}".strip() for n in nums]


# "PREFIX LOC1 DAN LOC2" -> pecah per cabang kota (lihat CATATAN PERLU
# DIKONFIRMASI -- hanya menangani daftar lokasi yg sudah dikenal)
_DAN_LOCATION_RE = re.compile(
    r"^(.*?)\s+(" + "|".join(_KNOWN_LOCATION_WORDS) + r")\s+DAN\s+("
    + "|".join(_KNOWN_LOCATION_WORDS) + r")\s*$",
    re.IGNORECASE,
)


def _expand_dan_location(val: str):
    m = _DAN_LOCATION_RE.match(val.strip())
    if not m:
        return None
    prefix, loc1, loc2 = m.groups()
    prefix = prefix.strip()
    if not prefix:
        return None
    return [f"{prefix} {loc1}".strip(), f"{prefix} {loc2}".strip()]


_AS_ROLE_RE = re.compile(
    r"\bAS\s+(?:THE\s+)?(?:OWNER|PRINCIPAL|INSURED|CONTRACTOR|OFF-TAKER)\b\.?",
    re.IGNORECASE,
)


def _initials_windows(name: str) -> set:
    words = [w for w in re.split(r"[\s\-]+", name.upper()) if w]
    variants = set()
    n = len(words)
    for start in range(n):
        for end in range(start + 1, n + 1):
            window = words[start:end]
            variants.add("".join(w[0] for w in window if w[0].isalpha()))
    return variants


def _is_abbreviation_of(short: str, long_name: str) -> bool:
    short_clean = re.sub(r"[^A-Z0-9]", "", short.upper())
    if not short_clean or len(short_clean) < 2 or len(short_clean) > 8:
        return False
    if short_clean in _initials_windows(long_name):
        return True
    words = [w for w in re.split(r"[\s\-]+", long_name.upper()) if w]
    if any(w.startswith(short_clean) and w != short_clean for w in words):
        return True
    return False


def _strip_parenthetical_abbrev(text: str) -> str:
    def repl(m):
        inner = m.group(1).strip()
        if re.search(r"\bPT\b|\bCV\b", inner, re.IGNORECASE):
            return " "
        inner_letters = re.sub(r"[^A-Za-z]", "", inner)
        before = text[:m.start()].strip()
        if before and inner_letters and _is_abbreviation_of(inner_letters, before):
            return " "
        return m.group(0)
    return re.sub(r"\(([^()]*)\)", repl, text)


_SHORT_AMP_RE = re.compile(r"\b[A-Z]{1,2}&[A-Z]{1,2}\b")

_KNOWN_INSURED_ABBREV = {
    "SEIN": "SAMSUNG ELECTRONICS INDONESIA",
}

_PT_CV_PREFIX_RE = re.compile(r"^\s*(?:PT|CV)(?:\.|(?=\s|$))\.?\s*", re.IGNORECASE)
_PT_CV_SUFFIX_RE = re.compile(r",?\s*(?:PT|CV)\.?\s*$", re.IGNORECASE)


def _clean_insured_name(name: str) -> str:
    name = _normalize_spaces(name)
    name = _GELAR_PREFIX_RE.sub("", name)
    for _ in range(3):
        stripped = _GELAR_SUFFIX_RE.sub("", name).strip()
        if stripped == name:
            break
        name = stripped
    name = re.sub(r",?\s*\(\s*PERSERO\s*\)\s*", " ", name, flags=re.IGNORECASE)
    name = re.sub(r"\bTBK\b\s*,\s*PT\.?\s*", " ", name, flags=re.IGNORECASE)
    name = re.sub(r",?\s*\bTBK\b\s*", " ", name, flags=re.IGNORECASE)
    name = re.sub(r"\bLTD\b\.?\s*$", "", name, flags=re.IGNORECASE)
    name = re.sub(r"\bPTE\b\.?\s*", "", name, flags=re.IGNORECASE)
    name = _PT_CV_SUFFIX_RE.sub("", name)
    name = _PT_CV_PREFIX_RE.sub("", name)
    name = re.sub(r"(?<!\d)\.(?!\d)", " ", name)
    return _normalize_spaces(name).strip().upper()


def _dedup_abbreviations(cleaned: list) -> list:
    final = []
    for name in cleaned:
        words = name.split()
        first_word = words[0] if words else ""
        is_dup = False
        for prior in final:
            if name in _KNOWN_INSURED_ABBREV and _KNOWN_INSURED_ABBREV[name] in prior:
                is_dup = True
                break
            if _is_abbreviation_of(name, prior):
                is_dup = True
                break
            if _SHORT_AMP_RE.search(name) and prior.split() and prior.split()[0] == first_word:
                is_dup = True
                break
        if not is_dup:
            final.append(name)
    return final


def clean_insured(val, polis_ori=None, slip_ori=None) -> list:
    if pd.isna(val):
        return []
    val = str(val).strip()
    if not val:
        return []
    if _SUSPENSE_KEEP_RE.search(val):
        return [_normalize_spaces(val).upper()]

    val = _remove_polis_slip_from_text(val, polis_ori, slip_ori)
    val = _strip_parenthetical_abbrev(val)
    val = _AS_ROLE_RE.sub(" ", val)

    for fn in (_expand_factory_multi, _expand_dan_location):
        r = fn(val)
        if r is not None:
            return [_normalize_spaces(x).upper() for x in r]

    parts = [p.strip() for p in _SPLIT_RE.split(val) if p.strip()]
    cleaned = []
    for p in parts:
        p_clean = _clean_insured_name(p)
        if not p_clean:
            continue
        merge_check = p_clean.strip("() ")
        if _is_location_merge_part(merge_check) and cleaned:
            cleaned[-1] = _normalize_spaces(cleaned[-1] + " " + p_clean)
            continue
        if len(p_clean) < 2 or p_clean in INSURED_JUNK_WORDS:
            continue
        cleaned.append(p_clean)

    if not cleaned:
        fallback = _clean_insured_name(val)
        return [fallback] if fallback else []

    cleaned = _dedup_abbreviations(cleaned)

    if len(cleaned) > MAX_SPLIT_COLS:
        return [", ".join(cleaned)]
    return cleaned


# ─────────────────────────────────────────────────────────────────────────────
# PROSES UTAMA
# ─────────────────────────────────────────────────────────────────────────────

def _force_text_format(output_file, target_cols_names):
    import openpyxl
    wb = openpyxl.load_workbook(output_file)
    ws = wb.active
    header = [c.value for c in ws[1]]
    target_idx = [i for i, h in enumerate(header, start=1) if h in target_cols_names]
    for col_idx in target_idx:
        for row in ws.iter_rows(min_row=2, min_col=col_idx, max_col=col_idx):
            for cell in row:
                cell.number_format = "@"
    wb.save(output_file)


def process_data(input_file: str, output_file: str) -> None:
    print(f"[1/5] Membaca data dari: {input_file} (sheet: {INPUT_SHEET}) ...")
    df = pd.read_excel(input_file, sheet_name=INPUT_SHEET, header=HEADER_ROW)
    print(f"      Total baris keseluruhan: {len(df):,}")

    if CEDANT_FILTER_COL not in df.columns:
        print(f"\n[ERROR] Kolom '{CEDANT_FILTER_COL}' tidak ditemukan!")
        print(f"        Kolom tersedia: {list(df.columns)}")
        return

    df[CEDANT_FILTER_COL] = df[CEDANT_FILTER_COL].astype(str).str.strip()
    df = df[df[CEDANT_FILTER_COL] == CEDANT_FILTER_VALUE].copy()
    print(f"[2/5] Filter cedant '{CEDANT_FILTER_VALUE}': {len(df):,} baris ditemukan.")

    if df.empty:
        print("\n[WARN] Tidak ada data setelah filter. Proses dihentikan.")
        return

    if STATUS_COL not in df.columns:
        print(f"\n[ERROR] Kolom '{STATUS_COL}' tidak ditemukan!")
        print(f"        Kolom tersedia: {list(df.columns)}")
        return

    before_status = len(df)
    df[STATUS_COL] = df[STATUS_COL].astype(str).str.strip()
    df = df[df[STATUS_COL] == STATUS_KEEP_VALUE].copy()
    print(f"[2b/5] Filter status '{STATUS_KEEP_VALUE}': {len(df):,} baris "
          f"(dibuang {before_status - len(df):,} baris ADJUSTED).")

    if df.empty:
        print("\n[WARN] Tidak ada data setelah filter status. Proses dihentikan.")
        return

    for col in [INSURED_COL, POLIS_COL, SLIP_COL]:
        if col not in df.columns:
            print(f"\n[ERROR] Kolom '{col}' tidak ditemukan di file!")
            print(f"        Kolom tersedia: {list(df.columns)}")
            return

    df.rename(columns={
        INSURED_COL: "INSURED_ORI",
        POLIS_COL:   "POLIS_ORI",
        SLIP_COL:    "SLIP_NO_ORI",
    }, inplace=True)

    print("[3/5] Menjalankan proses cleaning ...")

    polis_cln, slip_cln, all_certificate, all_insured_parts = [], [], [], []
    max_insured_parts = 1

    for _, row in df.iterrows():
        p_ori = row.get("POLIS_ORI", "")
        s_ori = row.get("SLIP_NO_ORI", "")
        i_ori = row.get("INSURED_ORI", "")

        polis_cln.append(refine_with_known_pattern(clean_polis(p_ori), KNOWN_POLIS_PATTERNS))
        slip_cln.append(refine_with_known_pattern(clean_slip(s_ori), KNOWN_SLIP_PATTERNS))
        all_certificate.append(extract_certificate(p_ori, s_ori))

        parts = clean_insured(i_ori, p_ori, s_ori)
        max_insured_parts = max(max_insured_parts, len(parts))
        all_insured_parts.append(parts)

    print("[4/5] Menyusun kolom output ...")

    df["POLIS_CLEAN_1"] = polis_cln
    df["CERTIFICATE_1"] = all_certificate
    df["SLIP_CLEAN_1"]  = slip_cln

    insured_cols = []
    for i in range(1, max_insured_parts + 1):
        col_name = f"INSURED_{i}"
        df[col_name] = [p[i - 1] if i - 1 < len(p) else None for p in all_insured_parts]
        insured_cols.append(col_name)

    new_columns = []
    for col in df.columns:
        if col in ("POLIS_CLEAN_1", "CERTIFICATE_1", "SLIP_CLEAN_1") or col in insured_cols:
            continue
        new_columns.append(col)
        if col == "INSURED_ORI":
            new_columns += insured_cols
        elif col == "POLIS_ORI":
            new_columns.append("POLIS_CLEAN_1")
            new_columns.append("CERTIFICATE_1")
        elif col == "SLIP_NO_ORI":
            new_columns.append("SLIP_CLEAN_1")

    df = df[new_columns]

    os.makedirs(os.path.dirname(output_file), exist_ok=True)
    print(f"[5/5] Menyimpan hasil ke: {output_file} ...")
    df.to_excel(output_file, index=False)
    _force_text_format(output_file, target_cols_names={"POLIS_ORI", "POLIS_CLEAN_1", "CERTIFICATE_1", "SLIP_NO_ORI", "SLIP_CLEAN_1"})

    print(f"\n{'=' * 55}")
    print("  [OK] Selesai!")
    print(f"  Total baris output  : {len(df):,}")
    print(f"  Total kolom output  : {len(df.columns)}")
    print(f"  File disimpan di    : {output_file}")
    print(f"{'=' * 55}")


if __name__ == "__main__":
    process_data(INPUT_FILE, OUTPUT_FILE)