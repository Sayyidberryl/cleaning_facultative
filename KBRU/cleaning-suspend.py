"""
Cleaning Data 3 - Suspend | Cedant: KBRU (PT. KBRU REINSURANCE BROKERS
(EX. SIMAS))

CATATAN VERSI: v5 -- rewrite setelah bug report ke-4 (15/09/26, file:
KBRU_BUG_150926.md). Perubahan dari v4:

  POLIS/SLIP:
  1. [BARU] Narasi/deskripsi MURNI yang nyasar di kolom polis/slip --
     (a) mengandung SUSPEN/UTANG PIUTANG (mis. "PENYELESAIAN HUTANG
     PIUTANG KBRU 2021 USD", "SIMASRB - PENYELESAIAN SUSPENSE 2018 -
     2019 VARIOUS CEDING"), atau (b) narasi murni BORDERO/SUMMARY SIMAS
     RE/nama bulan/tahun tanpa angka lain (mis. "BORDERO NOVEMBER 2017",
     "MEI - JUNI 2018") -> DIBIARKAN APA ADANYA, spasi TIDAK dihapus.
  2. [BARU] "TBA/" (placeholder + garis miring doang) -> garis miringnya
     dibuang, hasil cukup "TBA".
  3. [FIX] "No." SEKARANG dibuang di MANA PUN posisinya (sebelumnya cuma
     di AWAL value) -- mis. "0061-00/SFGI/...-0001/No.00087/S/06/2018/AB"
     -> "006100SFGI0006291SRB062018000100087S062018AB".
  4. [FIX] Nilai yang HANYA "CN : <nomor>" (tanpa nilai lain) -> cuma
     label "CN :"-nya yg dibuang, NOMORNYA dipertahankan sbg hasil
     (sebelumnya seluruhnya ikut terbuang jadi kosong). Notasi CN yang
     BERSANDINGAN dgn nilai lain tetap dibuang total spt v3/v4.
  5. [FIX] Pemangkasan ekor "/segmen-tanpa-digit" hanya dipakai kalau
     value MEMANG punya rantai '+' -- supaya ekor sah spt "/AB", "/BB",
     "/AC" (kode 2 huruf) pada value tunggal tidak ikut kebuang.
  6. [BARU] Dash NYASAR di dalam pola F/S 2-segmen (mis.
     "11055129F00000-275") dilebur jadi "11055129F00000275" sebelum
     diproses.
  7. [FIX] Target substitusi 'pengulangan' saat ada BEBERAPA segmen digit
     berpanjang SAMA -> pilih yg pertama yang BUKAN nol semua (mis.
     "0000-5747": target "5747", bukan "0000").
  8. [BARU] Ambang 'berstruktur' utk substitusi diperluas: >=1 pemisah
     (./-) DAN total digit >=7 juga memenuhi (mis. "0000-5747"), selain
     3 syarat lama.
  9. [BARU] Token substitusi yang SANGAT pendek (<=2 digit, mis.
     "+97+07+43") -> pola pengulangan dianggap tidak jelas, seluruh
     nilai dibiarkan apa adanya.
 10. [BARU] Token tambahan tanpa huruf tapi ada "/" atau "." nempel
     (mis. "04820/0001189") -> diambil potongan PERTAMA saja sbg
     kandidat substitusi; sisanya dibuang (KBRU tidak punya kolom
     sertifikat).
 11. [BARU] Suffix "PT"/"CV" ikut dibuang dari hasil yg berupa nama
     perusahaan (mis. "TBA / ASURANSI FPG INSURANCE, PT" -> "ASURANSI
     FPG INSURANCE").

  INSURED:
 12. [BARU] "PELINDO" dikenali sbg SINONIM "PELABUHAN INDONESIA", dan
     penyebutan berulang dgn pemisah apa pun (mis. "PELABUHAN INDONESIA
     I QQ PELINDO I QQ PELINDO II QQ ...") dikumpulkan + dinormalisasi
     ke "PELABUHAN INDONESIA <arab>" + dedup.
 13. [FIX] Pola inisial huruf-tunggal-bertitik (mis. "S.A.R.I") -> titik
     dibuang TANPA jadi spasi ("SARI"), bukan "S A R I".
 14. [BARU] Kata kunci ekspansi ditambah UNIT & MALL (mis. "DELTA DUNIA
     TEKSTIL/UNIT IV/...", "GALAXY MALL I & II").
 15. [FIX] Ekspansi keyword+daftar sekarang juga dicoba PER BAGIAN hasil
     split (bukan cuma whole-value), supaya entitas lain di depan tidak
     ikut kebawa jadi prefix (mis. "SINAR GALAXY, PT / GALAXY MALL I &
     II" -> 3 insured, bukan 2 dgn prefix salah).
 16. [BARU] "QQ" yang nempel langsung ke huruf lain tanpa spasi (mis.
     "QQBALI", "INDAHQQ") -> tanda data rusak/ambigu, SELURUH value
     dibiarkan apa adanya. (Rantai Q murni spt "QQQ" TIDAK dihitung
     rusak -- itu sudah ditangani sbg varian pemisah QQ.)

CATATAN VERSI (v4) -- bug report ke-3 (14/09/26, file:
KBRU_BUG_140926.md). Perubahan dari v3:
  1. [FIX INSURED, semua data] "SINAR MAS"/"SINARMAS" dikonsistenkan jadi
     "SINARMAS" (tanpa spasi) -- berlaku di kolom INSURED maupun di
     hasil POLIS/SLIP yang kebetulan berisi nama perusahaan (lihat poin
     2 & 3).
  2. [FIX POLIS/SLIP] Hasil yang ternyata berupa TEKS/NAMA (tanpa digit
     sama sekali, muncul lewat mekanisme buang TBA/P<n> yg bersandingan,
     mis. "TBA/MERITZ KORINDO INSURANCE") SEBELUMNYA ikut kena "hapus
     karakter" penuh (spasi antar kata ikut hilang, jadi
     "MERITZKORINDOINSURANCE") -- SEKARANG utk hasil semacam ini spasi
     DIPERTAHANKAN (cuma dirapikan + SINAR MAS->SINARMAS), BUKAN
     dihapus semua.
  3. [BARU] Rantai P<n> yang dipisah SPASI (bukan "+", mis. "SINARMAS P2
     P3 P4 P5 P6 P7") sekarang ikut dibuang (sebelumnya cuma pola
     berpisah "+" yg dikenali). Rantai P<n> yg diselingi anotasi
     "(ADD n)" (mis. "P2 (ADD 31) + P3 (ADD 32) + ...") juga sekarang
     dikenali & dibuang bersama anotasinya.
  4. [BARU] Segmen depan (sebelum "/" pertama) berisi RANTAI angka
     pendek digabung "-" (bukan "+") sebanyak >=5 kali -> dianggap
     'pengulangan berlebih', DIBIARKAN APA ADANYA (sama spt aturan >5
     via "+" yg sudah ada) -- sebelumnya pola dash-chain ini tidak
     terdeteksi (cuma pola "+"-chain yg dicek).
  5. [FIX] Target substitusi 'pengulangan' utk pola PERSIS "ANGKA+1-2
     HURUF+ANGKA" (mis. "F00000473" pada "110551216F00000473", tanpa
     karakter lain) SEKARANG benar menyasar SEGMEN TERAKHIR (setelah
     huruf) -- sebelumnya salah menyasar segmen PERTAMA krn kebetulan
     lebih panjang (heuristik "segmen terpanjang" cocok utk pola FPAR-
     dot-style tapi TIDAK utk pola F/S 2-segmen ini).
  (Perubahan khusus Osbal terkait file ini ada di CATATAN ARSITEKTUR
  KBRU_cleaning_osbal.py poin seleksi sumber SLIP_CLEAN.)

CATATAN VERSI (v3) -- rewrite setelah bug report ke-2 (10/09/26, file:
filter_slip_clean_1_yang_masih_belu.md, Filter_clsdt_polis_yang_sesuai_
kara.md, filter_polis_clean_yang_sesuai_kara.md, KBRU_BUG.md). File ini
sendiri (Suspend) TIDAK kena bug seleksi sumber CLSDT/FAC (itu khusus
Osbal), TAPI mesin cleaning karakter POLIS/SLIP & INSURED yang dipakai
DIROMBAK CUKUP BESAR, jadi versi ini ikut naik ke v3 supaya konsisten
dgn 2 file lain. Perubahan dari v2:

  MESIN POLIS/SLIP:
  1. [FIX] "TBA"/"P<n>" yang BERSANDINGAN dgn konten lain (mis. "TBA/
     SINARMAS", "SINARMAS/TBA") -- TBA/P<n>-nya DIBUANG, sisanya (SINARMAS)
     yang dipakai sbg nilai (SEBELUMNYA di v2 seluruh nilai malah
     dikembalikan apa adanya/"abaikan" -- ini salah). TBA/P<n> yang
     BERDIRI SENDIRI (tanpa konten lain) tetap "abaikan" apa adanya.
  2. [FIX] Notasi "CN : <nomor>" sekarang DIBUANG TOTAL bersama nilainya
     (bukan lagi dijadikan nilai tambahan di kolom dinamis spt v2).
  3. [FIX] 'Pengulangan' digit sekarang bisa mendeteksi suffix
     berstruktur yang nempel di TOKEN MANAPUN dalam rantai '+' (v2 cuma
     mendukung suffix di token PERTAMA), mis. "23018+23019+23023+23022/
     FIAR/001.0000344-SRB/06/2022" -> 4 nilai, suffix "/FIAR/...2022"
     diterapkan ke semua basis angka (23018/23019/23023/23022).
  4. [BARU] Substitusi 'pengulangan' HANYA jalan kalau nilai utama cukup
     'berstruktur' (ada huruf, ATAU >=2 pemisah '.'//', ATAU total
     digitnya panjang >=10) -- angka pendek polos yg digabung '+' tanpa
     struktur (mis. "3522 + 3846", "0011477 + 0012959") TETAP dibiarkan
     apa adanya (ambigu, BUKAN disubstitusi jadi 2 nilai terpisah).
  5. [FIX] Extra token yg kena tempelan tanda baca (mis. "9877." dgn
     titik nyempil) sekarang benar dikenali sbg token digit polos,
     bukan salah dianggap "nilai lengkap" krn ada titiknya.

  INSURED: lihat CATATAN ARSITEKTUR di bawah utk daftar lengkap
  perubahan (pola PELABUHAN INDONESIA/CLUSTER/TOWER/REGIONAL/DISTRICT,
  suffix UD/PD/PT/CV/gelar, delimiter QQQ/OO/*, dst).

CATATAN ARSITEKTUR:
- KBRU adalah BROKER, nomor polis tidak selalu ada & formatnya
  tergantung cedant aslinya. TIDAK ADA penukaran isi/posisi kolom POLIS
  & SLIP.
- Mesin cleaning POLIS & SLIP (clean_value(), 1 fungsi utk KEDUA kolom):
    1. Placeholder murni BERDIRI SENDIRI (TBA, P<n>, VARIOUS/VAR,
       SELISIH, END, "-", "&"-joined 2 angka polos) -> ABAIKAN = teks
       ORI PERSIS.
    2. Pola khusus "NNNNNNNN[A-Z]NNNNNNNN-NN" (mis. "11055124F00000144-
       03") -> ABAIKAN juga (dash dipertahankan), di mana pun ditemukan.
    3. TBA/P<n> yang bersandingan dgn konten lain -> dibuang, sisanya
       yang diproses (baik itu angka asli MAUPUN nama tanpa digit, mis.
       "SINARMAS").
    4. Notasi "CN : <nomor>" -> dibuang total.
    5. "Pengulangan" digit: token digit pendek (di rantai '+') meng-
       gantikan N digit terakhir dari SEGMEN DIGIT TERPANJANG pada
       bagian yang paling berstruktur (bisa di posisi mana pun dalam
       rantai '+', termasuk kalau suffix-nya nempel di token terakhir).
       Hanya jalan kalau bagian utama cukup berstruktur; kalau tidak
       (angka pendek polos) -> abaikan (seluruh rantai '+' dibiarkan
       apa adanya).
    6. TERAKHIR ("hapus karakter"): tiap hasil dibuang SEMUA karakter
       selain huruf/angka -- KECUALI kalau hasilnya sendiri cocok pola
       khusus poin 2 (dash dipertahankan).
- INSURED -- delimiter: QQ/QQQ (typo "OO" juga dikenali), AND/OR, "/",
  "//", ",", en-dash "–", "¿", "*", dan "&" HANYA kalau diikuti PT/CV
  (mis. "... & PT OKI PULP..."). "&" polos (mis. "PULP & PAPER") TETAP
  bukan pemisah. "+" dan " - " (spasi-dash-spasi) BUKAN pemisah (supaya
  "SK+SUMATERA" & "PULAU - PULAU MEDIA" tidak salah pecah).
  * ", PT"/", CV" (di mana pun, bukan cuma di akhir) dibuang total
    bersama komanya (cegah split palsu kalau ada isi lain menyertai).
  * ", UD"/", PD" (jenis badan usaha) TIDAK memicu split -- koma
    dibuang, kode dipertahankan menyatu (mis. "GAJAH MAS, UD" -> "GAJAH
    MAS UD").
  * Gelar/sapaan SUFFIX (BP, MR, MRS, NY, TN, DRS, S.E, S.H, dst) yg
    nempel via koma di belakang nama -> dibuang.
  * Pola khusus Pelabuhan Indonesia/Pelindo: "PELABUHAN INDONESIA
    <daftar romawi/angka>" -> pecah per nomor (romawi dikonversi ke
    angka arab).
  * Pola "PREFIX [/] KEYWORD tok1 tok2 ..." (KEYWORD = FACTORY/CLUSTER/
    TOWER/REGIONAL/DISTRICT, daftar >=2 token dipisah koma/&/DAN/spasi)
    -> pecah per token jadi "PREFIX KEYWORD tok". Kalau cuma 1 token
    (mis. "... / DISTRICT 8 / ...") -> TIDAK dipecah jadi entitas baru,
    malah DIGABUNG ke entitas SEBELUMNYA ("... DISTRICT 8").
  * Isi dalam kurung: singkatan dari teks sebelumnya / kata boilerplate
    peran (PERSERO/PRINCIPAL/OWNER/INSURED/CONTRACTOR) / mengandung
    PT-CV -> DIBUANG total. Isi kurung mengandung KOMA (daftar
    deskriptif, mis. "(REFINERY, PLANTATION, KEMITRAAN)") -> kurung
    DIPERTAHANKAN UTUH literal. Isi kurung lain (identitas cabang/pihak
    lain, mis. "(KSO)") -> kurung dibuang, isi dipertahankan.
  * Pola "(AS (THE) OWNER/PRINCIPAL OF <NAMA>)" -> <NAMA> jadi insured
    TERPISAH.
  * Data rusak (kata terduplikasi berturut-turut, mis. "WASKITA WASKITA
    KARYA KARYA", atau kurung tak seimbang) dinormalisasi dulu di awal.
  * Baris SUSPEN(SE)/(H) UTANG PIUTANG / deskripsi treaty-struktur
    (TREATY/STRUCTURE/EXCESS OF LOSS/WHOLE ACCOUNT) -> dibiarkan apa
    adanya (tidak diproses/di-split sama sekali).
  * TIDAK ADA lagi langkah dedup-abbreviation (beda dari KB INS) --
    singkatan yg muncul via delimiter (mis. "SINARMAS LAND GROUP / SML")
    tetap di-breakdown jadi entitas terpisah.
- Filter baris: CEDANT SHRT NAME == "KBRU", STATUS == "SUSPENSE".
- CERTIFICATE_1 SELALU KOSONG (cedant ini tidak punya konsep sertifikat),
  posisinya TEPAT setelah POLIS_CLEAN_1.
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

OUTPUT_FILE = os.path.join("output", "Kbru_output_suspend.xlsx")

CEDANT_FILTER_COL   = "CEDANT SHRT NAME"
CEDANT_FILTER_VALUE = "KBRU"

STATUS_COL          = "STATUS"
STATUS_KEEP_VALUE   = "SUSPENSE"

INSURED_COL = "INSURED"
POLIS_COL   = "POLIS"
SLIP_COL    = "SLIP NO"

MAX_SPLIT_COLS = 5


# ─────────────────────────────────────────────────────────────────────────────
# HELPER UMUM
# ─────────────────────────────────────────────────────────────────────────────

def _normalize_spaces(text: str) -> str:
    return re.sub(r"\s{2,}", " ", text).strip()


def _dedup_preserve_order(items: list) -> list:
    seen, out = set(), []
    for i in items:
        if i not in seen:
            seen.add(i)
            out.append(i)
    return out


def _cap_or_join(items: list) -> list:
    if len(items) > MAX_SPLIT_COLS:
        return [", ".join(items)]
    return items


# ─────────────────────────────────────────────────────────────────────────────
# MESIN POLIS/SLIP KBRU
# ─────────────────────────────────────────────────────────────────────────────

_JUNK_STANDALONE_WORDS = {"TBA", "VAR", "VARIOUS", "SELISIH", "END", "-", "NAN", "NONE", "NULL"}
_P_CODE_RE = re.compile(r"^P\d+$", re.IGNORECASE)

_MONTH_RE = (
    r"JANUARI|FEBRUARI|MARET|APRIL|MEI|JUNI|JULI|AGUSTUS|SEPTEMBER|OKTOBER|"
    r"NOVEMBER|DES[EA]?MBER|JANUARY|FEBRUARY|MARCH|MAY|JUNE|JULY|AUGUST|OCTOBER|DECEMBER"
)
_MONTH_YEAR_TAIL_RE = re.compile(
    r"\s*/?\s*(?:" + _MONTH_RE + r")\.?\s*(?:19|20)?\d{0,4}\s*$", re.IGNORECASE
)

_CN_MARKER_STRIP_RE_UNUSED = None
_NO_ANYWHERE_RE = re.compile(r"\bNO\.\s*", re.IGNORECASE)
_LEADING_ALPHA_PLUS_RE = re.compile(r"^[^\d/]{2,}?\+\s*")
_TRAILING_JUNK_WORD_RE = re.compile(r"[\s/]+(?:TBA|VAR|VARIOUS)\s*$", re.IGNORECASE)

# [BARU 15/09] deskripsi murni SUSPENSE/UTANG PIUTANG yang nyasar di
# kolom polis/slip -- dibiarkan apa adanya (spasi TIDAK dihapus), sama
# spt perlakuan di kolom INSURED.
_SUSPENSE_LIKE_RE = re.compile(r"SUSPEN|\(?H\)?\s*UTANG\s*PIUTANG", re.IGNORECASE)

# [BARU 15/09] teks narasi MURNI (BORDERO/SUMMARY SIMAS RE/nama bulan/
# tahun/pemisah saja, TANPA angka lain) -- dibiarkan apa adanya.
_NARRATIVE_ONLY_RE = re.compile(
    r"^(?:BORDERO|SUMMARY|BULAN|SIMAS\s*RE|" + _MONTH_RE + r"|(?:19|20)\d{2}|[-\s,])+$",
    re.IGNORECASE,
)

# [BARU] rantai P<n> (+ anotasi opsional "(ADD n)") yang digabung "+" --
# dibuang SELURUHNYA (termasuk "+"-nya), di mana pun posisinya (bukan
# cuma di akhir).
_STRAY_P_CHAIN_RE = re.compile(
    r"(?:\bP\d+\b\s*(?:\(ADD\s*\d+\))?\s*\+\s*)+", re.IGNORECASE
)
# [BARU] rantai P<n> yang dipisah SPASI (bukan "+") yang nempel di
# BELAKANG nilai -- dibuang.
_TRAILING_P_CHAIN_RE = re.compile(r"(?:\s+P\d+\b)+\s*$", re.IGNORECASE)

# [BARU] segmen pertama (sebelum "/" pertama) yang berisi RANTAI angka
# pendek digabung "-" (bukan "+") sebanyak >=5 kali -- dianggap
# 'pengulangan berlebih', DIBIARKAN APA ADANYA (raw), sama seperti
# aturan pengulangan >5 via "+".
def _has_excessive_dash_chain(value: str) -> bool:
    first_seg = value.split("/", 1)[0]
    return first_seg.count("-") >= 5


KEEP_AS_IS_PATTERN_RE = re.compile(r"^\d{7,10}[A-Z]{1,2}\d{6,9}-\d{2}$", re.IGNORECASE)


def _is_junk_token(tok: str) -> bool:
    t = tok.strip().upper()
    if not t:
        return True
    if t in _JUNK_STANDALONE_WORDS:
        return True
    if _P_CODE_RE.match(t):
        return True
    return False


def _value_has_real_digits(text: str) -> bool:
    return bool(re.search(r"\d{3,}", text))


def _has_junk_marker(text: str) -> bool:
    return bool(re.search(r"\bTBA\b|\bP\d+\b", text, re.IGNORECASE))


def _split_long_space_tail(val: str):
    parts = re.split(r"\s{2,}", val, maxsplit=1)
    if len(parts) == 1:
        return val, None
    return parts[0].strip(), parts[1].strip()


def _strip_leading_junk_wrapper(val: str):
    """Buang wrapper di depan via '/' SAMPAI ketemu nilai asli:
    - PASS 1: cari segmen PERTAMA yang punya digit -> itu awal nilai
      asli (buang semua sebelumnya, apa pun isinya -- nama cedant/broker
      mis. 'SINARMAS', placeholder TBA/P<n>, dst).
    - PASS 2 (kalau TIDAK ADA segmen berdigit sama sekali): hanya buang
      segmen placeholder BAKU (TBA/P<n>) di depan; segmen bukan-baku
      lain (mis. nama perusahaan tanpa digit) TETAP DIPERTAHANKAN kalau
      itu satu-satunya isi yang tersisa (mis. 'TBA/SINARMAS' -> 'SINARMAS').
    Kalau SEMUA segmen placeholder baku -> None (murni placeholder)."""
    if "/" not in val:
        return None if _is_junk_token(val) else val
    segs = val.split("/")
    for i, s in enumerate(segs):
        if re.search(r"\d", s):
            return "/".join(segs[i:]).strip()
    start_idx = None
    for i, s in enumerate(segs):
        if not _is_junk_token(s.strip()):
            start_idx = i
            break
    if start_idx is None:
        return None
    return "/".join(segs[start_idx:]).strip()


def _strip_trailing_junk_wrapper(val: str) -> str:
    """Buang segmen TANPA digit (nama cedant/broker/kode singkat, mis.
    'ACA', 'SINAR MAS') yang nempel di BELAKANG via '/', SAMPAI ketemu
    segmen yang punya digit (dianggap bagian nomor asli). BEDA dgn
    _strip_leading_junk_wrapper (berbasis _is_junk_token) -- di
    belakang, kriterianya 'ada digit atau tidak', krn tag broker/cedant
    yg nempel di ekor biasanya bukan kode baku (TBA/P<n>) tapi nama
    bebas apa saja (mis. 'ACA')."""
    if "/" not in val:
        return val
    segs = val.split("/")
    end_idx = None
    for i in range(len(segs) - 1, -1, -1):
        if re.search(r"\d", segs[i]):
            end_idx = i + 1
            break
    if end_idx is None:
        return val
    return "/".join(segs[:end_idx]).strip()


def _find_substitution_target(text: str):
    """Cari segmen digit yang jadi TARGET substitusi 'pengulangan'.
    Kasus khusus: pola 'ANGKA + 1-2 HURUF + ANGKA' PERSIS (mis.
    'S00000271' pada '110551216F00000473', tanpa karakter lain) ->
    target SEGMEN TERAKHIR (setelah huruf) -- meski segmen pertama
    (di depan huruf) sering lebih panjang, BUKAN itu yg dimaksud
    'pengulangan'. Selain pola ini -> segmen digit TERPANJANG; kalau
    ada BEBERAPA segmen dgn panjang SAMA (tie), pilih yg PERTAMA di
    antaranya yang BUKAN nol semua (mis. '0000-5747': '0000' dan '5747'
    sama panjang, tapi '0000' cuma nol -> target '5747')."""
    m = re.match(r"^(\d+)([A-Za-z]{1,2})(\d+)$", text)
    if m:
        return m.group(3), m.start(3), m.end(3)
    candidates = list(re.finditer(r"\d+", text))
    if not candidates:
        return None, None, None
    max_len = max(len(c.group(0)) for c in candidates)
    tied = [c for c in candidates if len(c.group(0)) == max_len]
    non_zero = [c for c in tied if set(c.group(0)) != {"0"}]
    best = non_zero[0] if non_zero else tied[0]
    return best.group(0), best.start(), best.end()


def _qualifies_for_substitution(main: str) -> bool:
    """Nilai utama cukup 'berstruktur' utk dijadikan target substitusi
    'pengulangan' kalau: ada huruf (pola FPAR-style), ATAU >=2 pemisah
    '.'/'/' (pola dot-style spt '12.100.0001.82288'), ATAU total digitnya
    panjang (>=10, pola digit polos panjang spt '0124012000618'), ATAU
    ada >=1 pemisah '.'/'/'/'-' DAN total digit >=7 (pola spt '0000-
    5747', beda dari angka pendek polos TANPA pemisah spt '3522' yang
    TETAP tidak memenuhi)."""
    if re.search(r"[A-Za-z]", main):
        return True
    sep_count = len(re.findall(r"[./]", main))
    if sep_count >= 2:
        return True
    digit_len = len(re.sub(r"\D", "", main))
    if digit_len >= 10:
        return True
    if len(re.findall(r"[.\-/]", main)) >= 1 and digit_len >= 7:
        return True
    return False


def _apply_segment_substitution(primary: str, extra_tokens: list) -> list:
    seg, start, end = _find_substitution_target(primary)
    if seg is None:
        return [primary]
    prefix, suffix = primary[:start], primary[end:]
    out = [primary]
    for tok in extra_tokens:
        tok = tok.strip()
        if not tok.isdigit():
            continue
        if len(tok) >= len(seg):
            new_seg = tok
        else:
            new_seg = seg[: -len(tok)] + tok
        out.append(prefix + new_seg + suffix)
    return _dedup_preserve_order(out)


# [BARU 15/09] dash NYASAR di dalam pola F/S 2-segmen (mis.
# '11055129F00000-275' -- seharusnya '11055129F00000275') -- dash
# dianggap noise/typo, dilebur.
_FS_STRAY_DASH_RE = re.compile(r"^(\d+[A-Za-z]{1,2}\d+)-(\d+)$")


def _merge_fs_stray_dash(token: str) -> str:
    m = _FS_STRAY_DASH_RE.match(token)
    if m:
        return m.group(1) + m.group(2)
    return token


_CN_TRAILING_RE = re.compile(r"\s*/?\s*CN\s*:\s*.*$", re.IGNORECASE)
_CN_LEADING_WITH_SLASH_RE = re.compile(r"^\s*CN\s*:\s*[^/]*/\s*", re.IGNORECASE)
_CN_LABEL_ONLY_RE = re.compile(r"^\s*CN\s*:\s*", re.IGNORECASE)
_CN_STARTS_RE = re.compile(r"^CN\s*:", re.IGNORECASE)


def _strip_cn_suffix(val: str) -> str:
    """Notasi 'CN : <nomor>':
    - Muncul di TENGAH/BELAKANG ('main / CN : cn') -> dibuang total
      bersama nilainya, sisakan main di depan.
    - Muncul di AWAL DIIKUTI nilai lain setelah '/' ('CN : cn / main')
      -> label+nilai CN dibuang, sisakan main setelah '/'.
    - HANYA 'CN : cn' saja (seluruh value, tanpa '/' sama sekali) ->
      cuma label 'CN :'-nya yg dibuang, nilai CN dipertahankan sbg
      hasil utama (BUKAN dibuang semua)."""
    stripped = val.strip()
    if not _CN_STARTS_RE.match(stripped):
        return _CN_TRAILING_RE.sub("", val).strip(" /-")
    m = _CN_LEADING_WITH_SLASH_RE.match(val)
    if m:
        return val[m.end():].strip(" /-")
    return _CN_LABEL_ONLY_RE.sub("", val).strip(" /-")


_EMBEDDED_SUFFIX_RE = re.compile(r"^(\d+)([./].*[A-Za-z].*)$")


def _split_embedded_suffix(token: str):
    """Deteksi token yg formatnya 'ANGKA' + 'SUFFIX BERSTRUKTUR' nempel
    jadi satu (mis. '23022/FIAR/001.0000344-SRB/06/2022' -> ('23022',
    '/FIAR/001.0000344-SRB/06/2022')). Dipakai utk kasus 'pengulangan'
    di mana suffix format cuma nempel di SALAH SATU token dlm rantai
    '+' (biasanya yg terakhir), bukan di token pertama."""
    m = _EMBEDDED_SUFFIX_RE.match(token)
    if m:
        return m.group(1), m.group(2)
    return token, None


def _strip_all_characters(value: str) -> str:
    return re.sub(r"[^A-Za-z0-9]", "", value)


_SINAR_MAS_RE = re.compile(r"\bSINAR\s+MAS\b", re.IGNORECASE)
_TEXT_PIECE_PTCV_RE = re.compile(r",?\s*\b(?:P\s*T|CV)\b\.?\s*$", re.IGNORECASE)


def _clean_text_piece(value: str) -> str:
    """Utk hasil yang ternyata berupa TEKS/NAMA (tanpa digit sama
    sekali, muncul lewat mekanisme buang TBA/P<n> yang bersandingan) --
    JANGAN di-'hapus karakter' penuh (spasi antar kata dipertahankan,
    beda dari nilai numerik biasa), cuma normalisasi spasi + standardisasi
    'SINAR MAS' -> 'SINARMAS' (konsisten dgn kolom INSURED) + buang
    suffix "PT"/"CV" di akhir (mis. "ASURANSI FPG INSURANCE, PT" ->
    "ASURANSI FPG INSURANCE")."""
    value = _SINAR_MAS_RE.sub("SINARMAS", value)
    value = _TEXT_PIECE_PTCV_RE.sub("", value)
    return _normalize_spaces(value).upper()


def _simplify_noletter_extra(e: str) -> str:
    """[BARU 15/09] Token tambahan (di rantai '+') yang TIDAK punya
    huruf tapi masih ada '/' atau '.' nempel (mis. '04820/0001189' dari
    '...+04820/0001189.') -- ambil cuma potongan PERTAMA sbg kandidat
    substitusi digit, sisanya (mis. sertifikat/nomor duplikat) dibuang
    (KBRU tidak punya kolom sertifikat, jadi info ini tidak tertampung)."""
    if re.search(r"[A-Za-z]", e):
        return e
    return re.split(r"[./]", e)[0]


def clean_value(val) -> list:
    if pd.isna(val):
        return []
    raw = str(val).strip()
    if not raw:
        return []

    if raw.upper() in _JUNK_STANDALONE_WORDS:
        return [raw]
    # [BARU 15/09] "TBA/" (placeholder + garis miring doang, tanpa isi
    # lain) -> garis miringnya dibuang, cukup "TBA".
    raw_no_slash = raw.strip("/ ")
    if raw_no_slash.upper() in _JUNK_STANDALONE_WORDS:
        return [raw_no_slash]
    if KEEP_AS_IS_PATTERN_RE.match(raw):
        return [raw]
    if re.fullmatch(r"\d+&\d+", raw):
        return [raw]
    if _has_excessive_dash_chain(raw):
        # segmen depan berisi >=5 angka digabung '-' -- pengulangan
        # berlebih, dibiarkan apa adanya (sama spt aturan >5 via '+')
        return [raw]
    if _SUSPENSE_LIKE_RE.search(raw) or _NARRATIVE_ONLY_RE.match(raw.strip()):
        # [BARU 15/09] deskripsi SUSPENSE/UTANG PIUTANG, atau narasi
        # murni (BORDERO/bulan/tahun/SUMMARY SIMAS RE) yg nyasar di
        # kolom polis/slip -- dibiarkan apa adanya, spasi TIDAK dihapus
        return [raw]

    if not _has_junk_marker(raw) and not _value_has_real_digits(raw):
        # teks tanpa placeholder (TBA/P<n>) & tanpa digit sama sekali
        # (mis. nama perusahaan nyasar, "ASURANSI SINAR MAS") -> tetap
        # "abaikan" (bukan diproses spt polis/slip), TAPI tetap lewat
        # normalisasi teks (spasi + SINAR MAS->SINARMAS) spy konsisten
        # dgn kasus yg lewat jalur TBA-adjacent.
        return [_clean_text_piece(raw)]

    val2 = _NO_ANYWHERE_RE.sub("", raw)
    val2 = _strip_cn_suffix(val2)
    val2 = _LEADING_ALPHA_PLUS_RE.sub("", val2)
    val2 = _STRAY_P_CHAIN_RE.sub("", val2)
    val2 = _TRAILING_P_CHAIN_RE.sub("", val2)

    pieces = []
    stripped = _strip_leading_junk_wrapper(val2)
    if stripped is not None:
        val2 = stripped
        val2 = _TRAILING_P_CHAIN_RE.sub("", val2)
        primary, _tail = _split_long_space_tail(val2)
        primary2 = _MONTH_YEAR_TAIL_RE.sub("", primary).strip(" -/,")
        primary = primary2 if primary2 else primary
        primary = _TRAILING_JUNK_WORD_RE.sub("", primary).strip()

        raw_parts = [p.strip() for p in re.split(r"\s*\+\s*", primary) if p.strip()]
        raw_parts = [_merge_fs_stray_dash(p) for p in raw_parts]
        # [FIX 15/09] "_strip_trailing_junk_wrapper" cuma dipakai kalau
        # MEMANG ada rantai '+' (>1 bagian) -- kalau cuma 1 bagian
        # (whole value, mis. hasil gabungan lewat "No." di tengah),
        # JANGAN dipangkas, supaya ekor sah spt "/AB"/"/BB" (kode akhir
        # 2 huruf tanpa digit) tidak ikut kebuang.
        apply_trailing_strip = len(raw_parts) > 1
        parts = []
        for p in raw_parts:
            if _is_junk_token(p):
                continue
            p2 = _strip_trailing_junk_wrapper(p) if apply_trailing_strip else p
            p2 = p2.strip()
            if p2 and not _is_junk_token(p2):
                parts.append(p2)

        if parts:
            if len(parts) > 1:
                # kalau ADA campuran (sebagian bagian punya digit,
                # sebagian tidak sama sekali, mis. 'UMUM MEGA' vs
                # '2010208012400001') -> bagian tanpa digit dibuang
                # (nama cedant/broker yg nyempil, bukan bagian nilai).
                digit_parts = [p for p in parts if re.search(r"\d", p)]
                if digit_parts:
                    parts = digit_parts

            # [CEK 1] suffix berstruktur nempel di SALAH SATU token
            # (biasanya token TERAKHIR, mis. '23018+23019+23023+23022/
            # FIAR/001.0000344-SRB/06/2022') -- kalau ketemu & SISANYA
            # semua digit polos, terapkan suffix yg sama ke semua nomor.
            suffix_found = None
            base_numbers = []
            embed_ok = True
            for p in parts:
                num_part, suf = _split_embedded_suffix(p)
                if suf:
                    if suffix_found is None:
                        suffix_found = suf
                    else:
                        embed_ok = False
                if not num_part.isdigit():
                    embed_ok = False
                base_numbers.append(num_part)

            if embed_ok and suffix_found and len(base_numbers) > 1:
                ref = max(base_numbers, key=len)
                combined = []
                for n in base_numbers:
                    if len(n) >= len(ref):
                        combined.append(n + suffix_found)
                    else:
                        combined.append(ref[: -len(n)] + n + suffix_found)
                pieces = _dedup_preserve_order(combined)
            else:
                # [CEK 2] main = bagian yg PALING berstruktur (ada digit
                # DAN (huruf ATAU >=2 pemisah)), di mana pun posisinya
                # dalam rantai '+' -- sisanya jadi kandidat token
                # substitusi/nilai lain.
                structured_idx = None
                for i, p in enumerate(parts):
                    if re.search(r"\d", p) and (
                        re.search(r"[A-Za-z]", p) or len(re.findall(r"[./]", p)) >= 2
                    ):
                        structured_idx = i
                        break
                if structured_idx is not None:
                    main = parts[structured_idx]
                    extras_raw = parts[:structured_idx] + parts[structured_idx + 1:]
                else:
                    main = parts[0]
                    extras_raw = parts[1:]

                extras_clean = [_simplify_noletter_extra(e.strip().rstrip(".").strip()) for e in extras_raw]
                whole_extras = [e for e in extras_clean if "/" in e or re.search(r"\d\.\d", e)]
                digit_extras = [e for e in extras_clean if e.isdigit()]

                # [BARU 15/09] token substitusi yg SANGAT pendek (<=2
                # digit, mis. "+97+07+43") -- pola pengulangannya tidak
                # jelas (berapa digit ekor yg diganti jadi ambigu) ->
                # seluruh nilai dibiarkan apa adanya.
                if digit_extras and any(len(e) < 3 for e in digit_extras):
                    pieces = []
                elif digit_extras and not _qualifies_for_substitution(main):
                    pieces = []  # ambigu -> jatuh ke abaikan (raw) di bawah
                elif digit_extras:
                    pieces = _apply_segment_substitution(main, digit_extras)
                    pieces += whole_extras
                else:
                    pieces = [main] + whole_extras

    pieces = _dedup_preserve_order([_normalize_spaces(p) for p in pieces if p.strip()])

    if not pieces or len(pieces) > MAX_SPLIT_COLS:
        return [raw]

    final = []
    for p in pieces:
        if KEEP_AS_IS_PATTERN_RE.match(p):
            final.append(p)
        elif not re.search(r"\d", p):
            final.append(_clean_text_piece(p))
        else:
            final.append(_strip_all_characters(p))
    return final


def clean_polis(val):
    return clean_value(val)


def clean_slip(val):
    return clean_value(val)


# ─────────────────────────────────────────────────────────────────────────────
# CLEAN INSURED
# ─────────────────────────────────────────────────────────────────────────────

_GELAR_PREFIX_RE = re.compile(
    r"^\s*(?:BAPAK|IBU|SDRI?|NYONYA|NY|TUAN|TN|MR|MRS|MS|DR|IR|PROF|HJ|H)\b\.?\s*",
    re.IGNORECASE,
)
_GELAR_SUFFIX_RE = re.compile(
    r"\s*,?\s*\b(?:S\.?E|S\.?H|S\.?SI|S\.?T|S\.?SOS|S\.?KOM|S\.?PD|"
    r"M\.?SI|M\.?M|M\.?H|M\.?KOM|SH|SE|BP|MRS|MR|NY|TN|DRS)\b\.?\s*$",
    re.IGNORECASE,
)
_SUSPENSE_KEEP_RE = re.compile(r"SUSPEN|\(?H\)?\s*UTANG\s*PIUTANG", re.IGNORECASE)
_TREATY_KEEP_RE = re.compile(
    r"\bTREATY\b|\bSTRUCTURE\b|EXCESS\s+OF\s+LOSS|WHOLE\s+ACCOUNT", re.IGNORECASE
)


def _is_keep_as_is_insured(val: str) -> bool:
    if _SUSPENSE_KEEP_RE.search(val):
        return True
    if _TREATY_KEEP_RE.search(val):
        return True
    return False


def _collapse_duplicate_words(text: str) -> str:
    """Normalisasi data rusak: kata terduplikasi persis berturut-turut
    (mis. 'WASKITA WASKITA KARYA KARYA' -> 'WASKITA KARYA'), kata
    gandeng tanpa spasi ('PERSEROPERSERO' -> 'PERSERO'), kurung ganda/
    tak seimbang."""
    text = re.sub(r"\(\s*\(", "(", text)
    text = re.sub(r"\)\s*\)", ")", text)
    text = re.sub(r"\b(\w{3,})\1\b", r"\1", text, flags=re.IGNORECASE)
    text = re.sub(r"\b(\w+)(?:\s+\1\b)+", r"\1", text, flags=re.IGNORECASE)
    if text.count("(") > text.count(")"):
        text = text + ")"
    return text


_KEEP_SUFFIX_COMMA_RE = re.compile(r",\s*(UD|PD)\b", re.IGNORECASE)
_REMOVE_COMMA_PTCV_RE = re.compile(r",\s*(?:P\s*T|CV)\b\.?", re.IGNORECASE)

# pemisah: QQ/QQQ, OO (typo QQ), AND/OR, "/", "//", ",", en-dash, "¿",
# "*", dan "&" HANYA kalau diikuti PT/CV. "+" dan " - " BUKAN pemisah.
_SPLIT_RE = re.compile(
    r"\bQ{2,3}\b|\bOO\b|\bAND\s*/\s*OR\.?\b|/{1,2}|,|–|¿|\*|&\s*(?=(?:PT|CV)\b)",
    re.IGNORECASE,
)

INSURED_JUNK_WORDS = frozenset({
    "PT", "CV", "TBK", "PERSERO", "LTD", "INC", "LLC",
    "AND", "OR", "THE", "OF", "AS", "VARIOUS", "VAR", "QQ", "TBA",
    "ASSOCIATED",
})

_MONTH_STANDALONE_RE = re.compile(
    r"^(?:" + _MONTH_RE + r")\.?\s*(?:19|20)?\d{2,4}$", re.IGNORECASE
)

_GELAR_PREFIX_RE = re.compile(
    r"^\s*(?:BAPAK|IBU|SDRI?|NYONYA|NY|TUAN|TN|MR|MRS|MS|DR|IR|PROF|HJ|H)\b\.?\s*",
    re.IGNORECASE,
)
# [BARU] tambah BP/MR/MRS/NY/TN/DRS sbg gelar SUFFIX (nempel via koma di
# BELAKANG nama, mis. "SUTJITA, BP" / "RUDY AKILI, MR") -- dibuang.
_GELAR_SUFFIX_RE = re.compile(
    r"\s*,?\s*\b(?:S\.?E|S\.?H|S\.?SI|S\.?T|S\.?SOS|S\.?KOM|S\.?PD|"
    r"M\.?SI|M\.?M|M\.?H|M\.?KOM|SH|SE|BP|MRS|MR|NY|TN|DRS)\b\.?\s*$",
    re.IGNORECASE,
)
_SUSPENSE_KEEP_RE = re.compile(r"SUSPEN|\(?H\)?\s*UTANG\s*PIUTANG", re.IGNORECASE)
_TREATY_KEEP_RE = re.compile(
    r"\bTREATY\b|\bSTRUCTURE\b|EXCESS\s+OF\s+LOSS|WHOLE\s+ACCOUNT", re.IGNORECASE
)


def _is_keep_as_is_insured(val: str) -> bool:
    if _SUSPENSE_KEEP_RE.search(val):
        return True
    if _TREATY_KEEP_RE.search(val):
        return True
    if _MALFORMED_QQ_RE.search(val):
        return True
    return False


# [BARU] normalisasi data rusak: kata terduplikasi persis berturut-turut
# (mis. "WASKITA WASKITA KARYA KARYA" -> "WASKITA KARYA"), kata gandeng
# tanpa spasi ("PERSEROPERSERO" -> "PERSERO"), kurung ganda/tak seimbang.
def _collapse_duplicate_words(text: str) -> str:
    text = re.sub(r"\(\s*\(", "(", text)
    text = re.sub(r"\)\s*\)", ")", text)
    text = re.sub(r"\b(\w{3,})\1\b", r"\1", text, flags=re.IGNORECASE)
    text = re.sub(r"\b(\w+)(?:\s+\1\b)+", r"\1", text, flags=re.IGNORECASE)
    if text.count("(") > text.count(")"):
        text = text + ")"
    return text


# [BARU] ", PT" / ", CV" (di mana pun, bukan cuma di akhir) DIBUANG
# TOTAL bersama komanya -- supaya tidak memicu split palsu saat ada isi
# lain menyertai PT/CV di segmen yang sama (mis. "SMAILING TOUR, PT &
# TRAVEL SERVICE, PT" -> "SMAILING TOUR & TRAVEL SERVICE", bukan
# terpecah jadi "SMAILING TOUR" + "& TRAVEL SERVICE").
_REMOVE_COMMA_PTCV_RE = re.compile(r",\s*(?:P\s*T|CV)\b\.?", re.IGNORECASE)

# ", UD" / ", PD" (jenis badan usaha) TIDAK dianggap pemisah entitas
# baru -- koma dibuang, kode dipertahankan menyatu dgn nama sebelumnya.
_KEEP_SUFFIX_COMMA_RE = re.compile(r",\s*(UD|PD)\b", re.IGNORECASE)

# pemisah: QQ/QQQ, OO (typo QQ), AND/OR, "/", "//", ",", en-dash, "¿",
# "*", dan "&" HANYA kalau diikuti PT/CV (JV baru, mis. "... & PT OKI
# PULP..."). "&" polos (mis. "PULP & PAPER") TETAP bukan pemisah. "+"
# TIDAK dipakai sbg pemisah (mis. "SK+SUMATERA" -> 1 entitas "SK
# SUMATERA", bukan 2) -- beda dgn mesin polis/slip. " - " (spasi-dash-
# spasi) juga TIDAK dipakai (bikin false split di nama spt "PULAU -
# PULAU MEDIA").
_SPLIT_RE = re.compile(
    r"\bQ{2,3}\b|\bOO\b|\bAND\s*/\s*OR\.?\b|/{1,2}|,|–|¿|\*|&\s*(?=(?:PT|CV)\b)",
    re.IGNORECASE,
)

INSURED_JUNK_WORDS = frozenset({
    "PT", "CV", "TBK", "PERSERO", "LTD", "INC", "LLC",
    "AND", "OR", "THE", "OF", "AS", "VARIOUS", "VAR", "QQ", "TBA",
    "ASSOCIATED",
})

_OWNER_OF_RE = re.compile(
    r"\(\s*AS\s+(?:THE\s+)?(?:OWNER|PRINCIPAL)\s+OF\s+([^()]+)\)", re.IGNORECASE
)

_ROLE_KEYWORDS_RE = re.compile(
    r"^(?:PERSERO|PRINCIPAL|OWNER|INSURED|CONTRACTOR)$", re.IGNORECASE
)

# [DIPERLONGGAR] toleransi "P T" (spasi nyempil di antara P & T, data
# rusak) selain "PT" biasa.
_PT_CV_PREFIX_RE = re.compile(r"^\s*(?:P\s*T|CV)(?:\.|(?=\s|$))\.?\s*", re.IGNORECASE)
_PT_CV_SUFFIX_RE = re.compile(r",?\s*(?:P\s*T|CV)\.?\s*$", re.IGNORECASE)

_MASK_COMMA = "\x01"
_MASK_SLASH = "\x02"


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


def _extract_owner_of_entities(text: str):
    extra = []

    def repl(m):
        extra.append(m.group(1).strip())
        return " "

    new_text = _OWNER_OF_RE.sub(repl, text)
    return new_text, extra


def _strip_parenthetical(text: str) -> str:
    """Isi dalam kurung yang singkatan dari teks sebelumnya, ATAU kata
    boilerplate peran (PERSERO/PRINCIPAL/OWNER/INSURED/CONTRACTOR), ATAU
    mengandung PT/CV -> kurung + isi DIBUANG total. Isi kurung yang
    mengandung KOMA (daftar deskriptif, mis. "(REFINERY, PLANTATION,
    KEMITRAAN)") -> kurung DIPERTAHANKAN UTUH (literal, koma di-mask
    dulu spy tdk ikut split). Isi kurung lain (identitas cabang/pihak
    lain, mis. "(KSO)") -> kurung dibuang, ISI dipertahankan."""

    def repl(m):
        inner = m.group(1).strip()
        if re.search(r"\bPT\b|\bCV\b", inner, re.IGNORECASE):
            return " "
        if _ROLE_KEYWORDS_RE.match(inner):
            return " "
        inner_letters = re.sub(r"[^A-Za-z]", "", inner)
        before = text[: m.start()].strip()
        if before and inner_letters and _is_abbreviation_of(inner_letters, before):
            return " "
        if "," in inner:
            masked = inner.replace(",", _MASK_COMMA).replace("/", _MASK_SLASH)
            return f" ({masked}) "
        masked = inner.replace(",", _MASK_COMMA).replace("/", _MASK_SLASH)
        return f" {masked} "

    return re.sub(r"\(([^()]*)\)", repl, text)


# ─────────────────────────────────────────────────────────────────────────────
# [BARU] EKSPANSI POLA "KEYWORD + DAFTAR" (FACTORY/CLUSTER/TOWER/
# REGIONAL/DISTRICT) & KHUSUS PELABUHAN INDONESIA/PELINDO
# ─────────────────────────────────────────────────────────────────────────────

_ROMAN_MAP = {"I": 1, "II": 2, "III": 3, "IV": 4, "V": 5, "VI": 6, "VII": 7, "VIII": 8, "IX": 9, "X": 10}
_ROMAN_ARABIC_ALT = r"I|II|III|IV|V|VI|VII|VIII|IX|X|\d+"

_PELINDO_RE = re.compile(
    r"^(PELABUHAN\s+INDONESIA)\s+((?:" + _ROMAN_ARABIC_ALT + r")"
    r"(?:\s*(?:,|&|DAN)\s*(?:" + _ROMAN_ARABIC_ALT + r"))+)\s*$",
    re.IGNORECASE,
)

# [BARU 15/09] "PELINDO" adalah SINONIM "PELABUHAN INDONESIA" -- bisa
# disebut BERULANG-ULANG dgn pemisah apa saja (biasanya QQ), mis.
# "PELABUHAN INDONESIA I QQ PELINDO I QQ PELINDO II QQ PELINDO III QQ
# PELINDO IV" -> semua kemunculan "PELABUHAN INDONESIA <n>"/"PELINDO <n>"
# dikumpulkan, dinormalisasi ke "PELABUHAN INDONESIA <n>" (arab), dedup.
_PELINDO_MENTION_RE = re.compile(
    r"\b(?:PELABUHAN\s+INDONESIA|PELINDO)\s+(" + _ROMAN_ARABIC_ALT + r")\b",
    re.IGNORECASE,
)


def _expand_pelindo(val: str):
    m = _PELINDO_RE.match(val.strip())
    if not m:
        return None
    prefix, nums_part = m.groups()
    tokens = [t.strip() for t in re.split(r"\s*(?:,|&|DAN)\s*", nums_part, flags=re.IGNORECASE) if t.strip()]
    out = []
    for t in tokens:
        num = _ROMAN_MAP.get(t.upper(), t)
        out.append(f"{prefix.upper()} {num}")
    return out


def _expand_pelindo_mentions(val: str):
    matches = _PELINDO_MENTION_RE.findall(val)
    if len(matches) < 2:
        return None
    seen = []
    for t in matches:
        num = _ROMAN_MAP.get(t.upper(), t)
        entry = f"PELABUHAN INDONESIA {num}"
        if entry not in seen:
            seen.append(entry)
    return seen if len(seen) >= 2 else None


_KEYWORD_MULTI_RE = re.compile(
    r"^([^/]*?)[/\s]*\b(FACTORY|CLUSTER|TOWER|REGIONAL|DISTRICT|UNIT|MALL)\s+([^/]+)$",
    re.IGNORECASE,
)


def _expand_keyword_multi(val: str):
    """'PREFIX [/] KEYWORD tok1 tok2 ...' atau 'PREFIX [/] KEYWORD tok1,
    tok2 & tok3' -> daftar 'PREFIX KEYWORD tok'. Butuh >=2 token supaya
    tidak salah nembak segmen tunggal (yg ditangani _merge_keyword_single
    di tempat lain)."""
    m = _KEYWORD_MULTI_RE.match(val.strip())
    if not m:
        return None
    prefix, keyword, tail = m.groups()
    tokens = [t for t in re.split(r"\s*(?:,|&|DAN)\s*|\s+", tail.strip(), flags=re.IGNORECASE) if t.strip()]
    tokens = [t.strip() for t in tokens if re.fullmatch(r"[A-Za-z0-9]{1,4}", t.strip())]
    if len(tokens) < 2:
        return None
    prefix = prefix.strip().rstrip(",").strip()
    prefix = re.sub(r",?\s*(?:P\s*T|CV)\.?\s*$", "", prefix, flags=re.IGNORECASE).strip()
    return [f"{prefix} {keyword.upper()} {t.upper()}".strip() for t in tokens]


_KEYWORD_SINGLE_RE = re.compile(
    r"^(?:FACTORY|CLUSTER|TOWER|REGIONAL|DISTRICT|UNIT|MALL)\s+[A-Za-z0-9]{1,4}$", re.IGNORECASE
)

# [BARU 15/09] "QQ" yang NYASAR nempel langsung ke huruf lain tanpa
# spasi/pemisah (mis. "QQBALI", "INDAHQQ") -- tanda data rusak/typo yg
# terlalu ambigu utk diproses -> seluruh value dibiarkan apa adanya.
# Rantai Q murni ("QQQ") TIDAK dihitung rusak -- itu cuma varian QQ
# dgn Q berlebih, sudah ditangani _SPLIT_RE (\bQ{2,3}\b).
_MALFORMED_QQ_RE = re.compile(r"[A-PR-Za-pr-z]QQ|QQ[A-PR-Za-pr-z]")


def _clean_insured_name(name: str) -> str:
    name = _normalize_spaces(name)
    name = _GELAR_PREFIX_RE.sub("", name)
    for _ in range(3):
        stripped = _GELAR_SUFFIX_RE.sub("", name).strip()
        if stripped == name:
            break
        name = stripped
    name = re.sub(r",?\s*\(\s*PERSERO\s*\)\s*", " ", name, flags=re.IGNORECASE)
    name = re.sub(r"\bTBK\b\s*,\s*P\s*T\.?\s*", " ", name, flags=re.IGNORECASE)
    name = re.sub(r",?\s*\bTBK\b\s*", " ", name, flags=re.IGNORECASE)
    name = re.sub(r"\bLTD\b\.?\s*$", "", name, flags=re.IGNORECASE)
    name = re.sub(r"\bPTE\b\.?\s*", "", name, flags=re.IGNORECASE)
    name = _PT_CV_SUFFIX_RE.sub("", name)
    name = _PT_CV_PREFIX_RE.sub("", name)
    # [FIX 15/09] pola inisial huruf-tunggal-bertitik berturut-turut
    # (mis. "S.A.R.I") -> titik dibuang TANPA jadi spasi ("SARI"),
    # BUKAN "S A R I". Titik lain (bukan pola ini, bukan antar digit)
    # baru diganti spasi spt biasa.
    name = re.sub(r"\b(?:[A-Za-z]\.){2,}[A-Za-z]?\b", lambda mo: mo.group(0).replace(".", ""), name)
    name = re.sub(r"(?<!\d)\.(?!\d)", " ", name)
    name = re.sub(r"\s*\+\s*", " ", name)
    name = _SINAR_MAS_RE.sub("SINARMAS", name)
    name = name.replace(_MASK_COMMA, ",").replace(_MASK_SLASH, "/")
    return _normalize_spaces(name).strip().upper()


def _try_expansions(segment: str):
    """Coba semua mekanisme ekspansi (pelindo list, pelindo scattered
    mentions, keyword+daftar) pada SATU string. None kalau tidak ada yg
    cocok."""
    r = _expand_pelindo(segment)
    if r is not None:
        return r
    r = _expand_pelindo_mentions(segment)
    if r is not None:
        return r
    r = _expand_keyword_multi(segment)
    if r is not None:
        return r
    return None


def clean_insured(val) -> list:
    if pd.isna(val):
        return []
    val = str(val).strip()
    if not val:
        return []

    val = _collapse_duplicate_words(val)

    if _is_keep_as_is_insured(val):
        return [_normalize_spaces(val).upper()]

    val, owner_extra = _extract_owner_of_entities(val)
    val = _strip_parenthetical(val)
    val = _KEEP_SUFFIX_COMMA_RE.sub(r" \1", val)
    val = _REMOVE_COMMA_PTCV_RE.sub("", val)

    # coba ekspansi pada SELURUH value dulu (kasus umum: satu pola
    # keyword/pelindo mencakup semua, mis. "BIO FARMA/CLUSTER B C D",
    # "PELABUHAN INDONESIA I QQ PELINDO II ...").
    whole = _try_expansions(val)
    if whole is not None:
        return [_normalize_spaces(x).upper() for x in whole]

    parts = [p.strip() for p in _SPLIT_RE.split(val) if p.strip()]
    cleaned = []
    for p in parts:
        # [FIX 15/09] kalau SATU bagian hasil split (bukan whole value)
        # sendiri cocok pola keyword+daftar (mis. "GALAXY MALL I & II"
        # setelah "SINAR GALAXY /" ikut terpisah lebih dulu) -> expand
        # jadi beberapa entitas, JANGAN diproses sbg 1 entitas biasa.
        r = _expand_keyword_multi(p)
        if r is not None and len(r) >= 2:
            cleaned.extend(_normalize_spaces(x).upper() for x in r)
            continue
        p_clean = _clean_insured_name(p)
        if not p_clean:
            continue
        if _MONTH_STANDALONE_RE.match(p_clean):
            continue
        if _KEYWORD_SINGLE_RE.match(p_clean) and cleaned:
            cleaned[-1] = _normalize_spaces(cleaned[-1] + " " + p_clean)
            continue
        if len(p_clean) < 2 or p_clean in INSURED_JUNK_WORDS:
            continue
        cleaned.append(p_clean)

    for extra in owner_extra:
        extra_clean = _clean_insured_name(extra)
        if extra_clean and extra_clean not in cleaned:
            cleaned.append(extra_clean)

    if not cleaned:
        fallback = _clean_insured_name(val)
        return [fallback] if fallback else []

    return _cap_or_join(cleaned)


# ─────────────────────────────────────────────────────────────────────────────
# PROSES UTAMA
# ─────────────────────────────────────────────────────────────────────────────

def _force_text_format(output_file, prefixes):
    import openpyxl
    wb = openpyxl.load_workbook(output_file)
    ws = wb.active
    header = [c.value for c in ws[1]]
    target_cols = [i for i, h in enumerate(header, start=1) if h and str(h).startswith(prefixes)]
    for col_idx in target_cols:
        for row in ws.iter_rows(min_row=2, min_col=col_idx, max_col=col_idx):
            for cell in row:
                cell.number_format = "@"
    wb.save(output_file)


def _insert_clean_columns(df, all_lists, prefix, max_cols):
    added = []
    actual_max_cols = max((len(lst) for lst in all_lists if isinstance(lst, list)), default=1)
    actual_max_cols = min(actual_max_cols, max_cols)
    for i in range(1, actual_max_cols + 1):
        col_name = f"{prefix}_{i}"
        df[col_name] = [lst[i - 1] if isinstance(lst, list) and i - 1 < len(lst) else None for lst in all_lists]
        added.append(col_name)
    return added


def _insert_polis_cert_columns(df, all_polis_lists, max_cols):
    """POLIS_CLEAN_i dgn CERTIFICATE_1 TEPAT setelah POLIS_CLEAN_1
    (SELALU kosong -- cedant ini tidak punya konsep sertifikat)."""
    added = []
    actual_max_cols = max((len(lst) for lst in all_polis_lists if isinstance(lst, list)), default=1)
    actual_max_cols = min(actual_max_cols, max_cols)
    for i in range(1, actual_max_cols + 1):
        polis_col = f"POLIS_CLEAN_{i}"
        df[polis_col] = [lst[i - 1] if isinstance(lst, list) and i - 1 < len(lst) else None for lst in all_polis_lists]
        added.append(polis_col)
        if i == 1:
            df["CERTIFICATE_1"] = None
            added.append("CERTIFICATE_1")
    return added


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

    all_polis, all_slip, all_insured = [], [], []
    max_polis = max_slip = max_ins = 1

    for _, row in df.iterrows():
        p_clean = clean_polis(row.get("POLIS_ORI", ""))
        s_clean = clean_slip(row.get("SLIP_NO_ORI", ""))
        parts = clean_insured(row.get("INSURED_ORI", ""))

        max_polis = max(max_polis, len(p_clean))
        max_slip = max(max_slip, len(s_clean))
        max_ins = max(max_ins, len(parts))

        all_polis.append(p_clean)
        all_slip.append(s_clean)
        all_insured.append(parts)

    print("[4/5] Menyusun kolom output ...")

    insured_cols = _insert_clean_columns(df, all_insured, "INSURED", max_ins)
    polis_cols = _insert_polis_cert_columns(df, all_polis, max_polis)
    slip_cols = _insert_clean_columns(df, all_slip, "SLIP_CLEAN", max_slip)

    new_columns = []
    for col in df.columns:
        if col in insured_cols + polis_cols + slip_cols:
            continue
        new_columns.append(col)
        if col == "INSURED_ORI":
            new_columns += insured_cols
        elif col == "POLIS_ORI":
            new_columns += polis_cols
        elif col == "SLIP_NO_ORI":
            new_columns += slip_cols

    df = df[new_columns]

    os.makedirs(os.path.dirname(output_file), exist_ok=True)
    print(f"[5/5] Menyimpan hasil ke: {output_file} ...")
    df.to_excel(output_file, index=False)
    _force_text_format(output_file, prefixes=("POLIS_CLEAN_", "CERTIFICATE_", "SLIP_CLEAN_"))

    print(f"\n{'=' * 55}")
    print("  [OK] Selesai!")
    print(f"  Total baris output  : {len(df):,}")
    print(f"  Total kolom output  : {len(df.columns)}")
    print(f"  File disimpan di    : {output_file}")
    print(f"{'=' * 55}")


if __name__ == "__main__":
    process_data(INPUT_FILE, OUTPUT_FILE)