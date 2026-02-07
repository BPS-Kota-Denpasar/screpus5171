# =========================
# READY-TO-COPY SCRIPT (tanpa menghapus fungsi yang ada)
# - Tambahan: FORCE STOP (Ctrl+C / STOP.txt)
# - Tambahan: AUTOSAVE aman (tmp -> replace) + autosave backup
# - Perbaikan: handle query kosong, save robust, folder screenshot, log lebih jelas
# =========================

import os
import time
import re
import pandas as pd
import signal
from urllib.parse import quote_plus
from difflib import SequenceMatcher

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import (
    TimeoutException,
    StaleElementReferenceException,
    WebDriverException,
)
from webdriver_manager.chrome import ChromeDriverManager

os.environ.setdefault("TF_CPP_MIN_LOG_LEVEL", "3")

# =========================
# FORCE STOP + AUTOSAVE
# =========================
STOP_FILE = "STOP.txt"            # buat file ini di folder script untuk stop halus
AUTOSAVE_EVERY_ROWS = 10          # autosave tiap N baris
AUTOSAVE_EVERY_SEC = 60           # atau tiap N detik (mana yang tercapai dulu)
AUTOSAVE_KEEP_COPY = True         # simpan juga file .autosave.xlsx
SCREENSHOT_DIR = "debug_screens"  # folder screenshot error

_stop_requested = False

def request_stop(reason=""):
    global _stop_requested
    _stop_requested = True
    if reason:
        print(f"\n🛑 STOP requested: {reason}", flush=True)

def _on_signal(sig, frame):
    # Ctrl+C / kill -> stop aman
    request_stop(f"signal={sig}")

signal.signal(signal.SIGINT, _on_signal)
try:
    signal.signal(signal.SIGTERM, _on_signal)
except Exception:
    pass

def should_stop() -> bool:
    if _stop_requested:
        return True
    try:
        if os.path.exists(STOP_FILE):
            request_stop(f"found {STOP_FILE}")
            return True
    except Exception:
        pass
    return False

def ensure_dir(path: str):
    try:
        os.makedirs(path, exist_ok=True)
    except Exception:
        pass

def safe_save_excel(df: pd.DataFrame, file_path: str, tag: str = ""):
    """
    Save aman: tulis ke tmp, lalu replace.
    Mengurangi risiko file corrupt kalau proses berhenti/PC mati.
    """
    try:
        base, ext = os.path.splitext(file_path)
        tmp_path = base + ".__tmp__" + (ext or ".xlsx")
        df.to_excel(tmp_path, index=False)
        os.replace(tmp_path, file_path)

        if AUTOSAVE_KEEP_COPY:
            bak_path = base + ".autosave" + (ext or ".xlsx")
            try:
                df.to_excel(bak_path, index=False)
            except Exception:
                pass

        if tag:
            print(f"💾 Saved {tag} -> {file_path}", flush=True)
        else:
            print(f"💾 Saved -> {file_path}", flush=True)
    except Exception as e:
        print(f"⚠️ Gagal save ({tag}): {e}", flush=True)

# =========================
# Denpasar bbox filter (WAJIB)
# =========================
DENPASAR_BBOX = {
    "lat_min": -8.752640565327079,
    "lat_max": -8.592768863988631,
    "lon_min": 115.17372541407495,
    "lon_max": 115.27485445343008,
}

def is_within_bbox(lat, lon, bbox=DENPASAR_BBOX):
    try:
        if lat is None or lon is None:
            return False
        lat = float(lat)
        lon = float(lon)
        return (bbox["lat_min"] <= lat <= bbox["lat_max"]) and (bbox["lon_min"] <= lon <= bbox["lon_max"])
    except Exception:
        return False

# =========================
# Helper: safe cell (NaN -> "")
# =========================
def s_cell(v) -> str:
    try:
        if v is None:
            return ""
        if pd.isna(v):
            return ""
        sv = str(v).strip()
        if sv.lower() == "nan":
            return ""
        return sv
    except Exception:
        return ""

# =========================
# Helper: browser
# =========================
def wait_document_ready(driver, timeout=12):
    WebDriverWait(driver, timeout).until(
        lambda d: d.execute_script("return document.readyState") in ("interactive", "complete")
    )

def click_consent_if_any(driver, timeout=2):
    xpaths = [
        "//button//*[contains(.,'Accept')]/ancestor::button[1]",
        "//button//*[contains(.,'I agree')]/ancestor::button[1]",
        "//button//*[contains(.,'Setuju')]/ancestor::button[1]",
        "//button//*[contains(.,'Terima')]/ancestor::button[1]",
        "//button//*[contains(.,'AGREE')]/ancestor::button[1]",
    ]
    for xp in xpaths:
        try:
            btn = WebDriverWait(driver, timeout).until(EC.element_to_be_clickable((By.XPATH, xp)))
            btn.click()
            time.sleep(0.15)
            return True
        except TimeoutException:
            pass
    return False

def safe_text(el):
    try:
        return el.text.strip()
    except Exception:
        return None

def open_home(driver):
    driver.get("https://www.google.com/maps")
    wait_document_ready(driver, 12)
    click_consent_if_any(driver, timeout=2)

# =========================
# Parsing coords (jangan ambil @latlon dari /maps/search)
# =========================
def _to_float(x):
    try:
        if x is None:
            return None
        return float(str(x).strip())
    except Exception:
        return None

def parse_coords_from_url(url: str):
    """
    Valid coords:
    - PRIORITAS: !3dLAT!4dLON
    - Jika URL /maps/place dan ada /@lat,lon => boleh
    - Jika /maps/search => @lat,lon viewport (JANGAN) => None
    """
    if not url:
        return None, None

    u = url.lower()

    m = re.search(r"!3d(-?\d+(?:\.\d+)?)!4d(-?\d+(?:\.\d+)?)", url)
    if m:
        return _to_float(m.group(1)), _to_float(m.group(2))

    if "/maps/place" in u and "/@" in url:
        try:
            part = url.split("/@")[1].split(",")
            return _to_float(part[0]), _to_float(part[1])
        except Exception:
            return None, None

    return None, None

# =========================
# Normalisasi teks
# =========================
def clean_text(s: str) -> str:
    s = str(s or "")
    s = s.replace("<", " ").replace(">", " ")
    s = re.sub(r"\bRT\s*\d+\/?\s*RW\s*\d+\b", " ", s, flags=re.I)
    s = re.sub(r"\bRT\s*\d+\b", " ", s, flags=re.I)
    s = re.sub(r"\bRW\s*\d+\b", " ", s, flags=re.I)
    s = re.sub(r"\s+", " ", s).strip()
    return s

# FIX penting: pecah "JalanIMAM" -> "Jalan IMAM", "No486A" -> "No 486 A"
def _split_stuck_words(s: str) -> str:
    if not s:
        return ""
    s = re.sub(r"([a-z])([A-Z])", r"\1 \2", s)
    s = re.sub(r"([A-Za-z])(\d)", r"\1 \2", s)
    s = re.sub(r"(\d)([A-Za-z])", r"\1 \2", s)
    return s
    

def compact_addr_for_query(addr: str) -> str:
    """
    Buat alamat ringkas untuk query:
    - Ambil 'Jalan ...' + nomor (kalau ada)
    - Buang token yang sering bikin Maps bingung
    """
    a = normalize_addr(addr or "")
    if not a:
        return ""

    low = a.lower()

    # Ambil bagian setelah "Jalan ..." kalau ada
    m = re.search(r"\b(jalan\s+[a-z0-9\s\-\.]{5,})", a, flags=re.I)
    base = m.group(1).strip() if m else a

    # Potong setelah koma kedua biar ringkas
    parts = [p.strip() for p in re.split(r"[,\|]", base) if p.strip()]
    base2 = ", ".join(parts[:2]) if parts else base

    # Cari "No ..." kalau ada dan belum masuk
    m2 = re.search(r"\b(no\.?|nomor)\s*([0-9]{1,4}\s*[a-z]?)\b", low, flags=re.I)
    if m2:
        no_txt = f"No {m2.group(2).strip().upper()}"
        if no_txt.lower() not in base2.lower():
            base2 = f"{base2} {no_txt}"

    # rapikan spasi
    base2 = re.sub(r"\s+", " ", base2).strip()
    return base2


def build_queries_adaptive(nama_in, alamat_in_raw, kec_in, city_context):
    """
    Generate query sedikit tapi efektif (urut dari paling informatif ke fallback).
    Output list query unik, max 4-5.
    """
    nama = clean_text(nama_in)
    kec = clean_text(kec_in)
    addr_compact = compact_addr_for_query(alamat_in_raw)

    # Varian wilayah
    kec_part = f", {kec}" if kec else ""

    q = []
    # 1) Nama + alamat ringkas + kec + city
    if nama and addr_compact:
        q.append(clean_text(f"{nama}, {addr_compact}{kec_part}, {city_context}"))

    # 2) Nama + kec + city (sering cepat & cukup)
    if nama:
        q.append(clean_text(f"{nama}{kec_part}, {city_context}"))

    # 3) Alamat ringkas + nama (dibalik) (kadang ranking beda)
    if nama and addr_compact:
        q.append(clean_text(f"{addr_compact}, {nama}{kec_part}, {city_context}"))

    # 4) Nama + city saja (fallback)
    if nama:
        q.append(clean_text(f"{nama}, {city_context}"))

    # unique & buang kosong
    seen = set()
    out = []
    for item in q:
        item = " ".join((item or "").split()).strip()
        if not item:
            continue
        key = item.lower()
        if key in seen:
            continue
        seen.add(key)
        out.append(item)

    return out[:4]


def normalize_addr(addr: str) -> str:
    a = clean_text(addr)
    a = _split_stuck_words(a)

    a = re.sub(r"\bJl\.?\b", "Jalan", a, flags=re.I)
    a = re.sub(r"\bJln\.?\b", "Jalan", a, flags=re.I)
    a = re.sub(r"\bGg\.?\b", "Gang", a, flags=re.I)
    a = re.sub(r"\bBr\.?\b", "Banjar", a, flags=re.I)
    a = re.sub(r"\bDs\.?\b", "Desa", a, flags=re.I)
    a = re.sub(r"\bKel\.?\b", "Kelurahan", a, flags=re.I)

    a = re.sub(r"\s+", " ", a).strip()
    return a

STOP_WORDS = {
    "jalan", "gang", "banjar", "br", "dk", "dusun",
    "denpasar", "bali", "indonesia",
    "kecamatan", "kec", "kelurahan", "kel", "desa", "rt", "rw",
    "kota", "kab", "kabupaten", "prov", "provinsi",
    "jl", "jln", "gg", "no", "nomor", "nmr",
    "gn", "gunung",
    "blok", "block", "lantai", "lt",
    "komplek", "kompleks", "kompleksnya", "kompleksperum",
    "perum", "perumahan",
    "ruko", "kav", "kavling",
    "km", "meter",
    "ggg",
}

GENERIC_NAME = {
    "warung", "toko", "ud", "cv", "pt", "resto", "restaurant", "cafe", "kedai", "depot",
    "laundry", "salon", "barber", "bengkel", "apotek", "klinik", "clinic", "fotocopy", "foto",
    "mart", "mini", "market", "shop", "store", "service", "jasa", "hasil",
    "sewa", "kost", "kos", "kontrakan", "homestay", "guesthouse", "villa", "hotel"
}

NAME_STOP_EXTRA = {
    "koperasi", "kpn", "ksp", "ksu", "kud", "lpd",
    "yayasan", "perkumpulan", "asosiasi",
    "serba", "usaha", "konsumen",
}

NAME_NOISE_PATTERNS = [
    r"\(.*?\)",
    r"\b(persero|tbk|t\.bk)\b",
    r"\b(denpasar|bali|indonesia)\b",
    r"\b(kantor|office|cabang|unit|pusat)\b.*$",
    r"\b(pemerintah|pemkot|pemkab)\b.*$",
]

def normalize_name(s: str) -> str:
    s = clean_text(s or "")
    s = _split_stuck_words(s)
    s = s.replace("&", " dan ")
    s = re.sub(r"[/,_\-]+", " ", s)
    s = re.sub(r"\s+", " ", s).strip()
    low = s.lower()
    for pat in NAME_NOISE_PATTERNS:
        low = re.sub(pat, " ", low, flags=re.I)
        low = re.sub(r"\s+", " ", low).strip()
    return low.strip()

def name_tokens2(s: str):
    s = normalize_name(s)
    s = re.sub(r"[^a-z0-9\s]", " ", s)
    toks = [t for t in s.split() if len(t) >= 2]
    out = set()
    for t in toks:
        if t in STOP_WORDS:
            continue
        if t in GENERIC_NAME:
            continue
        if t in NAME_STOP_EXTRA:
            continue
        out.add(t)
    return out

def addr_tokens(addr: str):
    a = normalize_addr(addr or "").lower()
    a = re.sub(r"[^a-z0-9\s]", " ", a)
    raw = [t for t in a.split() if t]
    out = set()
    for t in raw:
        if t in STOP_WORDS:
            continue
        if re.fullmatch(r"\d{1,4}", t):
            out.add(t)
            continue
        if re.fullmatch(r"\d{1,4}[a-z]{1,2}", t):
            out.add(t)
            continue
        if re.fullmatch(r"[a-z]{2,}", t):
            out.add(t)
            continue
        if re.fullmatch(r"[ivxlcdm]{2,}", t):
            out.add(t)
            continue
    return out

def addr_alpha_tokens(addr: str):
    a = normalize_addr(addr or "").lower()
    a = re.sub(r"[^a-z0-9\s]", " ", a)
    toks = [t for t in a.split() if t and t not in STOP_WORDS and len(t) >= 3]
    return {t for t in toks if re.fullmatch(r"[a-z]{3,}", t)}

def jaccard(a: set, b: set) -> float:
    if not a or not b:
        return 0.0
    return len(a & b) / max(1, len(a | b))

def fuzzy_ratio(a: str, b: str) -> float:
    a = (a or "").strip().lower()
    b = (b or "").strip().lower()
    if not a or not b:
        return 0.0
    return SequenceMatcher(None, a, b).ratio()

def strip_loc_words(s: str) -> str:
    s = (s or "").lower()
    s = re.sub(r"[^a-z0-9\s]", " ", s)
    s = " ".join([t for t in s.split() if t not in STOP_WORDS])
    return s.strip()

def acronym_of_words(s: str) -> str:
    toks = [t for t in re.sub(r"[^a-z0-9\s]", " ", (s or "").lower()).split() if len(t) >= 2]
    toks = [t for t in toks if t not in GENERIC_NAME and t not in STOP_WORDS]
    if not toks:
        return ""
    return "".join(t[0] for t in toks[:8]).upper()

def abbrev_input(s: str) -> str:
    s = (s or "").strip().upper()
    first = re.split(r"\s+", s)[0] if s else ""
    if 2 <= len(first) <= 6 and re.fullmatch(r"[A-Z0-9]+", first or ""):
        return first
    caps = re.sub(r"[^A-Z0-9]", "", s)
    if 2 <= len(caps) <= 6:
        return caps
    return ""

# =========================
# Infer kecamatan
# =========================
DENPASAR_KEC = [
    "Denpasar Selatan", "Denpasar Timur", "Denpasar Barat", "Denpasar Utara"
]

def extract_kec_from_gmaps(alamat_gmaps: str) -> str:
    a_low = (alamat_gmaps or "").lower()
    m = re.search(r"kec\.\s*([a-z\s]+)", a_low, flags=re.I)
    if m:
        guess = m.group(1).strip()
        guess = re.sub(r"[,].*$", "", guess).strip()
        guess = " ".join(guess.split()[:3])
        return guess.title()

    m = re.search(r"kecamatan\s+([a-z\s]+)", a_low, flags=re.I)
    if m:
        guess = m.group(1).strip()
        guess = re.sub(r"[,].*$", "", guess).strip()
        guess = " ".join(guess.split()[:3])
        return guess.title()

    for k in DENPASAR_KEC:
        if k.lower() in a_low:
            return k
    return ""

# =========================
# kualitas kandidat (echo query / generik)
# =========================
def looks_like_query_echo(nama_detail: str, query_used: str, city_context: str) -> bool:
    n = (nama_detail or "").strip().lower()
    q = (query_used or "").strip().lower()
    c = (city_context or "").strip().lower()
    if not n:
        return False

    if c and c in n and n.count(",") >= 2:
        if q and fuzzy_ratio(n, q) >= 0.78:
            return True
        if q and (q[: max(12, int(len(q) * 0.55))] in n):
            return True

    if q and fuzzy_ratio(n, q) >= 0.90:
        return True

    return False

COMMON_SUFFIX = {
    "jaya", "abadi", "makmur", "sentosa", "sejahtera", "berkah", "mulya", "utama", "sukses",
    "prima", "indo", "nusantara", "mandiri", "karya", "agung", "barokah", "barakah",
}

LEGAL_ENTITY_TOKENS = {"ud", "cv", "pt", "tbk", "t bk", "t.bk", "persero"}

def is_too_generic_name(nama: str) -> bool:
    n = normalize_name(nama or "")
    if not n:
        return True

    toks = [t for t in re.sub(r"[^a-z0-9\s]", " ", n.lower()).split() if t and t not in STOP_WORDS]
    toks_wo_legal = [t for t in toks if t not in LEGAL_ENTITY_TOKENS]

    if len(toks_wo_legal) <= 1:
        return True

    if len(toks_wo_legal) == 2 and (toks_wo_legal[0] in COMMON_SUFFIX or toks_wo_legal[1] in COMMON_SUFFIX):
        return True

    gen_hits = sum(1 for t in toks_wo_legal if (t in GENERIC_NAME or t in COMMON_SUFFIX))
    if gen_hits >= max(1, len(toks_wo_legal) - 1):
        return True

    return False

def is_generic_place_name(nama_g: str) -> bool:
    n = (nama_g or "").strip().lower()
    if not n:
        return True
    if n in {"hasil", "result", "results"}:
        return True
    if ("sewa" in n and ("kos" in n or "kost" in n)) and "denpasar" in n:
        return True
    if is_too_generic_name(nama_g):
        return True
    return False

def coords_only_guard_ok(best_name: str, dbg: dict) -> bool:
    if not dbg:
        return False
    if dbg.get("is_echo"):
        return False
    if dbg.get("is_generic"):
        return False

    ov_name = int(dbg.get("ov_name", 0) or 0)
    s_name = float(dbg.get("s_name", 0.0) or 0.0)
    s_fuz = float(dbg.get("s_name_fuzzy", 0.0) or 0.0)

    return (ov_name >= 1) or (s_name >= 0.50) or (s_fuz >= 0.55)

# =========================
# scoring helpers
# =========================
def containment_score(a: str, b: str) -> float:
    a2 = normalize_name(a)
    b2 = normalize_name(b)
    if not a2 or not b2:
        return 0.0
    if a2 == b2:
        return 1.0
    if a2 in b2:
        return min(1.0, len(a2) / max(1, len(b2)))
    return 0.0

def soft_token_overlap(a_set: set, b_set: set, sim_thr: float = 0.88) -> int:
    if not a_set or not b_set:
        return 0
    b_list = list(b_set)
    cnt = 0
    for ta in a_set:
        if ta in b_set:
            cnt += 1
            continue
        best = 0.0
        for tb in b_list:
            r = SequenceMatcher(None, ta, tb).ratio()
            if r > best:
                best = r
            if best >= sim_thr:
                break
        if best >= sim_thr:
            cnt += 1
    return cnt

def soft_jaccard(a_set: set, b_set: set, sim_thr: float = 0.88) -> float:
    if not a_set or not b_set:
        return 0.0
    inter = soft_token_overlap(a_set, b_set, sim_thr=sim_thr)
    union = len(a_set) + len(b_set) - inter
    return inter / max(1, union)

# =========================
# Scoring
# =========================
def score_candidate(nama_in, alamat_in, kec_in, nama_g, alamat_g, *, is_echo=False, is_generic=False):
    n_in = name_tokens2(nama_in)
    n_g = name_tokens2(nama_g)

    s_name_tok = jaccard(n_in, n_g)
    s_name_soft = soft_jaccard(n_in, n_g, sim_thr=0.88)
    s_name_fuz = fuzzy_ratio(strip_loc_words(normalize_name(nama_in)), strip_loc_words(normalize_name(nama_g)))
    s_name_cont = containment_score(nama_in, nama_g)

    abv = abbrev_input(nama_in)
    acr_g = acronym_of_words(nama_g)
    s_abbrev = 1.0 if (abv and acr_g and abv == acr_g) else 0.0

    s_name = max(s_name_tok, s_name_soft, s_name_fuz, s_name_cont, s_abbrev)
    ov_name = soft_token_overlap(n_in, n_g, sim_thr=0.88)

    a_in = addr_tokens(alamat_in)
    a_g = addr_tokens(alamat_g)
    s_addr = jaccard(a_in, a_g)
    ov_addr = len(a_in & a_g)

    ag_low = (alamat_g or "").lower()
    kec_raw = (kec_in or "").strip()
    kec_eff = (kec_raw or extract_kec_from_gmaps(alamat_g)).strip().lower()

    bonus = 0.0
    penalty = 0.0

    if "denpasar" in ag_low:
        bonus += 0.05

    if kec_eff:
        if kec_eff in ag_low:
            bonus += 0.06
        else:
            penalty -= 0.03

    if ov_addr >= 4:
        bonus += 0.18
    elif ov_addr == 3:
        bonus += 0.14
    elif ov_addr == 2:
        bonus += 0.10
    elif ov_addr == 1:
        bonus += 0.05

    if ov_name >= 3:
        bonus += 0.09
    elif ov_name == 2:
        bonus += 0.07
    elif ov_name == 1:
        bonus += 0.03

    if not (alamat_g or "").strip():
        penalty -= (0.08 if s_name >= 0.78 else 0.18)

    if is_echo:
        penalty -= 0.35
    if is_generic:
        penalty -= 0.18

    addr_in_weak = len(a_in) < 2
    if addr_in_weak:
        w_name, w_addr = 0.78, 0.22
    else:
        if s_name >= 0.75:
            w_name, w_addr = 0.60, 0.40
        else:
            w_name, w_addr = 0.35, 0.65

    score = (w_name * s_name) + (w_addr * s_addr) + bonus + penalty
    score = max(0.0, min(1.2, score))

    dbg = {
        "s_name": s_name,
        "s_name_tok": s_name_tok,
        "s_name_soft": s_name_soft,
        "s_name_fuzzy": s_name_fuz,
        "s_name_cont": s_name_cont,
        "s_abbrev": s_abbrev,
        "s_addr": s_addr,
        "ov_addr": ov_addr,
        "ov_name": ov_name,
        "kec_eff": kec_eff,
        "bonus": bonus,
        "penalty": penalty,
        "addr_in_weak": addr_in_weak,
        "w_name": w_name,
        "w_addr": w_addr,
        "abv": abv,
        "acr_g": acr_g,
        "is_echo": is_echo,
        "is_generic": is_generic,
    }
    return score, dbg

# =========================
# Search helpers (paksa buka place)
# =========================

def should_early_stop(best, threshold, bbox=DENPASAR_BBOX):
    if best["score"] < threshold:
        return False
    lat, lon = best.get("lat"), best.get("lon")
    if lat is None or lon is None:
        return False
    if not is_within_bbox(lat, lon, bbox=bbox):
        return False
    dbg = best.get("dbg") or {}
    if dbg.get("is_echo"):
        return False

    strong_name = (float(dbg.get("s_name", 0) or 0) >= 0.82) or (float(dbg.get("s_name_fuzzy", 0) or 0) >= 0.85)
    strong_addr = (int(dbg.get("ov_addr", 0) or 0) >= 3) or (float(dbg.get("s_addr", 0) or 0) >= 0.28)
    return strong_name or strong_addr


def get_list_candidates_fast(driver, limit=12):
    """
    Ambil kandidat dari list mode tanpa klik/buka detail dulu.
    Return list of dict: {href, name_hint, sub_hint}
    """
    out = []
    try:
        links = driver.find_elements(By.CSS_SELECTOR, "a.hfpxzc")
        for a in links[: max(limit, 1)]:
            href = None
            try:
                href = a.get_attribute("href")
            except Exception:
                href = None
            if not href:
                continue

            # Nama biasanya ada di aria-label
            name_hint = ""
            try:
                name_hint = (a.get_attribute("aria-label") or "").strip()
            except Exception:
                pass

            # Coba ambil “teks sekunder” dari container card terdekat
            # Ini agak dinamis, jadi kita buat best-effort, bukan wajib.
            sub_hint = ""
            try:
                card = a.find_element(By.XPATH, "./ancestor::div[contains(@role,'article') or contains(@class,'Nv2PK')][1]")
                # ambil beberapa teks yang mungkin berisi area/alamat singkat
                texts = []
                for sel in [
                    "div.W4Efsd",  # sering memuat meta
                    "div.W4Efsd span",
                    "div.qBF1Pd",  # variasi lain
                    "div.fontBodyMedium",
                ]:
                    try:
                        els = card.find_elements(By.CSS_SELECTOR, sel)
                        for e in els[:6]:
                            t = (e.text or "").strip()
                            if t and t.lower() not in {"hasil", "result", "results"}:
                                texts.append(t)
                    except Exception:
                        pass
                # gabungkan singkat saja
                sub_hint = " | ".join(list(dict.fromkeys(texts))[:3]).strip()
            except Exception:
                pass

            out.append({"href": href, "name_hint": name_hint, "sub_hint": sub_hint})
    except Exception:
        pass
    return out


def quick_score_from_list(nama_in, alamat_in, kec_in, cand_name, cand_sub):
    """
    Scoring cepat berbasis hint list (nama + sub text).
    Ini bukan final, hanya untuk ranking top-k agar hemat waktu.
    """
    # anggap cand_sub sebagai “alamat kasar” (best-effort)
    sc, dbg = score_candidate(
        nama_in, alamat_in, kec_in,
        cand_name or "", cand_sub or "",
        is_echo=False, is_generic=is_generic_place_name(cand_name or "")
    )
    # sedikiit bonus kalau subtext menyebut Denpasar / kecamatan
    low = (cand_sub or "").lower()
    bonus = 0.0
    if "denpasar" in low:
        bonus += 0.04
    if (kec_in or "").strip().lower() and (kec_in or "").strip().lower() in low:
        bonus += 0.05

    sc2 = max(0.0, min(1.2, sc + bonus))
    return sc2, dbg

def force_open_place_details(driver, timeout=8) -> bool:
    try:
        u = (driver.current_url or "").lower()
        if "/maps/place" in u:
            return True

        if "/maps/search" in u:
            WebDriverWait(driver, timeout).until(lambda d: len(d.find_elements(By.CSS_SELECTOR, "a.hfpxzc")) > 0)
            a = driver.find_elements(By.CSS_SELECTOR, "a.hfpxzc")[0]
            href = a.get_attribute("href")
            if href:
                driver.get(href)
            else:
                a.click()

            wait_document_ready(driver, 12)
            click_consent_if_any(driver, timeout=1)

            WebDriverWait(driver, timeout).until(
                lambda d: ("/maps/place" in (d.current_url or "").lower()) or len(d.find_elements(By.XPATH, "//h1")) > 0
            )
            return "/maps/place" in (driver.current_url or "").lower()

        WebDriverWait(driver, timeout).until(lambda d: len(d.find_elements(By.XPATH, "//h1")) > 0)
        return "/maps/place" in (driver.current_url or "").lower()

    except Exception:
        return False

def run_query_via_url(driver, query, timeout=18):
    q = " ".join(str(query).split()).strip()
    if not q:
        # penting: jangan lempar driver ke query kosong
        return ""

    url = "https://www.google.com/maps/search/?api=1&query=" + quote_plus(q)

    driver.get(url)
    wait_document_ready(driver, 12)
    click_consent_if_any(driver, timeout=1)

    def cond(d):
        u = d.current_url or ""
        if ("!3d" in u and "!4d" in u):
            return True
        if len(d.find_elements(By.CSS_SELECTOR, "a.hfpxzc")) > 0:
            return True
        if len(d.find_elements(By.XPATH, "//h1")) > 0:
            return True
        return False

    WebDriverWait(driver, timeout).until(cond)
    return url

def partial_match_detected(driver):
    try:
        return len(driver.find_elements(By.CSS_SELECTOR, "div.L5xkq.Hk4XGb")) > 0
    except Exception:
        return False

def wait_place_panel_ready(driver, timeout=8) -> bool:
    end = time.time() + timeout
    while time.time() < end:
        try:
            u = (driver.current_url or "").lower()
            if "/maps/place" not in u:
                time.sleep(0.15)
                continue

            h1s = driver.find_elements(By.XPATH, "//h1")
            for h in h1s[:2]:
                try:
                    t = (h.text or "").strip()
                    if t and t.lower() not in {"hasil", "result", "results"}:
                        return True
                except Exception:
                    pass

            t = (driver.title or "").strip()
            if t and " - " in t:
                base = t.split(" - ")[0].strip()
                if base and base.lower() not in {"hasil", "result", "results"}:
                    return True
        except Exception:
            pass

        time.sleep(0.2)
    return False

def get_place_title(driver, timeout=6):
    try:
        wait_place_panel_ready(driver, timeout=timeout)
    except Exception:
        pass

    try:
        h1s = driver.find_elements(By.XPATH, '//h1[contains(@class,"DUwDvf")] | //h1')
        for el in h1s[:3]:
            t = (el.text or "").strip()
            if t and t.lower() not in {"hasil", "result", "results"}:
                return t
    except Exception:
        pass

    try:
        t = (driver.title or "").strip()
        if " - " in t:
            t = t.split(" - ")[0].strip()
        if t and t.lower() not in {"hasil", "result", "results"}:
            return t
    except Exception:
        pass

    return ""

PLUS_CODE_RE = re.compile(r"\b[23456789CFGHJMPQRVWX]{4,8}\+[23456789CFGHJMPQRVWX]{2,4}\b", re.I)

def _clean_gmaps_address_text(t: str) -> str:
    t = (t or "").strip()
    if not t:
        return ""
    if t.strip().lower() in {"alamat", "address"}:
        return ""
    t = PLUS_CODE_RE.sub(" ", t)
    t = re.sub(r"\s+", " ", t).strip()
    return t

def get_address(driver, timeout=2):
    try:
        els = driver.find_elements(By.CSS_SELECTOR, "div.Io6YTe.fontBodyMedium.kR99db.fdkmkc")
        for e in els:
            t = _clean_gmaps_address_text(e.text)
            if t:
                return t
    except Exception:
        pass

    try:
        els = driver.find_elements(By.CSS_SELECTOR, '[data-item-id*="address"]')
        for e in els[:6]:
            t = _clean_gmaps_address_text(e.text)
            if t:
                return t
            aria = (e.get_attribute("aria-label") or "").strip()
            if aria:
                aria = re.sub(r"^(alamat|address)\s*:\s*", "", aria, flags=re.I).strip()
                aria = _clean_gmaps_address_text(aria)
                if aria:
                    return aria
    except Exception:
        pass

    try:
        els = driver.find_elements(By.XPATH, '//*[@aria-label[contains(., "Alamat") or contains(., "Address")]]')
        for e in els[:8]:
            aria = (e.get_attribute("aria-label") or "").strip()
            if aria:
                aria = re.sub(r"^(alamat|address)\s*:\s*", "", aria, flags=re.I).strip()
                aria = _clean_gmaps_address_text(aria)
                if aria:
                    return aria
    except Exception:
        pass

    try:
        html = driver.page_source or ""
        m = re.search(
            r'(Alamat|Address)\\u003c\\/span\\u003e\\s*\\u003cspan[^\\>]*\\u003e([^\\<]{10,200})\\u003c',
            html,
            flags=re.I,
        )
        if m:
            t = _clean_gmaps_address_text(m.group(2))
            if t:
                return t
    except Exception:
        pass

    return ""

def get_phone(driver):
    try:
        btns = driver.find_elements(
            By.XPATH,
            '//button[contains(@aria-label,"Telepon") or contains(@aria-label,"telepon") '
            'or contains(@aria-label,"Phone") or contains(@aria-label,"phone")]'
        )
        for b in btns:
            t = b.text.strip()
            if t:
                return t
    except Exception:
        pass
    return None

# =========================
# Closed detection
# =========================
CLOSED_PATTERNS = [
    r"\bpermanently closed\b",
    r"\btemporarily closed\b",
    r"\bclosed permanently\b",
    r"\bclosed temporarily\b",
    r"\btutup permanen\b",
    r"\bditutup permanen\b",
    r"\btutup sementara\b",
    r"\bditutup sementara\b",
    r"\bsecara permanen ditutup\b",
]

def detect_closed_status(driver):
    try:
        panel_text = ""
        candidates = []
        candidates += driver.find_elements(By.CSS_SELECTOR, "div.UGUb2e")
        candidates += driver.find_elements(By.CSS_SELECTOR, "div.fontBodyMedium")
        candidates += driver.find_elements(By.CSS_SELECTOR, "div.rogA2c")
        candidates += driver.find_elements(By.CSS_SELECTOR, "div[role='main']")

        for el in candidates[:6]:
            t = safe_text(el)
            if t:
                panel_text += "\n" + t

        text = (panel_text.strip() or driver.page_source).lower()

        for pat in CLOSED_PATTERNS:
            if re.search(pat, text, flags=re.I):
                if "temporary" in pat or "sementara" in pat:
                    return True, "temporary"
                if "permanent" in pat or "permanen" in pat:
                    return True, "permanent"
                return True, "unknown"
        return False, None
    except Exception:
        return False, None

# =========================
# Output mapping
# =========================
def apply_gc_fields(df, idx, status_kode, nama_usaha, alamat_usaha, lat, lon):
    if status_kode == 99:
        lat = None
        lon = None

    latlong_status = "valid" if (lat is not None and lon is not None) else "invalid"

    df.at[idx, "latitude"] = lat
    df.at[idx, "longitude"] = lon
    df.at[idx, "latlong_status"] = latlong_status

    df.at[idx, "gcs_result"] = status_kode
    df.at[idx, "latitude_gc"] = lat
    df.at[idx, "longitude_gc"] = lon
    df.at[idx, "latlong_status_gc"] = latlong_status

    df.at[idx, "nama_usaha_gc"] = nama_usaha
    df.at[idx, "alamat_usaha_gc"] = alamat_usaha

    df.at[idx, "hasilgc"] = status_kode

# =========================
# MAIN
# =========================
file_path = "test.xlsx"
df = pd.read_excel(file_path)

# FIX: pastikan kolom input yang dipakai memang ada
input_cols = ["nama_usaha", "alamat_usaha", "nmkec"]
for c in input_cols:
    if c not in df.columns:
        df[c] = ""

needed_cols = [
    "nama_gmaps", "alamat_gmaps", "nomor_telepon",
    "latitude", "longitude",
    "keterangan", "score_match",
    "status_bisnis", "status_kode", "status_tutup",
    "latlong_status", "gcs_result", "latitude_gc", "longitude_gc", "latlong_status_gc",
    "nama_usaha_gc", "alamat_usaha_gc", "hasilgc"
]
for col in needed_cols:
    if col not in df.columns:
        df[col] = pd.NA

# =========================
# FIX KRUSIAL: paksa dtype kolom output (hindari float64 -> string error)
# =========================
TEXT_COLS = [
    "nama_gmaps", "alamat_gmaps", "nomor_telepon",
    "keterangan", "status_bisnis", "status_tutup",
    "nama_usaha_gc", "alamat_usaha_gc", "latlong_status", "latlong_status_gc",
]
for c in TEXT_COLS:
    df[c] = df[c].astype("string")

NUM_COLS = ["latitude", "longitude", "latitude_gc", "longitude_gc", "score_match"]
for c in NUM_COLS:
    df[c] = pd.to_numeric(df[c], errors="coerce")

INT_COLS = ["status_kode", "gcs_result", "hasilgc"]
for c in INT_COLS:
    df[c] = pd.to_numeric(df[c], errors="coerce").astype("Int64")

# === Chrome options ===
options = webdriver.ChromeOptions()
options.page_load_strategy = "eager"
options.add_argument("--start-maximized")
options.add_argument("--log-level=3")
options.add_argument("--silent")
options.add_experimental_option("excludeSwitches", ["enable-logging", "enable-automation"])
options.add_argument("--disable-gpu")
options.add_argument("--disable-software-rasterizer")
options.add_argument("--disable-features=DirectComposition,UseSkiaRenderer")
options.add_argument("--no-sandbox")
options.add_argument("--disable-dev-shm-usage")
options.add_argument("--disable-notifications")
options.add_argument("--disable-popup-blocking")
options.add_argument("--lang=id-ID")

prefs = {
    "profile.managed_default_content_settings.images": 1,
    "profile.default_content_setting_values.notifications": 2,
    "profile.default_content_setting_values.geolocation": 2,
}
options.add_experimental_option("prefs", prefs)

options.add_argument(
    "user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
    "(KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
)
options.add_experimental_option("useAutomationExtension", False)

service = Service(ChromeDriverManager().install())
log_fh = None
try:
    log_fh = open(os.devnull, "w")
    service.log_output = log_fh
except Exception:
    pass

driver = webdriver.Chrome(service=service, options=options)
driver.implicitly_wait(0.4)

MAX_CANDIDATES = 10
THRESHOLD_OK = 0.45
THRESHOLD_EARLY_STOP = 0.70
CITY_CONTEXT = "Denpasar, Bali, Indonesia"
MAX_RETRY = 0

ALLOW_COORDS_ONLY_MATCH = True

ensure_dir(SCREENSHOT_DIR)

try:
    open_home(driver)

    _last_save_ts = time.time()
    total_rows = len(df)

    for idx, row in df.iterrows():
        # ---- stop check (STOP.txt / Ctrl+C) ----
        if should_stop():
            print(f"\n🛑 Berhenti aman di baris {idx}/{total_rows}.", flush=True)
            break

        # ---- autosave check ----
        now = time.time()
        if (idx > 0 and idx % AUTOSAVE_EVERY_ROWS == 0) or ((now - _last_save_ts) >= AUTOSAVE_EVERY_SEC):
            safe_save_excel(df, file_path, tag=f"(autosave row {idx})")
            _last_save_ts = now

        nama_usaha_raw = s_cell(row.get("nama_usaha"))
        alamat_usaha_raw = s_cell(row.get("alamat_usaha"))
        kec_in_raw = s_cell(row.get("nmkec"))

        # kalau input nama kosong total, skip cepat (menghindari query aneh)
        if not clean_text(nama_usaha_raw):
            df.at[idx, "keterangan"] = "Skip: nama_usaha kosong"
            df.at[idx, "status_bisnis"] = "Tidak ditemukan"
            df.at[idx, "status_kode"] = 99
            apply_gc_fields(df, idx, 99, nama_usaha_raw, alamat_usaha_raw, None, None)
            continue

        lat_existing = row.get("latitude")
        lon_existing = row.get("longitude")
        if pd.notnull(lat_existing) and pd.notnull(lon_existing):
            apply_gc_fields(df, idx, 1, nama_usaha_raw, alamat_usaha_raw, float(lat_existing), float(lon_existing))
            continue

        nama_in = clean_text(nama_usaha_raw)
        alamat_in = normalize_addr(alamat_usaha_raw)
        kec_in = clean_text(kec_in_raw)

        # kec_part = f", {kec_in}" if kec_in.strip() else ""
        # q_full = clean_text(f"{nama_in}, {alamat_in}{kec_part}, {CITY_CONTEXT}") if alamat_in.strip() else ""
        # q_name = clean_text(f"{nama_in}{kec_part}, {CITY_CONTEXT}")
        # queries = [q for q in [q_full, q_name] if q.strip()]

        queries = build_queries_adaptive(nama_in, alamat_usaha_raw, kec_in, CITY_CONTEXT)
        if not queries:
            queries = [clean_text(f"{nama_in}, {CITY_CONTEXT}")]


        print(f"\n🔍 Baris {idx} | mulai", flush=True)

        best = {
            "score": -1.0,
            "dbg": None,
            "nama": None,
            "alamat": None,
            "phone": None,
            "lat": None,
            "lon": None,
            "is_closed": False,
            "closed_type": None,
            "source": None,
        }

        try:
            # jika queries kosong (misal alamat kosong & city context somehow kosong) -> fallback minimal
            if not queries:
                queries = [clean_text(f"{nama_in}, {CITY_CONTEXT}")]

            stop_queries = False
            for q in queries:
                if should_stop():
                    print(f"\n🛑 Stop saat proses baris {idx}.", flush=True)
                    break

                print(f"   ▶ query: {q}", flush=True)

                last_search_url = None
                for attempt in range(MAX_RETRY + 1):
                    try:
                        last_search_url = run_query_via_url(driver, q, timeout=18)
                        force_open_place_details(driver, timeout=8)
                        wait_place_panel_ready(driver, timeout=6)
                        break
                    except (StaleElementReferenceException, TimeoutException):
                        if attempt == MAX_RETRY:
                            raise
                        open_home(driver)

                if partial_match_detected(driver):
                    print("   ⚠ partial match terdeteksi", flush=True)

                cur_url = (driver.current_url or "")
                cur_low = cur_url.lower()
                in_place = ("/maps/place" in cur_low)

                results_links = driver.find_elements(By.CSS_SELECTOR, "a.hfpxzc")

                # A) DIRECT PLACE
                if in_place:
                    nama_detail = get_place_title(driver, timeout=4) or ""
                    if nama_detail.strip().lower() in {"hasil", "result", "results"}:
                        nama_detail = ""

                    driver.execute_script("window.scrollBy(0, 300);")
                    time.sleep(0.2)

                    alamat_detail = get_address(driver, timeout=2) or ""
                    phone = get_phone(driver)
                    lat, lon = parse_coords_from_url(driver.current_url)

                    if not nama_detail:
                        t = (driver.title or "").strip()
                        if " - " in t:
                            t = t.split(" - ")[0].strip()
                        if t.lower() not in {"hasil", "result", "results"}:
                            nama_detail = t

                    is_closed, closed_type = detect_closed_status(driver)
                    is_echo = looks_like_query_echo(nama_detail or "", q, CITY_CONTEXT)
                    is_gen = is_generic_place_name(nama_detail or "")

                    sc, dbg = score_candidate(
                        nama_in, alamat_in, kec_in,
                        nama_detail or "", alamat_detail or "",
                        is_echo=is_echo, is_generic=is_gen
                    )

                    if sc <= 0 and lat is not None and lon is not None:
                        sc = 0.12
                        dbg = dbg or {}
                        dbg["coords_only_boost"] = True

                    print(
                        f"   • direct/place | score={sc:.2f} "
                        f"(ov_addr={dbg.get('ov_addr',0)}, ov_name={dbg.get('ov_name',0)}, "
                        f"s_name={dbg.get('s_name',0):.2f}, fuz={dbg.get('s_name_fuzzy',0):.2f}, "
                        f"s_addr={dbg.get('s_addr',0):.2f}, echo={dbg.get('is_echo')}, gen={dbg.get('is_generic')}) "
                        f"| url_latlon=({lat},{lon}) | nama={nama_detail} | alamat={alamat_detail}",
                        flush=True
                    )

                    if any([nama_detail, alamat_detail, lat, lon]) and sc > best["score"]:
                        best.update({
                            "score": sc,
                            "dbg": dbg,
                            "nama": nama_detail,
                            "alamat": alamat_detail,
                            "phone": phone,
                            "lat": lat,
                            "lon": lon,
                            "is_closed": is_closed,
                            "closed_type": closed_type,
                            "source": "direct/place",
                        })
                    if should_early_stop(best, THRESHOLD_EARLY_STOP):
                        stop_queries = True
                        break
    

                # B) LIST MODE
                elif results_links:
                    # 1) ambil kandidat banyak tapi tanpa buka detail
                    raw_cands = get_list_candidates_fast(driver, limit=max(8, MAX_CANDIDATES))
                    if not raw_cands:
                        raw_cands = [{"href": a.get_attribute("href"), "name_hint": a.get_attribute("aria-label") or "", "sub_hint": ""} 
                                    for a in results_links[:max(8, MAX_CANDIDATES)] if a.get_attribute("href")]

                    # 2) quick-score untuk ranking top-k
                    scored = []
                    for c in raw_cands:
                        qs, qdbg = quick_score_from_list(nama_in, alamat_in, kec_in, c.get("name_hint",""), c.get("sub_hint",""))
                        scored.append((qs, c))
                    scored.sort(key=lambda x: x[0], reverse=True)

                    # 3) buka detail hanya top_k (hemat waktu)
                    TOP_OPEN = 2  # <-- bisa 2 kalau mau lebih cepat
                    to_open = scored[:TOP_OPEN]

                    for ci, (qs, c) in enumerate(to_open, start=1):
                        if should_stop():
                            print(f"\n🛑 Stop saat proses kandidat baris {idx}.", flush=True)
                            break

                        href = c.get("href")
                        if not href:
                            continue

                        # buka detail kandidat pilihan
                        driver.get(href)
                        wait_document_ready(driver, 12)
                        click_consent_if_any(driver, timeout=1)
                        force_open_place_details(driver, timeout=6)
                        wait_place_panel_ready(driver, timeout=6)

                        nama_detail = get_place_title(driver, timeout=4) or ""
                        if nama_detail.strip().lower() in {"hasil", "result", "results"}:
                            nama_detail = ""

                        driver.execute_script("window.scrollBy(0, 300);")
                        time.sleep(0.2)
                        alamat_detail = get_address(driver, timeout=2) or ""
                        phone = get_phone(driver)
                        lat, lon = parse_coords_from_url(driver.current_url)

                        is_closed, closed_type = detect_closed_status(driver)
                        is_echo = looks_like_query_echo(nama_detail or "", q, CITY_CONTEXT)
                        is_gen = is_generic_place_name(nama_detail or "")

                        sc, dbg = score_candidate(
                            nama_in, alamat_in, kec_in,
                            nama_detail or "", alamat_detail or "",
                            is_echo=is_echo, is_generic=is_gen
                        )

                        if sc <= 0 and lat is not None and lon is not None:
                            sc = 0.12
                            dbg = dbg or {}
                            dbg["coords_only_boost"] = True

                        print(
                            f"   • cand#{ci} (pre={qs:.2f}) | score={sc:.2f} "
                            f"(ov_addr={dbg.get('ov_addr',0)}, ov_name={dbg.get('ov_name',0)}, "
                            f"s_name={dbg.get('s_name',0):.2f}, fuz={dbg.get('s_name_fuzzy',0):.2f}, "
                            f"s_addr={dbg.get('s_addr',0):.2f}, echo={dbg.get('is_echo')}, gen={dbg.get('is_generic')}) "
                            f"| latlon=({lat},{lon}) | nama={nama_detail} | alamat={alamat_detail}",
                            flush=True
                        )

                        if any([nama_detail, alamat_detail, lat, lon]) and sc > best["score"]:
                            best.update({
                                "score": sc,
                                "dbg": dbg,
                                "nama": nama_detail,
                                "alamat": alamat_detail,
                                "phone": phone,
                                "lat": lat,
                                "lon": lon,
                                "is_closed": is_closed,
                                "closed_type": closed_type,
                                "source": f"listTop#{ci}",
                            })

                        # stop dini kalau sudah sangat meyakinkan + coords valid di Denpasar
                        has_coords = (best["lat"] is not None and best["lon"] is not None)
                        in_den = is_within_bbox(best["lat"], best["lon"]) if has_coords else False
                        dbg_best = best.get("dbg") or {}

                        strong_name = (
                            float(dbg_best.get("s_name", 0.0) or 0.0) >= 0.82
                            or float(dbg_best.get("s_name_fuzzy", 0.0) or 0.0) >= 0.85
                        )

                        strong_addr = (
                            int(dbg_best.get("ov_addr", 0) or 0) >= 3
                            or float(dbg_best.get("s_addr", 0.0) or 0.0) >= 0.28
                        )

                        if best["score"] >= THRESHOLD_EARLY_STOP:
                            break

                        # ✅ tambahan: kalau sudah dapat coords Denpasar + (nama kuat atau alamat kuat), stop query berikutnya
                        if has_coords and in_den and (strong_name or strong_addr) and not dbg_best.get("is_echo"):
                            break


                        # kembali ke search list kalau masih perlu kandidat berikut
                        if last_search_url:
                            driver.get(last_search_url)
                            wait_document_ready(driver, 12)
                            click_consent_if_any(driver, timeout=1)

                        if should_early_stop(best, THRESHOLD_EARLY_STOP):
                            stop_queries = True

                        if stop_queries:
                            break
    


                # C) EMPTY FALLBACK
                else:
                    nama_detail = get_place_title(driver, timeout=3) or ""
                    if nama_detail.strip().lower() in {"hasil", "result", "results"}:
                        nama_detail = ""

                    driver.execute_script("window.scrollBy(0, 300);")
                    time.sleep(0.2)
                    alamat_detail = get_address(driver, timeout=2) or ""
                    phone = get_phone(driver)
                    lat, lon = parse_coords_from_url(driver.current_url)

                    is_closed, closed_type = detect_closed_status(driver)
                    is_echo = looks_like_query_echo(nama_detail or "", q, CITY_CONTEXT)
                    is_gen = is_generic_place_name(nama_detail or "")

                    sc, dbg = score_candidate(
                        nama_in, alamat_in, kec_in,
                        nama_detail or "", alamat_detail or "",
                        is_echo=is_echo, is_generic=is_gen
                    )

                    if sc <= 0 and lat is not None and lon is not None:
                        sc = 0.12
                        dbg = dbg or {}
                        dbg["coords_only_boost"] = True

                    print(
                        f"   • fallback/empty | score={sc:.2f} "
                        f"(ov_addr={dbg.get('ov_addr',0)}, ov_name={dbg.get('ov_name',0)}, "
                        f"s_name={dbg.get('s_name',0):.2f}, fuz={dbg.get('s_name_fuzzy',0):.2f}, "
                        f"s_addr={dbg.get('s_addr',0):.2f}, echo={dbg.get('is_echo')}, gen={dbg.get('is_generic')}) "
                        f"| latlon=({lat},{lon}) | nama={nama_detail} | alamat={alamat_detail}",
                        flush=True
                    )

                    if any([nama_detail, alamat_detail, lat, lon]) and sc > best["score"]:
                        best.update({
                            "score": sc,
                            "dbg": dbg,
                            "nama": nama_detail,
                            "alamat": alamat_detail,
                            "phone": phone,
                            "lat": lat,
                            "lon": lon,
                            "is_closed": is_closed,
                            "closed_type": closed_type,
                            "source": "fallback/empty",
                        })

                        

                        # stop dini kalau sudah sangat meyakinkan + coords valid di Denpasar
                        has_coords = (best["lat"] is not None and best["lon"] is not None)
                        in_den = is_within_bbox(best["lat"], best["lon"]) if has_coords else False
                        dbg_best = best.get("dbg") or {}

                        strong_name = (
                            float(dbg_best.get("s_name", 0.0) or 0.0) >= 0.82
                            or float(dbg_best.get("s_name_fuzzy", 0.0) or 0.0) >= 0.85
                        )

                        strong_addr = (
                            int(dbg_best.get("ov_addr", 0) or 0) >= 3
                            or float(dbg_best.get("s_addr", 0.0) or 0.0) >= 0.28
                        )

                        if best["score"] >= THRESHOLD_EARLY_STOP:
                            break

                        # ✅ tambahan: kalau sudah dapat coords Denpasar + (nama kuat atau alamat kuat), stop query berikutnya
                        if has_coords and in_den and (strong_name or strong_addr) and not dbg_best.get("is_echo"):
                            break

                        if should_early_stop(best, THRESHOLD_EARLY_STOP):
                            stop_queries = True
                            break


            # bila stop saat query loop, tetap simpan progres baris yg sudah ada
            if should_stop():
                print(f"\n🛑 Stop sebelum finalize scoring baris {idx}.", flush=True)
                break

            has_coords = (best["lat"] is not None and best["lon"] is not None)
            in_denpasar = is_within_bbox(best["lat"], best["lon"]) if has_coords else False

            dbg = best.get("dbg") or {}
            score_ok = best["score"] >= THRESHOLD_OK

            echo_bad = bool(dbg.get("is_echo")) and not (best.get("alamat") or "").strip()

            name_very_strong = (
                float(dbg.get("s_name", 0.0) or 0.0) >= 0.78
                or float(dbg.get("s_name_cont", 0.0) or 0.0) >= 0.70
                or float(dbg.get("s_name_fuzzy", 0.0) or 0.0) >= 0.80
            )

            name_signal_ok = (
                int(dbg.get("ov_name", 0) or 0) >= 1
                or float(dbg.get("s_name_fuzzy", 0.0) or 0.0) >= 0.55
                or float(dbg.get("s_name", 0.0) or 0.0) >= 0.50
            )

            alamat_g_ada = bool((best.get("alamat") or "").strip())
            alamat_in_ada = bool((alamat_in or "").strip())

            if has_coords and not in_denpasar:
                status_bisnis = f"Di luar Denpasar (lat={best['lat']}, lon={best['lon']})"
                status_kode = 0
                status_tutup = pd.NA
                lat_out = best["lat"]
                lon_out = best["lon"]

            elif (not has_coords) and best["score"] < 0:
                status_bisnis = "Tidak ditemukan"
                status_kode = 99
                status_tutup = pd.NA
                lat_out = None
                lon_out = None

            elif echo_bad:
                status_bisnis = f"Tidak ditemukan (echo_query, score={best['score']:.2f})"
                status_kode = 99
                status_tutup = pd.NA
                lat_out = None
                lon_out = None

            elif best["is_closed"] and has_coords and in_denpasar and (score_ok or name_very_strong):
                status_bisnis = "Tutup"
                status_kode = 3
                status_tutup = best["closed_type"] or "unknown"
                lat_out = best["lat"]
                lon_out = best["lon"]

            elif has_coords and in_denpasar and not dbg.get("is_echo") and (score_ok or name_very_strong):
                allow_even_if_generic = name_very_strong or (float(dbg.get("s_name_fuzzy", 0.0) or 0.0) >= 0.90)
                strong_addr = (int(dbg.get("ov_addr", 0) or 0) >= 3) or (float(dbg.get("s_addr", 0.0) or 0.0) >= 0.35)


                # if (dbg.get("is_generic") and not allow_even_if_generic):
                #     status_bisnis = f"Tidak ditemukan (nama_generik, score={best['score']:.2f})"
                #     status_kode = 99
                #     status_tutup = pd.NA
                #     lat_out = None
                #     lon_out = None

                if dbg.get("is_generic"):
                    # Tolak hanya kalau nama lemah DAN alamat juga lemah (jadi benar-benar generik)
                    if (not name_signal_ok) and (not strong_addr):
                        status_bisnis = f"Tidak ditemukan (nama_generik_lemah, score={best['score']:.2f})"
                        status_kode = 99
                        status_tutup = pd.NA
                        lat_out = None
                        lon_out = None
                    else:
                        # biarkan lanjut ke pemeriksaan alamat (atau accept)
                        pass


                else:
                    if alamat_in_ada and alamat_g_ada:
                        ov_addr = int(dbg.get("ov_addr", 0) or 0)
                        s_addr = float(dbg.get("s_addr", 0.0) or 0.0)

                        if float(dbg.get("s_name", 0.0) or 0.0) >= 0.92 or float(dbg.get("s_name_fuzzy", 0.0) or 0.0) >= 0.92:
                            status_bisnis = "Ditemukan (nama sangat kuat; alamat diabaikan)"
                            status_kode = 1
                            status_tutup = pd.NA
                            lat_out = best["lat"]
                            lon_out = best["lon"]
                        else:
                            a_in_alpha = addr_alpha_tokens(alamat_in)
                            a_g_alpha = addr_alpha_tokens(best.get("alamat") or "")
                            alpha_overlap = len(a_in_alpha & a_g_alpha)

                            addr_lock_ok = True
                            if len(a_in_alpha) >= 2:
                                addr_lock_ok = (alpha_overlap >= 1) or (ov_addr >= 2) or (s_addr >= 0.18)

                            if addr_lock_ok:
                                status_bisnis = "Ditemukan"
                                status_kode = 1
                                status_tutup = pd.NA
                                lat_out = best["lat"]
                                lon_out = best["lon"]
                            else:
                                status_bisnis = (
                                    f"Tidak ditemukan (alamat_tidak_match, score={best['score']:.2f}, "
                                    f"alpha_overlap={alpha_overlap}, ov_addr={ov_addr}, s_addr={s_addr:.2f})"
                                )
                                status_kode = 99
                                status_tutup = pd.NA
                                lat_out = None
                                lon_out = None

                    else:
                        if name_signal_ok:
                            status_bisnis = "Ditemukan (nama+coords; alamat_gmaps_kosong)"
                            status_kode = 1
                            status_tutup = pd.NA
                            lat_out = best["lat"]
                            lon_out = best["lon"]
                        else:
                            status_bisnis = f"Tidak ditemukan (alamat_gmaps_kosong & nama_lemah, score={best['score']:.2f})"
                            status_kode = 99
                            status_tutup = pd.NA
                            lat_out = None
                            lon_out = None

            elif ALLOW_COORDS_ONLY_MATCH and has_coords and in_denpasar:
                # pakai guard function yang sudah ada (biar fungsi kepakai, tidak cuma definisi)
                if coords_only_guard_ok(best.get("nama") or "", dbg):
                    status_bisnis = "Ditemukan (coords-only)"
                    status_kode = 5
                    status_tutup = pd.NA
                    lat_out = best["lat"]
                    lon_out = best["lon"]
                else:
                    status_bisnis = (
                        f"Tidak ditemukan (coords_only_ditolak, score={best['score']:.2f}, "
                        f"echo={dbg.get('is_echo')}, gen={dbg.get('is_generic')}, "
                        f"ov_name={dbg.get('ov_name',0)}, fuz={dbg.get('s_name_fuzzy',0.0):.2f})"
                    )
                    status_kode = 99
                    status_tutup = pd.NA
                    lat_out = None
                    lon_out = None

            else:
                status_bisnis = (
                    f"Tidak ditemukan (score_kurang, score={best['score']:.2f}, "
                    f"ov_addr={dbg.get('ov_addr',0)}, ov_name={dbg.get('ov_name',0)}, "
                    f"s_addr={dbg.get('s_addr',0.0):.2f})"
                )
                status_kode = 99
                status_tutup = best["closed_type"] if best["is_closed"] else pd.NA
                lat_out = None
                lon_out = None

            # =========================
            # SAFE ASSIGN (hindari dtype error)
            # =========================
            df.at[idx, "nama_gmaps"] = best["nama"] or ""
            df.at[idx, "alamat_gmaps"] = best["alamat"] or ""
            df.at[idx, "nomor_telepon"] = best["phone"] or ""
            df.at[idx, "score_match"] = round(best["score"], 4) if best["score"] >= 0 else pd.NA

            df.at[idx, "status_bisnis"] = status_bisnis or ""
            df.at[idx, "status_kode"] = int(status_kode) if status_kode is not None else pd.NA
            df.at[idx, "status_tutup"] = status_tutup if (status_tutup is not None and status_tutup is not pd.NA) else pd.NA

            apply_gc_fields(df, idx, int(status_kode) if status_kode is not None else 99,
                            nama_usaha_raw, alamat_usaha_raw, lat_out, lon_out)

            print(
                f"✅ Baris {idx} | best_score={best['score']:.2f} | source={best['source']} "
                f"| best_latlon=({best['lat']},{best['lon']}) | in_denpasar={in_denpasar} "
                f"| ov_addr={dbg.get('ov_addr',0)} ov_name={dbg.get('ov_name',0)} "
                f"| kode={status_kode} | {status_bisnis}",
                flush=True
            )

        except (TimeoutException, WebDriverException) as e:
            df.at[idx, "keterangan"] = f"Gagal diproses (timeout/driver): {e}"
            df.at[idx, "status_bisnis"] = "Gagal diproses"
            df.at[idx, "status_kode"] = 99
            apply_gc_fields(df, idx, 99, nama_usaha_raw, alamat_usaha_raw, None, None)
            try:
                driver.save_screenshot(os.path.join(SCREENSHOT_DIR, f"debug_row_{idx}.png"))
            except Exception:
                pass
            try:
                open_home(driver)
            except Exception:
                pass
            continue

        except Exception as e:
            df.at[idx, "keterangan"] = f"Gagal diproses: {e}"
            df.at[idx, "status_bisnis"] = "Gagal diproses"
            df.at[idx, "status_kode"] = 99
            apply_gc_fields(df, idx, 99, nama_usaha_raw, alamat_usaha_raw, None, None)
            try:
                driver.save_screenshot(os.path.join(SCREENSHOT_DIR, f"debug_row_{idx}.png"))
            except Exception:
                pass
            try:
                open_home(driver)
            except Exception:
                pass
            continue

    # final save (aman)
    safe_save_excel(df, file_path, tag="(final save)")
    print(f"\n✅ Proses selesai! File disimpan kembali ke: {file_path}", flush=True)

finally:
    try:
        driver.quit()
    except Exception:
        pass
    try:
        if log_fh:
            log_fh.close()
    except Exception:
        pass
