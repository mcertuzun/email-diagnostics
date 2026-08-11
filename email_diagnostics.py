# -*- coding: utf-8 -*-
"""
====================================================================
 EMAIL BOUNCE DIAGNOSTICS  -  Companies House assisted
====================================================================

Amac:
    Bounce eden (teslim edilemeyen) e-posta satirlarini analiz edip
    her satira TEK bir "result" teshisi yazmak.

Akis (STEP 1..7):
    1) Excel yukle + kolon dogrula
    2) Sadece problemli (bounce ailesi) statuleri filtrele
    3) first_name / last_name / company / email temizle
    4) TYPO kontrolu  -> typo bulunursa satir burada biter
    5) Typo yoksa Companies House officer sorgusu (paralel + throttle)
    6) Hata yonetimi
    7) TEK Excel ciktisi

ONEMLI KISIT:
    Bu script hicbir mail sunucusuna baglanmaz. SMTP handshake, RCPT TO
    probe, dogrulama servisi vb. YOKTUR. DNS/MX sorgusu da yapilmaz.
    Tek dis baglanti Companies House REST API'sidir.
    Dolayisiyla bir adresin gecerli oldugu ASLA kanitlanmaz; sadece
    isim/sirket ile tutarliligi cikarimlanir.

CALISMA ORTAMI:
    Python 3.6.5 / Windows uyumlu yazilmistir.
    - dataclasses (3.7+) KULLANILMAZ
    - pandas / numpy KULLANILMAZ  (regnum'un bastaki sifirlari korunsun diye)
    - rapidfuzz / python-Levenshtein KULLANILMAZ (Levenshtein elle yazildi)

GEREKLI PAKETLER:
    pip install openpyxl==3.0.10 requests==2.27.1

API ANAHTARI (kodda ASLA yazili degil):
    Windows kalici :   setx CH_API_KEY "buraya_anahtar"      (yeni terminal ac)
    Windows gecici :   set CH_API_KEY=buraya_anahtar
    PowerShell     :   $env:CH_API_KEY="buraya_anahtar"
    macOS / Linux  :   export CH_API_KEY="buraya_anahtar"

    Anahtari https://developer.company-information.service.gov.uk/ uzerinden
    bir "REST API" uygulamasi olusturarak alirsin.
"""

from __future__ import print_function

import argparse
import csv
import io
import logging
import os
import re
import sys
import threading
import time
import unicodedata
from collections import OrderedDict
from concurrent.futures import ThreadPoolExecutor

try:
    import openpyxl
except ImportError:  # pragma: no cover
    sys.exit("HATA: openpyxl kurulu degil.  ->  pip install openpyxl==3.0.10")

try:
    import requests
except ImportError:  # pragma: no cover
    sys.exit("HATA: requests kurulu degil.  ->  pip install requests==2.27.1")


# ====================================================================
# BOLUM 0 - AYARLAR  (tek degistirmen gereken yer burasi)
# ====================================================================

# --- Dosya yollari -------------------------------------------------
# Bunlar VARSAYILANDIR; komut satirindan --input / --output ile ezilebilir:
#   python email_diagnostics.py triage --input liste.csv --output sonuc.csv --verbose
INPUT_FILE = r"contacts.xlsx"                    # .xlsx / .xlsm / .csv / .tsv
OUTPUT_FILE = r"result_email_diagnostics.xlsx"   # yazilacak TEK dosya
INPUT_SHEET = None                               # None = ilk sayfa, ya da "Sheet1"
INPUT_DELIMITER = None                           # CSV ayraci; None = otomatik tahmin
INPUT_ENCODING = None                            # CSV kodlamasi; None = otomatik tespit

# --- Calisma modu --------------------------------------------------
DEBUG = False          # True -> yardimci/denetim kolonlari da yazilir
DRY_RUN = False        # True -> Companies House cagrilmaz (offline test)
MAX_ROWS = None        # Orn. 50 -> sadece ilk 50 problemli satir islenir (test)
MAX_COMPANIES = None   # Orn. 10 -> en fazla 10 BENZERSIZ regnum sorgulanir (kota testi)

# LOOKUP_MODE:
#   "typo_first" -> ONCE typo kontrolu, temiz satirlar API'ye gider (mevcut karar)
#   "ch_first"   -> ONCE Companies House, resmi isim alinir, email ona gore denetlenir
# Not: "ch_first" her satiri API'ye gonderir (daha yavas ama daha guvenilir isim).
LOOKUP_MODE = "typo_first"

# Sirket profili ek cagrisi (resmi sirket adi + dissolved durumu).
# Sirket basina +1 istek demektir, bu yuzden varsayilan kapali.
FETCH_COMPANY_PROFILE = False

# --- Zorunlu kolonlar ----------------------------------------------
REQUIRED_COLUMNS = ["first_name", "last_name", "email", "company", "regnum", "status"]

# --- STEP 2: status filtreleme -------------------------------------
# Bu kelimelerden herhangi biri statude geciyorsa satir ANALIZE ALINIR.
PROBLEMATIC_STATUS_KEYWORDS = [
    "bounce", "bounced", "bounced back", "hard bounce", "soft bounce",
    "undelivered", "not delivered", "delivery failed", "delivery failure",
    "failed", "failure", "rejected",
]
# Bu kelimeler statude geciyorsa satir DISLANIR (oncelikli).
# "blocked" bilerek disaridadir: spam filtresi/reputation kaynaklidir,
# adres yanlisligi veya kisinin ayrilmasiyla ilgisi yoktur.
EXCLUDED_STATUS_KEYWORDS = ["blocked", "block", "spam", "unsubscribed", "suppressed"]

# --- STEP 3: temizleme sozlukleri ----------------------------------
TITLES = {
    "mr", "mrs", "ms", "miss", "mx", "dr", "prof", "professor", "sir", "lord",
    "lady", "dame", "rev", "reverend", "capt", "captain", "hon", "madam", "master",
}
ROLE_WORDS = {
    "ceo", "cfo", "coo", "cto", "cio", "md", "director", "managing", "founder",
    "cofounder", "co-founder", "partner", "manager", "owner", "chairman",
    "chairwoman", "chair", "president", "vp", "head", "principal", "consultant",
    "secretary", "proprietor", "executive", "officer",
}
# UK post-nominal / unvan sonekleri (hem first_name hem last_name icinde olabilir)
POST_NOMINALS = {
    "obe", "mbe", "cbe", "kbe", "dbe", "bem", "jp", "dl",
    "fca", "aca", "acca", "fcca", "cima", "acma", "mrics", "frics", "mciob",
    "ceng", "mieee", "miet", "mba", "bsc", "msc", "ba", "ma", "beng", "meng",
    "phd", "dphil", "llb", "llm", "bcom", "cpa",
    "jr", "junior", "snr", "sr", "senior", "ii", "iii", "iv",
}
# Soyad on ekleri (surname'in parcasi olarak korunur)
SURNAME_PARTICLES = {
    "van", "von", "de", "del", "della", "der", "den", "di", "da", "du", "dos",
    "la", "le", "les", "mac", "mc", "st", "saint", "ter", "ten", "af", "al",
}
COMPANY_STOPWORDS = {
    "limited", "ltd", "ltd.", "plc", "llp", "llc", "incorporated", "inc",
    "corporation", "corp", "co", "company", "group", "holdings", "holding",
    "services", "service", "solutions", "uk", "gb", "the", "and", "international",
    "trading", "enterprises", "enterprise", "partners", "associates",
}
GENERIC_MAILBOXES = {
    "info", "information", "admin", "administrator", "office", "hello", "hi",
    "contact", "contactus", "enquiries", "enquiry", "inquiries", "inquiry",
    "sales", "support", "help", "helpdesk", "service", "customerservice",
    "accounts", "accounting", "account", "finance", "invoices", "invoice",
    "billing", "payments", "purchasing", "purchase", "orders", "order",
    "hr", "jobs", "careers", "recruitment", "marketing", "press", "media",
    "noreply", "no-reply", "donotreply", "mail", "email", "post", "reception",
    "team", "general", "main", "operations", "ops", "it", "webmaster", "director",
}
FREE_EMAIL_DOMAINS = {
    "gmail.com", "googlemail.com", "hotmail.com", "hotmail.co.uk", "outlook.com",
    "outlook.co.uk", "live.com", "live.co.uk", "msn.com", "yahoo.com",
    "yahoo.co.uk", "ymail.com", "aol.com", "icloud.com", "me.com", "mac.com",
    "btinternet.com", "sky.com", "virginmedia.com", "talktalk.net", "blueyonder.co.uk",
    "ntlworld.com", "protonmail.com", "proton.me", "gmx.com", "mail.com", "zoho.com",
}
# Cok parcali TLD'ler (domain kokunu dogru ayirmak icin)
MULTI_PART_TLDS = {
    "co.uk", "org.uk", "ltd.uk", "plc.uk", "me.uk", "net.uk", "sch.uk",
    "ac.uk", "gov.uk", "nhs.uk", "com.au", "co.nz", "co.za", "com.tr",
    "co.in", "com.sg", "co.jp", "com.br",
}
# Ingilizce lakap gruplari (iki yonlu genisletilir) - istedigin kadar ekleyebilirsin
NICKNAME_GROUPS = [
    ["robert", "rob", "bob", "bobby", "robbie"],
    ["william", "will", "bill", "billy", "willy"],
    ["richard", "rich", "rick", "dick", "richie"],
    ["john", "jon", "jack", "johnny", "jonny"],
    ["james", "jim", "jimmy", "jamie"],
    ["thomas", "tom", "tommy"],
    ["michael", "mike", "mick", "micky", "mikey"],
    ["david", "dave", "davey"],
    ["stephen", "steven", "steve", "stevie"],
    ["christopher", "chris", "kit"],
    ["matthew", "matt", "matty"],
    ["nicholas", "nick", "nicky"],
    ["anthony", "tony", "ant"],
    ["edward", "ed", "eddie", "ted", "teddy", "ned"],
    ["samuel", "sam", "sammy"],
    ["benjamin", "ben", "benny"],
    ["andrew", "andy", "drew"],
    ["daniel", "dan", "danny"],
    ["joseph", "joe", "joey"],
    ["peter", "pete", "petey"],
    ["gregory", "greg"],
    ["charles", "charlie", "chas", "chuck"],
    ["frederick", "fred", "freddie"],
    ["henry", "harry", "hal", "harold"],
    ["alexander", "alexandra", "alex", "sandy", "lex"],
    ["patrick", "pat", "paddy"],
    ["timothy", "tim", "timmy"],
    ["ronald", "ron", "ronnie"],
    ["donald", "don", "donnie"],
    ["kenneth", "ken", "kenny"],
    ["lawrence", "laurence", "larry", "laurie"],
    ["philip", "phillip", "phil"],
    ["raymond", "ray"],
    ["terence", "terry"],
    ["albert", "al", "bert", "bertie"],
    ["arthur", "art", "artie"],
    ["george", "geordie"],
    ["katherine", "catherine", "kathryn", "kate", "katie", "kathy", "cathy", "kay"],
    ["elizabeth", "liz", "lizzie", "beth", "betty", "eliza", "libby", "betsy"],
    ["susan", "sue", "susie", "suzy"],
    ["jennifer", "jen", "jenny"],
    ["margaret", "maggie", "peggy", "meg", "greta"],
    ["patricia", "pat", "patty", "trish", "tricia"],
    ["rebecca", "becky", "becca"],
    ["victoria", "vicky", "vicki", "tori"],
    ["christine", "christina", "chris", "chrissy", "tina"],
    ["deborah", "debbie", "deb", "debra"],
    ["sandra", "sandy"],
    ["barbara", "barb", "babs"],
    ["joanne", "joanna", "jo", "joan"],
    ["pamela", "pam"],
    ["angela", "angie"],
    ["theresa", "teresa", "terri", "tess"],
    ["cynthia", "cindy"],
    ["dorothy", "dot", "dottie"],
    ["frances", "francis", "fran", "frankie"],
    ["gillian", "gill"],
    ["helen", "ellie", "nell"],
    ["jacqueline", "jackie", "jacqui"],
    ["kimberley", "kim"],
    ["michelle", "shelly"],
    ["nicola", "nicky", "nikki"],
    ["samantha", "sam", "sammy"],
    ["sarah", "sara", "sally"],
    ["stephanie", "steph"],
    ["abigail", "abby"],
    ["amanda", "mandy"],
    ["caroline", "carolyn", "carrie", "caz"],
    ["eleanor", "ellie", "nora"],
    ["isabella", "isabel", "izzy", "bella"],
]

# --- STEP 4: esikler (hepsi ayarlanabilir) -------------------------
TYPO_MAX_DISTANCE_SHORT = 1   # local part <= 8 karakter ise izin verilen mesafe
TYPO_MAX_DISTANCE_LONG = 2    # local part > 8 karakter ise izin verilen mesafe
MIN_LOCAL_LEN_FOR_TYPO = 5    # bundan kisa local part'a asla "typo" denmez
DOMAIN_TYPO_MAX_DISTANCE = 2  # domain kokunde izin verilen mesafe
MIN_DOMAIN_LEN_FOR_TYPO = 4   # bundan kisa domain kokune asla "typo" denmez
SURNAME_MAX_DISTANCE = 1      # officer soyad eslesmesinde izin verilen mesafe

# --- STEP 5: Companies House ---------------------------------------
CH_API_BASE = "https://api.company-information.service.gov.uk"
CH_API_KEY_ENV = "CH_API_KEY"     # ortam degiskeni adi
CH_TIMEOUT = 20                   # saniye
CH_MAX_RETRIES = 3                # 429 / 5xx / timeout icin
CH_RATE_LIMIT_PER_SEC = 1.8       # resmi limit: 600 istek / 5 dk = 2.0/sn
CH_WORKERS = 4                    # paralel thread sayisi
CH_PAGE_SIZE = 100                # officers endpoint sayfa boyutu (UST SINIR, garanti degil)
CH_MAX_OFFICERS = 2000            # tek sirkette taranacak azami officer sayisi
CH_MAX_REQUESTS = None            # toplam HTTP istegi ust siniri (None = sinirsiz)

# --- Loglama -------------------------------------------------------
LOG_LEVEL = logging.INFO

logging.basicConfig(
    level=LOG_LEVEL,
    format="%(asctime)s  %(levelname)-7s  %(message)s",
    datefmt="%H:%M:%S",
)
log = logging.getLogger("diagnostics")


# ====================================================================
# BOLUM 1 - SABIT SONUC DEGERLERI
# ====================================================================

class R(object):
    """result kolonunda kullanilan sabit degerler."""
    MISSING_EMAIL = "missing_email"
    MALFORMED_EMAIL = "malformed_email"
    FIRST_NAME_TYPO = "first_name_typo"
    SURNAME_TYPO = "surname_typo"
    BOTH_TYPO = "first_name_and_surname_typo"
    DOMAIN_TYPO = "domain_typo"
    MISSING_REGNUM = "missing_regnum"
    COMPANY_NOT_FOUND = "company_not_found"
    CH_SKIPPED = "companies_house_skipped"
    COMPANY_DISSOLVED = "company_dissolved"
    LOOKUP_FAILED = "companies_house_lookup_failed"
    NO_OFFICER = "no_officer_match_found"
    ACTIVE = "active_officer_match"
    RESIGNED = "resigned_officer_match"
    POSSIBLE_ACTIVE = "possible_officer_match_active"
    POSSIBLE_RESIGNED = "possible_officer_match_resigned"


class RSN(object):
    """result_reason kolonunda kullanilan sabit degerler (kisa tutuldu)."""
    PATTERN_OK = "email_matches_expected_pattern"
    GENERIC = "generic_mailbox"
    PERSONAL_DOMAIN = "personal_email_domain"
    UNRECOGNISED = "email_pattern_unrecognised"
    DOMAIN_NOT_MATCHED = "domain_not_matched"
    CLOSE_TO_PATTERN = "close_to_expected_pattern"
    WITH_MIDDLE_NAME = "matched_including_middle_name"
    SURNAME_ONLY = "surname_only_match"
    MULTIPLE_OFFICERS = "multiple_possible_officers"
    API_ERROR = "api_error"
    NO_NAME = "name_fields_empty"


# ====================================================================
# BOLUM 2 - GENEL YARDIMCILAR
# ====================================================================

def strip_accents(text):
    """Sukru -> sukru, Ozturk -> ozturk, O'Brien -> o'brien (aksan katlama)."""
    if not text:
        return ""
    decomposed = unicodedata.normalize("NFKD", text)
    return "".join(ch for ch in decomposed if not unicodedata.combining(ch))


def normalize(text):
    """Karsilastirma icin: aksansiz, kucuk harf, tek bosluk."""
    if text is None:
        return ""
    text = strip_accents(str(text))
    text = text.replace("﻿", " ").replace(" ", " ")
    text = text.lower().strip()
    return re.sub(r"\s+", " ", text)


def alnum_only(text):
    """Sadece harf/rakam birak: 'john.smith' -> 'johnsmith'."""
    return re.sub(r"[^a-z0-9]", "", normalize(text))


def letters_only(text):
    """Sadece harf birak (rakam da atilir)."""
    return re.sub(r"[^a-z]", "", normalize(text))


def edit_distance(a, b, max_distance=None):
    """
    Damerau-Levenshtein (optimal string alignment) mesafesi.
    Saf Python - harici paket gerektirmez.

    Neden duz Levenshtein degil:
        En yaygin yazim hatasi iki harfin yer degistirmesidir
        ('jhon' <- 'john', 'smiht' <- 'smith').  Duz Levenshtein bunu
        2 hata sayar ve esigi asar; Damerau 1 hata sayar.

    max_distance verilirse asildigi anda erken cikar (hiz icin).
    """
    if a == b:
        return 0
    if not a:
        return len(b)
    if not b:
        return len(a)
    if max_distance is not None and abs(len(a) - len(b)) > max_distance:
        return max_distance + 1

    len_b = len(b)
    before_previous = None
    previous = list(range(len_b + 1))

    for i in range(1, len(a) + 1):
        current = [i] + [0] * len_b
        row_min = i
        for j in range(1, len_b + 1):
            cost = 0 if a[i - 1] == b[j - 1] else 1
            value = min(previous[j] + 1,          # silme
                        current[j - 1] + 1,       # ekleme
                        previous[j - 1] + cost)   # degistirme
            # yer degistirme (transposition)
            if (i > 1 and j > 1
                    and a[i - 1] == b[j - 2]
                    and a[i - 2] == b[j - 1]):
                value = min(value, before_previous[j - 2] + cost)
            current[j] = value
            if value < row_min:
                row_min = value
        if max_distance is not None and row_min > max_distance:
            return max_distance + 1
        before_previous = previous
        previous = current
    return previous[len_b]


def build_nickname_map():
    """['robert','bob'] -> {'robert': {'robert','bob'}, 'bob': {'robert','bob'}}"""
    mapping = {}
    for group in NICKNAME_GROUPS:
        members = set(normalize(name) for name in group if name)
        for member in members:
            mapping.setdefault(member, set()).update(members)
    return mapping


NICKNAME_MAP = build_nickname_map()


def expand_nicknames(name):
    """Bir ada karsilik gelen tum lakap/resmi varyantlari dondurur."""
    key = normalize(name)
    if not key:
        return set()
    return set(NICKNAME_MAP.get(key, {key}))


def cell_to_text(value):
    """
    openpyxl hucre degerini guvenli metne cevirir.
    1234567.0 gibi float'lari 1234567 yapar (regnum icin kritik).
    """
    if value is None:
        return ""
    if isinstance(value, bool):
        return "TRUE" if value else "FALSE"
    if isinstance(value, float):
        if value.is_integer():
            return str(int(value))
        return repr(value)
    if isinstance(value, int):
        return str(value)
    return str(value).replace("﻿", "").strip()


# ====================================================================
# BOLUM 3 - STEP 1: YUKLEME VE KOLON DOGRULAMA
# ====================================================================

def normalize_header(header):
    """'  First Name ' -> 'first_name',  BOM ve NBSP temizlenir."""
    text = (header or "")
    text = text.replace("﻿", "").replace(" ", " ")
    text = strip_accents(text).strip().lower()
    text = re.sub(r"[\s\-]+", "_", text)
    text = re.sub(r"[^a-z0-9_]", "", text)
    return re.sub(r"_+", "_", text).strip("_")


def is_csv_path(path):
    """Uzantiya bakarak CSV/TSV mi Excel mi oldugunu soyler."""
    return os.path.splitext(path)[1].lower() in (".csv", ".tsv", ".txt")


def _sniff_delimiter(header_line):
    """
    Ayraci baslik satirindan tahmin eder.
    Turkiye/Avrupa Excel ciktilari genelde ';' kullanir, bu yuzden ',' varsayimi yetmez.
    """
    counts = [(header_line.count(sep), sep) for sep in [";", ",", "\t", "|"]]
    counts.sort(reverse=True)
    return counts[0][1] if counts[0][0] > 0 else ","


def _read_text_lines(path, encoding=None):
    """
    CSV'yi dogru kodlamayla okur.
    Windows ciktilarinda utf-8-sig (BOM'lu) ve cp1254/cp1252 cok yaygindir.
    """
    encodings = [encoding] if encoding else ["utf-8-sig", "utf-8", "cp1254", "cp1252", "latin-1"]
    last_error = None
    for candidate in encodings:
        try:
            with io.open(path, "r", encoding=candidate, newline="") as handle:
                content = handle.read()
            if candidate != encodings[0]:
                log.warning("Kodlama '%s' olarak algilandi.", candidate)
            return content, candidate
        except (UnicodeDecodeError, LookupError) as exc:
            last_error = exc
            continue
    raise ValueError("Dosya kodlamasi cozulemedi: {}  ({})\n"
                     "--encoding ile elle belirtebilirsin.".format(path, last_error))


def _load_csv(path, delimiter=None, encoding=None):
    """CSV/TSV okur. Tum degerler metin olarak kalir (regnum sifirlari korunur)."""
    content, used_encoding = _read_text_lines(path, encoding)
    if not content.strip():
        raise ValueError("CSV dosyasi bos: {}".format(path))

    first_line = content.split("\n", 1)[0]
    sep = delimiter or _sniff_delimiter(first_line)
    log.info("CSV okunuyor (ayrac=%r, kodlama=%s)", sep, used_encoding)

    reader = csv.reader(io.StringIO(content), delimiter=sep)
    try:
        header_row = next(reader)
    except StopIteration:
        raise ValueError("CSV dosyasi bos: {}".format(path))

    original_headers = [cell_to_text(cell) for cell in header_row]
    while original_headers and not original_headers[-1]:
        original_headers.pop()
    if not original_headers:
        raise ValueError("Baslik satiri bos.")
    width = len(original_headers)

    rows = []
    for raw_row in reader:
        values = [cell_to_text(cell) for cell in raw_row[:width]]
        if len(values) < width:
            values.extend([""] * (width - len(values)))
        if not any(values):
            continue
        rows.append(values)

    return original_headers, [normalize_header(h) for h in original_headers], rows, sep, used_encoding


def load_data(path, sheet_name=None, delimiter=None, encoding=None):
    """
    Girdi dosyasini okur. Uzantiya gore CSV/TSV veya Excel.

    pandas KULLANMAZ - tum degerler metin olarak alinir, boylece regnum'un
    bastaki sifirlari ve tarih benzeri alanlar bozulmaz.

    Doner: (original_headers, normalized_headers, rows, meta)
           rows -> her biri list(str), header ile ayni uzunlukta
           meta -> {"delimiter": ..., "encoding": ...}  (CSV icin, Excel'de None)
    """
    if not os.path.isfile(path):
        raise IOError("Girdi dosyasi bulunamadi: {}".format(os.path.abspath(path)))

    if is_csv_path(path):
        headers, normalized, rows, sep, enc = _load_csv(path, delimiter, encoding)
        log.info("Yuklendi: %s satir, %s kolon  (%s)",
                 len(rows), len(headers), os.path.basename(path))
        return headers, normalized, rows, {"delimiter": sep, "encoding": enc}

    if not path.lower().endswith((".xlsx", ".xlsm")):
        raise ValueError(
            "Desteklenmeyen uzanti: {}\n"
            "Okunabilenler: .xlsx, .xlsm, .csv, .tsv, .txt  (.xls destegi yok)".format(path))

    workbook = openpyxl.load_workbook(path, read_only=True, data_only=True)
    try:
        if sheet_name:
            if sheet_name not in workbook.sheetnames:
                raise ValueError("Sayfa bulunamadi: '{}'. Mevcut sayfalar: {}"
                                 .format(sheet_name, ", ".join(workbook.sheetnames)))
            sheet = workbook[sheet_name]
        else:
            sheet = workbook[workbook.sheetnames[0]]

        row_iter = sheet.iter_rows(values_only=True)
        try:
            header_row = next(row_iter)
        except StopIteration:
            raise ValueError("Excel dosyasi bos.")

        original_headers = [cell_to_text(cell) for cell in header_row]
        # Sondaki tamamen bos kolonlari kirp
        while original_headers and not original_headers[-1]:
            original_headers.pop()
        width = len(original_headers)
        if width == 0:
            raise ValueError("Baslik satiri bos.")

        normalized_headers = [normalize_header(h) for h in original_headers]

        rows = []
        for raw_row in row_iter:
            values = [cell_to_text(cell) for cell in raw_row[:width]]
            if len(values) < width:
                values.extend([""] * (width - len(values)))
            if not any(values):      # tamamen bos satiri atla
                continue
            rows.append(values)
    finally:
        workbook.close()

    log.info("Yuklendi: %s satir, %s kolon  (%s)", len(rows), width, os.path.basename(path))
    return original_headers, normalized_headers, rows, {"delimiter": None, "encoding": None}


def validate_columns(normalized_headers, required=REQUIRED_COLUMNS):
    """
    Zorunlu kolonlarin varligini dogrular.
    Doner: {kolon_adi: indeks}
    Eksik varsa net bir hata firlatir.
    """
    index_map = {}
    for position, name in enumerate(normalized_headers):
        if name and name not in index_map:
            index_map[name] = position

    missing = [name for name in required if name not in index_map]
    if missing:
        raise ValueError(
            "Zorunlu kolon(lar) eksik: {}\nDosyada bulunanlar: {}".format(
                ", ".join(missing),
                ", ".join(h for h in normalized_headers if h) or "(yok)"
            )
        )
    return index_map


# ====================================================================
# BOLUM 4 - STEP 2: STATU FILTRELEME
# ====================================================================

def is_problematic_status(status_text):
    """
    Buyuk/kucuk harf ve kucuk yazim farklarina dayanikli statu kontrolu.
    Once dislama listesine (blocked vb.) bakar, sonra problem listesine.
    """
    text = normalize(status_text)
    if not text:
        return False
    flat = re.sub(r"[^a-z ]", " ", text)
    tokens = set(flat.split())

    for keyword in EXCLUDED_STATUS_KEYWORDS:
        key = normalize(keyword)
        if key in tokens or key in flat:
            return False

    for keyword in PROBLEMATIC_STATUS_KEYWORDS:
        if normalize(keyword) in flat:
            return True
    return False


def filter_problematic_statuses(rows, status_index):
    """Sadece problemli statulu satirlari dondurur."""
    kept = [row for row in rows if is_problematic_status(row[status_index])]
    log.info("Statu filtresi: %s satirdan %s tanesi problemli.", len(rows), len(kept))
    return kept


# ====================================================================
# BOLUM 5 - STEP 3: VERI TEMIZLEME
# ====================================================================

_NOISE_WORDS = TITLES | ROLE_WORDS | POST_NOMINALS


def _tokenize_person_field(raw):
    """Isim alanini parcalara ayirir; parantez icindeki lakaplari ayri dondurur."""
    text = (raw or "").replace("﻿", " ")
    nicknames = [normalize(m) for m in re.findall(r"[\(\[\"']([^\)\]\"']+)[\)\]\"']", text)]
    text = re.sub(r"[\(\[\"'][^\)\]\"']*[\)\]\"']", " ", text)   # parantezleri sil
    text = text.replace(",", " , ")
    tokens = [t for t in re.split(r"\s+", strip_accents(text).strip()) if t]
    return tokens, [n for n in nicknames if n]


def _is_noise_token(token):
    """Unvan / rol / post-nominal mi?"""
    key = re.sub(r"[^a-z]", "", normalize(token))
    return bool(key) and key in _NOISE_WORDS


def _clean_name_tokens(tokens):
    """
    Unvan, rol, post-nominal ve noktalama artiklarini atar.
    Geriye (gercek_isim_parcalari, bas_harfler) dondurur.
    """
    words = []
    initials = []
    for token in tokens:
        if token == ",":
            words.append(",")
            continue
        if _is_noise_token(token):
            continue
        core = re.sub(r"[^a-z\-']", "", normalize(token))
        if not core:
            continue
        if len(core) == 1:
            initials.append(core)
        else:
            words.append(core)
    return words, initials


def clean_first_name(raw):
    """
    first_name alanini temizler.

    Ele aldigi durumlar:
      "Mr John Smith"        -> first='john', middle=[], extra_surname='smith'
      "Dr. Jane (Janie)"     -> first='jane', nicknames=['janie']
      "John, CEO"            -> first='john'
      "J. Michael"           -> first='michael', initials=['j']

    Doner: dict(first, middles, nicknames, initials, trailing)
    """
    tokens, nicknames = _tokenize_person_field(raw)
    words, initials = _clean_name_tokens(tokens)
    words = [w for w in words if w != ","]

    first = words[0] if words else ""
    middles = words[1:] if len(words) > 1 else []
    return {
        "first": first,
        "middles": middles,
        "nicknames": [n for n in nicknames if n and n != first],
        "initials": initials,
    }


def clean_last_name(raw):
    """
    last_name alanini temizler.

    Ele aldigi durumlar:
      "Smith"                -> surname='smith'
      "John Smith"           -> surname='smith', extra_firsts=['john']
      "SMITH, John"          -> surname='smith', extra_firsts=['john']   (virgullu format)
      "van der Berg"         -> surname='van der berg' (+ 'vanderberg' varyanti)
      "Smith MBE"            -> surname='smith'
      "Mr Smith"             -> surname='smith'   (unvan soyadda da olabilir)

    Doner: dict(surname, extra_firsts, initials)
    """
    tokens, _ = _tokenize_person_field(raw)
    words, initials = _clean_name_tokens(tokens)

    if "," in words:
        pivot = words.index(",")
        before = [w for w in words[:pivot] if w != ","]
        after = [w for w in words[pivot + 1:] if w != ","]
        # "SMITH, John" -> soyad once
        if before:
            return {"surname": " ".join(before), "extra_firsts": after, "initials": initials}
        words = after

    words = [w for w in words if w != ","]
    if not words:
        return {"surname": "", "extra_firsts": [], "initials": initials}

    # Soyad on eklerini (van/de/mc...) soyada dahil et
    surname_start = len(words) - 1
    while surname_start > 0 and words[surname_start - 1] in SURNAME_PARTICLES:
        surname_start -= 1

    surname = " ".join(words[surname_start:])
    extra_firsts = words[:surname_start]
    return {"surname": surname, "extra_firsts": extra_firsts, "initials": initials}


def resolve_person_name(raw_first, raw_last):
    """
    first_name ve last_name alanlarini birlikte degerlendirip
    nihai ad / orta ad / soyad adaylarini uretir.

    Kritik durumlar:
      - first_name bos, last_name = "John Smith"   -> first='john', surname='smith'
      - first_name = "John Smith", last_name bos   -> first='john', surname='smith'
      - her ikisi de dolu                          -> dogrudan kullanilir
    """
    first_data = clean_first_name(raw_first)
    last_data = clean_last_name(raw_last)

    first = first_data["first"]
    middles = list(first_data["middles"])
    surname = last_data["surname"]
    extra_firsts = list(last_data["extra_firsts"])

    # last_name tam ad iceriyorsa ("John Smith") ilk parcalari ad adayina cevir
    if extra_firsts:
        if not first:
            first = extra_firsts[0]
            middles.extend(extra_firsts[1:])
        else:
            middles.extend(extra_firsts)

    # first_name tam ad iceriyorsa ("John Smith") ve soyad bossa
    if not surname and middles:
        surname = middles[-1]
        middles = middles[:-1]

    # Ad adaylari: ad + lakaplar + lakap sozlugu genislemesi
    first_candidates = set()
    for value in [first] + first_data["nicknames"]:
        if value:
            first_candidates.update(expand_nicknames(value))
    first_candidates = set(letters_only(c) for c in first_candidates if letters_only(c))

    # Soyad varyantlari: bosluklu, bitisik, tireli, tire parcalari
    surname_variants = set()
    if surname:
        base = letters_only(surname)
        if base:
            surname_variants.add(base)
        surname_variants.add(re.sub(r"[^a-z\-]", "", normalize(surname)).replace("-", ""))
        for part in re.split(r"[\s\-]+", normalize(surname)):
            part = letters_only(part)
            if len(part) >= 3 and part not in SURNAME_PARTICLES:
                surname_variants.add(part)
    surname_variants = set(v for v in surname_variants if v)

    initials = set(first_data["initials"]) | set(last_data["initials"])

    return {
        "first": letters_only(first),
        "middles": [letters_only(m) for m in middles if letters_only(m)],
        "surname": letters_only(surname.replace(" ", "")),
        "surname_display": surname,
        "first_candidates": first_candidates,
        "surname_variants": surname_variants,
        "nicknames": first_data["nicknames"],
        "initials": initials,
    }


def clean_company_name(raw):
    """
    Sirket adindan hukuki ekleri ve durak kelimeleri atar.
    "Ali Veli Zeynep Trading Ltd." -> ['ali','veli','zeynep']

    Orijinal deger ASLA degistirilmez; bu sadece eslestirme icindir.
    """
    text = normalize(raw)
    text = text.replace("&", " and ")
    text = re.sub(r"[^a-z0-9]+", " ", text)
    tokens = [t for t in text.split() if t]
    meaningful = [t for t in tokens if t not in COMPANY_STOPWORDS and len(t) > 1]
    return meaningful if meaningful else tokens


def generate_company_domain_candidates(tokens, extra_names=None):
    """
    Sirket adindan olasi domain koklerini uretir.
    ['ali','veli','zeynep'] -> {'alivelizeynep','avz','ali','aliveli','ali-veli-zeynep', ...}
    """
    candidates = set()
    if extra_names:
        for name in extra_names:
            for token_list in [clean_company_name(name)]:
                candidates.update(generate_company_domain_candidates(token_list))

    if not tokens:
        return candidates

    joined = "".join(tokens)
    if joined:
        candidates.add(joined)
        candidates.add("-".join(tokens))
    acronym = "".join(t[0] for t in tokens if t)
    if len(acronym) >= 2:
        candidates.add(acronym)
    candidates.add(tokens[0])
    if len(tokens) >= 2:
        candidates.add(tokens[0] + tokens[1])
        candidates.add(tokens[0] + "-" + tokens[1])
        candidates.add(tokens[0] + tokens[1][0])
        candidates.add(tokens[0][0] + tokens[1])
    if len(tokens) >= 3:
        candidates.add(tokens[0] + tokens[1] + tokens[2])
    return set(c for c in candidates if c and len(c) >= 2)


EMAIL_RE = re.compile(r"^[a-z0-9!#$%&'*+/=?^_`{|}~\-]+(\.[a-z0-9!#$%&'*+/=?^_`{|}~\-]+)*"
                      r"@[a-z0-9]([a-z0-9\-]*[a-z0-9])?(\.[a-z0-9]([a-z0-9\-]*[a-z0-9])?)+$")


def split_domain(domain):
    """
    'mail.acme.co.uk' -> brand='acme', root='mail.acme'
    Cok parcali TLD'ler (co.uk, org.uk ...) dogru soyulur.
    """
    parts = [p for p in domain.split(".") if p]
    if len(parts) < 2:
        return domain, domain
    tld_length = 2 if ".".join(parts[-2:]) in MULTI_PART_TLDS and len(parts) >= 3 else 1
    name_parts = parts[:-tld_length]
    if not name_parts:
        name_parts = [parts[0]]
    brand = name_parts[-1]
    return brand, ".".join(name_parts)


def parse_email(raw):
    """
    E-postayi kucuk harfe cevirir, bosluklari atar, yapisini dogrular.
    Doner: dict(ok, status, raw, email, local, domain, brand)
      status -> "ok" | "missing" | "malformed"
    """
    text = (raw or "").strip()
    text = re.sub(r"\s+", "", text)
    text = text.replace("﻿", "").strip("<>").strip(";,")
    lowered = strip_accents(text).lower()

    if not lowered:
        return {"ok": False, "status": "missing", "raw": raw, "email": "",
                "local": "", "domain": "", "brand": ""}

    if lowered.count("@") != 1 or not EMAIL_RE.match(lowered) or ".." in lowered:
        return {"ok": False, "status": "malformed", "raw": raw, "email": lowered,
                "local": "", "domain": "", "brand": ""}

    local, domain = lowered.split("@", 1)
    brand, _root = split_domain(domain)
    return {"ok": True, "status": "ok", "raw": raw, "email": lowered,
            "local": local, "domain": domain, "brand": brand}


# ====================================================================
# BOLUM 6 - STEP 4: BEKLENEN E-POSTA KALIPLARI VE TYPO TESPITI
# ====================================================================

SEPARATORS = ["", ".", "_", "-"]


class Pattern(object):
    """Beklenen bir local-part kalibi + hangi parcadan olustugu bilgisi."""

    __slots__ = ("text", "parts", "roles", "sep")

    def __init__(self, parts, roles, sep):
        self.parts = parts
        self.roles = roles
        self.sep = sep
        self.text = sep.join(parts)


def generate_expected_email_patterns(name_data):
    """
    Ad/soyad/orta ad/lakap/bas harflerden olasi local-part kaliplarini uretir.

    Uretilenler (her ayrac icin):
        first.last, last.first, f.last, first.l, first.m.last,
        firstlast, lastfirst, flast, firstl, fmlast
    Ayrac gerektirmeyenler:
        first, last, fl, fm, fml  (bas harf kombinasyonlari)
    """
    patterns = []
    firsts = set(f for f in name_data["first_candidates"] if f)
    surnames = set(s for s in name_data["surname_variants"] if s)
    middles = [m for m in name_data["middles"] if m]

    if not firsts and not surnames:
        return patterns

    def add(parts, roles, sep):
        parts = [p for p in parts if p]
        if len(parts) != len(roles) or not parts:
            return
        patterns.append(Pattern(parts, roles, sep))

    for surname in surnames:
        add([surname], ["surname"], "")
        for first in firsts:
            for sep in SEPARATORS:
                add([first, surname], ["first", "surname"], sep)
                add([surname, first], ["surname", "first"], sep)
                add([first[0], surname], ["first", "surname"], sep)
                add([first, surname[0]], ["first", "surname"], sep)
                add([surname, first[0]], ["surname", "first"], sep)
                for middle in middles:
                    add([first, middle, surname], ["first", "middle", "surname"], sep)
                    add([first, middle[0], surname], ["first", "middle", "surname"], sep)
                    add([first[0], middle[0], surname], ["first", "middle", "surname"], sep)

    for first in firsts:
        add([first], ["first"], "")

    # Bas harf kombinasyonlari - "js@" gibi kisa adresler icin
    # (orn. Jasmine Susanne Smith -> js, jss, ss)
    initial_sources = []
    for first in firsts:
        initial_sources.append(first[0])
    middle_initials = [m[0] for m in middles]
    surname_initials = [s[0] for s in surnames]

    for fi in set(initial_sources):
        for si in set(surname_initials):
            add([fi, si], ["first", "surname"], "")
            add([si, fi], ["surname", "first"], "")
        for mi in set(middle_initials):
            add([fi, mi], ["first", "middle"], "")
            for si in set(surname_initials):
                add([fi, mi, si], ["first", "middle", "surname"], "")
    for extra in name_data["initials"]:
        for si in set(surname_initials):
            add([extra, si], ["first", "surname"], "")

    return patterns


def _strip_trailing_digits(text):
    """'john.smith2' -> ('john.smith', True)"""
    stripped = re.sub(r"\d+$", "", text)
    return (stripped, stripped != text) if stripped else (text, False)


def _score_pattern(local, pattern):
    """Local part ile kalip arasindaki duzenleme mesafesi."""
    limit = TYPO_MAX_DISTANCE_LONG + 1
    return edit_distance(local, pattern.text, max_distance=limit)


def _attribute_difference(local, pattern):
    """
    Sapmanin hangi isim parcasinda oldugunu bulur.
    Doner: set(['first'] / ['surname'] / ikisi birden)
    """
    bad = set()
    if pattern.sep:
        local_parts = local.split(pattern.sep)
        if len(local_parts) == len(pattern.parts):
            for value, expected, role in zip(local_parts, pattern.parts, pattern.roles):
                if edit_distance(value, expected, max_distance=3) > 0:
                    bad.add("first" if role in ("first", "middle") else "surname")
            return bad

    # Ayracsiz (bitisik) kalip: en iyi bolunme noktasini ara
    if len(pattern.parts) == 2:
        expected_a, expected_b = pattern.parts
        best = None
        for cut in range(1, max(2, len(local))):
            left, right = local[:cut], local[cut:]
            total = (edit_distance(left, expected_a, max_distance=4)
                     + edit_distance(right, expected_b, max_distance=4))
            if best is None or total < best[0]:
                best = (total, edit_distance(left, expected_a, max_distance=4),
                        edit_distance(right, expected_b, max_distance=4))
        if best:
            _total, da, db = best
            if da > 0:
                bad.add("first" if pattern.roles[0] in ("first", "middle") else "surname")
            if db > 0:
                bad.add("first" if pattern.roles[1] in ("first", "middle") else "surname")
            return bad

    return {"first", "surname"}


def check_domain(email_info, company_candidates):
    """
    Domain kokunu sirket adi adaylariyla karsilastirir.

    Kontrol SIRASI onemlidir:
      1) Tam eslesme                      -> ok
      2) Anlamli ek/on ek ile icerme      -> ok   ('acme' -> 'acmegroup', 'acmeuk')
      3) 1-2 karakter mesafe              -> typo ('acme' -> 'acmee')
      4) Digerleri                        -> unmatched (typo DEGIL, CH'ye devam)

    2. adim 3. adimdan ONCE gelmelidir, aksi halde 'acmegroup' typo sanilir.
    Ama 2. adimdaki ek en az 2 karakter olmalidir, aksi halde 'acmee'
    (tek harf fazlasi) yanlislikla gecerli sayilir.

    Doner: (verdict, reason)
      verdict -> "ok" | "personal" | "typo" | "unmatched"
    """
    domain = email_info["domain"]
    brand = email_info["brand"]

    if domain in FREE_EMAIL_DOMAINS:
        return "personal", RSN.PERSONAL_DOMAIN
    if not company_candidates:
        return "unmatched", RSN.DOMAIN_NOT_MATCHED

    brand_flat = re.sub(r"[^a-z0-9]", "", brand)
    if not brand_flat:
        return "unmatched", RSN.DOMAIN_NOT_MATCHED

    flats = []
    for candidate in company_candidates:
        flat = re.sub(r"[^a-z0-9]", "", candidate)
        if flat:
            flats.append(flat)

    # 1) Tam eslesme
    if brand_flat in flats:
        return "ok", None

    # 2) Icerme - fark en az 2 karakter olmali ('acmegroup' evet, 'acmee' hayir)
    for flat in flats:
        if len(flat) < MIN_DOMAIN_LEN_FOR_TYPO:
            continue     # 'ali' gibi kisa tokenlar 'alibaba'ya eslesmesin
        if flat in brand_flat and len(brand_flat) - len(flat) >= 2:
            return "ok", None
        if brand_flat in flat and len(flat) - len(brand_flat) >= 2:
            return "ok", None

    # 3) Yakin ama esit degil -> typo
    if len(brand_flat) >= MIN_DOMAIN_LEN_FOR_TYPO:
        best = None
        for flat in flats:
            if len(flat) < MIN_DOMAIN_LEN_FOR_TYPO:
                continue
            distance = edit_distance(brand_flat, flat,
                                     max_distance=DOMAIN_TYPO_MAX_DISTANCE + 1)
            if best is None or distance < best:
                best = distance
        if best is not None and 0 < best <= DOMAIN_TYPO_MAX_DISTANCE:
            return "typo", RSN.CLOSE_TO_PATTERN

    # 4) Alakasiz domain -> typo DEMEYIZ (holding/grup domaini olabilir)
    return "unmatched", RSN.DOMAIN_NOT_MATCHED


def detect_email_typo(email_info, name_data, company_candidates):
    """
    STEP 4 ana fonksiyonu.

    Doner: dict(terminal, result, reason, best_pattern, best_distance, domain_verdict)
      terminal=True  -> satir burada biter, Companies House cagrilmaz
      terminal=False -> satir Companies House kontrolune devam eder

    Tasarim ilkesi: ZAYIF farklara "typo" DEMEZ.
      - Tam eslesme veya sadece ayrac farki  -> tutarli
      - 1-2 karakter mesafe                  -> typo
      - Hicbir kalibin yakinina dusmuyorsa   -> 'email_pattern_unrecognised' (typo degil)
      - 4 karakterden kisa local part        -> asla typo degil (bas harf olabilir)
    """
    outcome = {"terminal": False, "result": None, "reason": None,
               "best_pattern": "", "best_distance": "", "domain_verdict": ""}

    local = email_info["local"]
    local_letters = re.sub(r"[^a-z0-9._\-]", "", local)

    domain_verdict, domain_reason = check_domain(email_info, company_candidates)
    outcome["domain_verdict"] = domain_verdict

    # --- Generic mailbox: kisiye ait degil, typo denemez ---
    if re.sub(r"[^a-z]", "", local) in GENERIC_MAILBOXES or local in GENERIC_MAILBOXES:
        if domain_verdict == "typo":
            outcome.update(terminal=True, result=R.DOMAIN_TYPO, reason=RSN.GENERIC)
        else:
            outcome["reason"] = RSN.GENERIC
        return outcome

    # --- Isim yoksa typo kontrolu anlamsiz ---
    patterns = generate_expected_email_patterns(name_data)
    if not patterns:
        outcome["reason"] = RSN.NO_NAME
        if domain_verdict == "typo":
            outcome.update(terminal=True, result=R.DOMAIN_TYPO,
                           reason=RSN.CLOSE_TO_PATTERN)
        return outcome

    compare_local, _had_digits = _strip_trailing_digits(local_letters)
    compare_flat = alnum_only(compare_local)

    # 1) Tam eslesme
    for pattern in patterns:
        if compare_local == pattern.text:
            outcome["best_pattern"] = pattern.text
            outcome["best_distance"] = 0
            outcome["reason"] = RSN.PATTERN_OK
            if domain_verdict == "typo":
                outcome.update(terminal=True, result=R.DOMAIN_TYPO,
                               reason=RSN.CLOSE_TO_PATTERN)
            elif domain_verdict == "unmatched":
                outcome["reason"] = RSN.DOMAIN_NOT_MATCHED
            elif domain_verdict == "personal":
                outcome["reason"] = RSN.PERSONAL_DOMAIN
            return outcome

    # 2) Ayrac farki (john.smith vs johnsmith) -> typo DEGIL
    for pattern in patterns:
        if compare_flat and compare_flat == alnum_only(pattern.text):
            outcome["best_pattern"] = pattern.text
            outcome["best_distance"] = 0
            outcome["reason"] = RSN.PATTERN_OK
            if domain_verdict == "typo":
                outcome.update(terminal=True, result=R.DOMAIN_TYPO,
                               reason=RSN.CLOSE_TO_PATTERN)
            elif domain_verdict == "unmatched":
                outcome["reason"] = RSN.DOMAIN_NOT_MATCHED
            elif domain_verdict == "personal":
                outcome["reason"] = RSN.PERSONAL_DOMAIN
            return outcome

    # 3) En yakin kalip
    best_pattern = None
    best_distance = None
    for pattern in patterns:
        distance = _score_pattern(compare_local, pattern)
        if best_distance is None or distance < best_distance:
            best_distance = distance
            best_pattern = pattern
            if distance == 0:
                break

    outcome["best_pattern"] = best_pattern.text if best_pattern else ""
    outcome["best_distance"] = best_distance if best_distance is not None else ""

    threshold = (TYPO_MAX_DISTANCE_SHORT if len(compare_local) <= 8
                 else TYPO_MAX_DISTANCE_LONG)

    too_short = len(compare_local) < MIN_LOCAL_LEN_FOR_TYPO
    is_typo = (best_pattern is not None
               and best_distance is not None
               and 0 < best_distance <= threshold
               and not too_short)

    if is_typo:
        bad_parts = _attribute_difference(compare_local, best_pattern)
        if bad_parts == {"first"}:
            result = R.FIRST_NAME_TYPO
        elif bad_parts == {"surname"}:
            result = R.SURNAME_TYPO
        else:
            result = R.BOTH_TYPO
        outcome.update(terminal=True, result=result, reason=RSN.CLOSE_TO_PATTERN)
        return outcome

    # 4) Hicbir kalibin yakininda degil -> typo DEMEYIZ, CH'ye devam
    if domain_verdict == "typo":
        outcome.update(terminal=True, result=R.DOMAIN_TYPO, reason=RSN.UNRECOGNISED)
        return outcome

    if domain_verdict == "personal":
        outcome["reason"] = RSN.PERSONAL_DOMAIN
    elif domain_verdict == "unmatched":
        outcome["reason"] = RSN.DOMAIN_NOT_MATCHED
    else:
        outcome["reason"] = RSN.UNRECOGNISED
    return outcome


# ====================================================================
# BOLUM 7 - STEP 5: COMPANIES HOUSE ISTEMCISI
# ====================================================================

class CompaniesHouseAuthError(Exception):
    """401/403 - anahtar hatali. Tum calisma durdurulur."""


class CompanyNotFound(Exception):
    """404 - sirket numarasi bulunamadi."""


class LookupFailed(Exception):
    """Yeniden denemelerden sonra basarisiz."""


class RateLimiter(object):
    """
    Thread'ler arasi ortak hiz sinirlayici.
    Companies House limiti: 5 dakikada 600 istek (= 2.0 istek/sn).
    """

    def __init__(self, rate_per_second):
        self._interval = 1.0 / float(rate_per_second)
        self._lock = threading.Lock()
        self._next_slot = 0.0

    def acquire(self):
        with self._lock:
            now = time.time()
            wait = self._next_slot - now
            if wait > 0:
                time.sleep(wait)
                now = time.time()
            self._next_slot = max(now, self._next_slot) + self._interval


class CompaniesHouseClient(object):
    """Companies House REST istemcisi: throttle + retry + pagination."""

    def __init__(self, api_key, rate_per_second=CH_RATE_LIMIT_PER_SEC,
                 max_requests=CH_MAX_REQUESTS):
        self._session = requests.Session()
        self._session.auth = (api_key, "")     # Basic auth: kullanici=anahtar, sifre bos
        self._session.headers.update({"Accept": "application/json"})
        self._limiter = RateLimiter(rate_per_second)
        self.last_status = {}                  # regnum -> son HTTP kodu (DEBUG icin)

        # Kota takibi: kac HTTP istegi gittigini TAHMIN etmek yerine sayariz.
        self._max_requests = max_requests
        self._stats_lock = threading.Lock()
        self.stats = {"requests": 0, "retries": 0, "rate_limited": 0,
                      "pages": 0, "failed": 0}

    def _bump(self, key, amount=1):
        with self._stats_lock:
            self.stats[key] += amount

    def _reserve_request(self):
        """
        Istek sayacini artirir. Ust sinir asildiysa False doner ve
        istek HIC gonderilmez - kotayi korumak icin sert fren.
        """
        with self._stats_lock:
            if self._max_requests is not None and self.stats["requests"] >= self._max_requests:
                return False
            self.stats["requests"] += 1
            return self.stats["requests"]

    def _request(self, path, params=None):
        url = CH_API_BASE + path
        last_error = "bilinmeyen hata"

        for attempt in range(1, CH_MAX_RETRIES + 1):
            sequence = self._reserve_request()
            if sequence is False:
                raise LookupFailed(
                    "istek ust sinirina ulasildi ({}). --max-requests ile artirabilirsin."
                    .format(self._max_requests))
            if attempt > 1:
                self._bump("retries")

            self._limiter.acquire()
            log.debug("CH istek #%s (deneme %s): %s %s",
                      sequence, attempt, path, params or "")
            try:
                response = self._session.get(url, params=params, timeout=CH_TIMEOUT)
            except requests.exceptions.RequestException as exc:
                last_error = "baglanti: {}".format(exc)
                time.sleep(min(2 ** attempt, 10))
                continue

            code = response.status_code

            if code == 200:
                self._bump("pages")
                try:
                    return response.json(), code
                except ValueError:
                    last_error = "gecersiz JSON"
                    time.sleep(1)
                    continue

            if code in (401, 403):
                raise CompaniesHouseAuthError(
                    "Companies House kimlik dogrulama hatasi (HTTP {}). "
                    "{} ortam degiskenindeki anahtari kontrol et.".format(code, CH_API_KEY_ENV)
                )
            if code == 404:
                raise CompanyNotFound("HTTP 404")
            if code == 429:
                retry_after = response.headers.get("Retry-After")
                try:
                    delay = float(retry_after) if retry_after else 5.0
                except (TypeError, ValueError):
                    delay = 5.0
                self._bump("rate_limited")
                log.warning("Rate limit (429) - %.1f sn bekleniyor...", delay)
                time.sleep(min(delay, 60))
                last_error = "HTTP 429"
                continue
            if 500 <= code < 600:
                last_error = "HTTP {}".format(code)
                time.sleep(min(2 ** attempt, 10))
                continue

            last_error = "HTTP {}".format(code)
            break

        self._bump("failed")
        raise LookupFailed(last_error)

    def get_companies_house_officers(self, company_number):
        """
        Sirketin TUM officer kayitlarini dondurur (sayfalama dahil).
        35'ten fazla officer'i olan sirketlerde sayfalama sart.
        """
        officers = []
        start_index = 0
        while True:
            params = {"items_per_page": CH_PAGE_SIZE, "start_index": start_index}
            data, code = self._request("/company/{}/officers".format(company_number), params)
            self.last_status[company_number] = code

            if not isinstance(data, dict):
                raise LookupFailed("beklenmeyen yanit yapisi")
            items = data.get("items")
            if not isinstance(items, list):
                items = []
            officers.extend(items)

            total = data.get("total_results")
            try:
                total = int(total)
            except (TypeError, ValueError):
                total = len(officers)

            # Sunucu istedigimizden AZ kayit dondurebilir (items_per_page bir
            # ust sinirdir, garanti degil). Bu yuzden ilerleme, istenen sayfa
            # boyutu kadar degil, GERCEKTEN donen kayit sayisi kadar olmali.
            # Aksi halde hem bosuna fazladan istek atilir hem de aradaki
            # kayitlar sessizce atlanir.
            received = len(items)
            if received == 0:
                break                       # sonsuz donguye karsi emniyet
            start_index += received
            if len(officers) >= total or start_index >= CH_MAX_OFFICERS:
                break
        return officers

    def get_company_profile(self, company_number):
        """Resmi sirket adi, durumu ve onceki adlari."""
        data, code = self._request("/company/{}".format(company_number))
        self.last_status[company_number] = code
        if not isinstance(data, dict):
            raise LookupFailed("beklenmeyen yanit yapisi")
        return data


def _clean_key_value(value):
    """
    Anahtarin etrafindaki bosluk ve tirnaklari atar.
    Windows'ta `set CH_API_KEY="abc"` yazildiginda tirnaklar degerin
    ICINE girer ve API 401 doner; bu sessiz hatanin onune gecer.
    """
    value = (value or "").strip()
    if len(value) >= 2 and value[0] == value[-1] and value[0] in ("\"", "'"):
        value = value[1:-1].strip()
    return value


def read_env_file(path):
    """
    Basit .env okuyucu (harici paket gerektirmez).
    KEY=VALUE satirlarini dondurur; '#' ile baslayan satirlar yorumdur.
    """
    values = {}
    if not os.path.isfile(path):
        return values
    try:
        with io.open(path, "r", encoding="utf-8-sig") as handle:
            for line in handle:
                line = line.strip()
                if not line or line.startswith("#") or "=" not in line:
                    continue
                key, _sep, value = line.partition("=")
                key = key.strip()
                if key.lower().startswith("export "):
                    key = key[7:].strip()
                values[key] = _clean_key_value(value)
    except (IOError, UnicodeDecodeError):
        pass
    return values


def get_api_key():
    """
    Companies House anahtarini sirayla arar:
        1) Ortam degiskeni
        2) Calisilan klasordeki .env
        3) Scriptin yanindaki .env

    Doner: (anahtar, kaynak_aciklamasi).  Bulunamazsa ("", "").
    Anahtar hicbir zaman kod icine yazilmaz.
    """
    value = _clean_key_value(os.environ.get(CH_API_KEY_ENV, ""))
    if value:
        return value, "ortam degiskeni"

    for path in env_file_candidates():
        value = _clean_key_value(read_env_file(path).get(CH_API_KEY_ENV, ""))
        if value:
            return value, path
    return "", ""


def env_file_candidates():
    """Aranacak .env yollari (ayni klasordeysek tekrar etmez)."""
    paths = [
        os.path.join(os.getcwd(), ".env"),
        os.path.join(os.path.dirname(os.path.abspath(__file__)), ".env"),
    ]
    unique = []
    for path in paths:
        if path not in unique:
            unique.append(path)
    return unique


def _lookalike_env_vars():
    """
    Yazim/buyuk-kucuk harf hatasini yakalar: 'ch_api_key', 'CH_APIKEY',
    'CH_API_KEY ' gibi degiskenler varsa kullaniciya soyler.
    """
    target = CH_API_KEY_ENV.replace("_", "").upper()
    hits = []
    for name in os.environ:
        flat = name.strip().replace("_", "").replace("-", "").upper()
        if name != CH_API_KEY_ENV and (flat == target or "COMPANIESHOUSE" in flat):
            hits.append(name)
    return sorted(hits)


def api_key_help():
    """Anahtar bulunamadiginda gosterilecek, tanilama iceren yardim metni."""
    similar = _lookalike_env_vars()
    lines = [
        "Companies House anahtari bulunamadi.",
        "",
        "Arandigi yerler:",
        "  1) {} ortam degiskeni".format(CH_API_KEY_ENV),
    ]
    for number, path in enumerate(env_file_candidates(), start=2):
        lines.append("  {}) {}".format(number, path))
    lines += [
        "",
        "Nasil ayarlanir:",
        "  Windows (kalici) : setx {} \"anahtar\"".format(CH_API_KEY_ENV),
        "                     >>> SONRA TERMINALI/VS CODE'U TAMAMEN KAPATIP AC <<<",
        "                     setx yalnizca YENI acilan islemleri etkiler; VS Code",
        "                     ortamini acildigi anda alir, yeni sekme yetmez.",
        "  Windows (gecici) : set {}=anahtar        (tirnak KOYMA)".format(CH_API_KEY_ENV),
        "  PowerShell       : $env:{}=\"anahtar\"".format(CH_API_KEY_ENV),
        "  macOS / Linux    : export {}=\"anahtar\"".format(CH_API_KEY_ENV),
        "",
        "Alternatif - bu klasore '.env' adli bir dosya olusturup icine yaz:",
        "  {}=anahtar".format(CH_API_KEY_ENV),
        "  (.env dosyasi .gitignore'da, repoya gitmez)",
        "",
        "Anahtari almak icin: https://developer.company-information.service.gov.uk/",
        "",
        "Kontrol icin: python email_diagnostics.py check",
    ]
    if similar:
        lines.insert(3, "  >>> Benzer isimli degisken(ler) BULUNDU: {}".format(", ".join(similar)))
        lines.insert(4, "  >>> Isim tam olarak '{}' olmali (buyuk harf, alt cizgi)."
                        .format(CH_API_KEY_ENV))
    return "\n".join(lines)


def mask_key(value):
    """Anahtari loglarken maskele: ilk 4 ve son 2 karakter disinda gizle."""
    if not value:
        return "(yok)"
    if len(value) <= 8:
        return value[:2] + "*" * (len(value) - 2)
    return "{}{}{}".format(value[:4], "*" * (len(value) - 6), value[-2:])


def normalize_regnum(raw):
    """
    UK sirket numarasi 8 karakterdir. Excel '01234567' degerini sayiya
    cevirip '1234567' yapabilir -> API 404 doner. Basa sifir eklenir.
    SC / NI / OC / FC gibi harfli onekler zaten 8 karakterdir, etkilenmez.
    """
    text = re.sub(r"[^A-Za-z0-9]", "", (raw or "")).upper()
    if not text:
        return ""
    if text.isdigit():
        return text.zfill(8)
    return text


# ====================================================================
# BOLUM 8 - OFFICER ESLESTIRME
# ====================================================================

def _officer_name_parts(officer):
    """
    Officer kaydindan (forename, other_forenames, surname, gorunen_ad) cikarir.
    Yapisal name_elements tercih edilir; yoksa "SURNAME, Forename" formatindan
    ayristirilir. Isim uretilemezse (orn. kurumsal officer) None doner.
    """
    if not isinstance(officer, dict):
        return None

    elements = officer.get("name_elements")
    display = (officer.get("name") or "").strip()

    if isinstance(elements, dict) and elements.get("surname"):
        forename = normalize(elements.get("forename") or "")
        others = normalize(elements.get("other_forenames") or "")
        surname = normalize(elements.get("surname") or "")
        pretty_parts = [p for p in [(elements.get("forename") or "").strip(),
                                    (elements.get("other_forenames") or "").strip(),
                                    (elements.get("surname") or "").strip()] if p]
        pretty = " ".join(pretty_parts) or display
        return forename, others, surname, pretty

    if display and "," in display:
        surname_raw, forenames_raw = display.split(",", 1)
        surname = normalize(surname_raw)
        forenames = normalize(forenames_raw).split()
        if surname:
            forename = forenames[0] if forenames else ""
            others = " ".join(forenames[1:])
            pretty = " ".join([p for p in forenames_raw.split()] + [surname_raw.strip()])
            return forename, others, surname, pretty.strip() or display

    return None   # kurumsal officer veya ayristirilamayan kayit


def _surname_match_level(contact_variants, officer_surnames):
    """2 = tam eslesme, 1 = <=1 mesafe, 0 = eslesme yok."""
    best = 0
    for officer_surname in officer_surnames:
        flat_officer = letters_only(officer_surname.replace(" ", ""))
        if not flat_officer:
            continue
        for variant in contact_variants:
            if not variant:
                continue
            if variant == flat_officer:
                return 2
            if (len(variant) >= 4 and len(flat_officer) >= 4
                    and edit_distance(variant, flat_officer,
                                    max_distance=SURNAME_MAX_DISTANCE) <= SURNAME_MAX_DISTANCE):
                best = max(best, 1)
    return best


def _forename_match_level(contact_firsts, contact_initials, officer_forenames):
    """2 = ad/lakap eslesti, 1 = sadece bas harf, 0 = eslesme yok."""
    officer_set = set()
    for value in officer_forenames:
        for token in normalize(value).split():
            token = letters_only(token)
            if token:
                officer_set.add(token)
    if not officer_set or not contact_firsts:
        return 0

    officer_expanded = set()
    for token in officer_set:
        officer_expanded.update(expand_nicknames(token))
    officer_expanded.update(officer_set)

    for first in contact_firsts:
        if first in officer_expanded:
            return 2
        for token in officer_expanded:
            if len(first) >= 4 and len(token) >= 4 and edit_distance(first, token, max_distance=1) <= 1:
                return 2

    initials = set(i[0] for i in contact_firsts if i) | set(contact_initials)
    for token in officer_expanded:
        if token and token[0] in initials:
            return 1
    return 0


def match_contact_to_officers(name_data, officers):
    """
    Kisiyi officer listesiyle eslestirir.  SOYAD CAPADIR:
    once soyad tutmali, sonra ad degerlendirilir.

    Doner: dict(status, officer_name, reason, confidence)
      status -> "active" | "resigned" | "possible_active" | "possible_resigned" | "none"
    """
    empty = {"status": "none", "officer_name": "", "reason": None, "confidence": "none"}
    if not officers:
        return empty

    contact_surnames = set(name_data["surname_variants"])
    if name_data["surname"]:
        contact_surnames.add(name_data["surname"])
    contact_surnames = set(s for s in contact_surnames if s)
    if not contact_surnames:
        return empty

    contact_firsts = set(f for f in name_data["first_candidates"] if f)
    contact_initials = set(name_data["initials"])

    matches = []
    for officer in officers:
        parsed = _officer_name_parts(officer)
        if not parsed:
            continue
        forename, others, surname, pretty = parsed

        # former_names: evlilik vb. nedeniyle degismis soyadlari da kapsar
        former_surnames, former_forenames = [], []
        for former in (officer.get("former_names") or []):
            if isinstance(former, dict):
                if former.get("surname"):
                    former_surnames.append(normalize(former["surname"]))
                if former.get("forenames"):
                    former_forenames.append(normalize(former["forenames"]))

        surname_level = _surname_match_level(contact_surnames, [surname] + former_surnames)
        if surname_level == 0:
            continue

        forename_level = _forename_match_level(
            contact_firsts, contact_initials, [forename, others] + former_forenames
        )

        if forename_level == 2:
            confidence = "confident"
        elif forename_level == 1:
            confidence = "possible"
        else:
            confidence = "possible"

        matches.append({
            "pretty": pretty,
            "key": letters_only(pretty.replace(" ", "")),
            "confidence": confidence,
            "surname_level": surname_level,
            "forename_level": forename_level,
            "resigned": bool(officer.get("resigned_on")),
        })

    if not matches:
        return empty

    rank = {"confident": 2, "possible": 1}
    best_rank = max(rank[m["confidence"]] for m in matches)
    finalists = [m for m in matches if rank[m["confidence"]] == best_rank]

    distinct_people = set(m["key"] for m in finalists)
    ambiguous = len(distinct_people) > 1

    # Ayni kisinin hem istifa hem aktif kaydi varsa AKTIF kazanir
    chosen_key = sorted(distinct_people)[0]
    same_person = [m for m in finalists if m["key"] == chosen_key]
    is_active = any(not m["resigned"] for m in same_person)
    display_name = same_person[0]["pretty"]

    confident = (best_rank == 2) and not ambiguous
    if confident:
        status = "active" if is_active else "resigned"
    else:
        status = "possible_active" if is_active else "possible_resigned"

    reason = None
    if ambiguous:
        reason = RSN.MULTIPLE_OFFICERS
    elif best_rank == 1 and finalists[0]["forename_level"] == 0:
        reason = RSN.SURNAME_ONLY
    elif finalists[0]["forename_level"] == 2 and len(display_name.split()) >= 3:
        reason = RSN.WITH_MIDDLE_NAME

    return {"status": status, "officer_name": display_name,
            "reason": reason, "confidence": "confident" if confident else "possible"}


# ====================================================================
# BOLUM 9 - SATIR ISLEME VE ORKESTRASYON
# ====================================================================

def build_row_context(row, index_map):
    """Bir satirin temizlenmis (orijinali bozmadan) analiz baglamini kurar."""
    raw_first = row[index_map["first_name"]]
    raw_last = row[index_map["last_name"]]
    raw_email = row[index_map["email"]]
    raw_company = row[index_map["company"]]
    raw_regnum = row[index_map["regnum"]]

    name_data = resolve_person_name(raw_first, raw_last)
    company_tokens = clean_company_name(raw_company)
    email_info = parse_email(raw_email)
    regnum = normalize_regnum(raw_regnum)

    return {
        "row": row,
        "name": name_data,
        "company_tokens": company_tokens,
        "company_candidates": generate_company_domain_candidates(company_tokens),
        "email": email_info,
        "regnum": regnum,
        "result": None,
        "reason": None,
        "officer_name": "",
        "officer_status": "not_checked",
        "company_name": "",
        "typo": None,
    }


def run_email_stage(context):
    """STEP 4'u calistirir ve terminal olup olmadigini isaretler."""
    email_info = context["email"]

    if email_info["status"] == "missing":
        context["result"] = R.MISSING_EMAIL
        context["reason"] = None
        return True
    if email_info["status"] == "malformed":
        context["result"] = R.MALFORMED_EMAIL
        context["reason"] = None
        return True

    outcome = detect_email_typo(email_info, context["name"], context["company_candidates"])
    context["typo"] = outcome
    context["reason"] = outcome["reason"]
    if outcome["terminal"]:
        context["result"] = outcome["result"]
        return True
    return False


def fetch_company_data(client, company_numbers):
    """
    Benzersiz sirket numaralari icin officer (ve istege bagli profil) verisini
    PARALEL ceker. Ayni regnum bir kez sorgulanir.

    Doner: {regnum: {"officers": [...], "profile": {...}, "error": None|str}}
    """
    results = {}
    if not company_numbers:
        return results

    total = len(company_numbers)
    log.info("Companies House: %s benzersiz sirket, %s thread, ~%.1f istek/sn",
             total, CH_WORKERS, CH_RATE_LIMIT_PER_SEC)
    estimated = total * (2 if FETCH_COMPANY_PROFILE else 1) / CH_RATE_LIMIT_PER_SEC
    log.info("Tahmini sure: ~%.1f dakika", estimated / 60.0)

    counter = {"done": 0}
    counter_lock = threading.Lock()
    fatal = {"error": None}

    def worker(company_number):
        if fatal["error"]:
            return company_number, {"officers": [], "profile": None, "error": "aborted"}
        entry = {"officers": [], "profile": None, "error": None}
        before = client.stats["requests"]
        try:
            if FETCH_COMPANY_PROFILE:
                entry["profile"] = client.get_company_profile(company_number)
            entry["officers"] = client.get_companies_house_officers(company_number)
        except CompaniesHouseAuthError as exc:
            fatal["error"] = str(exc)
            entry["error"] = "auth"
        except CompanyNotFound:
            entry["error"] = "not_found"
        except LookupFailed as exc:
            entry["error"] = "failed: {}".format(exc)
        except Exception as exc:                      # beklenmeyen yanit yapisi vb.
            entry["error"] = "failed: {}".format(exc)

        spent = client.stats["requests"] - before
        if spent > 1:
            log.debug("  %s: %s istek harcadi (sayfalama veya yeniden deneme)",
                      company_number, spent)
        with counter_lock:
            counter["done"] += 1
            if counter["done"] % 25 == 0 or counter["done"] == total:
                log.info("  ... %s/%s sirket sorgulandi", counter["done"], total)
        return company_number, entry

    with ThreadPoolExecutor(max_workers=CH_WORKERS) as pool:
        for company_number, entry in pool.map(worker, company_numbers):
            results[company_number] = entry

    # Kota raporu: tahmin degil, gercekten gonderilen istek sayisi.
    stats = client.stats
    per_company = (float(stats["requests"]) / total) if total else 0.0
    log.info("Companies House: %s HTTP istegi / %s sirket  (sirket basina %.2f)",
             stats["requests"], total, per_company)
    if stats["retries"] or stats["rate_limited"] or stats["failed"]:
        log.info("  bunun %s tanesi yeniden deneme, %s tanesi rate limit (429), "
                 "%s sirket basarisiz", stats["retries"], stats["rate_limited"],
                 stats["failed"])
    if per_company > 1.2 and not FETCH_COMPANY_PROFILE:
        log.warning("  Sirket basina 1'den fazla istek gitti. Sebebi genelde cok "
                    "officer'i olan sirketlerde sayfalamadir.")

    if fatal["error"]:
        raise CompaniesHouseAuthError(fatal["error"])
    return results


def apply_companies_house(context, company_data):
    """CH sonucunu satira uygular (STEP 5)."""
    regnum = context["regnum"]

    if not regnum:
        context["result"] = R.MISSING_REGNUM
        context["officer_status"] = "not_checked"
        return

    entry = company_data.get(regnum)
    if entry is None:
        context["result"] = R.LOOKUP_FAILED
        context["officer_status"] = "lookup_failed"
        context["reason"] = context["reason"] or RSN.API_ERROR
        return

    profile = entry.get("profile")
    if isinstance(profile, dict):
        context["company_name"] = profile.get("company_name") or ""

    if entry["error"] == "not_found":
        context["result"] = R.COMPANY_NOT_FOUND
        context["officer_status"] = "not_found"
        return
    if entry["error"]:
        context["result"] = R.LOOKUP_FAILED
        context["officer_status"] = "lookup_failed"
        context["reason"] = RSN.API_ERROR
        return

    # Sirket kapanmissa bounce'un en guclu aciklamasi budur
    if isinstance(profile, dict):
        status = (profile.get("company_status") or "").lower()
        if status in ("dissolved", "liquidation", "receivership", "administration",
                      "converted-closed", "closed"):
            context["result"] = R.COMPANY_DISSOLVED
            context["officer_status"] = "not_checked"
            return

    officers = entry.get("officers") or []
    if not officers:
        context["result"] = R.NO_OFFICER
        context["officer_status"] = "not_found"
        return

    match = match_contact_to_officers(context["name"], officers)
    context["officer_name"] = match["officer_name"]
    if match["reason"]:
        context["reason"] = match["reason"]

    mapping = {
        "active": (R.ACTIVE, "active"),
        "resigned": (R.RESIGNED, "resigned"),
        "possible_active": (R.POSSIBLE_ACTIVE, "possible_active"),
        "possible_resigned": (R.POSSIBLE_RESIGNED, "possible_resigned"),
    }
    if match["status"] in mapping:
        label, officer_status = mapping[match["status"]]
        context["result"] = "{}: {}".format(label, match["officer_name"])
        context["officer_status"] = officer_status
    else:
        context["result"] = R.NO_OFFICER
        context["officer_status"] = "not_found"


# ====================================================================
# BOLUM 10 - STEP 7: CIKTI YAZIMI
# ====================================================================

OUTPUT_COLUMNS = ["result", "result_reason", "ch_officer_name", "ch_officer_status"]

DEBUG_COLUMNS = [
    "dbg_clean_first", "dbg_clean_middles", "dbg_clean_surname", "dbg_nicknames",
    "dbg_company_tokens", "dbg_email_local", "dbg_email_domain",
    "dbg_best_pattern", "dbg_best_distance", "dbg_domain_verdict", "dbg_regnum_used",
]


def _build_output_headers(original_headers):
    headers = list(original_headers) + list(OUTPUT_COLUMNS)
    if FETCH_COMPANY_PROFILE:
        headers.append("ch_company_name")
    if DEBUG:
        headers.extend(DEBUG_COLUMNS)
    return headers


def _build_output_row(context):
    values = list(context["row"])
    values.extend([
        context["result"] or "",
        context["reason"] or "",
        context["officer_name"] or "",
        context["officer_status"] or "",
    ])
    if FETCH_COMPANY_PROFILE:
        values.append(context["company_name"] or "")
    if DEBUG:
        typo = context["typo"] or {}
        name = context["name"]
        values.extend([
            name["first"],
            " ".join(name["middles"]),
            name["surname"],
            " ".join(name["nicknames"]),
            " ".join(context["company_tokens"]),
            context["email"]["local"],
            context["email"]["domain"],
            typo.get("best_pattern", ""),
            typo.get("best_distance", ""),
            typo.get("domain_verdict", ""),
            context["regnum"],
        ])
    return values


def write_output(path, original_headers, contexts, meta=None):
    """
    TEK cikti dosyasi yazar. Uzantiya gore CSV veya Excel.
    Orijinal kolonlar aynen korunur, sonuna yeni kolonlar eklenir.
    Girdi dosyasi ASLA degistirilmez.

    meta -> {"delimiter":..., "encoding":...}; CSV ciktisinda girdiyle
    ayni ayrac kullanilir, kodlama Excel uyumlulugu icin utf-8-sig olur.
    """
    headers = _build_output_headers(original_headers)
    meta = meta or {}

    try:
        if is_csv_path(path):
            sep = meta.get("delimiter") or ","
            # utf-8-sig: Excel'in Turkce karakterleri dogru acmasi icin BOM sart
            with io.open(path, "w", encoding="utf-8-sig", newline="") as handle:
                writer = csv.writer(handle, delimiter=sep, quoting=csv.QUOTE_MINIMAL)
                writer.writerow(headers)
                for context in contexts:
                    writer.writerow([("" if v is None else v)
                                     for v in _build_output_row(context)])
        else:
            workbook = openpyxl.Workbook()
            sheet = workbook.active
            sheet.title = "diagnostics"
            sheet.append(headers)
            for context in contexts:
                sheet.append(_build_output_row(context))
            sheet.freeze_panes = "A2"
            workbook.save(path)
    except IOError as exc:
        raise IOError(
            "Cikti dosyasi yazilamadi: {}\n"
            "Dosya Excel'de acik olabilir - kapatip tekrar dene.\nDetay: {}"
            .format(os.path.abspath(path), exc)
        )
    log.info("Cikti yazildi: %s  (%s satir)", os.path.abspath(path), len(contexts))


def print_summary(contexts):
    """Konsola result dagilimini basar (ekstra dosya olusturmaz)."""
    counts = OrderedDict()
    for context in contexts:
        label = (context["result"] or "").split(":")[0].strip() or "(bos)"
        counts[label] = counts.get(label, 0) + 1

    log.info("-" * 52)
    log.info("SONUC DAGILIMI (%s satir)", len(contexts))
    for label, count in sorted(counts.items(), key=lambda kv: -kv[1]):
        share = 100.0 * count / len(contexts) if contexts else 0.0
        log.info("  %-38s %5s  (%4.1f%%)", label, count, share)
    log.info("-" * 52)


# ====================================================================
# BOLUM 11 - MAIN
# ====================================================================

def main():
    start_time = time.time()

    # --- Guvenlik: girdi dosyasinin uzerine yazma ---
    if os.path.abspath(INPUT_FILE) == os.path.abspath(OUTPUT_FILE):
        raise ValueError("Cikti dosyasi girdi dosyasiyla ayni olamaz. Girdi korunmalidir.")

    # --- STEP 1 ---
    original_headers, normalized_headers, rows, meta = load_data(
        INPUT_FILE, INPUT_SHEET, INPUT_DELIMITER, INPUT_ENCODING)
    index_map = validate_columns(normalized_headers)

    # --- STEP 2 ---
    problematic = filter_problematic_statuses(rows, index_map["status"])
    if MAX_ROWS:
        problematic = problematic[:MAX_ROWS]
        log.warning("MAX_ROWS aktif: sadece ilk %s satir islenecek.", MAX_ROWS)
    if not problematic:
        log.warning("Problemli statuye sahip satir yok. Bos cikti yazilacak.")
        write_output(OUTPUT_FILE, original_headers, [], meta)
        return

    # --- STEP 3 + 4 ---
    contexts = [build_row_context(row, index_map) for row in problematic]

    pending = []
    for context in contexts:
        if LOOKUP_MODE == "ch_first":
            pending.append(context)                 # herkes API'ye gider
        else:
            if not run_email_stage(context):        # typo yoksa API'ye gider
                pending.append(context)

    log.info("Typo asamasi: %s satirdan %s tanesi Companies House'a gidecek.",
             len(contexts), len(pending))

    # --- STEP 5 ---
    # DIKKAT: --limit SATIR sayisini sinirlar, sirket sayisini degil.
    # 10 satir 3 farkli regnum tasiyorsa yalnizca 3 sirket sorgulanir.
    # Kotayi dogrudan sinirlamak icin --limit-companies kullanilir.
    company_numbers = sorted(set(c["regnum"] for c in pending if c["regnum"]))

    if MAX_COMPANIES and len(company_numbers) > MAX_COMPANIES:
        kept = set(company_numbers[:MAX_COMPANIES])
        dropped = len(company_numbers) - len(kept)
        company_numbers = sorted(kept)
        skipped = 0
        for context in pending:
            if context["regnum"] not in kept:
                context["result"] = context["result"] or R.CH_SKIPPED
                context["officer_status"] = "not_checked"
                skipped += 1
        pending = [c for c in pending if c["regnum"] in kept]
        log.warning("--limit-companies %s: %s sirket kapsam disinda birakildi "
                    "(%s satir '%s' olarak isaretlendi).",
                    MAX_COMPANIES, dropped, skipped, R.CH_SKIPPED)

    log.info("Kapsam: %s satir -> %s benzersiz sirket "
             "(sirket basina en az 1 istek gidecek)",
             len(pending), len(company_numbers))

    if DRY_RUN:
        # Companies House BILEREK atlandi. Bu bir hata degil, bu yuzden
        # 'lookup_failed' yazmak yanlis olur - ayri bir deger kullaniyoruz.
        # E-posta asamasindan gelen reason korunur, boylece typo analizi
        # tam olarak elde kalir; sadece officer kontrolu yapilmamis olur.
        log.warning("Companies House atlandi (--dry-run / --no-ch): %s satir "
                    "officer kontrolu yapilmadan isaretlendi.", len(pending))
        for context in pending:
            context["result"] = context["result"] or R.CH_SKIPPED
            context["officer_status"] = "not_checked"
    elif company_numbers or pending:
        api_key, key_source = get_api_key()
        if not api_key:
            raise EnvironmentError(api_key_help())
        log.info("Companies House anahtari bulundu (kaynak: %s)", key_source)
        client = CompaniesHouseClient(api_key, CH_RATE_LIMIT_PER_SEC, CH_MAX_REQUESTS)
        company_data = fetch_company_data(client, company_numbers)

        for context in pending:
            apply_companies_house(context, company_data)
            # ch_first modunda: resmi isim alindiktan SONRA email kontrolu
            if LOOKUP_MODE == "ch_first" and context["officer_name"]:
                verified = resolve_person_name(context["officer_name"], "")
                context["name"]["first_candidates"] |= verified["first_candidates"]
                if not context["result"] or context["result"].startswith(R.ACTIVE):
                    run_email_stage(context)

    # Emniyet: hicbir satir bos result ile kalmasin
    for context in contexts:
        if not context["result"]:
            context["result"] = R.NO_OFFICER

    # --verbose: her satirin karari tek tek loglanir
    for context in contexts:
        log.debug("%-40s -> %-34s %s",
                  context["email"]["email"] or "(email yok)",
                  context["result"], context["reason"] or "")

    # --- STEP 7 ---
    write_output(OUTPUT_FILE, original_headers, contexts, meta)
    print_summary(contexts)
    log.info("Tamamlandi: %.1f saniye", time.time() - start_time)


# ====================================================================
# BOLUM 12 - KURULUM KONTROLU  (check komutu)
# ====================================================================

# Anahtarin gecerliligini sinamak icin kullanilan sirket numarasi.
# Hangi numara oldugu onemli degil: 401/403 disinda HERHANGI bir yanit
# (200 de 404 de) kimlik dogrulamanin gectigini kanitlar.
CHECK_TEST_COMPANY = "00000006"


def _mark(ok):
    return "[OK]  " if ok else "[HATA]"


def check_setup(input_path=None, skip_api=False, delimiter=None, encoding=None):
    """
    Kurulumu bastan sona dogrular ve nerede takildigini soyler.

    Sirasiyla: Python surumu -> paketler -> API anahtari -> anahtar gecerli mi
    -> (istege bagli) girdi dosyasi okunabiliyor mu, kolonlar tam mi.

    Doner: 0 = her sey hazir, 1 = eksik var.
    """
    problems = []

    print("=" * 68)
    print(" KURULUM KONTROLU")
    print("=" * 68)

    # --- 1) Python ---
    version = ".".join(str(n) for n in sys.version_info[:3])
    ok_version = sys.version_info >= (3, 6)
    print("%s Python %s  (%s)" % (_mark(ok_version), version, sys.platform))
    if not ok_version:
        problems.append("Python 3.6 veya uzeri gerekiyor.")

    # --- 2) Paketler ---
    for module, name in [(openpyxl, "openpyxl"), (requests, "requests")]:
        installed = getattr(module, "__version__", "?")
        print("%s %-9s %s" % (_mark(True), name, installed))

    # --- 3) API anahtari ---
    api_key, source = get_api_key()
    if api_key:
        print("%s Anahtar bulundu: %s   (kaynak: %s)"
              % (_mark(True), mask_key(api_key), source))
    else:
        print("%s Anahtar bulunamadi" % _mark(False))
        problems.append("api_key")

    # --- 4) Anahtar gercekten calisiyor mu ---
    if api_key and not skip_api:
        print("      Companies House'a tek test istegi gonderiliyor...")
        client = CompaniesHouseClient(api_key)
        try:
            client.get_company_profile(CHECK_TEST_COMPANY)
            print("%s Anahtar gecerli, API erisimi calisiyor." % _mark(True))
        except CompanyNotFound:
            # 404 da kimlik dogrulamanin gectigini kanitlar
            print("%s Anahtar gecerli, API erisimi calisiyor." % _mark(True))
        except CompaniesHouseAuthError:
            print("%s Anahtar REDDEDILDI (HTTP 401/403)." % _mark(False))
            print("      Anahtari yanlis kopyalamis olabilirsin, ya da 'Live' yerine")
            print("      'Test' anahtari kullaniyorsundur. Developer Hub'dan kontrol et.")
            problems.append("api_key_invalid")
        except LookupFailed as exc:
            print("%s API'ye ulasilamadi: %s" % (_mark(False), exc))
            print("      Internet baglantisi, proxy veya guvenlik duvari olabilir.")
            print("      Kurumsal agdaysan HTTPS_PROXY ortam degiskenini ayarla.")
            problems.append("api_unreachable")
    elif api_key and skip_api:
        print("      (--skip-api verildi, test istegi gonderilmedi)")

    # --- 5) Girdi dosyasi ---
    if input_path:
        print("-" * 68)
        print(" GIRDI DOSYASI: %s" % input_path)
        print("-" * 68)
        try:
            _orig, normalized, rows, _meta = load_data(input_path, None, delimiter, encoding)
            print("%s Dosya okundu: %s satir" % (_mark(True), len(rows)))

            missing = [c for c in REQUIRED_COLUMNS if c not in normalized]
            if missing:
                print("%s Eksik kolon(lar): %s" % (_mark(False), ", ".join(missing)))
                print("      Dosyadaki kolonlar: %s"
                      % ", ".join(h for h in normalized if h))
                problems.append("columns")
            else:
                print("%s Zorunlu kolonlarin hepsi var." % _mark(True))

                status_index = normalized.index("status")
                counts = {}
                for row in rows:
                    key = (row[status_index] or "(bos)").strip()
                    counts[key] = counts.get(key, 0) + 1
                problematic = sum(1 for row in rows
                                  if is_problematic_status(row[status_index]))
                print("%s Analiz edilecek satir: %s / %s"
                      % (_mark(problematic > 0), problematic, len(rows)))
                if problematic == 0:
                    print("      Hicbir satir bounce ailesinde degil. Statu degerlerin:")
                    problems.append("no_rows")
                print("      Statu dagilimi:")
                for value, count in sorted(counts.items(), key=lambda kv: -kv[1])[:12]:
                    flag = "analiz edilir" if is_problematic_status(value) else "-"
                    print("        %-28s %5s  %s" % (value[:28], count, flag))

                regnum_index = normalized.index("regnum")
                empty_regnum = sum(1 for row in rows if not (row[regnum_index] or "").strip())
                if empty_regnum:
                    print("      Uyari: %s satirda regnum bos." % empty_regnum)
        except (IOError, ValueError) as exc:
            print("%s Dosya okunamadi: %s" % (_mark(False), exc))
            problems.append("input")

    # --- Ozet ---
    print("=" * 68)
    if not problems:
        print(" HER SEY HAZIR. Calistirmak icin:")
        print("   python email_diagnostics.py triage --input liste.csv"
              " --output sonuc.csv --verbose")
        print("=" * 68)
        return 0

    print(" EKSIKLER VAR")
    print("=" * 68)
    if "api_key" in problems:
        print()
        print(api_key_help())
    return 1


# ====================================================================
# BOLUM 13 - KOMUT SATIRI ARAYUZU
# ====================================================================

def build_arg_parser():
    """
    Kullanim:
        python email_diagnostics.py triage --input liste.csv --output sonuc.csv --verbose

    Girdi ve cikti .xlsx veya .csv olabilir; uzantiya gore otomatik secilir.
    """
    parser = argparse.ArgumentParser(
        prog="email_diagnostics.py",
        description="Bounce eden e-posta satirlarini teshis eder ve TEK bir cikti dosyasi yazar.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=(
            "ornekler:\n"
            "  python email_diagnostics.py check\n"
            "  python email_diagnostics.py check --input liste.csv\n"
            "  python email_diagnostics.py triage --input liste.csv --output sonuc.csv --verbose\n"
            "  python email_diagnostics.py triage -i liste.xlsx -o sonuc.xlsx --limit 50 --dry-run\n\n"
            "Yeni bir bilgisayarda ONCE 'check' calistir: neyin eksik oldugunu soyler.\n\n"
            "Companies House anahtari {env} ortam degiskeninden ya da .env dosyasindan okunur.\n"
            "  Windows    : setx {env} \"anahtar\"   (sonra VS Code'u TAMAMEN kapatip ac)\n"
            "  PowerShell : $env:{env}=\"anahtar\"\n"
            "  macOS/Linux: export {env}=\"anahtar\"".format(env=CH_API_KEY_ENV)
        ),
    )
    subparsers = parser.add_subparsers(dest="command", metavar="komut")

    check = subparsers.add_parser(
        "check",
        help="Kurulumu dogrular: paketler, API anahtari, girdi dosyasi.",
        description="Kurulumu bastan sona dogrular ve nerede takildigini soyler. "
                    "Yeni bir bilgisayarda once bunu calistir.",
    )
    check.add_argument("-i", "--input", default=None, metavar="YOL",
                       help="Verilirse girdi dosyasi da kontrol edilir: okunabiliyor mu, "
                            "zorunlu kolonlar var mi, kac satir analiz edilecek.")
    check.add_argument("--skip-api", "--no-ch", "--dry-run", dest="skip_api",
                       action="store_true",
                       help="Anahtari sinamak icin test istegi gonderme (tamamen offline).")
    check.add_argument("--delimiter", default=None, metavar="KARAKTER",
                       help="CSV ayraci (varsayilan: otomatik tahmin).")
    check.add_argument("--encoding", default=None, metavar="KODLAMA",
                       help="CSV kodlamasi (varsayilan: otomatik tespit).")
    check.add_argument("-v", "--verbose", action="store_true",
                       help="Ayrintili loglama.")

    triage = subparsers.add_parser(
        "triage",
        help="Girdi dosyasini analiz edip sonuc dosyasini yazar.",
        description="Girdi dosyasini analiz edip sonuc dosyasini yazar.",
    )
    triage.add_argument("-i", "--input", default=INPUT_FILE, metavar="YOL",
                        help="Girdi dosyasi (.xlsx / .xlsm / .csv / .tsv). "
                             "Varsayilan: %(default)s")
    triage.add_argument("-o", "--output", default=OUTPUT_FILE, metavar="YOL",
                        help="Cikti dosyasi. Uzantiya gore CSV veya Excel yazilir. "
                             "Varsayilan: %(default)s")
    triage.add_argument("--sheet", default=INPUT_SHEET, metavar="AD",
                        help="Excel sayfa adi (varsayilan: ilk sayfa).")
    triage.add_argument("--delimiter", default=None, metavar="KARAKTER",
                        help="CSV ayraci. Verilmezse baslik satirindan tahmin edilir "
                             "(; , sekme | arasindan).")
    triage.add_argument("--encoding", default=None, metavar="KODLAMA",
                        help="CSV kodlamasi. Verilmezse utf-8-sig, cp1254, cp1252 sirayla denenir.")

    triage.add_argument("-v", "--verbose", action="store_true",
                        help="Her satirin kararini tek tek yazar (DEBUG loglama).")
    triage.add_argument("--debug", action="store_true",
                        help="Cikti dosyasina denetim kolonlarini da ekler.")
    triage.add_argument("--dry-run", "--no-ch", "--skip-api", dest="dry_run",
                        action="store_true",
                        help="Companies House'a hic gitmez; sadece typo analizi yapilir. "
                             "Atlanan satirlara '{}' yazilir.".format(R.CH_SKIPPED))
    triage.add_argument("--limit", type=int, default=MAX_ROWS, metavar="N",
                        help="Sadece ilk N problemli satiri isler. Kota yakmadan deneme icin.")

    triage.add_argument("--mode", choices=["typo_first", "ch_first"], default=LOOKUP_MODE,
                        help="typo_first: once typo kontrolu, temiz satirlar API'ye gider. "
                             "ch_first: once resmi isim alinir, e-posta ona gore denetlenir. "
                             "Varsayilan: %(default)s")
    triage.add_argument("--company-profile", action="store_true", default=FETCH_COMPANY_PROFILE,
                        help="Resmi sirket adi ve dissolved durumunu da ceker "
                             "(sirket basina +1 istek).")
    triage.add_argument("--workers", type=int, default=CH_WORKERS, metavar="N",
                        help="Paralel Companies House thread sayisi. Varsayilan: %(default)s")
    triage.add_argument("--limit-companies", type=int, default=MAX_COMPANIES, metavar="N",
                        help="En fazla N BENZERSIZ regnum sorgulanir. --limit satir sayisini "
                             "sinirlar, bu ise sirket sayisini - kotayi belirleyen budur.")
    triage.add_argument("--max-requests", type=int, default=CH_MAX_REQUESTS, metavar="N",
                        help="Toplam HTTP istegi ust siniri. Asilirsa istek gonderilmez. "
                             "Kotayi korumak icin sert fren.")
    triage.add_argument("--rate", type=float, default=CH_RATE_LIMIT_PER_SEC, metavar="N",
                        help="Saniyedeki istek ust siniri. Resmi limit 600/5dk = 2.0. "
                             "Varsayilan: %(default)s")
    return parser


def apply_cli_args(args):
    """CLI argumanlarini modul ayarlarina uygular."""
    global INPUT_FILE, OUTPUT_FILE, INPUT_SHEET, INPUT_DELIMITER, INPUT_ENCODING
    global DEBUG, DRY_RUN, MAX_ROWS, MAX_COMPANIES, LOOKUP_MODE, FETCH_COMPANY_PROFILE
    global CH_WORKERS, CH_RATE_LIMIT_PER_SEC, CH_MAX_REQUESTS

    INPUT_FILE = args.input
    OUTPUT_FILE = args.output
    INPUT_SHEET = args.sheet
    INPUT_DELIMITER = args.delimiter
    INPUT_ENCODING = args.encoding

    DEBUG = bool(args.debug)
    DRY_RUN = bool(args.dry_run)
    MAX_ROWS = args.limit
    MAX_COMPANIES = args.limit_companies
    LOOKUP_MODE = args.mode
    FETCH_COMPANY_PROFILE = bool(args.company_profile)
    CH_WORKERS = max(1, int(args.workers))
    CH_RATE_LIMIT_PER_SEC = max(0.1, float(args.rate))
    CH_MAX_REQUESTS = args.max_requests

    if args.verbose:
        log.setLevel(logging.DEBUG)
        logging.getLogger().setLevel(logging.DEBUG)
        log.debug("Ayarlar: mode=%s workers=%s rate=%.1f/sn debug=%s dry_run=%s limit=%s",
                  LOOKUP_MODE, CH_WORKERS, CH_RATE_LIMIT_PER_SEC, DEBUG, DRY_RUN, MAX_ROWS)


def cli(argv=None):
    parser = build_arg_parser()
    args = parser.parse_args(argv)

    # Alt komut verilmediyse yardimi goster (Python 3.6'da subparser zorunlu degil)
    if not getattr(args, "command", None):
        parser.print_help()
        return 2

    if args.verbose:
        log.setLevel(logging.DEBUG)
        logging.getLogger().setLevel(logging.DEBUG)

    if args.command == "check":
        try:
            return check_setup(args.input, args.skip_api, args.delimiter, args.encoding)
        except KeyboardInterrupt:
            log.error("Kullanici tarafindan durduruldu.")
            return 130

    apply_cli_args(args)
    try:
        main()
    except (ValueError, IOError, EnvironmentError, CompaniesHouseAuthError) as error:
        log.error("%s", error)
        return 1
    except KeyboardInterrupt:
        log.error("Kullanici tarafindan durduruldu.")
        return 130
    return 0


if __name__ == "__main__":
    sys.exit(cli())
