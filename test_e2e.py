# -*- coding: utf-8 -*-
"""Ucdan uca test: gercek .xlsx olustur, pipeline'i calistir, ciktiyi dogrula."""
import csv
import io
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
os.chdir(HERE)

import openpyxl
import email_diagnostics as ED

# ------------------------------------------------------------------
# 1) Gercekci bir girdi dosyasi olustur (kirli veriyle)
# ------------------------------------------------------------------
SRC = os.path.join(HERE, "contacts_test.xlsx")
OUT = os.path.join(HERE, "out_test.xlsx")

wb = openpyxl.Workbook()
ws = wb.active
# Bilerek kirli basliklar: bosluk, buyuk harf, BOM
ws.append(["﻿First Name ", "last_name", "Email", "company", "regnum", " status"])
ROWS = [
    ("Mr John",        "Smith",         "jhon.smith@acme.co.uk",  "Acme Trading Ltd",  "01234567", "Bounced"),
    ("Dr. Jane (Janie)", "Doe MBE",     "jane.doe@acme.co.uk",    "Acme Trading Ltd",  1234567,    "bounced back"),
    ("",               "Robert Brown",  "r.brown@acmee.co.uk",    "Acme Trading Ltd",  "01234567", "Delivery Failed"),
    ("Sarah",          "Jones",         "info@acme.co.uk",        "Acme Trading Ltd",  "SC123456", "hard bounce"),
    ("Peter",          "Nobody",        "",                       "Acme Trading Ltd",  "01234567", "undelivered"),
    ("Ali",            "Veli",          "not-an-email",           "Ali Veli Zeynep Ltd", "07654321", "failed"),
    ("Mehmet",         "Ozturk",        "m.ozturk@avz.co.uk",     "Ali Veli Zeynep Ltd", "07654321", "bounced"),
    ("Kate",           "Wilson",        "kate.wilson@acme.co.uk", "Acme Trading Ltd",  "01234567", "Delivered"),   # filtrelenmeli
    ("Tom",            "Baker",         "tom.baker@acme.co.uk",   "Acme Trading Ltd",  "01234567", "Blocked"),     # filtrelenmeli
]
for row in ROWS:
    ws.append(list(row))
wb.save(SRC)
print("Girdi olusturuldu: %s (%d veri satiri)" % (os.path.basename(SRC), len(ROWS)))

# ------------------------------------------------------------------
# 2) Ayarlari test icin degistir ve calistir
# ------------------------------------------------------------------
ED.INPUT_FILE = SRC
ED.OUTPUT_FILE = OUT
ED.DRY_RUN = True          # Companies House cagrilmaz
ED.DEBUG = True            # denetim kolonlari da yazilsin
if os.path.exists(OUT):
    os.remove(OUT)

before_mtime = os.path.getmtime(SRC)
before_size = os.path.getsize(SRC)

print("-" * 78)
ED.main()
print("-" * 78)

# ------------------------------------------------------------------
# 3) Dogrulamalar
# ------------------------------------------------------------------
problems = []

# a) Girdi dosyasi degismemis olmali
if os.path.getmtime(SRC) != before_mtime or os.path.getsize(SRC) != before_size:
    problems.append("GIRDI DOSYASI DEGISTIRILDI!")
else:
    print("OK   girdi dosyasi degistirilmedi")

# b) Sadece TEK cikti dosyasi olmali
xlsx_files = sorted(f for f in os.listdir(HERE) if f.endswith(".xlsx"))
if xlsx_files != ["contacts_test.xlsx", "out_test.xlsx"]:
    problems.append("Beklenmeyen dosyalar: %s" % xlsx_files)
else:
    print("OK   sadece 1 cikti dosyasi olusturuldu")

# c) Ciktiyi oku
wb2 = openpyxl.load_workbook(OUT)
ws2 = wb2.active
rows = list(ws2.iter_rows(values_only=True))
header = list(rows[0])
data = [list(r) for r in rows[1:]]

print("OK   cikti kolonlari: %s" % ", ".join(str(h) for h in header[:10]))

# d) Orijinal kolonlar korunmus mu (BOM/bosluk dahil aynen)
if [str(h).strip() for h in header[:6]] != ["First Name", "last_name", "Email", "company", "regnum", "status"]:
    problems.append("Orijinal basliklar korunmadi: %r" % header[:6])
else:
    print("OK   orijinal kolon basliklari aynen korundu")

# e) Filtreleme: 9 satirdan 7'si kalmali (Delivered + Blocked disarida)
if len(data) != 7:
    problems.append("Beklenen 7 satir, gelen %d" % len(data))
else:
    print("OK   statu filtresi: 9 -> 7 satir (Delivered ve Blocked disarida)")

result_idx = header.index("result")
reason_idx = header.index("result_reason")
regnum_dbg_idx = header.index("dbg_regnum_used")

expected = {
    "jhon.smith@acme.co.uk": ED.R.FIRST_NAME_TYPO,
    "not-an-email":          ED.R.MALFORMED_EMAIL,
    "":                      ED.R.MISSING_EMAIL,
    "r.brown@acmee.co.uk":   ED.R.DOMAIN_TYPO,
}
email_idx = 2
print()
print("%-26s %-24s %-30s %s" % ("EMAIL", "RESULT", "REASON", "REGNUM"))
print("-" * 100)
for row in data:
    email = row[email_idx] or ""
    print("%-26s %-24s %-30s %s" % (email[:26], str(row[result_idx])[:24],
                                    str(row[reason_idx])[:30], row[regnum_dbg_idx]))
    if email in expected and row[result_idx] != expected[email]:
        problems.append("%r -> beklenen %r, gelen %r" % (email, expected[email], row[result_idx]))

print()
# f) regnum sifir dolgusu: Excel'de 1234567 (sayi) yazan satir 01234567 olmali
regnums = set(r[regnum_dbg_idx] for r in data)
if "1234567" in regnums:
    problems.append("regnum sifir dolgusu calismadi: %s" % regnums)
else:
    print("OK   regnum sifir dolgusu calisti: %s" % sorted(regnums))

# g) Hicbir satir bos result ile kalmamali
if any(not r[result_idx] for r in data):
    problems.append("Bos result iceren satir var")
else:
    print("OK   her satirin bir result degeri var")

# ------------------------------------------------------------------
# 4) CSV yolu + komut satiri arayuzu
#    Windows tarzi zor bir CSV: ';' ayrac, cp1254 kodlama, Turkce karakter
# ------------------------------------------------------------------
print()
print("=" * 78)
print("CSV + CLI TESTI")
print("=" * 78)

CSV_IN = os.path.join(HERE, "liste_test.csv")
CSV_OUT = os.path.join(HERE, "liste_test_sonuc.csv")
csv_rows = [
    "first_name;last_name;email;company;regnum;status",
    "Mr John;Smith;jhon.smith@acme.co.uk;Acme Trading Ltd;01234567;Bounced",
    "Mehmet;Öztürk;m.ozturk@avz.co.uk;Ali Veli Zeynep Ltd;07654321;bounced back",
    "Şükrü;Yılmaz;info@acme.co.uk;Acme Trading Ltd;01234567;hard bounce",
    "Kate;Wilson;kate.wilson@acme.co.uk;Acme Trading Ltd;01234567;Delivered",
]
io.open(CSV_IN, "w", encoding="cp1254", newline="").write("\r\n".join(csv_rows) + "\r\n")

exit_code = ED.cli(["triage", "--input", CSV_IN, "--output", CSV_OUT,
                    "--verbose", "--dry-run"])
if exit_code != 0:
    problems.append("CLI sifirdan farkli cikis kodu dondurdu: %s" % exit_code)
else:
    print("OK   CLI cikis kodu 0")

with io.open(CSV_OUT, "r", encoding="utf-8-sig", newline="") as handle:
    out_rows = list(csv.reader(handle, delimiter=";"))

if len(out_rows) != 4:              # 1 baslik + 3 problemli satir
    problems.append("CSV cikti satir sayisi: beklenen 4, gelen %d" % len(out_rows))
else:
    print("OK   ';' ayraci otomatik algilandi ve ciktida korundu")

if out_rows[0][:6] != ["first_name", "last_name", "email", "company", "regnum", "status"]:
    problems.append("CSV basliklari korunmadi: %r" % out_rows[0][:6])
else:
    print("OK   CSV orijinal kolonlari korundu")

turkish = [r[1] for r in out_rows[1:]]
if "Öztürk" not in turkish or "Yılmaz" not in turkish:
    problems.append("Turkce karakterler bozuldu: %r" % turkish)
else:
    print("OK   cp1254 -> utf-8-sig donusumunde Turkce karakterler korundu")

csv_result_idx = out_rows[0].index("result")
if out_rows[1][csv_result_idx] != ED.R.FIRST_NAME_TYPO:
    problems.append("CSV teshisi hatali: %r" % out_rows[1][csv_result_idx])
else:
    print("OK   CSV yolunda teshis dogru uretildi")

os.remove(CSV_IN)
os.remove(CSV_OUT)

print()
print("=" * 78)
if problems:
    print("%d SORUN:" % len(problems))
    for p in problems:
        print("   -", p)
    sys.exit(1)
print("UCDAN UCA TEST GECTI")
