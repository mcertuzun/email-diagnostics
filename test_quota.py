# -*- coding: utf-8 -*-
"""
'--limit 10 dedim ama 25 istek gitti' sorusunun olculmesi.

Sahte bir Companies House sunucusuyla, satir sayisi / benzersiz sirket sayisi
/ gonderilen HTTP istegi zincirini uctan uca dogrular. Ag baglantisi yok.
"""
import io
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
os.chdir(HERE)

import openpyxl
import email_diagnostics as ED


class FakeResponse(object):
    def __init__(self, payload, status=200):
        self._payload, self.status_code, self.headers = payload, status, {}

    def json(self):
        return self._payload


class FakeCH(object):
    """Sayfa basina en fazla 35 kayit donen, gercege yakin bir sunucu."""
    PAGE_CAP = 35

    def __init__(self, officers_per_company):
        self.officers_per_company = officers_per_company
        self.requests = []

    def get(self, url, params=None, timeout=None):
        params = params or {}
        number = url.rstrip("/").split("/")[-2] if url.endswith("officers") else "?"
        self.requests.append(url)
        count = self.officers_per_company.get(number, 3)
        start = int(params.get("start_index", 0))
        page = min(int(params.get("items_per_page", 35)), self.PAGE_CAP)
        items = [{"name_elements": {"forename": "Officer%d" % i, "surname": "Smith"}}
                 for i in range(start, min(start + page, count))]
        return FakeResponse({"items": items, "total_results": count,
                             "items_per_page": page, "start_index": start})


def build_input(path, rows):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["first_name", "last_name", "email", "company", "regnum", "status"])
    for row in rows:
        ws.append(list(row))
    wb.save(path)


def run(rows, officers_per_company, limit=None, limit_companies=None):
    src = os.path.join(HERE, "q_in.xlsx")
    out = os.path.join(HERE, "q_out.xlsx")
    build_input(src, rows)

    server = FakeCH(officers_per_company)
    real_client_cls = ED.CompaniesHouseClient
    captured = {}

    class PatchedClient(real_client_cls):
        def __init__(self, *a, **kw):
            real_client_cls.__init__(self, *a, **kw)
            self._session.get = server.get
            captured["client"] = self

    ED.CompaniesHouseClient = PatchedClient
    os.environ["CH_API_KEY"] = "sahte"
    argv = ["triage", "-i", src, "-o", out]
    if limit:
        argv += ["--limit", str(limit)]
    if limit_companies:
        argv += ["--limit-companies", str(limit_companies)]
    try:
        ED.cli(argv)
    finally:
        ED.CompaniesHouseClient = real_client_cls
        for f in (src, out):
            if os.path.exists(f):
                os.remove(f)
    return len(server.requests)


# 20 satir, 20 farkli regnum, hepsi kucuk sirket (3 officer)
ROWS_UNIQUE = [("John", "Smith", "j.smith@acme.co.uk", "Acme Ltd",
                "%08d" % (i + 1), "Bounced") for i in range(20)]
SMALL = dict(("%08d" % (i + 1), 3) for i in range(20))

# 20 satir ama sadece 4 farkli regnum
ROWS_SHARED = [("John", "Smith", "j.smith@acme.co.uk", "Acme Ltd",
                "%08d" % ((i % 4) + 1), "Bounced") for i in range(20)]

# 10 farkli regnum, hepsi BUYUK sirket (70 officer -> 2 sayfa)
BIG = dict(("%08d" % (i + 1), 70) for i in range(20))

failures = []
print("=" * 76)
print(" ISTEK SAYISI OLCUMU  (sunucu sayfa basina en fazla 35 kayit donuyor)")
print("=" * 76)
print("%-46s %-10s %s" % ("SENARYO", "ISTEK", "SONUC"))
print("-" * 76)

cases = [
    ("--limit 10, 10 farkli sirket, kucuk", ROWS_UNIQUE, SMALL, 10, None, 10),
    ("--limit 10, satirlar 4 sirketi paylasiyor", ROWS_SHARED, SMALL, 10, None, 4),
    ("--limit 10, 10 farkli sirket, 70 officer", ROWS_UNIQUE, BIG, 10, None, 20),
    ("--limit-companies 3 (buyuk sirketler)", ROWS_UNIQUE, BIG, None, 3, 6),
    ("--limit-companies 5 (kucuk sirketler)", ROWS_UNIQUE, SMALL, None, 5, 5),
]
for label, rows, officers, limit, limit_companies, expected in cases:
    sent = run(rows, officers, limit, limit_companies)
    ok = sent == expected
    if not ok:
        failures.append("%s -> %s istek, %s bekleniyordu" % (label, sent, expected))
    print("%-46s %-10s %s" % (label, sent, "OK" if ok else "HATA (beklenen %d)" % expected))

print()
print("=" * 76)
if failures:
    print("%d SORUN:" % len(failures))
    for f in failures:
        print("   -", f)
    sys.exit(1)
print("KOTA TESTLERI GECTI")
