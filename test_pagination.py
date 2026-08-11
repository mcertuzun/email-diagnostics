# -*- coding: utf-8 -*-
"""
Sayfalama hatasini kanitlar: Companies House istenen items_per_page'den
DAHA AZ kayit dondurdugunde kac istek gidiyor ve kayit kaybi var mi?

Ag baglantisi yok - requests.Session.get sahte bir yanit uretecek sekilde
degistirilir.
"""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import email_diagnostics as ED


class FakeResponse(object):
    def __init__(self, payload, status=200):
        self._payload = payload
        self.status_code = status
        self.headers = {}

    def json(self):
        return self._payload


class FakeServer(object):
    """
    Gercek Companies House davranisi: items_per_page bir UST SINIRDIR.
    Bu sunucu istenen ne olursa olsun sayfa basina en fazla PAGE_CAP dondurur.
    """
    PAGE_CAP = 35

    def __init__(self, officer_count):
        self.officer_count = officer_count
        self.calls = []

    def get(self, url, params=None, timeout=None):
        params = params or {}
        start = int(params.get("start_index", 0))
        asked = int(params.get("items_per_page", 35))
        self.calls.append(start)

        page = min(asked, self.PAGE_CAP)
        items = []
        for i in range(start, min(start + page, self.officer_count)):
            items.append({
                "name": "SMITH, Officer%d" % i,
                "name_elements": {"forename": "Officer%d" % i, "surname": "Smith"},
                "officer_role": "director",
            })
        return FakeResponse({
            "items": items,
            "total_results": self.officer_count,
            "items_per_page": page,
            "start_index": start,
        })


def run(officer_count):
    server = FakeServer(officer_count)
    client = ED.CompaniesHouseClient("sahte-anahtar", rate_per_second=10000)
    client._session.get = server.get
    officers = client.get_companies_house_officers("01234567")
    return len(server.calls), len(officers), client.stats["requests"]


print("=" * 74)
print(" SAYFALAMA TESTI  (sunucu sayfa basina en fazla 35 kayit donuyor)")
print("=" * 74)
print("%-12s %-10s %-14s %-14s %s" % ("OFFICER", "ISTEK", "BEKLENEN", "BULUNAN", "SONUC"))
print("-" * 74)

failures = []
for officer_count, expected_requests in [(3, 1), (35, 1), (36, 2), (70, 2), (80, 3)]:
    calls, found, counted = run(officer_count)
    ok_count = (found == officer_count)
    ok_requests = (calls == expected_requests)
    verdict = "OK" if (ok_count and ok_requests) else "HATA"
    if not ok_count:
        failures.append("%d officer'dan sadece %d tanesi bulundu" % (officer_count, found))
    if not ok_requests:
        failures.append("%d officer icin %d istek gitti, %d bekleniyordu"
                        % (officer_count, calls, expected_requests))
    print("%-12s %-10s %-14s %-14s %s"
          % (officer_count, calls, expected_requests, found, verdict))
    assert counted == calls, "sayac tutmuyor: %s vs %s" % (counted, calls)

print()
print("Istek sayaci her vakada gercek istek sayisiyla birebir tutuyor.")

# --max-requests sert freni
print()
print("=" * 74)
print(" --max-requests SERT FRENI")
print("=" * 74)
server = FakeServer(500)
client = ED.CompaniesHouseClient("sahte-anahtar", rate_per_second=10000, max_requests=3)
client._session.get = server.get
try:
    client.get_companies_house_officers("01234567")
    failures.append("--max-requests siniri uygulanmadi")
    print("HATA: sinir asildi")
except ED.LookupFailed as exc:
    sent = len(server.calls)
    ok = sent <= 3
    print("%s Sinir 3 iken gonderilen istek: %s  ->  %s"
          % ("OK" if ok else "HATA", sent, exc))
    if not ok:
        failures.append("sinir 3 iken %d istek gitti" % sent)

print()
print("=" * 74)
if failures:
    print("%d SORUN:" % len(failures))
    for f in failures:
        print("   -", f)
    sys.exit(1)
print("SAYFALAMA TESTLERI GECTI")
