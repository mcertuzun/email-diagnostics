# -*- coding: utf-8 -*-
"""
Measures the chain behind "I said --limit 10 but 25 requests went out".

Against a fake Companies House server, verifies rows -> distinct companies
-> HTTP requests sent, end to end. No network access.
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
    """A realistic server that returns at most 35 records per page."""
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


def run(rows, officers_per_company, limit=None, limit_companies=None, profile=True):
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
    os.environ["CH_API_KEY"] = "fake"
    argv = ["triage", "-i", src, "-o", out]
    if limit:
        argv += ["--limit", str(limit)]
    if limit_companies:
        argv += ["--limit-companies", str(limit_companies)]
    if not profile:
        argv += ["--no-company-profile"]
    try:
        ED.cli(argv)
    finally:
        ED.CompaniesHouseClient = real_client_cls
        for f in (src, out):
            if os.path.exists(f):
                os.remove(f)
    return len(server.requests)


# 20 rows, 20 distinct regnums, all small companies (3 officers)
ROWS_UNIQUE = [("John", "Smith", "j.smith@acme.co.uk", "Acme Ltd",
                "%08d" % (i + 1), "Bounced") for i in range(20)]
SMALL = dict(("%08d" % (i + 1), 3) for i in range(20))

# 20 rows but only 4 distinct regnums
ROWS_SHARED = [("John", "Smith", "j.smith@acme.co.uk", "Acme Ltd",
                "%08d" % ((i % 4) + 1), "Bounced") for i in range(20)]

# Large companies: 70 officers each, so two pages
BIG = dict(("%08d" % (i + 1), 70) for i in range(20))

failures = []
print("=" * 76)
print(" REQUEST COUNT MEASUREMENT  (server caps pages at 35 records)")
print("=" * 76)
print("%-46s %-10s %s" % ("SCENARIO", "REQUESTS", "RESULT"))
print("-" * 76)

# The company profile call is on by default because it fills the
# companyhouse_names column, so a small company costs 2 requests:
# one profile + one page of officers.
cases = [
    # label, rows, officers, limit, limit_companies, profile, expected
    ("--limit 10, 10 companies, small", ROWS_UNIQUE, SMALL, 10, None, True, 20),
    ("--limit 10, rows share 4 companies", ROWS_SHARED, SMALL, 10, None, True, 8),
    ("--limit 10, 10 companies of 70 officers", ROWS_UNIQUE, BIG, 10, None, True, 30),
    ("--limit-companies 3 (70 officers each)", ROWS_UNIQUE, BIG, None, 3, True, 9),
    ("--limit-companies 5 (small)", ROWS_UNIQUE, SMALL, None, 5, True, 10),
    ("--no-company-profile, 10 small", ROWS_UNIQUE, SMALL, 10, None, False, 10),
    ("--no-company-profile, 10 x 70 officers", ROWS_UNIQUE, BIG, 10, None, False, 20),
]
for label, rows, officers, limit, limit_companies, profile, expected in cases:
    sent = run(rows, officers, limit, limit_companies, profile)
    ok = sent == expected
    if not ok:
        failures.append("%s -> %s requests, expected %s" % (label, sent, expected))
    print("%-46s %-10s %s" % (label, sent, "OK" if ok else "FAIL (expected %d)" % expected))

print()
print("=" * 76)
if failures:
    print("%d FAILURE(S):" % len(failures))
    for f in failures:
        print("   -", f)
    sys.exit(1)
print("QUOTA TESTS PASSED")
