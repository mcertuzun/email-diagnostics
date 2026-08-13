# -*- coding: utf-8 -*-
"""
Proves the pagination behaviour: when Companies House returns FEWER
records than the requested items_per_page, how many requests are sent and
are any records lost?

No network access; requests.Session.get is replaced with a fake server.
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
    Mirrors the real behaviour: items_per_page is a CEILING, not a promise.
    This server returns at most PAGE_CAP records per page whatever is asked.
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
    client = ED.CompaniesHouseClient("fake-key", rate_per_second=10000)
    client._session.get = server.get
    officers = client.get_companies_house_officers("01234567")
    return len(server.calls), len(officers), client.stats["requests"]


print("=" * 74)
print(" PAGINATION TEST  (server returns at most 35 records per page)")
print("=" * 74)
print("%-12s %-10s %-14s %-14s %s" % ("OFFICERS", "REQUESTS", "EXPECTED", "FOUND", "RESULT"))
print("-" * 74)

failures = []
for officer_count, expected_requests in [(3, 1), (35, 1), (36, 2), (70, 2), (80, 3)]:
    calls, found, counted = run(officer_count)
    ok_count = (found == officer_count)
    ok_requests = (calls == expected_requests)
    verdict = "OK" if (ok_count and ok_requests) else "FAIL"
    if not ok_count:
        failures.append("only %d of %d officers were found" % (found, officer_count))
    if not ok_requests:
        failures.append("%d officers took %d requests, expected %d"
                        % (officer_count, calls, expected_requests))
    print("%-12s %-10s %-14s %-14s %s"
          % (officer_count, calls, expected_requests, found, verdict))
    assert counted == calls, "counter mismatch: %s vs %s" % (counted, calls)

print()
print("The request counter matches the real request count in every case.")

# --max-requests sert freni
print()
print("=" * 74)
print(" --max-requests HARD CEILING")
print("=" * 74)
server = FakeServer(500)
client = ED.CompaniesHouseClient("fake-key", rate_per_second=10000, max_requests=3)
client._session.get = server.get
try:
    client.get_companies_house_officers("01234567")
    failures.append("--max-requests ceiling was not enforced")
    print("FAIL: ceiling exceeded")
except ED.LookupFailed as exc:
    sent = len(server.calls)
    ok = sent <= 3
    print("%s Ceiling 3, requests sent: %s  ->  %s"
          % ("OK" if ok else "FAIL", sent, exc))
    if not ok:
        failures.append("ceiling was 3 but %d requests were sent" % sent)

print()
print("=" * 74)
if failures:
    print("%d FAILURE(S):" % len(failures))
    for f in failures:
        print("   -", f)
    sys.exit(1)
print("PAGINATION TESTS PASSED")
