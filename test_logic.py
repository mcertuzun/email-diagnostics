# -*- coding: utf-8 -*-
"""Functional test harness for email_diagnostics.py. No network calls."""
import os
import sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import email_diagnostics as ED

FAIL = []

def check(label, got, want):
    ok = got == want
    if not ok:
        FAIL.append("%s -> got %r want %r" % (label, got, want))
    print("%-4s %-52s %s" % ("OK" if ok else "FAIL", label, got))


print("=" * 78)
print("1) NAME CLEANING")
print("=" * 78)

cases = [
    ("Mr John", "Smith",            "john", "smith"),
    ("Dr. Jane (Janie)", "Doe MBE", "jane", "doe"),
    ("", "John Smith",              "john", "smith"),
    ("John Smith", "",              "john", "smith"),
    ("Robert, CEO", "van der Berg", "robert", "vanderberg"),
    ("J. Michael", "O'Brien",       "michael", "obrien"),
    ("Mrs", "SMITH, John",          "john", "smith"),
    ("Ahmet", "Sukru Ozturk",       "ahmet", "ozturk"),
]
for f, l, want_first, want_surname in cases:
    nd = ED.resolve_person_name(f, l)
    check("first_name=%r last_name=%r" % (f, l),
          (nd["first"], nd["surname"]), (want_first, want_surname))

nd = ED.resolve_person_name("Mr John", "Smith")
print("     nickname expansion for 'john':", sorted(nd["first_candidates"]))
assert "jack" in nd["first_candidates"], "nickname map broken"

print()
print("=" * 78)
print("2) COMPANY -> DOMAIN CANDIDATES")
print("=" * 78)
tok = ED.clean_company_name("Ali Veli Zeynep Trading Ltd.")
check("clean_company_name", tok, ["ali", "veli", "zeynep"])
cand = ED.generate_company_domain_candidates(tok)
for expected in ["alivelizeynep", "avz", "ali", "aliveli"]:
    check("candidate contains %r" % expected, expected in cand, True)

print()
print("=" * 78)
print("3) EMAIL PARSING")
print("=" * 78)
check("parse ' John.Smith@Acme.CO.UK '", ED.parse_email(" John.Smith@Acme.CO.UK ")["email"],
      "john.smith@acme.co.uk")
check("brand of acme.co.uk", ED.parse_email("a@acme.co.uk")["brand"], "acme")
check("brand of mail.acme.co.uk", ED.parse_email("a@mail.acme.co.uk")["brand"], "acme")
check("brand of acme.com", ED.parse_email("a@acme.com")["brand"], "acme")
check("malformed 'john.smith'", ED.parse_email("john.smith")["status"], "malformed")
check("malformed 'a@@b.com'", ED.parse_email("a@@b.com")["status"], "malformed")
check("missing ''", ED.parse_email("")["status"], "missing")

print()
print("=" * 78)
print("4) TYPO DETECTION  (name: John Smith @ Acme Trading Ltd)")
print("=" * 78)
nd = ED.resolve_person_name("Mr John", "Smith")
cc = ED.generate_company_domain_candidates(ED.clean_company_name("Acme Trading Ltd"))

typo_cases = [
    ("john.smith@acme.co.uk",   None,               ED.RSN.PATTERN_OK,  "exact match"),
    ("johnsmith@acme.co.uk",    None,               ED.RSN.PATTERN_OK,  "separator only -> NOT a typo"),
    ("j.smith@acme.co.uk",      None,               ED.RSN.PATTERN_OK,  "initial pattern"),
    ("jack.smith@acme.co.uk",   None,               ED.RSN.PATTERN_OK,  "nickname Jack=John"),
    ("jhon.smith@acme.co.uk",   ED.R.FIRST_NAME_TYPO, None,             "forename typo"),
    ("john.smiith@acme.co.uk",  ED.R.SURNAME_TYPO,  None,               "surname typo"),
    ("jhon.smiht@acme.co.uk",   ED.R.BOTH_TYPO,     None,               "both parts"),
    ("john.smith@acmee.co.uk",  ED.R.DOMAIN_TYPO,   None,               "domain typo"),
    ("info@acme.co.uk",         None,               ED.RSN.GENERIC,     "generic mailbox"),
    ("john.smith@gmail.com",    None,               ED.RSN.PERSONAL_DOMAIN, "personal domain"),
    ("xq7z@acme.co.uk",         None,               ED.RSN.UNRECOGNISED, "short/unrelated -> NOT a typo"),
    ("john.smith@othercorp.com", None,              ED.RSN.DOMAIN_NOT_MATCHED, "different domain -> NOT a typo"),
]
for email, want_result, want_reason, note in typo_cases:
    info = ED.parse_email(email)
    out = ED.detect_email_typo(info, nd, cc)
    got = out["result"] if out["terminal"] else None
    label = "%-28s (%s)" % (email, note)
    if want_result is not None:
        check(label, got, want_result)
    else:
        check(label, (got, out["reason"]), (None, want_reason))

print()
print("     -- the 'js@' case: Jasmine Susanne Smith --")
nd2 = ED.resolve_person_name("Jasmine Susanne", "Smith")
out = ED.detect_email_typo(ED.parse_email("js@acme.co.uk"), nd2,
                           ED.generate_company_domain_candidates(ED.clean_company_name("Acme Ltd")))
check("js@acme.co.uk  (initials, NOT a typo)", (out["terminal"], out["reason"]),
      (False, ED.RSN.PATTERN_OK))

print()
print("=" * 78)
print("5) OFFICER MATCHING  (fake Companies House response)")
print("=" * 78)

officers = [
    {"name": "SMITH, John Andrew",
     "name_elements": {"forename": "John", "other_forenames": "Andrew", "surname": "Smith"},
     "officer_role": "director", "appointed_on": "2015-01-01"},
    {"name": "JONES, Sarah",
     "name_elements": {"forename": "Sarah", "surname": "Jones"},
     "officer_role": "director", "resigned_on": "2022-06-30"},
    {"name": "BIG CORP LIMITED", "officer_role": "corporate-director",
     "identification": {"legal_form": "ltd"}},
    {"name": "TAYLOR, Elizabeth",
     "name_elements": {"forename": "Elizabeth", "surname": "Taylor"},
     "former_names": [{"forenames": "Elizabeth", "surname": "Brown"}],
     "officer_role": "director"},
]

m = ED.match_contact_to_officers(ED.resolve_person_name("John", "Smith"), officers)
check("John Smith -> active, with middle name", (m["status"], m["officer_name"]),
      ("active", "John Andrew Smith"))
check("  reason", m["reason"], ED.RSN.WITH_MIDDLE_NAME)

m = ED.match_contact_to_officers(ED.resolve_person_name("Sarah", "Jones"), officers)
check("Sarah Jones -> resigned", m["status"], "resigned")

m = ED.match_contact_to_officers(ED.resolve_person_name("Michael", "Jones"), officers)
check("Michael Jones -> surname only -> possible", m["status"], "possible_resigned")
check("  reason", m["reason"], ED.RSN.SURNAME_ONLY)

m = ED.match_contact_to_officers(ED.resolve_person_name("Liz", "Brown"), officers)
check("Liz Brown -> found via former_names and nickname",
      (m["status"], m["officer_name"]), ("active", "Elizabeth Taylor"))

m = ED.match_contact_to_officers(ED.resolve_person_name("Peter", "Nobody"), officers)
check("Peter Nobody -> no match", m["status"], "none")

m = ED.match_contact_to_officers(ED.resolve_person_name("Big", "Corp"), officers)
check("corporate officer does not break parsing", m["status"], "none")

print()
print("     -- same person resigned then reappointed: ACTIVE must win --")
dual = [
    {"name_elements": {"forename": "Mark", "surname": "White"}, "resigned_on": "2019-01-01"},
    {"name_elements": {"forename": "Mark", "surname": "White"}},
]
m = ED.match_contact_to_officers(ED.resolve_person_name("Mark", "White"), dual)
check("Mark White -> active", m["status"], "active")

print()
print("     -- family company: two different John Smiths -> ambiguous --")
family = [
    {"name_elements": {"forename": "John", "surname": "Smith"}},
    {"name_elements": {"forename": "John", "surname": "Smith", "other_forenames": "Peter"}},
]
m = ED.match_contact_to_officers(ED.resolve_person_name("John", "Smith"), family)
check("ambiguity flagged", m["reason"], ED.RSN.MULTIPLE_OFFICERS)
check("  confidence downgraded", m["confidence"], "possible")

print()
print("=" * 78)
print("6) REGNUM NORMALISATION")
print("=" * 78)
check("1234567 (Excel dropped the zero)", ED.normalize_regnum("1234567"), "01234567")
check("01234567", ED.normalize_regnum("01234567"), "01234567")
check("SC123456", ED.normalize_regnum("sc123456"), "SC123456")
check("with spaces ' 12 34 567 '", ED.normalize_regnum(" 12 34 567 "), "01234567")
check("empty", ED.normalize_regnum(""), "")

print()
print("=" * 78)
print("7) STATUS FILTER")
print("=" * 78)
for status, want in [("Bounced", True), ("bounced back", True), ("Hard Bounce", True),
                     ("DELIVERY FAILED", True), ("undelivered", True), ("failed", True),
                     ("Delivered", False), ("Opened", False), ("Blocked", False),
                     ("blocked - spam", False), ("", False)]:
    check("status=%r" % status, ED.is_problematic_status(status), want)

print()
print("=" * 78)
if FAIL:
    print("%d TEST FAILURE(S):" % len(FAIL))
    for f in FAIL:
        print("   -", f)
    sys.exit(1)
print("ALL TESTS PASSED")
