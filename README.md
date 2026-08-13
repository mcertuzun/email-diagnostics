# email-diagnostics

Analyses bounced email rows and writes **one diagnosis per row**. It first verifies the contact against **Companies House** using the company number, then checks whether the email address is consistent with the verified name and the company.

The output is **a single file**. The input is never modified.

---

## Hard constraint: no mail server is ever contacted

No SMTP handshake, no `RCPT TO` probe, no verification service, not even a DNS/MX lookup. The Companies House REST API is the only outbound connection.

The direct consequence: **an address can never be proven valid.** The `result` column is the most likely explanation, not proof.

---

## Install

```bash
pip install -r requirements.txt
```

Two packages: `openpyxl` and `requests`.

pandas and numpy are deliberately avoided. They are awkward to install on Python 3.6 under Windows, and pandas turns `regnum` into a number and drops the leading zeros (`01234567` → `1234567`), which becomes a silent 404 from Companies House. There is no fuzzy-matching library either; the Damerau-Levenshtein distance is implemented in the script, so no compiler is needed.

Runs on Python 3.6.5 and above. No `dataclasses`, f-strings or walrus operators.

## API key

The key is **never written into the code**. It is looked for, in order:

1. the `CH_API_KEY` environment variable
2. a `.env` file in the working directory
3. a `.env` file next to the script

| Shell | Command |
|---|---|
| Windows, persistent | `setx CH_API_KEY "your_key"` |
| Windows, session | `set CH_API_KEY=your_key` — no quotes |
| PowerShell | `$env:CH_API_KEY="your_key"` |
| macOS / Linux | `export CH_API_KEY="your_key"` |

> **The most common Windows problem:** `setx` only affects **newly started** processes. If you use VS Code, opening a new terminal tab is not enough — VS Code inherits its environment at launch, so you must **close and reopen VS Code itself**.

The simpler alternative is a `.env` file next to the script. No restart needed, and it is in `.gitignore` so it never reaches the repository.

```
CH_API_KEY=your_key
```

> **The second most common Windows problem:** Notepad's *Save As* silently appends `.txt`, and Explorer hides extensions, so a file you see as `.env` is `.env.txt` on disk. The script accepts `.env.txt` and `env.txt` as well and warns you which file it used. To create it cleanly in PowerShell:
>
> ```powershell
> Set-Content -Path .env -Value "CH_API_KEY=your_key"
> ```

If no key is found, the script reports **what it actually found** rather than just failing:

```
[FAIL] No key found
  [none] CH_API_KEY is not set
  [!]    C:\Users\...\.env exists but has no CH_API_KEY.
         Keys it does contain: COMPANIES_HOUSE_KEY
```

Get a key from the [Companies House Developer Hub](https://developer.company-information.service.gov.uk/) by creating a **REST API** application.

---

## Which version am I running?

```bash
python email_diagnostics.py --version
```

An unexpected `unrecognized arguments` error for a subcommand that exists almost always means a stale file. Pull and check the version.

## Verify the setup first

On a new machine, run this before anything else:

```bash
python email_diagnostics.py check
```

It checks the Python version, the packages, the TLS environment and the API key, then sends **one** request to Companies House to prove the key is actually accepted. If something is missing it says exactly what and how to fix it.

You can have it check your input file too:

```bash
python email_diagnostics.py check --input list.csv
```

```
[OK]   File read: 1240 rows
[OK]   All required columns are present.
[OK]   Rows that will be analysed: 183 / 1240
      Status distribution:
        Delivered                      890  -
        Opened                         167  -
        Bounced                        142  analysed
        Blocked                         41  -
        hard bounce                     41  analysed
      Warning: regnum is empty on 3 rows.
```

`--skip-api` keeps it fully offline.

## Explore one company

```bash
python email_diagnostics.py inspect 17107304
```

Lists every field Companies House returns for that company number, marking with `*` the ones the diagnosis actually uses. `--raw` prints the untouched JSON. Costs two requests.

---

## Run

```bash
python email_diagnostics.py triage --input list.csv --output result.csv --verbose
```

On Windows:

```bash
python email_diagnostics.py triage --input "C:\Users\you\Downloads\list.csv" --output "C:\Users\you\Downloads\result.csv" --verbose
```

Input and output may each be `.xlsx` or `.csv`, chosen by extension, and they do not have to match — read CSV, write Excel if you like.

### Flags

| Flag | Effect |
|---|---|
| `-i`, `--input` | Input file (`.xlsx` / `.xlsm` / `.csv` / `.tsv`) |
| `-o`, `--output` | Output file; CSV or Excel by extension |
| `-v`, `--verbose` | Log the decision for every row |
| `--debug` | Add 11 audit columns to the output |
| `--plain` | Excel only: no sorting, colouring or extra sheets; keeps the input row order |
| `--dry-run` (`--no-ch`, `--skip-api`) | Never call Companies House; run the email analysis only. Skipped rows are marked `companies_house_skipped` |
| `--limit N` | Process only the first N bounced **rows** |
| `--limit-companies N` | Query at most N distinct **regnums** — this is what costs quota |
| `--max-requests N` | Hard ceiling on total HTTP requests; once reached, nothing further is sent |
| `--mode` | `ch_first` (default) or `typo_first` |
| `--no-company-profile` | Skip the profile request. Saves one request per company, but leaves `companyhouse_names` empty and cannot detect dissolved companies |
| `--sheet NAME` | Excel sheet name (default: the first sheet) |
| `--delimiter` | CSV delimiter — guessed from the header row if omitted |
| `--encoding` | CSV encoding — tries `utf-8-sig`, `cp1254`, `cp1252` in turn |
| `--ca-bundle PATH` | Corporate root certificate, for networks that intercept TLS |
| `--insecure` | Turn off certificate verification. Unsafe; for diagnosis only |
| `--workers N` | Parallel threads (default 4) |
| `--rate N` | Requests per second (default 1.8) |

A good first run, which spends no quota at all:

```bash
python email_diagnostics.py triage -i list.csv -o trial.csv --limit 50 --no-ch --debug --verbose
```

### Input file

First row is the header. These columns are required; case, spacing and BOM differences are tolerated:

`first_name` · `last_name` · `email` · `company` · `regnum` · `status`

`regnum` is the Companies House company number. `status` is the delivery status.

European Excel exports work as they are: a `;` delimiter and `cp1254`/`cp1252` encoding are detected automatically, and the output is written as `utf-8-sig` so Excel renders accented characters correctly.

> **Note:** if you open the CSV output in Excel, Excel may again turn `01234567` into a number and hide the leading zero. That is Excel's behaviour, not a problem with the file — write `.xlsx` output if it bothers you.

---

## Pipeline

1. **Load and validate** — a missing required column produces a clear error naming exactly which ones are missing.
2. **Status filter** — only the bounce family is processed. `blocked` is excluded deliberately: it comes from spam filtering or IP reputation and says nothing about a wrong address or a departed person.
3. **Clean** — titles (`Mr`, `Dr`), role words (`CEO`, `Director`) and UK post-nominals (`MBE`, `FCA`) are stripped from **both** name fields; bracketed nicknames (`John (Jack)`) are kept as candidates; a `last_name` holding a full name (`John Smith`) is split; surname particles (`van der`, `Mc`) are preserved; legal suffixes are removed from the company name. **The original values are never modified** — cleaning is only for matching.
4. **Companies House** — the officer list is fetched for each row and the contact is matched **surname first**, yielding the official full name including middle names.
5. **Email check** — the address is judged against that verified name.
6. **One output file**, plus a result distribution printed to the console.

### The two modes

**`ch_first` (default).** Every row goes to Companies House: **N rows = N lookups**, with each distinct regnum queried only once. The person's official name is fetched and the email judged against it. The accuracy gain is concrete — for a contact stored as `John Smith` whose address is `john.andrew.smith@`:

| Mode | Result |
|---|---|
| `typo_first` | `email_pattern_unrecognised` — the middle name is unknown |
| `ch_first` | `email_matches_expected_pattern` — Companies House supplied `Andrew` |

**`typo_first`.** Typo check first; a row with a typo never reaches the API. Cheaper, but the name check relies on whatever the input file contains.

### Result priority (`ch_first`)

`data problem` (missing/malformed email) → `company closed / API failure` → **`resigned`** → `typo` → `active`

Resignation outranks a typo: once someone has left, a correctly spelled address bounces too. `result_reason` carries the email verdict on every row, so you get both facts.

### How typo detection works

Patterns like `first.last`, `flast`, `f.last`, `lastfirst` and `f.m.last` are generated from the forename, surname, middle names, nicknames and initials, across four separators (`.`, `_`, `-`, and none). Nicknames expand both ways through a dictionary (`Bob`↔`Robert`, `Jack`↔`John`, `Liz`↔`Elizabeth` …).

The comparison is deliberately **conservative** — a weak difference is never called a typo:

| Situation | Verdict |
|---|---|
| Exact match | consistent |
| Separator difference only (`johnsmith` ≡ `john.smith`) | consistent, **not** a typo |
| Nickname match (`jack.smith` for John Smith) | consistent |
| 1–2 characters apart (`jhon.smith`) | **typo** |
| Nowhere near any pattern | `email_pattern_unrecognised` — **not** a typo, continues to CH |
| Local part shorter than 5 characters (`js@`) | never a typo; initial combinations are tried |

The distance is **Damerau-Levenshtein**: the most common typing mistake is two letters swapping places (`jhon` for `john`), and plain Levenshtein scores that as 2 edits and misses the threshold.

### Domain matching

Checked in order: exact match → containment with a real affix (`acme` → `acmegroup` ✔) → 1–2 characters apart (`acmee` ✘).

Acronyms are generated for companies whose name has **two or more words**, on their own and followed by a company word: *Ali Veli Zeynep Ltd* → `avz`, `avzltd`, `avzgroup`, `avzuk`, `avzholdings`, alongside `alivelizeynep` and `aliveli`. A single-word company is excluded, since a one-letter acronym would match almost anything.

Acronym forms are used for **exact matching only, never for distance comparison**. `avzgroup` and `xyzgroup` are two characters apart but are unrelated companies, so allowing acronyms into the fuzzy pass would report any similar acronym as a domain typo.

An unrelated domain is **not** called a typo — it may be a group or parent-company domain.

### Officer matching

The surname is the anchor: it must match first (exactly, or within one edit), and only then is the forename considered. The structured `name_elements` field is used rather than string splitting, and `former_names` is searched too, which is what catches surnames changed by marriage.

- surname **and** forename match → **confident**
- initial only, or surname only → **possible**
- several officers share the surname (common in family companies) → flagged ambiguous, confidence downgraded
- the same person appears as both resigned and active → **active wins**
- corporate officers (`corporate-director`) are skipped

---

## The output workbook

An `.xlsx` output has three sheets. This is still **one file** — the original constraint was one file, not one sheet.

**`Results`** — one row per contact. Sorted by `action` so the work queues come first and `investigate-all-correct` sits at the bottom, since those are the rows you do not need to read. The header is frozen and filterable, columns are sized to their content, and the `action` cell is colour-coded. `source_row` records the original position in the input, so sorting never costs you traceability.

**`Summary`** — the action, result and reason distributions plus the run metadata: when it ran, which mode, how many rows, how many companies, how many HTTP requests. The console prints this too, but there it scrolls away.

**Work queue sheets** — one per queue that has rows: `Fix address (23)`, `Find new contact (11)`, `Investigate (7)`, `Fix data (2)`. The point is that a sheet can be handed to one person: the address corrections go to whoever maintains the list, the replacements go to whoever does the research. Filtering cannot do that.

These are narrow **views** of `Results`, not copies of it. Each carries only the columns its queue needs — `Find new contact` has the company and the suggestions, not the email pattern reasoning — and `source_row` is the first column on every one, so any line maps straight back. That also settles which sheet is authoritative: edit `Results`.

Two deliberate choices:

- **`investigate-all-correct` is not in the `Investigate` sheet.** It is usually the largest group and there is nothing to do with it, so including it would bury the rows that do need looking at. `Investigate` holds the three that need a human: `investigate-mismatched`, `investigate-uncertain-match`, `investigate-non-company-domain`. The all-correct rows stay in `Results`.
- **A queue with no rows gets no sheet**, and the row count is in the tab name. The tab bar then tells you the workload without opening anything.

**`Companies`** — one row per company, not per contact. A different grain, not a copy: regnum, official name, status, how many contacts bounced there and how many of them have gone, with the active officers. Sorted so the companies with the most departures come first. This is the sheet that answers "which companies do I need to re-contact wholesale".

`--plain` turns all of it off — one sheet, input order, no colour — if you are feeding the output into something else.

CSV output cannot carry sheets or formatting. It still gets `action` and `source_row`; if you want the rest, write `.xlsx`.

## Output columns

Every original column is preserved, with these appended:

| Column | Contents |
|---|---|
| `action` | What to do about the row — see below |
| `result` | The primary diagnosis |
| `result_reason` | A short supporting reason |
| `ch_officer_name` | The official full name from Companies House, middle names included |
| `ch_officer_status` | `active` / `resigned` / `possible_active` / `possible_resigned` / `not_found` / `lookup_failed` / `not_checked` |
| `companyhouse_names` | The current registered company name |
| `source_row` | The row's position in the input file, so sorted output stays traceable |
| `active_officer_suggestions` | Who is currently in post, with their role, whenever this contact cannot be reached — either nobody matched or the person matched but has resigned. Capped at `ACTIVE_SUGGESTION_LIMIT` (2) names, with `+N more` when there are others |

`--debug` adds 11 more audit columns: cleaned name parts, nicknames, company tokens, the matched pattern, the distance, the domain verdict and the regnum actually used.

### `action` values

`result` has 16 values, which is too many to work through by hand. `action`
collapses them into seven queues, so you can filter the sheet by what needs
doing rather than by diagnosis.

| `action` | Covers | Next step |
|---|---|---|
| `fix-address` | `*_typo`, `malformed_email`, `missing_email` | Correct the address yourself |
| `find-new-contact` | `resigned_officer_match`, `no_officer_match_found`, `company_dissolved` | The person is gone — use `active_officer_suggestions` |
| `investigate-all-correct` | Active officer, address consistent with the name, domain is the company's | Nothing wrong was found. The bounce is a mail-system matter: full mailbox, deleted mailbox, spam filter, server fault |
| `investigate-non-company-domain` | Active officer, but the address is on a personal provider (gmail, hotmail…) | The domain carries no company signal, so nothing can be concluded from it |
| `investigate-mismatched` | Active officer, but the domain is unrelated to the company, or the address bears no resemblance to the name | Not close enough to call a typo; needs a human eye |
| `investigate-uncertain-match` | `possible_officer_match_*` — only the surname or only an initial matched | Confirm who this person is before concluding anything |
| `fix-data` | `missing_regnum`, `company_not_found`, `companies_house_lookup_failed`, `companies_house_skipped`, empty name fields | Fix the input or rerun |

Most rows usually land in `investigate-all-correct`, and those are the ones you do **not** need to read. The other six are the real work queue.

> `investigate-all-correct` does **not** mean the address is valid. Since no mail server is ever contacted, that can never be established — it means nothing was found wrong in the places this tool can look.

**A note on how the classification is derived.** It reads three things: `result`, `result_reason` and the recorded domain verdict. The verdict is needed because `detect_email_typo` returns early for a generic mailbox and reports `result_reason` as `generic_mailbox` whatever the domain turned out to be. By reason alone, `info@acme.co.uk`, `info@gmail.com` and `info@totallyunrelated.com` are indistinguishable; the verdict is what sends them to `investigate-all-correct`, `investigate-non-company-domain` and `investigate-mismatched` respectively.

### `result` values

`missing_email` · `malformed_email` · `first_name_typo` · `surname_typo` · `first_name_and_surname_typo` · `domain_typo` · `active_officer_match: <name>` · `resigned_officer_match: <name>` · `possible_officer_match_active: <name>` · `possible_officer_match_resigned: <name>` · `no_officer_match_found` · `company_not_found` · `company_dissolved` · `companies_house_lookup_failed` · `companies_house_skipped` · `missing_regnum`

`companies_house_lookup_failed` is a real failure — the API could not be reached. `companies_house_skipped` means it was skipped on purpose with `--no-ch`. Do not confuse the two.

### `result_reason` values

`email_matches_expected_pattern` · `generic_mailbox` · `personal_email_domain` · `email_pattern_unrecognised` · `domain_not_matched` · `close_to_expected_pattern` · `matched_including_middle_name` · `surname_only_match` · `multiple_possible_officers` · `api_error` · `name_fields_empty`

---

## Settings

Most settings have a command line equivalent. These live only in the file, at the top:

| Setting | Purpose |
|---|---|
| `PROBLEMATIC_STATUS_KEYWORDS` | Which statuses are analysed |
| `EXCLUDED_STATUS_KEYWORDS` | Which statuses are excluded (`blocked` is here) |
| `NICKNAME_GROUPS` | The nickname dictionary — extend it freely |
| `ACRONYM_DOMAIN_SUFFIXES` | Words that follow an acronym in a domain |
| `TYPO_MAX_DISTANCE_*` | Typo thresholds |
| `ACTIVE_SUGGESTION_LIMIT` | How many names `active_officer_suggestions` lists (default 2) |
| `GENERIC_MAILBOXES` | Non-personal mailboxes such as `info@`, `accounts@` |
| `COMPANY_STOPWORDS` | Legal suffixes stripped from company names |
| `FREE_EMAIL_DOMAINS` | Personal email providers |

---

## `SSL: bad handshake`

Common on corporate networks. A firewall (Zscaler, Netskope, Fortinet, Cisco Umbrella) or antivirus SSL scanning (Kaspersky, ESET, Avast) intercepts TLS and presents its own certificate, which is not in Python's store. The failure is permanent, so the script stops immediately rather than retrying through it.

**1. Use the Windows certificate store** — cleanest on a corporate network. Your organisation's root is already registered there:

```bash
pip install pip-system-certs
```

**2. Supply the root certificate as a file:**

```bash
python email_diagnostics.py triage -i list.csv -o result.csv --ca-bundle C:\path\corp-root.pem
```

Ask IT for the `.pem`, or export it from `certmgr.msc` → Trusted Root Certification Authorities. `REQUESTS_CA_BUNDLE` works too.

**3. If a proxy is required:** `set HTTPS_PROXY=http://proxy.company.local:8080`

**4. If the certificate store is stale:** `pip install --upgrade certifi`

**Last resort:** `--insecure` turns verification off. Your API key then travels over an unverified connection — use it to confirm the diagnosis, not as a fix. It prints a warning on every run.

`check` reports the OpenSSL version and any proxy or CA settings in effect.

## API failures

When you see `companies_house_lookup_failed`, the script now names the real cause:

```
WARNING  01234567 lookup failed: HTTP 500
WARNING  07654321 lookup failed: connection: read timeout
WARNING    Failed lookups (cause -> count):
WARNING      failed: HTTP 500                          12
WARNING      not_found                                  3
```

| Symptom | Meaning |
|---|---|
| `HTTP 401` / `403` | Key rejected — the run stops immediately. Verify with `check` |
| `HTTP 404` → `company_not_found` | That regnum is not in the register. Check `dbg_regnum_used` with `--debug` for a lost leading zero |
| `HTTP 429` | Rate limited. Lower `--rate` (default 1.8/s) |
| `HTTP 5xx` | A transient fault at Companies House; retried three times |
| `connection: ...` | Network, proxy or firewall. Set `HTTPS_PROXY` on a corporate network |

---

## Why one company at a time?

**There is no way to bulk-download officers from Companies House.** `/company/{regnum}/officers` is per company by design; there is no endpoint that takes a list of companies. The type of API key does not change this — it is the shape of the API, not a permission level.

The free bulk products exist but contain no officers:

| Product | Contents | Officers? |
|---|---|---|
| [Free Company Data Product](https://download.companieshouse.gov.uk/en_output.html) | Every live company: number, name, status, address, SIC | **No** |
| [PSC Data Product](https://download.companieshouse.gov.uk/en_pscdata.html) | Persons with Significant Control snapshot | No (PSC ≠ officer, though they overlap heavily in small companies) |
| [Accounts Data Product](https://download.companieshouse.gov.uk/en_accountsdata.html) | Electronically filed accounts | No |

The Free Company Data Product is still worth knowing about: it is a free monthly CSV carrying the number, official name and status (`active` / `dissolved`) of every company. Reading it locally would do the job of `--company-profile` at **zero API cost**. That integration is not written yet.

## How many requests?

**One request per company for the officers, plus one for the profile** (the profile call is on by default because it fills `companyhouse_names`; `--no-company-profile` turns it off).

Beyond that, only two things add requests:

1. **Pagination.** `total_results` counts **resigned officers too**, so a long-established company does not fit in one page. This is irreducible.

   The maximum for `items_per_page` is **not documented**. The script does not assume one: it reads the page size the server reports, advances by the records actually returned, and prints the largest page seen (`Largest page the server returned: N records`).

2. **Retries** after a timeout, a 429 or a 5xx.

`--limit` bounds **rows**, not companies — ten rows carrying three distinct regnums means only three companies are queried. Use `--limit-companies N` to bound what costs quota, and `--max-requests N` for a hard ceiling that stops requests being sent at all.

Measured against a fake server that caps pages at 35 records (`python test_quota.py`):

| Scenario | Requests |
|---|---|
| `--limit 10`, 10 small companies | 20 |
| `--limit 10`, rows sharing 4 companies | 8 |
| `--limit 10`, 10 companies of 70 officers | 30 |
| `--limit-companies 3`, 70-officer companies | 9 |
| `--no-company-profile`, 10 small companies | 10 |

The Companies House limit is **600 requests per 5 minutes** (2.0/s). The script runs 4 threads behind a shared 1.8/s throttle — the threads exist to reach that ceiling reliably, not to exceed it. A `401` aborts the whole run rather than writing thousands of pointless failures.

Rough guide: 1,000 distinct companies ≈ 20 minutes with the profile call, ≈ 10 minutes without.

---

## How the code works

One file, `email_diagnostics.py`, about 3,400 lines in fifteen sections numbered 0 to 14. Everything you would normally change lives in section 0 at the top; the rest is ordered the way the data flows through it.

| Section | Contains |
|---|---|
| 0 | Settings: paths, thresholds, vocabularies, rate limits |
| 1 | The fixed `result`, `action` and `result_reason` values, as classes `R`, `A`, `RSN` |
| 2 | Helpers: accent folding, `edit_distance`, the nickname map |
| 3 | Loading Excel/CSV and validating columns |
| 4 | Status filtering |
| 5 | Cleaning names, companies and email addresses |
| 6 | Generating expected email patterns and detecting typos |
| 7 | The Companies House client, and finding the API key |
| 8 | Officer matching |
| 9 | Per-row orchestration |
| 10 | Writing the workbook |
| 11 | `main()` |
| 12–14 | The `check` and `inspect` commands, and the CLI |

### The path one row takes

```
load_data                 read the file, every value as text
  validate_columns        confirm the six required columns exist
filter_problematic_statuses   keep the bounce family, drop blocked
  build_row_context       one dict per row; the original values are never touched
    resolve_person_name   forename / middle names / surname, plus nickname candidates
    clean_company_name    strip legal suffixes, leaving matchable tokens
    parse_email           lower-case, validate, split into local part and domain
    normalize_regnum      restore the leading zeros Excel dropped
fetch_company_data        one thread pool, one request per distinct regnum
  match_contact_to_officers   surname first, then forename
finalise_ch_first         merge the verified name, then judge the email
  detect_email_typo       compare the local part against the expected patterns
  check_domain            compare the domain against the company name
classify_action           collapse the diagnosis into one next step
write_output              Results, the work queues, Summary, Companies
```

`build_row_context` produces a dict per row that carries the raw values, the cleaned ones and the verdicts. Nothing mutates the input row, which is what keeps the original columns intact in the output.

### The parts that are not obvious

These are the decisions that took a wrong turn first, and would look arbitrary without the reason.

**`edit_distance` is Damerau, not plain Levenshtein.** The most common typing error is two letters swapping (`jhon` for `john`). Plain Levenshtein scores that as 2 edits, so `jhon.smiht@` sat outside the threshold and no typo was reported. Damerau scores a transposition as 1.

**Pagination advances by `len(items)`, not by the requested page size.** `items_per_page` is a ceiling, not a promise, and its maximum is undocumented. Advancing by what was *asked for* sent an extra request per company and, worse, silently skipped every officer past the first page — a contact in the unread part came back as `no_officer_match_found`, which is a confidently wrong answer. The client also records the page size the server actually used, so the real limit comes from your own data rather than a guess.

**`DomainCandidates` keeps two sets.** An acronym is information dense: `avzgroup` and `xyzgroup` are two characters apart and are unrelated companies. Acronym forms therefore go in the exact-match set only, and only full-word forms like `alivelizeynep` are compared by edit distance. Without the split, any similar acronym was reported as a domain typo.

**`classify_action` reads the domain verdict, not just the reason.** `detect_email_typo` returns early for a generic mailbox and reports `result_reason` as `generic_mailbox` whatever the domain turned out to be, so `info@acme.co.uk`, `info@gmail.com` and `info@totallyunrelated.com` are identical by reason alone. The verdict is recorded separately, which is what keeps an unrelated domain out of the all-correct bucket.

**`normalize_regnum` pads to eight characters.** Excel stores `01234567` as the number 1234567. Companies House then returns 404, and the row is reported as a lookup failure with no hint of the cause. This is also why the loader reads every value as text and why pandas is not used.

**Officer matching anchors on the surname.** A forename may be an abbreviation, a nickname or one of several middle names, so it cannot carry the match. The surname must match first, exactly or within one edit, and only then is the forename considered. `former_names` is searched as well, which is what finds someone whose surname changed on marriage. Where the same person appears as both resigned and active, active wins.

**The verified name is only trusted on a confident match.** In `finalise_ch_first` the official name is merged into the matching pool only when both surname and forename matched. Judging an email against the wrong person's name would produce a confidently wrong typo verdict, which is worse than no verdict.

**A weak difference is never called a typo.** A separator difference, a nickname, or a local part under five characters is not a typo. Anything that resembles no pattern at all becomes `email_pattern_unrecognised` and continues to the officer check rather than being guessed at.

**An SSL failure is not retried.** A handshake failure is permanent, so retrying burns time and quota and buries the cause. `SSLHandshakeError` is raised on the first attempt with the fixes listed in order.

**A 401 aborts the entire run.** Otherwise a bad key writes `companies_house_lookup_failed` across thousands of rows and looks like a data problem.

**The rate limiter sleeps inside the lock.** That is deliberate: it serialises the *pacing* across threads while each HTTP call happens outside the lock. Four workers exist to reach the 1.8/second ceiling reliably, not to exceed it.

**Command line defaults are captured once at import.** `apply_cli_args` writes back to the module globals that the parser reads its defaults from, so a second `cli()` call in the same process inherited the first one's settings and silently queried fewer companies. `_capture_defaults()` freezes them before anything can be overwritten.

**`.env.txt` is accepted.** Notepad's Save As appends `.txt` and Explorer hides extensions, so a file the user believes is `.env` is not. The loader accepts the common wrong names and says which file it used.

### Where to change things safely

| To change | Edit |
|---|---|
| Which statuses are analysed | `PROBLEMATIC_STATUS_KEYWORDS`, `EXCLUDED_STATUS_KEYWORDS` |
| Nickname coverage | `NICKNAME_GROUPS` — groups, expanded both ways automatically |
| How eagerly a typo is called | `TYPO_MAX_DISTANCE_*`, `MIN_LOCAL_LEN_FOR_TYPO` |
| Domain acronym suffixes | `ACRONYM_DOMAIN_SUFFIXES` |
| What each queue shows | `WORKLIST_SHEETS` — the sheet title, its actions and its columns |
| How a diagnosis maps to a next step | `classify_action` |
| Sheet order and colours | `ACTION_ORDER`, `ACTION_FILL` |

Adding a `result` value means adding it to `R`, mapping it in `classify_action`, and adding a case to `test_action.py`. Leaving it out of `classify_action` sends it to `investigate-mismatched` rather than crashing, but that is a silent misfile rather than a useful default.

---

## Tests

```bash
python test_logic.py        # cleaning, typo detection, officer matching, regnum, status
python test_e2e.py          # end to end: builds a real file, runs, verifies the output
python test_pagination.py   # pagination against a fake server, and --max-requests
python test_quota.py        # rows -> companies -> requests, measured
python test_ch_first.py     # ch_first priority ladder and name matching
python test_features.py     # acronym domain rule and the two added columns
python test_action.py       # the action column, including the generic-mailbox split
python test_sheets.py       # sorting, colouring, Summary and Companies sheets
```

None of them need network access.

---

## Known limits

- `result` is the most likely explanation, not proof. A bounce may be neither a typo nor a resignation — a full mailbox, a deleted mailbox, spam filtering or a server fault.
- Domain matching works from the **registered** name. Where the trading name differs, the result is `domain_not_matched`, which is deliberately not treated as a typo.
- Companies House holds **officers** only: directors, secretaries and LLP members. If your list is not made of officers, `no_officer_match_found` will dominate and mean little.
- The output contains personal data. Retention and sharing are your responsibility under GDPR.

## Licence

MIT
