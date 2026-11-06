# LLB Reseller Portal Automation Toolkit

Python CLI tools for the Limerick Broadband reseller portal
(Height8 iCRM, `https://reseller.limerickbroadband.in/Partner/`) — they drive
the portal's own pages/endpoints the way a browser would; no browser needed.

| Script | What it does | Safe by default? |
|---|---|---|
| **`crm_customers.py`** | Fetch customer data: **all** customers (incl. numeric `CustomerID`), cleaned & structured (table/CSV/JSON), plus full detail for one customer (plan info + address) | **Yes — 100% read-only.** |
| **`auto_renew.py`** | Trigger a **plan renewal (recharge)** for one customer — confirm-first, anti-double-recharge gate, submits exactly once, verifies on the report, keeps a recharge log | Guarded: shows a summary, asks for confirmation, and submits **only once** |

Both scripts read credentials silently from a `.env` file next to them
(nothing hard-coded) and parse every token / customer / plan value **fresh
from the portal's live responses** on every run — stale values are never reused.

---

## Table of contents

- [Requirements](#requirements)
- [Setup (both scripts)](#setup-both-scripts)
- [Script 1 — `crm_customers.py` (customer data)](#script-1--crm_customerspy-customer-data)
  - [Usage](#usage)
  - [Commands & flags](#commands--flags)
  - [Output fields](#output-fields)
  - [What "cleaned data" means](#what-cleaned-data-means)
  - [How it works (endpoints used)](#how-it-works-endpoints-used)
- [Script 2 — `auto_renew.py` (guarded recharge)](#script-2--auto_renewpy-guarded-recharge)
  - [Usage](#usage-1)
  - [Customer selectors](#customer-selectors)
  - [Flags](#flags)
  - [What it does, step by step](#what-it-does-step-by-step)
  - [The anti-double-recharge gate](#the-anti-double-recharge-gate)
  - [Exit codes](#exit-codes)
  - [Recharge log (`recharge_log.csv`)](#recharge-log-recharge_logcsv)
  - [Generating a cURL for Postman (`--curl`)](#generating-a-curl-for-postman---curl)
- [Troubleshooting](#troubleshooting)
- [Sessions & security notes](#sessions--security-notes)
- [Related files](#related-files)

---

## Requirements

- **Python 3.8+** (developed and tested on 3.12 / Windows)
- Python packages:
  - `requests`
  - `beautifulsoup4`
  - `lxml` (optional but recommended — the BS4 parser used)
- A working reseller login for the portal

### Install the dependencies

```powershell
pip install --user requests beautifulsoup4 lxml
```

---

## Setup (both scripts)

Both scripts read credentials from a file named **`.env`** in the script
directory (or pass a custom path with `--env`). Expected format:

```
username :- YOUR_RESELLER_USERNAME
password :- YOUR_RESELLER_PASSWORD
```

Careful: the separator is `:-` (colon-dash), not `=`. Leading/trailing spaces
are fine. Never commit or share this file — it is your portal password.

---

# Script 1 — `crm_customers.py` (customer data)

## Usage

```powershell
# show built-in help
python crm_customers.py --help

# every customer, table limited to the first 20 rows (default limit)
python crm_customers.py all

# show everything (no limit)
python crm_customers.py all --limit 0

# export the cleaned list
python crm_customers.py all --csv customers.csv --json customers.json

# full details for one customer id
python crm_customers.py get 123456

# same, saved as JSON
python crm_customers.py get 123456 --json profile.json
```

Progress messages (`[login] ok`, `[fetch] ...`, summary lines) go to
**stderr**; the table/JSON result goes to **stdout**, so piping/redirecting
works: `python crm_customers.py all --limit 0 > all_customers.txt`.

## Commands & flags

### Global flag

| Flag         | Default                   | Description                                                                                     |
|--------------|---------------------------|-------------------------------------------------------------------------------------------------|
| `--env PATH` | `crm_customers.py` directory | Path to the credential file (must use the `username :-` / `password :-` format).              |

### `all` — fetch & clean all customers

| Flag          | Default  | Description                                                                                                                                                      |
|---------------|----------|-------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| `--limit N`   | `20`     | How many rows of the table to *display* (does **not** affect CSV/JSON export — the export always contains all fetched customers). `--limit 0` shows every row. |
| `--csv FILE`  | *(none)* | Write the full cleaned record set to a CSV file (UTF-8 with BOM, opens correctly in Excel).                                                                       |
| `--json FILE` | *(none)* | Write the full cleaned record set to a JSON file (UTF-8, pretty-printed).                                                                                         |

### `get CID` — detailed info for one customer

`CID` (positional) is the numeric **customer id** (the `customer id` column of
`all`, the same id used by `auto_renew.py`).

| Argument           | Description                                   |
|--------------------|-----------------------------------------------|
| `cid` (positional) | Numeric customer id, e.g. `123456`.           |
| `--json FILE`      | Additionally save the profile JSON to a file. |

It returns pretty-printed JSON with:

- **`summary`** — the customer's Accounts-grid row (id, user id, name, mobile,
  status, expiry, last renewal).
- **`plan_info_from_renew_dialog`** — user id, partner name, partner balance,
  current plan name and plan MRP (read-only; **no recharge happens**).
- **`address_info`** — address-form values from the customer detail page.

Unknown ids abort with an error instead of guessing.

## Output fields

| Column              | Meaning                                                    |
|---------------------|------------------------------------------------------------|
| `customer_id`       | Numeric portal CustomerID (needed for renewals).           |
| `user_id`           | Login/username shown in the grid.                          |
| `name`              | Customer's full name.                                      |
| `mobile`            | Mobile number, digits only (empty if unparseable).         |
| `status`            | `Active` / `Suspended` / `Inactive`.                       |
| `expiry`            | Account expiry date, `dd mmm yyyy` as shown in the portal. |
| `last_renewal`      | Date of the last renewal, same format.                     |
| `expiry_iso`        | Expiry as `YYYY-MM-DD`.                                    |
| `last_renewal_iso`  | Last renewal as `YYYY-MM-DD`.                              |

## What "cleaned data" means

- Whitespace collapsed; `&nbsp;`, zero-width and control characters removed;
  mojibake/replacement characters dropped.
- Placeholder values (`-`, `--`, `N/A`, empty) become empty values.
- Mobile numbers reduced to digits (rejected if not a plausible number).
- Status title-cased; duplicate `customer_id` rows dropped.
- Sorted **Active first, soonest expiry at the top**; a status summary and
  expiry range are printed after each fetch.

## How it works (endpoints used)

1. **Login** — `GET /Partner/crmhomepage.aspx` for the ASP.NET hidden fields,
   then `POST /Partner/PortalLogin.aspx?h8=1`. Success = `AuthToken` cookie.
2. **Accounts search** — `POST /Partner/Accounts.aspx?DshID=10026` with the
   full WebForms payload (all hidden inputs **including
   `__VIEWSTATEENCRYPTED`**), wide date window `01 Jan 2000`–`31 Dec 2027`,
   `rdblst=ALL`, `ddlPageSize=500`, Search button, plus the
   `X-MicrosoftAjax: Delta=true` / `X-Requested-With` headers (missing either
   = server 500). Rows extracted from the UpdatePanel delta.
3. **Renew dialog (read-only)** — `GET /Partner/Account/PurchaseRenewPlan/{cid}`
   with the CSRF token; used only to read plan info in `get`.
4. **Customer detail** — `GET /Partner/Account/CustomerDetail/{cid}` parsed
   with BeautifulSoup.

Tokens are always taken from the current response — never reused, never
hard-coded.

---

# Script 2 — `auto_renew.py` (guarded recharge)

## Usage

```powershell
# by numeric customer id (normal flow: summary -> confirm -> submit once -> verify)
python auto_renew.py 123456

# by customer NAME (unique match required; lists candidates & stops if ambiguous)
python auto_renew.py --by-name "full customer name"

# by USER ID / login name (case-insensitive, unique match required)
python auto_renew.py --userid ABCD123

# safe test: opens the dialog and builds the full 92-field payload, but sends NOTHING
python auto_renew.py --userid ABCD123 --dry-run

# skip the confirmation prompt (still confirm-first gates; still submits only once)
python auto_renew.py 123456 --force

# print a ready-to-import cURL for Postman instead of submitting
python auto_renew.py 123456 --curl

# deliberately bypass the anti-double-recharge gate (only after verifying the report!)
python auto_renew.py 123456 --bypass-gate

# custom .env location
python auto_renew.py 123456 --env D:\secure\.env
```

## Customer selectors

Exactly **one** selector may be given:

| Selector                        | Matches on                    | Rule                                    |
|---------------------------------|-------------------------------|------------------------------------------|
| `ID` (positional)               | numeric customer id           | exact                                    |
| `--by-name "text"`              | name **or** user id           | substring on name / exact on user id     |
| `--userid TEXT`                 | user id (login name)          | exact, case-insensitive                  |

All selectors resolve against the **live customer grid**. A selector that
matches **no** customer aborts (exit 1, nothing submitted). A selector that
matches **several** customers **refuses to guess**, prints the candidate list
(customer id / user / name / status / expiry) and aborts — re-run with a
more specific phrase or the numeric id.

## Flags

| Flag              | Default | Description                                                                                                        |
|-------------------|---------|---------------------------------------------------------------------------------------------------------------------|
| `ID` / `--by-name` / `--userid` | — | Customer selector (see above). **Exactly one required.**                  |
| `--force`         | off     | Skip the `yes/NO` confirmation prompt (gates and sanity checks still run).                                          |
| `--dry-run`       | off     | Open the renew dialog and BUILD the payload, then stop — nothing is submitted or logged.                            |
| `--curl`          | off     | Instead of submitting, print a complete one-shot cURL command (for Postman) built from live values.                 |
| `--bypass-gate`   | off     | Skip the anti-double-recharge gate (local marker + today's report check). Use ONLY after verifying no recharge was made. |
| `--env PATH`      | next to script | Path to the credential file.                                                                                 |

## What it does, step by step

```
1. LOGIN          postback login -> AuthToken cookie (aborts if absent)
2. TOKEN          GET Accounts page -> __RequestVerificationToken
3. DIALOG         GET PurchaseRenewPlan/<customerId> -> parse all 92 payload values
                  (customer, plan, cost, dates, payment mode, tokens) fresh
4. SAFETY CHECKS  summary must match the requested/resolved customer id
                  plan + cost must exist and be numeric
                  partner balance must cover the cost
                  anti-double-recharge gate (see below)
5. CONFIRM        summary + yes/NO prompt (skip with --force)
6. SUBMIT ONCE    POST DoPurchaseRenewPlan exactly once
                  success  -> record in renew_gate.json + recharge_log.csv
                  rejected -> FAILED-FLAGGED message (exit 2), never retried
                  network error -> OUTCOME UNKNOWN (exit 3), NEVER auto-retried
7. VERIFY         GET Recent Purchase/Renew Report (ResReport.aspx) and
                  print the matching ledger entry as the independent authority
```

Scope: same-plan renew, payment mode = the dialog's first option (Cash),
`NoOfRenewal` as-is. Plan-change, wallet top-up and cheque/EDC modes are **not
implemented**.

## The anti-double-recharge gate

The vendor API is **not idempotent** — a second submission = a second debit.
The gate has two layers:

1. **Local marker (`renew_gate.json`)** — an *"attempted"* record (customer id,
   date/time) is written **before the POST fires**. This blocks re-runs if the
   script crashed/ended without a confirmed outcome, or if the same customer
   already got a confirmed recharge **today**. A previous `attempted`/
   `unknown` entry blocks regardless of the day until resolved.
2. **Live report check** — before confirming, the Recent Purchase/Renew Report
   is read: if the portal already shows a transaction for this customer
   **dated today**, the run is refused (exit 4).

To re-run deliberately after verifying the report yourself: `--bypass-gate`.
Failed attempts (`exit 2`) do not block later days; the report check still
protects same-day duplicates.

## Exit codes

| Code | Meaning |
|------|---------|
| `0`  | OK (or aborted at the confirmation prompt — safe) |
| `1`  | Aborted **before** submit (data/balance sanity, selector resolution) — safe |
| `2`  | Submission rejected by the portal (FAILED-FLAGGED; verify the report) |
| `3`  | Outcome UNKNOWN (network error/timeout; verify the report before re-running) |
| `4`  | Blocked by the anti-double-recharge gate |

## Recharge log (`recharge_log.csv`)

One line per **confirmed success**, inserted at the TOP (newest first; older
entries shift down). Deliberately minimal fields only:

```
customer_id,customer_name,recharge_date,activation_date,expiry_date
```

- `recharge_date` — the day the recharge was made (portal payment date)
- `activation_date` — when the new plan activates (server-refreshed value)
- `expiry_date` — the new expiry (server-refreshed value)

Dry-runs, prompt-aborts and failed submissions are **never** logged.

## Generating a cURL for Postman (`--curl`)

```powershell
python auto_renew.py 123456 --curl
```

Prints a complete `curl --request POST ... --data-raw '...'` command built
live (92 field values, session cookies, CSRF token) — nothing is submitted.
Copy it → Postman → **Import → Raw text → paste** → **Send**.

Caveats — the cURL is a **one-shot, session-bound artefact**:

- Cookies + token belong to the generating session; they expire (the portal
  session is swept by the server shortly after it stops being used — treat an
  exported cURL as dead after ~20–30 idle minutes).
- Customer/plan state may change if you wait — prefer re-generating right
  before sending.
- **Sending it = one real recharge. Send it ONCE.** Postman has none of the
  script's safety gates; delete the request from Postman history after sending.

The generic shape (placeholders — do NOT reuse stale values):

```
curl --request POST \
  'https://reseller.limerickbroadband.in/Partner/Account/DoPurchaseRenewPlan' \
  --header 'Content-Type: application/x-www-form-urlencoded; charset=UTF-8' \
  --header 'X-Requested-With: XMLHttpRequest' \
  --header 'x-requestverificationtoken: <token-from-Accounts-page>' \
  --header 'Referer: https://reseller.limerickbroadband.in/Partner/Accounts.aspx?DshID=10026' \
  --header 'Accept: application/json, text/javascript, */*; q=0.01' \
  --cookie 'ASP.NET_SessionId=<...>; AuthToken=<...>' \
  --data-raw 'CustomerId=...&Username=...&PlanId=...&PaymentMode=285&...&__RequestVerificationToken=<same-token>'
```

---

## Troubleshooting

### Both scripts

| Symptom | Cause / fix |
|---|---|
| `credentials not found in .env` | `.env` missing, or lines don't match `username :- value` / `password :- value`. Check the exact `:-` separator. |
| `login failed (no AuthToken cookie)` | Wrong credentials, or the portal changed its login flow. Verify manually in a browser. |
| Mojibake (`锟`) in the console | Windows console code-page issue; use `chcp 65001` or rely on the CSV/JSON export which is proper UTF-8. |

### `crm_customers.py` only

| Symptom | Cause / fix |
|---|---|
| `[warn] grid returned no rows (delta head: '0\|error\|500\|\|')` | The search POST was rejected — historically caused by a missing `X-MicrosoftAjax: Delta=true` header or a missing ASP.NET hidden field (esp. `__VIEWSTATEENCRYPTED`). The script includes both; occasional transient failures also occur — re-run. |
| `ModuleNotFoundError: No module named 'requests'` | `python -m pip install --user requests beautifulsoup4 lxml` with the same Python you invoke. |
| `customer N not found in Accounts grid` | Id doesn't exist or the grid search failed that run. Run `all` first. |

### `auto_renew.py` only

| Symptom | Cause / fix |
|---|---|
| `SAFETY GATE blocked this run` (exit 4) | See [the gate](#the-anti-double-recharge-gate) — verify the Recent Purchase/Renew Report, then `--bypass-gate` if certain. |
| `dialog customer id (X) doesn't match requested (Y)` | The dialog belongs to a different customer than requested — the script refuses; re-resolve the id. |
| `partner balance cannot cover cost` | Wallet balance too low; top up or use a different customer. |
| `no customer found matching ...` / `N customers match ...` | Selector problem (exit 1). Use `crm_customers.py all` to find the right id/user id. |
| `OUTCOME UNKNOWN` (exit 3) | POST failed mid-flight. **Do NOT retry.** Check the report; the gate blocks re-runs until resolved. |
| `SUBMISSION FAILED-FLAGGED` (exit 2) | Portal rejected it. Verify on the report; later days work normally. |

---

## Sessions & security notes

- The `.env` file holds your real password and stays local; the scripts never
  print, log or persist it.
- **Login lifetime:** cookies (`AuthToken`, `ASP.NET_SessionId`) live in the
  process RAM only. When the script ends, your side of the session ends; the
  server cleans the idle session shortly after (~20–30 idle minutes,
  sliding). There is no explicit logout — nothing persists between runs.
- Every token (`__VIEWSTATE`, `__EVENTVALIDATION`, CSRF) is parsed fresh per
  run; no captured values are ever reused.
- Files that contain customer data: `recharge_log.csv` (names), `renew_gate.json`
  (ids), and any exported `--curl` output (cookies + customer data). Store and
  share them carefully; the scripts print reminders when creating files.
- `crm_customers.py` is 100% read-only (renew dialog fetched but never
  submitted). Submitting renewals happens **only** in `auto_renew.py`, behind
  the confirm/gate layers described above.
- A `.gitignore` is provided — `.env`, `renew_gate.json`, `recharge_log.csv`
  and research captures never get committed.

## Related files

| File | Purpose |
|---|---|
| `crm_customers.py` | Customer data fetcher (read-only). |
| `auto_renew.py` | Guarded one-shot renewal trigger (Python). |
| `.env` | Credentials — local only (format `username :-` / `password :-`). |
| `auto_renew.ps1` | PowerShell renewal trigger (original; same flow, kept for reference). |
| `renew_gate.json` | Anti-double-recharge marker file written by `auto_renew.py`. Created automatically; deleting entries weakens the gate — only remove after checking the Recent Purchase/Renew Report. |
| `recharge_log.csv` | Written by `auto_renew.py` on every confirmed recharge. Limited fields (customer id/name, recharge date, activation date, expiry date), newest-first. Dry-runs/aborts/failures are never logged. Contains customer names — store carefully. |
| `RENEWAL_ENDPOINTS.md` | Full documentation of the renewal endpoints and payload. |
| `AUTOMATION_ASSESSMENT.md` | Safety assessment of automating renewals. |
