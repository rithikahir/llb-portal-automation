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
- [Setup (all scripts)](#setup-all-scripts)
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
  - [Recharge log (`logs\recharge_log.csv`)](#recharge-log-logsrecharge_logcsv)
  - [Generating a cURL for Postman (`--curl`)](#generating-a-curl-for-postman---curl)
- [Script 3 — `transfer_balance.py` (guarded wallet transfer)](#script-3--transfer_balancepy-guarded-wallet-transfer)
  - [Usage](#usage-2)
  - [Partner selectors](#partner-selectors)
  - [Flags](#flags-1)
  - [What it does, step by step](#what-it-does-step-by-step-1)
  - [The anti-double-transfer gate](#the-anti-double-transfer-gate)
  - [Transfer log (`logs\transfer_log.csv`)](#transfer-log-logstransfer_logcsv)
- [Script 4 — `run_billing.py` (GST billing — driver + bill engine in one)](#script-4--run_billingpy-gst-billing--driver--bill-engine-in-one)
  - [How the pieces fit](#how-the-pieces-fit)
  - [What gets updated in a bill](#what-gets-updated-in-a-bill)
  - [One-shot mode (bill one customer by hand)](#one-shot-mode-bill-one-customer-by-hand)
  - [Safety & outputs](#safety--outputs)
  - [One-time setup](#one-time-setup)
  - [Running it](#running-it)
  - [What you see in the sheet](#what-you-see-in-the-sheet)
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

## Setup (all scripts)

The scripts read credentials from a file named **`.env`** in the script
directory (or pass a custom path with `--env`). Expected format — the plain
pair is the **default login** (called `reseller`):

```
username :- YOUR_RESELLER_USERNAME
password :- YOUR_RESELLER_PASSWORD
```

Extra portal logins are stored as suffixed pairs and selected per run with
`--account NAME` (all three scripts support this):

```
username_shreeganesh :- ...
password_shreeganesh :- ...
username_shree :- ...
password_shree :- ...
```

- `python crm_customers.py --account shree all`
- `python auto_renew.py --account shree 123456`
- `python transfer_balance.py --account shree ...`

Any future login works the same way: add `username_<NAME>` / `password_<NAME>`
to `.env` and run with `--account <NAME>`. If a name is wrong, the script
aborts immediately and lists the accounts it *did* find (values are never
printed).

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

| Flag              | Default                      | Description                                                                                     |
|-------------------|------------------------------|-------------------------------------------------------------------------------------------------|
| `--env PATH`      | `crm_customers.py` directory | Path to the credential file (must use the `username :-` / `password :-` format).                |
| `--account NAME`  | `reseller` (plain pair)      | Which login in `.env` to use: the plain pair by default, or `username_<NAME>/password_<NAME>` (e.g. `--account shree`). |

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
| `--account NAME`  | `reseller` (plain pair) | Which login in `.env` to use (e.g. `--account shree`). The customer grid, dialog, balance and gate then all refer to that account. |
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
                  success  -> record in logs\renew_gate.json + logs\recharge_log.csv
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

1. **Local marker (`logs\renew_gate.json`)** — an *"attempted"* record (customer id,
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

## Recharge log (`logs\recharge_log.csv`)

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

# Script 3 — `transfer_balance.py` (guarded wallet transfer)

Moves **real money** inside the partner tree by driving the portal's own
"Manage Transfer Balance" page (`/Partner/TransferBalance.aspx`): it debits the
source account's wallet (default: the `shreeganesh` login = distributor
account) and instantly credits a sub-partner (e.g. the reseller's own child
account). The result appears on both sides of the Balance Transfer History
ledger as a same-timestamp `Dr`/`Cr` pair.

> ⚠️ **Transfers are effectively one-way.** Only Limerick staff accounts have
> ever been able to run `Reverse Balance` corrections. Treat every transfer
> with the same care as a recharge — verify the summary before confirming.

> ⚠️ The source wallet is shared money: it also funds recharges for *your*
> customers and carries the running balance owed to Limerick.

## Usage

```powershell
# always `python <script> <flags>` — never `.\transfer_balance.py`

# normal flow: summary -> type yes -> submit exactly once -> verify on the ledger
python transfer_balance.py --partner-name KHADKI --amount 100

# by numeric partner id from the live dropdown (safer when names repeat)
python transfer_balance.py --partner-id 54 --amount 100

# safe test: resolves everything, builds and prints the payload — sends NOTHING
python transfer_balance.py --partner-name KHADKI --amount 100 --dry-run

# no confirmation prompt (page + gates + single-shot submit still apply)
python transfer_balance.py --partner-id 54 --amount 100 --force

# print a one-shot cURL instead of submitting (session-bound — see Script 2 notes)
python transfer_balance.py --partner-id 54 --amount 100 --curl

# use the plain reseller login as the source instead of shreeganesh
python transfer_balance.py --partner-id 54 --amount 100 --account reseller

# optional knobs
python transfer_balance.py --partner-id 54 --amount 100 \
  --remark "topup for recharges" --payment-received no --date "07 Nov 2026"
```

## Partner selectors

- `--partner-name <text>` — case-insensitive substring; must match **exactly
  one** live dropdown entry or the script refuses to guess (exit 1).
- `--partner-id <n>` — the numeric id shown in the dropdown
  (e.g. `54` = the Khadki account).
- The target list is fetched **live** from the portal every run
  (`--level sub` = Sub Distributor, the default; `--level lco` = LCO).
  No partner ids are ever hard-coded.

## Flags

| Flag | Meaning |
|---|---|
| `--account {shreeganesh,reseller,...}` | source login; looks up `username_<NAME>/password_<NAME>` in `.env` (`reseller` = plain pair). Default `shreeganesh`. Only use a login that legitimately owns the wallet being debited. |
| `--level {sub,lco}` | Entity Level for the recipient dropdown (default `sub`). |
| `--partner-id` / `--partner-name` | (required, pick one) the transfer target. |
| `--amount` | (required) amount in ₹; must be > 0 and ≤ the live wallet balance. |
| `--date` | transfer date `dd MMM yyyy`; default = the page's own default (today). |
| `--remark` | free-text remark recorded in the ledger. |
| `--payment-received {yes,no}` | portal radio; default `yes` (the page's own default). |
| `--dry-run` | build everything, submit nothing, touch no gate file. |
| `--curl` | print a one-shot cURL instead of submitting. |
| `--force` | skip the `yes/NO` confirmation prompt. |
| `--bypass-gate` | skip the anti-double-transfer gate — **only after verifying the ledger yourself**. |
| `--env <path>` | alternate .env location. |

## What it does, step by step

1. Live login as the source account.
2. Opens the Transfer Balance page; reads the **live wallet balance**
   (`Current Balance:` label) and the default transfer date.
3. Async-refreshes the Entity Level / partner dropdown **read-only** and
   resolves the target (unique match or refuse).
4. Sanity checks: amount > 0, amount ≤ balance.
5. Gate: local marker + **same-day ledger check** on Balance Transfer History
   (a matching `Transfer Balance To` row already on the ledger blocks the run).
6. Prints the full summary → `yes/NO` prompt (skipped with `--force`).
7. Writes the *"attempted"* marker to `logs\transfer_gate.json` **before** the POST.
8. Submits the WebForms POST **exactly once** — no retry after timeout or
   HTTP error (outcome UNKNOWN → verify the ledger manually; the gate stays
   closed until resolved).
9. Verifies on the portal's own **Balance Transfer History** report: a
   `Transfer Balance To <partner>` debit row dated today with the exact
   amount (checked up to 4 times, 15 s apart). Only a confirmed match updates
   the gate to `confirmed` and writes `logs\transfer_log.csv`.
10. Prints the confirmed ledger row (including the closing wallet balance).

## The anti-double-transfer gate

Same two-layer design as the recharge gate:

1. **Local marker (`logs\transfer_gate.json`)** — the *"attempted"* record is
   written **before** the POST. Any `attempted`/`unknown` record blocks all
   re-runs (any partner, any amount) until you resolve it; a *confirmed*
   record blocks the same partner + amount + day duplicate.
2. **Live ledger check** — before submitting, Balance Transfer History is
   queried: a matching `Transfer Balance` row dated today for this
   partner/amount refuses the run (exit 4).

To re-run deliberately after checking the ledger yourself:
`--bypass-gate`. Resolving stuck `attempted`/`unknown` entries: delete the
entry from `logs\transfer_gate.json` (after verifying the report) or use
`--bypass-gate` once.

## Transfer log (`logs\transfer_log.csv`)

One line per **confirmed transfer**, inserted at the TOP (newest first).
Only after ledger verification — dry-runs, prompt-aborts, `--curl` and
unknown/failed outcomes are never logged:

```
transfer_datetime,from_account,partner_id,partner_name,amount,remark
```

---

# Script 4 — `run_billing.py` (GST billing — driver + bill engine in one)

The one script for everything billing (the old `update_bills.py` was merged
into it and deleted). **You** decide when to bill; run it from the project
folder. It asks the Apps Script Web App who is due, edits each customer's
bill `.docx` in place, produces the PDF, and reports back so the sheet is
updated — all in one run.

```powershell
python run_billing.py --dry-run    # preview: shows who is due, edits nothing
python run_billing.py              # the real run
```

## How the pieces fit

```
you run run_billing.py (PC)
        │  1. POST get_due  ──────────────►  Apps Script Web App
        │                                    reads UPCOMING BILLS, returns
        │                                    the due rows (read-only)
        │         ◄──── list of due bills ──┘
        │  2. for each bill (all on YOUR PC):
        │        pick the .docx (fuzzy or 'Bill File' column)
        │        rewrite Invoice Date / Invoice No / Bill Period
        │        docx saved in place → PDF via LibreOffice
        │  3. POST complete ──────────────►  Apps Script Web App
        │                                    appends GST RECORDS rows,
        ◄──────────────────────────────────► advances Expire Dates,
                                             re-sorts UPCOMING BILLS
```

No tunnel and no always-on server are needed — the connection lives only for
the few seconds the script runs.

## What gets updated in a bill

| Field | Before → After |
|---|---|
| `Invoice Date : - 18-07-26` | today's date, **in the bill's own format** (e.g. `11-11-26`) |
| `Invoice No : - ST/023/2026-27` | the **exact next global number** Apps Script computed (e.g. `ST/024/2026-27`) |
| `Bill Period:-17/07/2026 to 16/01/2027` | advanced by the bill's **own duration** in whole calendar months (`auto`, default): 1-month bill → next month, 3-month → 3, 13-month → 13. Alternatives: `next` (day-count), `year` (both dates +1 year) |

Everything else in the document (amounts, GSTIN, buyer address, item rows) is
untouched — edits are label-anchored (`Invoice No`, `Invoice Date`,
`Bill Period`) so only those values are rewritten. All repeated copies
(merged cells) are updated too.

**Idempotency:** the same run re-executed never edits twice — Python's
`logs\bills_state.json` dedupes by key (second attempt answers from cache
with `duplicate: true`). This protects against re-running after a crash
(which would otherwise bump the invoice number and shift the period twice).

## One-shot mode (bill one customer by hand)

No sheet contact — useful for manual fixes and testing:

```powershell
# preview (changes nothing)
python run_billing.py --customer "SKIN CITY" --invoice-no "ST/023/2026-27" --dry-run

# real run: updates the docx in place + writes Desktop\gst_bills\SKIN_CITY_BILL.pdf
python run_billing.py --customer "SKIN CITY" --invoice-no "ST/023/2026-27"

# pick the file directly, shift the period to the consecutive one
python run_billing.py --bill-file SKIN_CITY_BILL.docx --period-shift next
```

Flags: `--customer`, `--bill-file`, `--invoice-no`, `--new-invoice-no`,
`--invoice-date`, `--period-shift {auto,next,year}`, `--pdf-dir`, `--bills-dir`,
`--dry-run`.

## Safety & outputs

- The `.docx` is rewritten **in place** — no backup copies are kept (the
  docx write itself is atomic: the file is assembled as a temp file and
  swapped in only when complete).
- PDF → `Desktop\gst_bills\<bill name>.pdf` (LibreOffice headless; MS Word via
  COM as fallback). If PDF conversion fails, the docx update still stands and
  the result carries `pdf_error`.
- `logs\bills_update_log.csv` — one line per successful update (old→new invoice no,
  periods, pdf path, key). `logs\bills_state.json` — the idempotency cache.
  Both are gitignored.
- Windows quirk: run as `python run_billing.py ...`, never `.\run_billing.py`.

## Apps Script side (`Code.js`)

`Code.js` (in this repo) is pasted into Extensions > Apps Script of the GST
sheet. It is deployed as a **Web App** — the local `run_billing.py` script
calls it; nothing on Google's side runs by itself (no trigger, no tunnel).

What it does when called:

- **`get_due`** — returns every `UPCOMING BILLS` row whose Expire Date is on
  or before a cutoff sent by Python (default: yesterday). It reads the
  optional `Bill File` column but **never writes anything to the sheet** in
  this step.
- **`complete`** — after Python has edited the bills, this appends the new
  rows to `GST RECORDS` (copying GSTIN / description / amounts / payment
  mode from the customer's **last** invoice, fuzzy name match; `Payment
  Status` = Pending) **only if that invoice number is not already there**
  (crash-safe), then **advances that customer's Expire Date to the new
  period end** and **re-sorts UPCOMING BILLS by Expire Date** (nearest
  first). Nothing is written to the `Bill File` / `Bill Status` columns;
  failed customers keep their old Expire Date and are simply retried by the
  next run.

Invoice numbers are **global**: `ST/046/2026-27 -> ST/047/2026-27`,
restarting at `ST/001/2027-28` after each 31 March. Python writes the exact
number Apps Script computed.

The sheet ID and a shared secret are constants in `Code.js`; the web app URL
and the same secret live in `.env` (`appscript_url`, `appscript_secret`).
See the one-time setup below for the full details.

Optional helper column: `Bill File` (col E of UPCOMING BILLS) — exact docx
name for customers whose name doesn't uniquely map to a file (e.g. the Skin
City variants, Spes Manning vs Router, Station HQ, Airan 1/2).

---

## One-time setup

1. Open the sheet → Extensions → Apps Script → paste this repo's `Code.js`.
2. Seed the secrets: in `Code.js`'s editor open `setupSecrets()`, paste the
   values of `appscript_secret` / `sheet_id` from `.env` into its two
   placeholder arguments, run it once and authorize. (Or POST
   `{action:'set_props', secret:<current>, props:{SECRET:'…', SHEET_ID:'…'}}`
   to the `/exec` URL.) Secrets live in **Script Properties**, never in code.
3. Deploy → New deployment → type **Web app** → Execute as: **Me** →
   Who has access: **Anyone** → Deploy → copy the URL ending in `/exec`.
4. Put that URL in `.env`:  `appscript_url :- https://script.google.com/.../exec`
5. Test in the editor with `dryRunOnce()` (logs a preview, changes nothing).
6. Test locally with `python run_billing.py --dry-run`.

`.env` keys used by this script:

| key | meaning |
|---|---|
| `appscript_url :-` | the Web App `/exec` URL |
| `appscript_secret :-` | shared secret — matches the `SECRET` Script Property (rotate via `set_props`) |
| `sheet_id :-` | the GST Google Sheet ID (matches the `SHEET_ID` Script Property) |
| `pdf_dir :-` | **directory controller** for the PDFs (default `Desktop\gst_bills`) |
| `bills_dir :-` | optional: where the bill `.docx` templates live |

Remember: every time you edit `Code.js`, create a new deployment version
(Deploy → Manage deployments → Edit → Version: New), or the old code keeps
running. From this PC that is automated:  `python deploy_apps_script.py`

## Running it

```powershell
python run_billing.py                    # bills everything that expired yesterday or earlier
python run_billing.py --include-today    # also rows that expire today
python run_billing.py --days 7           # everything that expired in the last week
python run_billing.py --dry-run          # preview only
```

Exit codes: `0` all good, `2` config/connection problem, `3` some bills
failed (nothing was changed for those customers in the sheet — they stay
due and are retried automatically on the next run). Re-runs are safe:
Python's `bills_state.json` (in `logs\`) dedupes by key, Apps Script only appends a GST
row when the invoice number is not already in the sheet, and Python refuses
to edit a bill whose current invoice number doesn't match the sheet
(double-bill guard).

## What you see in the sheet

- `UPCOMING BILLS`: a billed customer's **Expire Date moves forward** to the
  new period end, and the tab stays **sorted by Expire Date, nearest first**.
  The `Bill File` / `Bill Status` columns are never written by the
  automation (Bill File is only read if you create and fill it yourself).
- `GST RECORDS` gains the new invoice rows with `Payment Status` = Pending —
  fill Payment Date / GST PAID after collection.
- PDFs land in the `pdf_dir` folder from `.env` (`Desktop\gst_bills` by
  default). Every edited docx is updated **in place** — no backups are kept.

---

## Troubleshooting

### All scripts

| Symptom | Cause / fix |
|---|---|
| `credentials not found in .env` | `.env` missing, or lines don't match `username :- value` / `password :- value`. Check the exact `:-` separator. |
| `credentials for account 'X' not found ... accounts available: ...` | No `username_X` / `password_X` pair in `.env` (or the name is misspelled). The error lists the accounts it found — values are never printed. |
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
- Files that contain customer/partner data (all under `logs\`):
  `recharge_log.csv`, `transfer_log.csv` (names), `renew_gate.json` /
  `transfer_gate.json` (ids), `bills_update_log.csv` / `bills_state.json`
  (bill history), and any exported `--curl` output (cookies + context).
  Store and share them carefully; the scripts print reminders when creating
  files.
- `crm_customers.py` is 100% read-only (renew dialog fetched but never
  submitted). Submitting renewals happens **only** in `auto_renew.py`, behind
  the confirm/gate layers described above.
- Submitting wallet transfers happens **only** in `transfer_balance.py`,
  behind its confirm/gate layers; every investigation step (dropdown
  refreshes, balance reads, ledger queries) is read-only.
- A `.gitignore` is provided — `.env`, the gate/log files and research
  captures never get committed.

## Related files

| File | Purpose |
|---|---|
| `crm_customers.py` | Customer data fetcher (read-only). |
| `auto_renew.py` | Guarded one-shot renewal trigger (Python). |
| `transfer_balance.py` | Guarded one-shot partner-to-partner wallet transfer (Python). |
| `.env` | Credentials — local only (format `username :-` / `password :-`). |
| `auto_renew.ps1` | PowerShell renewal trigger (original; same flow, kept for reference). |
| `logs\renew_gate.json` | Anti-double-recharge marker file written by `auto_renew.py`. Created automatically; deleting entries weakens the gate — only remove after checking the Recent Purchase/Renew Report. |
| `logs\recharge_log.csv` | Written by `auto_renew.py` on every confirmed recharge. Limited fields (customer id/name, recharge date, activation date, expiry date), newest-first. Dry-runs/aborts/failures are never logged. Contains customer names — store carefully. |
| `logs\transfer_gate.json` | Anti-double-transfer marker file written by `transfer_balance.py`. Same handling rules as `logs\renew_gate.json`. |
| `logs\transfer_log.csv` | Written by `transfer_balance.py` on every ledger-confirmed transfer (datetime, source account, target partner, amount, remark), newest-first. |
| `logs\bills_update_log.csv` / `logs\bills_state.json` | Bill update history / dedupe cache for `run_billing.py`. |
| `RENEWAL_ENDPOINTS.md` | Full documentation of the renewal endpoints and payload. |
| `AUTOMATION_ASSESSMENT.md` | Safety assessment of automating renewals. |
