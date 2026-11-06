# crm_customers.py â€” Customer Fetcher for the Limerick Broadband Reseller Portal

A self-contained Python CLI that logs into the Height8 iCRM reseller portal
(`https://reseller.limerickbroadband.in/Partner/`) and fetches customer data
through the portal's own pages/endpoints â€” no browser needed.

It can:

1. **Fetch all customers** from the Accounts grid (including the numeric
   `CustomerID` that the portal needs for renewals), clean the data, and show it
   as a structured table â€” with optional CSV/JSON export.
2. **Fetch detailed info for one customer**: grid summary, current plan info
   (from the renewal dialog, read-only â€” no recharge is performed) and address
   info.```

Credentials are read silently from a `.env` file next to the script; nothing is
hard-coded and all tokens/values are taken from live server responses on every
run.

---

## Table of contents

- [Requirements](#requirements)
- [Setup](#setup)
- [Usage](#usage)
- [Commands & flags](#commands--flags)
- [Output fields](#output-fields)
- [What "cleaned data" means](#what-cleaned-data-means)
- [How it works (endpoints used)](#how-it-works-endpoints-used)
- [Troubleshooting](#troubleshooting)
- [Security & privacy notes](#security--privacy-notes)
- [Related files](#related-files)

---

## Requirements

- **Python 3.8+** (developed and tested on 3.12 / Windows)
- Python packages:
  - `requests`
  - `beautifulsoup4`
  - `lxml` (optional but recommended â€” the BS4 parser used)
- A working reseller login for the portal

### Install the dependencies

```powershell
pip install --user requests beautifulsoup4 lxml
```

---

## Setup

The script reads credentials from a file named **`.env`** located in the same
directory as the script (or pass a custom path with `--env`, see below).
The expected format is:

```
username :- YOUR_RESELLER_USERNAME
password :- YOUR_RESELLER_PASSWORD
```

Careful: the separator is `:-` (colon-dash), not `=`. Extra blank lines are fine;
the script matches each line with a regex, so leading/trailing whitespace is ok.

Never commit or share this file â€” it is your portal password.

---

## Usage

From the folder containing the script:

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
python crm_customers.py get 1234

# same, saved as JSON
python crm_customers.py get 1234 --json profile.json
```

Progress messages (`[login] ok`, `[fetch] ...`, summary lines) go to **stderr**;
the table/JSON result goes to **stdout**, so you can pipe or redirect the result:
`python crm_customers.py all --limit 0 > all_customers.txt`.

---

## Commands & flags

### Global flags (before the sub-command)

| Flag         | Default                      | Description                                                                      |
| ------------ | ---------------------------- | -------------------------------------------------------------------------------- |
| `--env PATH` | `crm_customers.py` directory | Path to the credential file (must use the `username :-` / `password :-` format). |

### `all` â€” fetch & clean all customers

| Flag          | Default  | Description                                                                                                                                                    |
| ------------- | -------- | -------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `--limit N`   | `20`     | How many rows of the table to _display_ (does **not** affect CSV/JSON export â€” the export always contains all fetched customers). `--limit 0` shows every row. |
| `--csv FILE`  | _(none)_ | Write the full cleaned record set to a CSV file (UTF-8 with BOM, opens correctly in Excel).                                                                    |
| `--json FILE` | _(none)_ | Write the full cleaned record set to a JSON file (UTF-8, pretty-printed).                                                                                      |

### `get CID` â€” detailed info for one customer

`CID` (positional) is the numeric **customer id** (shown in the `customer id`
column of `all`, the same id used by `auto_renew.ps1 -CustomerId`).

| Argument           | Description                                   |
| ------------------ | --------------------------------------------- |
| `cid` (positional) | Numeric customer id, e.g. `1234`.             |
| `--json FILE`      | Additionally save the profile JSON to a file. |

What it returns (as pretty-printed JSON on stdout):

- **`summary`** â€” the customer's row from the Accounts grid (id, user id, name,
  mobile, status, expiry, last renewal).
- **`plan_info_from_renew_dialog`** â€” user id, partner name, partner balance,
  current plan name and plan MRP. This is fetched read-only; no recharge happens.
- **`address_info`** â€” address-form values from the customer detail page
  (name, title, present address, city, pincode, district, state, country,
  payment address fields when filled).

If the id is not found in the grid, the script exits with an error message
instead of guessing.

---

## Output fields

The cleaned customer list contains these columns:

| Column             | Meaning                                                    |
| ------------------ | ---------------------------------------------------------- |
| `customer_id`      | Numeric portal CustomerID (needed for renewals).           |
| `user_id`          | Login/username shown in the grid (e.g. `ABCD123`).         |
| `name`             | Customer's full name.                                      |
| `mobile`           | Mobile number, digits only (empty if unparseable).         |
| `status`           | `Active` / `Suspended` / `Inactive`.                       |
| `expiry`           | Account expiry date, `dd mmm yyyy` as shown in the portal. |
| `last_renewal`     | Date of the last renewal, same format.                     |
| `expiry_iso`       | Expiry as `YYYY-MM-DD` (empty if unparseable).             |
| `last_renewal_iso` | Last renewal as `YYYY-MM-DD`.                              |

---

## What "cleaned data" means

The raw grid HTML is messy; the script normalises it:

- Whitespace is collapsed, `&nbsp;`, zero-width and control characters removed,
  and mojibake/replacement characters dropped.
- Placeholder values (`-`, `--`, `N/A`, empty) are turned into empty values
  rather than being kept as text.
- Mobile numbers are reduced to digits and rejected if they don't look like a
  number (7+ digits).
- Status values are title-cased (`Active`, `Suspended`, `Inactive`).
- Duplicate rows (same `customer_id`) are removed.
- Rows are sorted **Active customers first, soonest expiry at the top**, then
  everything else.
- A status summary (`Active=54, Suspended=46, ...`) and the overall expiry range
  are printed after each fetch.

---

## How it works (endpoints used)

Everything is driven by the live portal, parsed fresh each run:

1. **Login** â€” `GET /Partner/crmhomepage.aspx` for ViewState fields, then
   `POST /Partner/PortalLogin.aspx?h8=1`. Success is detected by the `AuthToken`
   cookie; anything else aborts with an error.
2. **Accounts search** â€” `GET` then `POST /Partner/Accounts.aspx?DshID=10026`
   with the full ASP.NET WebForms payload: all hidden inputs
   (`__VIEWSTATE`, `__VIEWSTATEGENERATOR`, `__EVENTVALIDATION`,
   **`__VIEWSTATEENCRYPTED`**, â€¦), both date textboxes set to a wide window
   (`01 Jan 2000` â€“ `31 Dec 2027`), radio filter `rdblst=ALL`,
   `ddlPageSize=500`, and the `Search` button â€” plus the
   `X-MicrosoftAjax: Delta=true` / `X-Requested-With: XMLHttpRequest` headers
   (missing either produces a server 500). The rows are extracted from the
   returned UpdatePanel delta.
3. **Renew dialog (read-only)** â€” `GET /Partner/Account/PurchaseRenewPlan/{cid}`
   with the `__RequestVerificationToken` + `X-Requested-With` headers, used only
   to read current plan info in `get`. **No recharge is ever submitted by this
   script.**
4. **Customer detail** â€” `GET /Partner/Account/CustomerDetail/{cid}`, parsed
   with BeautifulSoup for the address-form values.

Tokens (`__VIEWSTATE`, `__EVENTVALIDATION`, CSRF token) are always taken from
the current response, never reused or hard-coded.

---

## Troubleshooting

| Symptom                                                          | Cause / fix                                                                                                                                                                                                                                                                                                                                     |
| ---------------------------------------------------------------- | ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `credentials not found in .env`                                  | The `.env` is missing, or lines don't match `username :- value` / `password :- value`. Check the exact `:-` separator.                                                                                                                                                                                                                          |
| `login failed (no AuthToken cookie)`                             | Wrong credentials, or the portal changed its login flow. Verify manually in a browser.                                                                                                                                                                                                                                                          |
| `[warn] grid returned no rows (delta head: '0\|error\|500\|\|')` | The search POST was rejected. Historically this happened when the `X-MicrosoftAjax: Delta=true` header or any ASP.NET hidden field (esp. `__VIEWSTATEENCRYPTED`) was missing from the payload. The script includes both; if it appears, re-run â€” occasional transient failures also occur (newer versions log this reason in the warning line). |
| `ModuleNotFoundError: No module named 'requests'`                | Run the pip install command in [Requirements](#requirements) with the same Python you invoke (`python -m pip install --user requests beautifulsoup4 lxml`).                                                                                                                                                                                     |
| `customer N not found in Accounts grid`                          | The id doesn't exist or the grid search failed that run. Run `all` first to see valid ids.                                                                                                                                                                                                                                                      |
| Mojibake (`é”Ÿ`) in the console                                   | Windows console code-page issue; set `chcp 65001` or rely on the CSV/JSON export which is proper UTF-8.                                                                                                                                                                                                                                         |

---

## Security & privacy notes

- The `.env` file holds your real password â€” keep it local. The script never
  prints or logs it.
- Exported CSV/JSON files and the on-screen table contain **personal customer
  data** â€” store and share them carefully. The script prints a reminder when you
  save a file.
- All portal interactions here are **read-only**. Submitting renewals is out of
  scope for this script; see `auto_renew.ps1` and `RENEWAL_ENDPOINTS.md` for the
  (guarded, confirm-first) renewal flow.

---

## Generating a cURL for Postman (`auto_renew.py --curl`)

```powershell
python auto_renew.py 123456 --curl
```

Opens the portal live for the customer, parses all 92 payload values, the
session cookies and the CSRF token, and prints a complete
`curl --request POST ... --data-raw '...'` command (nothing is submitted).
**Copy the whole command** â†’ in Postman: **Import â†’ Raw text â†’ paste â†’
Continue â†’ Import**, then Send.

Important caveats â€” the cURL is a **one-shot, session-bound artefact**:

- The cookies + `__RequestVerificationToken` belong to the session the script
  created; they expire and a later send without them fails.
- Every customer/plan value is parsed at generation time. If you wait long,
  customer/plan state may change â€” prefer re-generating right before sending.
- **Sending it = one real recharge.** Send it ONCE. Postman has none of the
  script's safety gates; sending twice would double-bill.

The generic shape (placeholders, do NOT reuse stale values):

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

## Related files

| File                       | Purpose                                                                                                            |
| -------------------------- | ------------------------------------------------------------------------------------------------------------------ |
| `crm_customers.py`         | This script.                                                                                                       |
| `.env`                     | Credentials (not committed â€” format `username :-` / `password :-`).                                                |
| `auto_renew.py` | **Python renewal trigger** â€” `python auto_renew.py <CustomerId> [--force] [--dry-run] [--curl] [--bypass-gate]` or by name: `python auto_renew.py --by-name "customer name"` or by login name: `--userid ROHITT` (case-insensitive; all refusals on ambiguous matches). Confirm-first, anti-double-recharge gate, submits exactly once, verifies on the report, newest-first recharge log. Replaces `auto_renew.ps1`. |
| `auto_renew.ps1` | PowerShell renewal trigger (original; same flow, kept for reference). |
| `renew_gate.json` | Anti-double-recharge marker file written by `auto_renew.py` (attempt date/status per customer). Created automatically; deleting entries weakens the gate â€” only remove after checking the Recent Purchase/Renew Report. |
| `recharge_log.csv` | Written by `auto_renew.py` on every confirmed recharge. One line per recharge with limited fields only: customer id, customer name, recharge date, activation date, expiry date. **Newest entry at the top** (older ones shift down). Dry-runs/aborts/failures are never logged. Contains customer names â€” store carefully. |
| `RENEWAL_ENDPOINTS.md`     | Full documentation of the renewal endpoints and payload.                                                           |
| `AUTOMATION_ASSESSMENT.md` | Safety assessment of automating renewals.                                                                          |
