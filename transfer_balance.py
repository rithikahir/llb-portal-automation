#!/usr/bin/env python3
"""
transfer_balance.py -- GUARDED one-shot wallet transfer on the Limerick reseller
portal (Height8 iCRM).

Moves REAL money inside the partner tree: it debits the wallet of the source
account (login `shreeganesh` = LT@GANESH ENTERPRISES by default) and credits a
sub-partner (e.g. LT@KHADKI BROADBAND), by driving the portal's own
"Manage Transfer Balance" page (/Partner/TransferBalance.aspx) the way a
browser would.

Safety model (mirrors auto_renew.py):
  * live everything -- partner list, balance, tokens are parsed fresh each run
  * resolves the target from the LIVE dropdown; refuses ambiguity
  * sanity checks: amount > 0, amount <= current balance
  * double-transfer gate: transfer_gate.json "attempted" marker written
    BEFORE the POST + live Balance Transfer History same-day check
  * submit exactly ONCE: no retry after timeout/error (outcome UNKNOWN ->
    verify the report manually; the gate then blocks re-runs)
  * verify on the portal's own ledger (Balance Transfer History, Self scope)
    and only log to transfer_log.csv on CONFIRMED success

Windows quirk (same as auto_renew.py): run with  `python <script> <flags>`;
never `.\transfer_balance.py` (.py is associated with an Electron app here).
"""

import argparse
import csv
import json
import re
import sys
import time
from datetime import date
from pathlib import Path
from urllib.parse import urlencode

import requests
from bs4 import BeautifulSoup

import crm_customers as C

BASE = C.BASE
URL_TRANSFER = f"{BASE}/Partner/TransferBalance.aspx"
URL_HIST = f"{BASE}/Partner/BalTransHist.aspx"

SCRIPT_DIR = Path(__file__).resolve().parent
GATE_FILE = SCRIPT_DIR / "transfer_gate.json"
LOG_FILE = SCRIPT_DIR / "transfer_log.csv"

# entity-level dropdown values (verified live; options are re-checked each run)
LEVEL_OPTIONS = {"sub": "22602", "lco": "22603"}

UA = C.UA


def log(msg: str) -> None:
    print(msg, file=sys.stderr)


def die(msg: str, code: int) -> "None":
    print(f"[ERROR] {msg}", file=sys.stderr)
    sys.exit(code)


# ---------------------------------------------------------------- credentials
def load_account(env_path: Path, account: str) -> dict:
    """Read the chosen credential pair from .env. Format: `<key> :- <value>`."""
    keys = {
        "shreeganesh": ("username_shreeganesh", "password_shreeganesh"),
        "reseller": ("username", "password"),
    }[account]
    creds = {}
    for line in env_path.read_text(encoding="utf-8", errors="ignore").splitlines():
        for key in keys:
            m = re.match(rf"\s*{re.escape(key)}\s*:-\s*(.+)", line)
            if m:
                creds[key.split("_")[0]] = m.group(1).strip()
    if len(creds) != 2:
        raise RuntimeError(f"credentials for account '{account}' not found in {env_path}")
    return {"username": creds["username"], "password": creds["password"]}


# ---------------------------------------------------------------- page helpers
def collect_fields(soup: BeautifulSoup) -> list[tuple[str, str]]:
    """All form fields in DOM order, mimicking what a browser would submit."""
    form = soup.find("form")
    out: list[tuple[str, str]] = []
    seen: set[str] = set()

    def add(k: str, v: str):
        if k and k not in seen:
            seen.add(k)
            out.append((k, v))

    for inp in form.find_all("input"):
        t = (inp.get("type") or "text").lower()
        name = inp.get("name")
        if not name:
            continue
        if t in ("submit", "button", "image"):
            continue                          # buttons appear only when we ask for them
        if t == "checkbox" and "checked" not in inp.attrs:
            continue
        if t == "radio" and "checked" not in inp.attrs:
            continue                          # only send the checked radio (browser-like)
        add(name, inp.get("value") or "")

    for sel in form.find_all("select"):
        name = sel.get("name")
        if not name:
            continue
        opted = sel.find("option", selected=True) or sel.find("option")
        add(name, opted.get("value", "") if opted else "")

    for ta in form.find_all("textarea"):
        add(ta.get("name"), (ta.get_text() or "").strip())

    # __EVENTTARGET etc. are rendered on this page -> force them to the front
    front = [("__EVENTTARGET", ""), ("__EVENTARGUMENT", ""), ("__LASTFOCUS", "")]
    rest = [(k, v) for k, v in out if k not in ("__EVENTTARGET", "__EVENTARGUMENT", "__LASTFOCUS")]
    return front + rest


def set_field(fields: list[tuple[str, str]], name: str, value: str) -> None:
    for i, (k, _) in enumerate(fields):
        if k == name:
            fields[i] = (name, value)
            return
    fields.append((name, value))


def get_field(fields: list[tuple[str, str]], name: str) -> str:
    for k, v in fields:
        if k == name:
            return v
    return ""


class Page:
    """State of /Partner/TransferBalance.aspx for one run."""

    def __init__(self, portal: "C.Portal"):
        self.portal = portal
        r = portal.s.get(URL_TRANSFER, headers={"Referer": f"{BASE}/Partner/Default.aspx"}, timeout=60)
        if r.status_code != 200:
            die(f"TransferBalance.aspx returned HTTP {r.status_code} (login may have expired)", 1)
        if "ucTransferBalancePartner" not in r.text:
            die("this login does not render the Transfer Balance widget — account lacks permission", 1)
        soup = BeautifulSoup(r.text, "lxml")
        self.fields = collect_fields(soup)

        # live wallet: "Current Balance: 5916.55" in span#lblCreditBalance
        bal = None
        el = soup.find(id="lblCreditBalance")
        if el is not None:
            t = el.get_text(" ", strip=True).replace(",", "")
            m = re.search(r"\d+\.\d{2}", t or "")
            if m:
                bal = float(m.group(0))
        if bal is None:
            m = re.search(r'id="lblCreditBalance"[^>]*>\s*([\d,]+\.\d{2})', r.text)
            if m:
                bal = float(m.group(1).replace(",", ""))
        if bal is None:
            die("could not read the live balance from the Transfer page", 1)
        self.balance = bal

        sel = soup.find("select", attrs={"name": "ctl00$ContentPlaceHolder1$ucTransferBalancePartner$ddlEntityLevel"})
        if not sel:
            die("Entity Level dropdown missing on the Transfer page", 1)
        self.entity_options = [(o.get("value", ""), o.get_text(" ", strip=True)) for o in sel.find_all("option")]

        self.date_default = get_field(self.fields, "ctl00$ContentPlaceHolder1$ucTransferBalancePartner$txtTDate")


# ------------------------------------------------------------------ AJAX step
def expand_entity_list(page: Page, level_value: str) -> list[tuple[str, str]]:
    """AutoPostBack of ddlEntityLevel (async) -> partner dropdown options +
    refreshed hidden fields. Read-only: no transfer happens here."""
    P = page
    pairs = list(P.fields)
    set_field(pairs, "ctl00$ContentPlaceHolder1$ucTransferBalancePartner$ddlEntityLevel", level_value)
    set_field(pairs, "__EVENTTARGET", "ctl00$ContentPlaceHolder1$ucTransferBalancePartner$ddlEntityLevel")
    r = P.portal.s.post(URL_TRANSFER, data=pairs, headers={
        "Referer": URL_TRANSFER, "X-MicrosoftAjax": "Delta=true", "X-Requested-With": "XMLHttpRequest",
    }, timeout=120)
    if r.status_code != 200:
        die(f"entity-list refresh failed (HTTP {r.status_code})", 1)
    raw = r.text

    # updated hidden fields from the delta (ViewState changes after every postback)
    for m in re.finditer(r"\|\d+\|hiddenField\|([^|]+)\|([^|]*)\|", raw):
        set_field(P.fields, m.group(1), m.group(2))

    panels = []
    for m in re.finditer(r"\|(\d+)\|updatePanel\|([^|]+)\|", raw):
        panels.append((m.group(2), int(m.group(1)), m.end()))
    if not panels:
        die("entity-list refresh returned no update panel (portal changed?)", 1)
    seg = ""
    for mid, size, start in panels:
        seg += raw[start: start + size]
    seg = seg.replace("\\\"", '"')

    soup = BeautifulSoup(seg, "lxml")
    sel = soup.find("select", attrs={"name": "ctl00$ContentPlaceHolder1$ucTransferBalancePartner$ddlSubDistributor"})
    if sel is None:
        die("partner dropdown not returned after entity-level refresh", 1)
    opts = [(o.get("value", ""), o.get_text(" ", strip=True)) for o in sel.find_all("option")]
    return opts


# ------------------------------------------------------------- gate/log utils
def gate_load() -> list[dict]:
    if GATE_FILE.exists():
        try:
            data = json.loads(GATE_FILE.read_text(encoding="utf-8"))
            return data.get("records", [])
        except Exception:
            print("[warn] transfer_gate.json unreadable — treating as empty", file=sys.stderr)
            return []
    return []


def gate_save(records: list[dict]) -> None:
    GATE_FILE.write_text(json.dumps({"records": records}, indent=2), encoding="utf-8")


def gate_check(records: list[dict], partner_id: int, amount: float) -> str | None:
    today = date.today().isoformat()
    for rec in records:
        if rec.get("outcome") in ("attempted", "unknown"):
            return (f"previous transfer to partner {rec.get('partner_id')} on {rec.get('ts')} "
                    f"is still '{rec.get('outcome')}' — verify the Balance Transfer History "
                    f"report, then either edit transfer_gate.json or use --bypass-gate")
        if (rec.get("outcome") == "confirmed" and rec.get("partner_id") == partner_id
                and abs(float(rec.get("amount", 0)) - amount) < 0.005
                and rec.get("date_iso") == today):
            return (f"a CONFIRMED transfer of {amount:.2f} to partner {partner_id} was already "
                    f"made today ({rec.get('ts')}) — same-day duplicate refused")
    return None


def gate_record(records: list[dict], rec: dict) -> None:
    records.append(rec)
    gate_save(records)


def log_transfer(row: list[str]) -> None:
    """Insert newest-first (line 2) into transfer_log.csv."""
    header = ["transfer_datetime", "from_account", "partner_id", "partner_name", "amount", "remark"]
    lines = []
    if LOG_FILE.exists():
        lines = LOG_FILE.read_text(encoding="utf-8-sig").splitlines()
        if lines and lines[0].lower().startswith("transfer_datetime"):
            lines = lines[1:]
    buf = [header] + [row] + lines
    with LOG_FILE.open("w", encoding="utf-8-sig", newline="") as f:
        csv.writer(f).writerows(buf)
    print(f"[log] confirmed transfer written to {LOG_FILE.name}", file=sys.stderr)


# ---------------------------------------------------------------- history API
def history_rows(portal: "C.Portal", start: str = "01 Jan 2025", end: str = "31 Dec 2027",
                 page_size: str = "500") -> list[str]:
    """Balance Transfer History (Self scope) as flattened row texts. Read-only."""
    g = portal.s.get(URL_HIST, headers={"Referer": URL_HIST}, timeout=60)
    soup = BeautifulSoup(g.text, "lxml")
    form = soup.find("form")
    pairs, seen = [], set()

    def add(k: str, v: str):
        if k and k not in seen:
            seen.add(k)
            pairs.append((k, v))

    for inp in form.find_all("input"):
        t = (inp.get("type") or "text").lower()
        name = inp.get("name")
        if not name or t in ("submit", "button", "image"):
            continue
        if t in ("checkbox", "radio") and "checked" not in inp.attrs:
            continue
        if re.search(r"(txtStartDate|txtEndDate)$", name):
            continue
        add(name, inp.get("value") or "")
    for sel in form.find_all("select"):
        name = sel.get("name")
        if name:
            opted = sel.find("option", selected=True) or sel.find("option")
            add(name, opted.get("value", "") if opted else "")

    out = [("__EVENTTARGET", ""), ("__EVENTARGUMENT", ""), ("__LASTFOCUS", "")]
    rest = [(k, v) for k, v in pairs
            if k not in ("__EVENTTARGET", "__EVENTARGUMENT", "__LASTFOCUS")
            and not re.search(r"(?i)(txtStartDate|txtEndDate)$", k)
            and "$ddlresname" not in k and "$ddlPartner" not in k and "$ddlPageSize" not in k]
    out += rest
    idx = next(i for i, (k, _) in enumerate(out) if k.startswith("__EVENTVALIDATION")) + 1
    out[idx:idx] = [
        ("ctl00$ContentPlaceHolder1$ddlresname", "0"),
        ("ctl00$ContentPlaceHolder1$ddlPartner", "0"),
        ("ctl00$ContentPlaceHolder1$ddlPageSize", page_size),
        ("ctl00$ContentPlaceHolder1$txtStartDate", start),
        ("ctl00$ContentPlaceHolder1$txtEndDate", end),
    ]
    out.append(("ctl00$ContentPlaceHolder1$btnserch", "View"))

    r = portal.s.post(URL_HIST, data=out, headers={
        "Referer": URL_HIST, "X-MicrosoftAjax": "Delta=true", "X-Requested-With": "XMLHttpRequest",
    }, timeout=240)
    raw = r.text
    if r.status_code != 200 or raw.startswith("0|error|"):
        log("[warn] Balance Transfer History query failed this round (transient?)")
        return []
    m = re.search(r"\|(\d+)\|updatePanel\|ContentPlaceHolder1_UpdatePanel1\|", raw)
    if not m:
        return []
    seg = raw[m.end(): m.end() + int(m.group(1))].replace("\\\"", '"')
    grid = BeautifulSoup(seg, "lxml")
    tbl = None
    for t in grid.find_all("table"):
        trs = t.find_all("tr")
        if trs and "SrNo" in trs[0].get_text(" ", strip=True):
            tbl = t
            break
    if tbl is None:
        return []
    rows = []
    for tr in tbl.find_all("tr")[1:]:
        cells = [td.get_text(" ", strip=True) or "-" for td in tr.find_all("td")]
        if len(cells) >= 8:
            rows.append(" | ".join(cells))
    return rows


def history_match(rows: list[str], partner_label_fragment: str, amount: float, date_prefix: str) -> list[str]:
    """Outgoing Transfer Balance rows on/after date, to this partner, amount +/-."""
    hits = []
    frag = partner_label_fragment.upper()
    for row in rows:
        u = row.upper()
        if "TRANSFER BALANCE" not in u:
            continue
        if frag not in u:
            continue
        if "TRANSFER BALANCE TO" not in u:
            continue
        if date_prefix and not re.search(re.escape(date_prefix.upper()), u):
            continue
        amounts = [float(a.replace(",", "")) for a in re.findall(r"-?\d[\d,]*\.\d{2}", row)]
        if not amounts:
            continue
        # primary hit: outgoing debit of exactly the transferred amount
        if any(abs(a + amount) < 0.005 for a in amounts):
            hits.append(row)
    return hits


# ------------------------------------------------------------------ formatting
def curl_output(portal: "C.Portal", pairs: list[tuple[str, str]]) -> None:
    cookies = "; ".join(f"{c.name}={c.value}" for c in portal.s.cookies
                        if c.name in ("AuthToken", "ASP.NET_SessionId"))
    body = urlencode(pairs)
    cmd = ("curl --request POST "
           f"'{URL_TRANSFER}' "
           "--header 'Content-Type: application/x-www-form-urlencoded; charset=UTF-8' "
           "--header 'X-Requested-With: XMLHttpRequest' "
           "--header 'X-MicrosoftAjax: Delta=true' "
           f"--header 'Referer: {URL_TRANSFER}' "
           f"--cookie '{cookies}' "
           f"--data-raw '{body}'")
    print("\n--- one-shot cURL (session-bound; sending it = one real transfer) ---")
    print(cmd)
    print("------------------------------------------------------------------\n")


# ------------------------------------------------------------------------ main
def main() -> None:
    ap = argparse.ArgumentParser(
        description="GUARDED one-shot balance transfer on the Limerick reseller portal "
                    "(accounts tree: distributor -> sub-partner). Submits exactly once.")
    ap.add_argument("--account", choices=("shreeganesh", "reseller"), default="shreeganesh",
                    help="source login; default shreeganesh (= username_shreeganesh key in .env)")
    ap.add_argument("--level", choices=("sub", "lco"), default="sub",
                    help="Entity Level to pick the recipient dropdown from (default sub=Sub Distributor)")
    sel = ap.add_mutually_exclusive_group(required=True)
    sel.add_argument("--partner-id", type=int, help="numeric partner id from the live dropdown (e.g. 54)")
    sel.add_argument("--partner-name", help="part of the partner name, must match exactly one dropdown item")
    ap.add_argument("--amount", type=float, required=True, help="amount to move (2 decimals)")
    ap.add_argument("--date", help="transfer date, 'dd MMM yyyy' (default: the page's own default = today)")
    ap.add_argument("--remark", default="", help="optional remark for the transfer")
    ap.add_argument("--payment-received", choices=("yes", "no"), default="yes",
                    help="Payment Received radio (portal default: Yes)")
    ap.add_argument("--dry-run", action="store_true",
                    help="walk every step and BUILD the payload, but submit NOTHING")
    ap.add_argument("--curl", action="store_true",
                    help="print a one-shot cURL instead of submitting (session-bound!)")
    ap.add_argument("--force", action="store_true", help="skip the yes/NO confirmation")
    ap.add_argument("--bypass-gate", action="store_true",
                    help="skip the anti-double-transfer gate (ONLY after the report was verified)")
    ap.add_argument("--env", default=str(SCRIPT_DIR / ".env"), help="path to the credential file")
    args = ap.parse_args()

    if args.partner_name is not None and not args.partner_name.strip():
        die("--partner-name must not be empty", 1)

    env_path = Path(args.env)
    try:
        creds = load_account(env_path, args.account)
    except Exception as e:
        die(str(e), 1)

    portal = C.Portal(env_path)
    portal.env = creds
    try:
        portal.login()
    except Exception as e:
        die(f"login failed: {e}", 1)
    log(f"[login] ok — account mode: {args.account}")

    # ---- 1. open the Transfer page, read live balance + default date ----
    page = Page(portal)
    balance = page.balance
    amount = round(args.amount, 2)
    if amount <= 0:
        die("amount must be > 0", 1)
    if amount > balance + 0.004:
        die(f"current balance {balance:.2f} cannot cover amount {amount:.2f} — aborting (exit 1)", 1)

    # ---- 2. Entity Level dropdown -> live partner list ----
    level_value = LEVEL_OPTIONS[args.level]
    labels = {v: t for v, t in page.entity_options}
    if level_value not in {v for v, _ in page.entity_options}:
        die(f"Entity Level option for {args.level} not offered to this account "
            f"(available: {page.entity_options})", 1)
    options = expand_entity_list(page, level_value)
    partners = [(int(v), t.replace("- ", "").strip()) for v, t in options
                if v and v != "0"]
    if not partners:
        die(f"no sub-partners listed under '{labels.get(level_value, args.level)}' for this account", 1)

    # ---- 3. resolve the target partner (unique match or refuse) ----
    if args.partner_id is not None:
        match = [p for p in partners if p[0] == args.partner_id]
        if not match:
            die(f"partner id {args.partner_id} not in the live dropdown "
                f"(available: {partners})", 1)
        partner_id, partner_name = match[0]
    else:
        frag = args.partner_name.strip().lower()
        match = [p for p in partners if frag in p[1].lower()]
        if not match:
            die(f"no partner matches '{args.partner_name}' (live options: {partners})", 1)
        if len(match) > 1:
            print("several partners match — refusing to guess:", file=sys.stderr)
            for pid, nm in match:
                print(f"    {pid}  {nm}", file=sys.stderr)
            die("re-run with a more specific --partner-name or the numeric --partner-id", 1)
        partner_id, partner_name = match[0]

    transfer_date = (args.date or page.date_default or
                     date.today().strftime("%d %b %Y")).strip()
    date_prefix = transfer_date  # the ledger prints this exact date string first

    # ---- 4. anti-double-transfer gate (static + live report) ----
    records = gate_load()
    reason = gate_check(records, partner_id, amount)
    if reason and not args.bypass_gate:
        print(f"[GATE] blocked: {reason}", file=sys.stderr)
        sys.exit(4)
    if not args.dry_run and not args.curl:
        today_rows = history_rows(portal)
        hits = history_match(today_rows, partner_name, amount, date_prefix)
        if hits and not args.bypass_gate:
            print("[GATE] the ledger already shows a Transfer Balance row for this partner "
                  f"dated {date_prefix} — same-day duplicate refused:", file=sys.stderr)
            for rw in hits[:3]:
                print("    " + rw[:220], file=sys.stderr)
            sys.exit(4)

    # ---- 5. payload ----
    pairs = list(page.fields)
    set_field(pairs, "__EVENTTARGET", "")
    set_field(pairs, "__EVENTARGUMENT", "")
    set_field(pairs, "__LASTFOCUS", "")
    set_field(pairs, "ctl00$ContentPlaceHolder1$ucTransferBalancePartner$ddlEntityLevel", level_value)
    set_field(pairs, "ctl00$ContentPlaceHolder1$ucTransferBalancePartner$ddlSubDistributor", str(partner_id))
    set_field(pairs, "ctl00$ContentPlaceHolder1$ucTransferBalancePartner$txtAmount", f"{amount:.2f}")
    set_field(pairs, "ctl00$ContentPlaceHolder1$ucTransferBalancePartner$txtTDate", transfer_date)
    set_field(pairs, "ctl00$ContentPlaceHolder1$ucTransferBalancePartner$txtRemark", args.remark.strip())
    set_field(pairs, "ctl00$ContentPlaceHolder1$ucTransferBalancePartner$rdbTPaymentReceived",
              "1" if args.payment_received == "yes" else "0")

    summary = [
        f"source account    : {args.account}",
        f"target partner    : {partner_id}  {partner_name}",
        f"amount            : {amount:.2f}",
        f"current balance   : {balance:.2f}  ->  after transfer {balance - amount:.2f}",
        f"transfer date     : {transfer_date}",
        f"payment received  : {args.payment_received}",
        f"remark            : {args.remark.strip() or '(empty)'}",
        f"submit target     : POST {URL_TRANSFER} (single WebForms post, gates active)",
    ]
    print("\n=== TRANSFER SUMMARY ===", file=sys.stderr)
    for s in summary:
        print(s, file=sys.stderr)
    print("========================", file=sys.stderr)

    if args.dry_run:
        print("[dry-run] payload built and NOT submitted; no gate file touched.", file=sys.stderr)
        named = [(k.split("$")[-1], v) for k, v in pairs if "$ucTransferBalancePartner$" in k]
        for k, v in named:
            shown = (v[:120] + "…") if len(v) > 120 else v
            print(f"[dry-run]   {k} = {shown!r}", file=sys.stderr)
        print("[dry-run] (plus the page's hidden ASP.NET state fields, untouched)", file=sys.stderr)
        return

    if args.curl:
        curl_output(portal, pairs)
        return

    if not args.force:
        try:
            ans = input("Submit this transfer? This moves real money! (yes/NO): ")
        except EOFError:
            ans = ""
        if ans.strip().lower() not in ("y", "yes"):
            print("[abort] nothing submitted (exit 0)", file=sys.stderr)
            return

    # ---- 6. gate marker BEFORE the POST ----
    rec = {"ts": time.strftime("%Y-%m-%d %H:%M:%S"), "date_iso": date.today().isoformat(),
           "date": transfer_date, "account": args.account, "partner_id": partner_id,
           "partner_name": partner_name, "amount": amount, "outcome": "attempted"}
    gate_record(records, rec)
    log("[gate] 'attempted' marker written")

    # ---- 7. SUBMIT EXACTLY ONCE ----
    headers = {"Referer": URL_TRANSFER, "X-MicrosoftAjax": "Delta=true", "X-Requested-With": "XMLHttpRequest"}
    log("[submit] firing the transfer POST (exactly once, no retry on failure)")
    try:
        rp = portal.s.post(URL_TRANSFER, data=pairs, headers=headers, timeout=120)
        status = rp.status_code
        body = rp.text
    except requests.RequestException as e:
        log(f"[submit] connection problem: {e.__class__.__name__}: {e}")
        log("[submit] OUTCOME UNKNOWN — do NOT re-run; verify the Balance Transfer "
            "History report, then resolve transfer_gate.json manually (--bypass-gate after)")
        sys.exit(3)

    if status != 200:
        log(f"[submit] HTTP {status} — UPDATE: outcome unknown; the click may or may not "
            "have executed. Verify the Balance Transfer History report before any re-run.")
        sys.exit(3)

    if "|error|500|" in body[:40]:
        log("[submit] portal error packet — outcome unknown; verify the report")
        sys.exit(3)

    # ---- 8. verify on the ledger (the only source of truth) ----
    confirmed = None
    for attempt in range(4):
        rows = history_rows(portal)
        hits = history_match(rows, partner_name, amount, date_prefix)
        if hits:
            confirmed = hits[0]
            break
        if attempt < 3:
            log(f"[verify] not visible yet ({attempt + 1}/4) — waiting 15s...")
            time.sleep(15)

    if confirmed is None:
        log("[verify] no Transfer Balance row found on the ledger for today "
            "with this partner and amount after 4 checks.")
        log("[result] FAILED-FLAGGED or never visible — verify manually; the gate "
            "stays 'attempted' so re-runs are blocked until resolved.")
        # conservative: leave the marker as 'attempted'; the report decides
        rec["outcome"] = "unknown"
        gate_save(records)
        sys.exit(2)

    # success -> update the gate + write the log (newest first)
    rec["outcome"] = "confirmed"
    gate_save(records)
    log_transfer([time.strftime("%Y-%m-%d %H:%M:%S"), args.account,
                  str(partner_id), partner_name, f"{amount:.2f}", args.remark.strip()])
    mclose = re.search(r"([-\d,]+\.\d{2})\s*\|\s*Transfer Balance", confirmed)
    print("\n[confirmed] transfer verified on the portal's own ledger:", file=sys.stderr)
    print("    " + confirmed[:240], file=sys.stderr)
    if mclose:
        print(f"    closing wallet balance: {mclose.group(1)}", file=sys.stderr)


if __name__ == "__main__":
    main()
