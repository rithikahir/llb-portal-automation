#!/usr/bin/env python3
"""
auto_renew.py — trigger a portal plan renewal (recharge) for a customer id, no UI.

Python port of auto_renew.ps1. Usage:

Usage:

    python auto_renew.py <CustomerId> [--force] [--curl] [--dry-run] [--env PATH]
    python auto_renew.py --by-name "customer name" [--force] [--curl] [--dry-run]

Reads credentials from .env next to this script (format: `username :- value`
/ `password :- value`). No values are hard-coded: customer, plan, cost, dates
and tokens are all parsed fresh from the portal's live responses every run.

SAFETY RULES (the vendor API is NOT idempotent):
  * Shows a summary and asks for confirmation (skip with --force).
  * Aborts BEFORE submitting if the data doesn't look sane or the partner
    balance cannot cover the cost.
  * ANTI-DOUBLE-RECHARGE GATE: a marker (renew_gate.json, next to this script)
    records "attempted" BEFORE the POST fires; an unconfirmed attempt, or a
    successful recharge for the same customer on the same day, blocks any new
    run (exit 4). Additionally the Recent Purchase/Renew Report is checked
    live: if the portal already shows a transaction for this customer today,
    the run is refused. To re-run deliberately after verifying the report,
    use --bypass-gate.
  * Submits DoPurchaseRenewPlan EXACTLY ONCE. On a network error / timeout the
    outcome is UNKNOWN — the script exits and tells you to check the Recent
    Purchase/Renew Report. It NEVER retries automatically.
  * A one-line record (customer name/id, recharge date, activation date, expiry
    date) is added to recharge_log.csv for every CONFIRMED recharge. New
    records are inserted at the top (newest first; older ones shift down).
    Dry-runs, prompt-aborts and failed submissions are never logged.
  * Same-plan renew only: plan taken from the dialog's selected option,
    payment mode = the dialog's first option (Cash), NoOfRenewal as-is.

Restricted renewal scope (matches auto_renew.ps1): plan-change, wallet top-up
and cheque/EDC payment modes are NOT implemented.

Exit codes: 0 ok | 1 aborted before submit (safe) | 2 submission failed-flagged
| 3 outcome unknown (verify on report before any re-run) | 4 blocked by the
anti-double-recharge gate (verify the report, then --bypass-gate if certain).
"""
from __future__ import annotations

import argparse
import csv
import html as ihtml
import json
import re
import sys
from datetime import datetime
from pathlib import Path
from urllib.parse import quote

# reuse the login/session machinery that is already working in crm_customers.py
sys.path.insert(0, str(Path(__file__).resolve().parent))
try:
    from crm_customers import Portal, hidden, parse_renew_dialog, clean_records, parse_rows, update_panel
except ImportError as exc:  # pragma: no cover
    raise SystemExit(f"error: need crm_customers.py in the same folder ({exc})")

BASE = "https://reseller.limerickbroadband.in"
ACCT_URL = f"{BASE}/Partner/Accounts.aspx?DshID=10026"
RENEW_GET = f"{BASE}/Partner/Account/PurchaseRenewPlan/{{cid}}"
RENEW_POST = f"{BASE}/Partner/Account/DoPurchaseRenewPlan"
REPORT_URL = f"{BASE}/Partner/ResReport.aspx"

AJAX_HEADERS = {
    "X-Requested-With": "XMLHttpRequest",
    "Accept": "text/html, */*; q=0.01",
}


# ------------------------------------------------------------------ parsing helpers (mirror of the .ps1 functions)
def hval(html: str, fid: str) -> str:
    """Value of an <input> located by its id or name attribute ('' if absent)."""
    m = re.search(r'<input[^>]*(?:id|name)="' + re.escape(fid) + r'"[^>]*value="([^"]*)"', html)
    if not m:
        m = re.search(r'<input[^>]*value="([^"]*)"[^>]*(?:id|name)="' + re.escape(fid) + r'"', html)
    return m.group(1).replace("&amp;", "&") if m else ""


def js_lit(html: str, lbl: str) -> str:
    """Inline-JS literal `Lbl: 'value'` rendered by the server in the dialog."""
    m = re.search(r"(?<![A-Za-z])" + re.escape(lbl) + r"\s*:\s*'([^']*)'", html)
    return m.group(1) if m else ""


def span_html(html: str, cls: str) -> str:
    """Text content of a <span class="..."> (tags stripped)."""
    m = re.search(r'(?s)class="' + re.escape(cls) + r'"[^>]*>(.*?)</span>', html)
    if not m:
        return ""
    return ihtml.unescape(re.sub(r"<[^>]+>", "", m.group(1))).strip()


def checked_radio(html: str, name: str) -> str:
    """Value of the CHECKED radio in a group ('' if none checked)."""
    for m in re.finditer(r'<input[^>]*name="' + re.escape(name) + r'"[^>]*>', html):
        if re.search(r"\bchecked\b", m.group(0), re.I):
            v = re.search(r'value="([^"]*)"', m.group(0))
            return v.group(1) if v else ""
    return ""


def selected_option(html: str, sel_id: str) -> str:
    """Selected option's value inside <select id=...>...</select>."""
    m = re.search(r'(?s)<select[^>]*id="' + re.escape(sel_id) + r'".*?</select>', html)
    if not m:
        return ""
    for o in re.finditer(r"<option[^>]*\bselected\b[^>]*>", m.group(0)):
        v = re.search(r'value="([^"]*)"', o.group(0))
        if v:
            return v.group(1)
    return ""


def first_option(html: str, sel_id: str) -> str:
    """First option's value inside <select id=...>...</select>."""
    m = re.search(r'(?s)<select[^>]*id="' + re.escape(sel_id) + r'".*?</select>', html)
    if not m:
        return ""
    v = re.search(r'<option[^>]*value="([^"]*)"', m.group(0))
    return v.group(1) if v else ""


# ------------------------------------------------------------------ anti-double-recharge gate
GATE_FILE = Path(__file__).resolve().parent / "renew_gate.json"


def gate_load() -> list[dict]:
    try:
        data = json.loads(GATE_FILE.read_text(encoding="utf-8"))
        return data if isinstance(data, list) else []
    except (FileNotFoundError, json.JSONDecodeError):
        return []          # no marker file -> no restriction, the report check still applies


def gate_save(entries: list[dict]) -> None:
    GATE_FILE.write_text(json.dumps(entries, indent=2, ensure_ascii=False), encoding="utf-8")




def gate_block_reason(entries: list[dict], cust_id: str, today: str) -> str | None:
    """Return a blocking reason for this customer today, or None."""
    for e in entries:
        if str(e.get("customer_id", "")) != cust_id:
            continue
        st = e.get("status")
        if st in ("attempted", "unknown"):
            return (f"an earlier run on {e.get('date', '?')} started a recharge for customer "
                    f"{cust_id} and its outcome was never confirmed ({st}). The recharge may "
                    f"still have gone through. Check the Recent Purchase/Renew Report "
                    f"({REPORT_URL}); re-run with --bypass-gate ONLY if you have verified no "
                    "recharge was made.")
        if st == "success" and e.get("date") == today:
            return (f"THIS SCRIPT already completed a recharge for customer {cust_id} today "
                    f"({e.get('message') or 'confirmed'}). Submitting again today would DOUBLE "
                    "the recharge and the debit. To charge again on purpose (e.g. you verified "
                    "the entry was for a different plan), run with --bypass-gate.")
    return None


def gate_write_attempt(entries: list[dict], cust_id: str, username: str, today: str) -> list[dict]:
    """Record 'attempted' BEFORE the POST fires, superseding stale unresolved/failed entries."""
    entries = [e for e in entries
               if not (str(e.get("customer_id", "")) == cust_id and e.get("status") in ("attempted", "unknown", "failed"))]
    entries.append({"customer_id": cust_id, "username": username, "date": today,
                    "ts": datetime.now().isoformat(timespec="seconds"), "status": "attempted"})
    gate_save(entries)
    return entries


def gate_close(entries: list[dict], cust_id: str, status: str, message: str = "") -> None:
    """Update this customer's most recent 'attempted' entry with the final outcome."""
    for e in reversed(entries):
        if str(e.get("customer_id", "")) == cust_id and e.get("status") == "attempted":
            e["status"] = status
            if message:
                e["message"] = message
            break
    gate_save(entries)


def report_recharged_today(portal: Portal, username: str, today: str) -> str | None:
    """Look on the Recent Purchase/Renew Report (current-month page) for a
    transaction row for `username` dated today. Returns the date/time line or None."""
    rp = portal.s.get(REPORT_URL, headers={"Referer": ACCT_URL}, timeout=60)
    txt = re.sub(r"(?is)<script.*?</script>", " ", rp.text)
    txt = re.sub(r"<[^>]+>", "\n", txt)
    lines = [re.sub(r"\s+", " ", ihtml.unescape(l).replace("&nbsp;", " ")).strip()
             for l in txt.split("\n")]
    lines = [l for l in lines if l]
    for i, line in enumerate(lines):
        if line == username:                     # report rows start with the user id
            window = lines[i:i + 25]             # datetime line sits a few rows below
            for w in window:
                if w.startswith(today + " "):
                    return w
    return None


# ------------------------------------------------------------------ recharge log (limited fields only)
RECHARGE_LOG = Path(__file__).resolve().parent / "recharge_log.csv"
LOG_FIELDS = ["customer_id", "customer_name", "recharge_date", "activation_date", "expiry_date"]


def log_recharge(customer_id: str, customer_name: str, recharge_date: str,
                 activation_date: str, expiry_date: str) -> None:
    """Add ONE line to recharge_log.csv per confirmed recharge, inserted at the
    TOP of the data (newest first; older logs shift down). Deliberately minimal:
    name/id, when charged, when activation starts, when it expires."""
    header, data = LOG_FIELDS, []
    if RECHARGE_LOG.exists():
        try:
            rows = list(csv.reader(RECHARGE_LOG.open(encoding="utf-8-sig")))
            rows = [r for r in rows if r]
            if rows:
                header, data = rows[0], rows[1:]
        except OSError:
            pass  # unreadable file -> start a fresh log rather than crash a success path
    new_row = [customer_id, customer_name, recharge_date, activation_date, expiry_date]
    with RECHARGE_LOG.open("w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(header)
        w.writerow(new_row)
        w.writerows(data)



# ------------------------------------------------------------------ main flow
def main() -> int:
    ap = argparse.ArgumentParser(description="One-shot portal renewal for a customer id (confirm-first, single submit).")
    ap.add_argument("custid", nargs="?", default=None,
                    help="numeric customer id (see: python crm_customers.py all); "
                         "omit if using --by-name/--userid")
    ap.add_argument("--by-name", help='resolve the customer by NAME or user id instead of an id, '
                                      'e.g. --by-name "rohit tambe". Only proceeds when the match '
                                      'is unique; lists all matches otherwise.')
    ap.add_argument("--userid", help='resolve the customer by USER ID (login name), case-insensitive, '
                                     'e.g. --userid ROHITT or --userid rohitt. Only proceeds when '
                                     'the match is unique.')
    ap.add_argument("--force", action="store_true", help="skip the yes/no confirmation prompt")
    ap.add_argument("--bypass-gate", action="store_true",
                    help="skip the anti-double-recharge gate (only after verifying the Recent "
                         "Purchase/Renew Report yourself)")
    ap.add_argument("--dry-run", action="store_true",
                    help="open the dialog and BUILD the payload but do not submit (safe test)")
    ap.add_argument("--curl", action="store_true",
                    help="instead of submitting, print a ready-to-import cURL command (e.g. for "
                         "Postman) with the freshly parsed values, cookies and token")
    ap.add_argument("--env", help="path to .env (default: next to this script)")
    args = ap.parse_args()

    selectors = sum(1 for x in (args.custid, args.by_name, args.userid) if x)
    if selectors > 1:
        ap.error("give only ONE customer selector: a customer id, --by-name, OR --userid")
    if selectors == 0:
        ap.error("provide a customer id, --by-name, or --userid")

    cust_id = args.custid
    if cust_id:
        cust_id = cust_id.strip()
        if not cust_id.isdigit():
            ap.error("customer id must be a number")
    by_name_mode = bool(args.by_name)
    userid_mode = bool(args.userid)

    env_path = Path(args.env) if args.env else Path(__file__).resolve().parent / ".env"

    gate_entries = []
    step = 0
    today = datetime.now().strftime("%d %b %Y")   # same format the portal report uses
    gate_reason = None

    # --- 0. anti-double-recharge gate (id-mode: local marker, no network) ---
    if cust_id:
        gate_entries = gate_load()
        gate_reason = gate_block_reason(gate_entries, cust_id, today)
    if gate_reason and not args.bypass_gate:
        print("SAFETY GATE blocked this run:", file=sys.stderr)
        print("  " + gate_reason, file=sys.stderr)
        return 4
    if gate_reason:
        print("[gate] --bypass-gate given: local gate cleared for this run "
              "(ensure you verified the Recent Purchase/Renew Report first)")

    # --- 1. login (reuses the verified Portal class) -------------------------
    step += 1
    portal = Portal(env_path)
    portal.login()

    # --- 1b. name/userid-mode: resolve the customer id from the Accounts grid
    if by_name_mode or userid_mode:
        delta = portal.accounts_search()
        records = clean_records(parse_rows(update_panel(delta)))
        if userid_mode:
            q = args.userid.strip().lower()
            matches = [r for r in records if (r.get("user_id") or "").lower() == q]
            sel_label = f"user id {args.userid!r}"
        else:
            q = args.by_name.strip().lower()
            matches = [r for r in records
                       if q in (r.get("name") or "").lower() or q == (r.get("user_id") or "").lower()]
            sel_label = f"name {args.by_name!r}"
        if not matches:
            print(f"[resolve] no customer found matching {sel_label}; aborting (nothing submitted)", file=sys.stderr)
            return 1
        if len(matches) > 1:
            print(f"[resolve] {len(matches)} customers match {sel_label} — refusing to guess:", file=sys.stderr)
            for r in sorted(matches, key=lambda x: str(x.get("name") or "")):
                print(f"            customer id={r.get('customer_id')}  user={r.get('user_id')}  "
                      f"name={r.get('name')}  status={r.get('status')}  expiry={r.get('expiry')}", file=sys.stderr)
            print("            re-run with the exact user id, or a longer --by-name phrase "
                  "(e.g. full name).", file=sys.stderr)
            return 1
        m = matches[0]
        cust_id = str(m.get("customer_id"))
        print(f"[resolve] unique match -> customer id {cust_id} "
              f"({m.get('user_id')} / {m.get('name')}, status {m.get('status')}, expiry {m.get('expiry')})")
        # run the local gate now that we know the id
        gate_entries = gate_load()
        gate_reason = gate_block_reason(gate_entries, cust_id, today)
        if gate_reason and not args.bypass_gate:
            print("SAFETY GATE blocked this run:", file=sys.stderr)
            print("  " + gate_reason, file=sys.stderr)
            return 4
        if gate_reason:
            print("[gate] --bypass-gate given: local gate cleared for this run "
                  "(ensure you verified the Recent Purchase/Renew Report first)")

    # --- 2. accounts page -> request verification token ---------------------
    step += 1
    acc = portal.s.get(ACCT_URL, headers={"Referer": f"{BASE}/Partner/Default.aspx"}, timeout=60)
    acc.raise_for_status()
    token = hidden(acc.text, "__RequestVerificationToken")
    print(f"[{step}] ACCOUNTS -> HTTP {acc.status_code}, token len={len(token)}")
    if not token:
        print("     no request verification token; aborting", file=sys.stderr)
        return 1

    # --- 3. open the renew dialog (read-only) -------------------------------
    step += 1
    dlg = portal.s.get(
        RENEW_GET.format(cid=cust_id),
        headers={"Referer": ACCT_URL, "x-requestverificationtoken": token, **AJAX_HEADERS},
        timeout=60,
    )
    print(f"[{step}] DIALOG PurchaseRenewPlan/{cust_id} -> HTTP {dlg.status_code}, bytes={len(dlg.text)}")
    if dlg.status_code != 200:
        print("     cannot open renewal dialog; aborting", file=sys.stderr)
        return 1
    d = dlg.text

    # --- 3b. sanity + summary for the operator ------------------------------
    js_cust = js_lit(d, "CustomerId")
    js_user = js_lit(d, "Username")
    plan_id = selected_option(d, "ddl_Plan")
    plan_name = hval(d, "hdnPlanName")
    cost = hval(d, "TotalPurchaseCost")
    act_date = hval(d, "ActivationDate")
    pay_date = hval(d, "PaymentDate")

    bal_info = parse_renew_dialog(d)          # reuse tested parser: partner_bal, plan_name, plan_mrp ...
    balance = bal_info.get("partner_bal", "")
    no_of_renewal = hval(d, "NoOfRenewal")

    print(f"     SUMMARY: customer={js_cust}/{js_user} plan={plan_id} ({plan_name}) "
          f"cost={cost} activation={act_date} paymentDate={pay_date} "
          f"noOfRenewal={no_of_renewal} balance={balance or '?'}")

    if not js_cust or js_cust != cust_id:
        print(f"     dialog customer id ({js_cust or 'empty'}) doesn't match requested ({cust_id}); aborting", file=sys.stderr)
        return 1
    if not plan_id or not re.fullmatch(r"\d+(\.\d+)?", cost or ""):
        print("     dialog did not yield a plan/cost; aborting", file=sys.stderr)
        return 1

    # pre-submit balance check: balance must cover the cost (nothing submitted yet)
    try:
        if balance and float(balance) < float(cost):
            print(f"     partner balance {balance} cannot cover cost {cost}; aborting before submit", file=sys.stderr)
            return 1
    except ValueError:
        pass  # unparseable balance: leave the decision to the confirmation prompt

    # --- 3c-2. server-side anti-double check: already recharged today? -------
    rep_when = report_recharged_today(portal, js_user, today)
    if rep_when and not args.bypass_gate:
        print(f"SAFETY GATE: the Recent Purchase/Renew Report already shows a transaction for "
              f"{js_user} dated today ({rep_when}). Submitting now risks a DOUBLE recharge.", file=sys.stderr)
        print("  Check the report / the plan of that transaction; re-run with --bypass-gate "
              "ONLY if you are sure today's entry is not this renewal.", file=sys.stderr)
        return 4
    if rep_when:
        print(f"[gate] --bypass-gate given: report shows a transaction today ({rep_when}) - proceeding anyway")

    if not args.force:
        try:
            conf = input(f"Submit renewal for {js_user} (plan {plan_id}, cost {cost})? [yes/NO] ")
        except EOFError:
            conf = ""
        if conf.strip().lower() not in ("y", "yes"):
            print("aborted by user")
            return 0

    # --- 3c. build the payload (same 92-field model as auto_renew.ps1) ------
    adv = checked_radio(d, "RenewalType") or "true"
    renew_type = checked_radio(d, "RenewalTypePurchaseRenew")
    wrr = re.search(r'<input[^>]*id="RenewWithoutRefund"[^>]*>', d)
    without_refund = "true" if (wrr and re.search(r"checked", wrr.group(0), re.I)) else "false"

    model = {
        "CustomerId": js_lit(d, "CustomerId"), "Username": js_lit(d, "Username"),
        "CustomerStatus": js_lit(d, "CustomerStatus"), "CustomerPartner": js_lit(d, "CustomerPartner"),
        "CustomerReseller": js_lit(d, "CustomerReseller"), "ButtonSaveText": "Save",
        "ActivationDate": hval(d, "ActivationDate"), "PlanName": hval(d, "hdnPlanName"),
        "NewPlanMRP": hval(d, "NewPlanMRP"), "OldPlanId": js_lit(d, "OldPlanId"),
        "OldPlanMRP": js_lit(d, "OldPlanMRP"), "PlanId": plan_id,
        "Validity": hval(d, "hdnValidity"), "VolumeQuota": hval(d, "hdnVolumeQuota"),
        "ProratedPlanValue": hval(d, "hdnProratedPlanValue"), "ProratedQuota": hval(d, "hdnProratedQuota"),
        "ProratedExpirayDateString": hval(d, "hdnProratedExpirayDate"),
        "NewExpirayDateString": hval(d, "hdnNewExpirayDate"),
        "AdvanceRenewal": adv, "RenewalTypePurchaseRenew": renew_type,
        "NoOfDiscount": hval(d, "hdnNoOfDiscount"), "Discount": hval(d, "Discount"),
        "DiscountReasonId": "", "DiscountPerAbs": hval(d, "hdnDiscountPerAbs"),
        "DiscountPerValue": hval(d, "hdnDiscountPerValue"),
        "MRP": hval(d, "MRP"), "Amount": hval(d, "Amount"), "RenewWithoutRefund": without_refund,
        "RefundCharge": hval(d, "hdnRefundCharge"), "PreFundCharge": hval(d, "hdnPreFundCharge"),
        "RefundSrNo": hval(d, "hdnRefundSrNo"), "CustomerRefundCharge": hval(d, "CustomerRefundCharge"),
        "BodCustRefundcharge": hval(d, "hdnBodCustRefundcharge"), "BodRefundcharge": hval(d, "hdnBodRefundcharge"),
        "BodPREFUND": hval(d, "hdnBodPREFUND"), "PaymentMode": first_option(d, "ddlPaymentMode"),
        "EDCPaymentPurchaseRenew": checked_radio(d, "EDCPaymentPurchaseRenew") or "1",
        "PaymentRemark": "", "IsUpgrade": hval(d, "hdnIsUpgrade"), "PlanType": hval(d, "hdnPlanType"),
        "IsBase": hval(d, "hdnIsBase"), "ChequeNo": "", "ChequeDate": "", "Branch": "", "Bank": "", "City": "",
        "PaymentDate": hval(d, "PaymentDate"), "MachineID": "", "TransactionID": "", "TransId": "", "InvoiceNO": "",
        "CAFNo": hval(d, "CAFNo"), "Remark": span_html(d, "Remark"),
        "PaymentRemarkReqired": hval(d, "hdnPaymentRemarkReqired"), "SelectCouponCode": hval(d, "SelectCouponCode"),
        "ButtonSave": hval(d, "hdnButtonSave"), "RenewalAdditionalPlan": hval(d, "RenewalAdditionalPlan"),
        "StaticIPRecurringCharge": span_html(d, "StaticIPRecurringCharge"),
        "StaticIPRecurringChargePR": span_html(d, "StaticIPRecurringChargePR"),
        "PlanValueChange": "0", "SelectVASDetailString": hval(d, "SelectVASDetailString"),
        "VasRevertDetailString": hval(d, "hdnVasRevertDetailString"),
        "RevertChargeDetailString": hval(d, "hdnRevertChargeDetailString"),
        "NoOfRenewal": hval(d, "NoOfRenewal"), "CashReceiptNo": "",
        "SelectVoiceDetailString": hval(d, "hdnSelectVoiceDetailString"),
        "SelectItemDetailString": hval(d, "hdnSelectItemDetailString"),
        "chargeId": hval(d, "hdnChargeID"), "taxGroupId": hval(d, "hdntaxGroupId"), "Price": hval(d, "hdnPrice"),
        "Usageprice": hval(d, "hdnUsageprice"), "TotalOfferPrice": hval(d, "hdnTotalOfferPrice"),
        "isfirstCusotmer": hval(d, "hdnisfirstCusotmer"), "CustomerName": hval(d, "hdnCustomerName"),
        "PlanChargeListString": hval(d, "hdnPlanChargeListString"), "LastModifiedDate": hval(d, "hdnLastModifiedDate"),
        "CustomerBillingStatus": hval(d, "hdnCustomerBillingStatus"), "__RequestVerificationToken": token,
        "Name": js_lit(d, "Name"), "Mobile": js_lit(d, "Mobile"), "Email": js_lit(d, "Email"), "OTP": "",
        "isOfferChargeRenew": hval(d, "hdnisOfferChargeRenew"),
        "SelectVideoDetailString": hval(d, "hdnSelectVideoDetailString"),
        "VideoPlanId": hval(d, "hdnVideoPlanId"), "offervariantId": hval(d, "hdnOfferVariantId"),
        "IsPrePaid": hval(d, "hdnIsPrePaid"), "TotalPurchaseCost": hval(d, "TotalPurchaseCost"),
        "IsMigrationbenefit": hval(d, "hdnIsMigrationbenefit"),
        "VideoCustRefundcharge": hval(d, "hdnVideoCustRefundcharge"),
        "RenewPartnerPriceDetailInfo": "True", "stbrefid": hval(d, "myHiddenField"),
    }
    # encode exactly like the browser: every key and value percent-encoded
    body = "&".join(f"{quote(str(k), safe='')}={quote(str(v), safe='')}" for k, v in model.items())
    print(f"     PAYLOAD: {len(model)} fields, {len(body)} bytes")

    if args.dry_run:
        show = ("CustomerId", "Username", "PlanId", "OldPlanId", "PaymentMode",
                "RenewalTypePurchaseRenew", "AdvanceRenewal", "NoOfRenewal",
                "ActivationDate", "PaymentDate", "TotalPurchaseCost")
        for k in show:
            print(f"     {k}={model.get(k, '')!r}")
        print("[dry-run] payload built OK; submission skipped (nothing was sent)")
        return 0

    if args.curl:
        cookie = "; ".join(f"{c.name}={c.value}" for c in portal.s.cookies)
        curl_cmd = (
            "curl --request POST \\\n"
            f"  '{RENEW_POST}' \\\n"
            "  --header 'Content-Type: application/x-www-form-urlencoded; charset=UTF-8' \\\n"
            "  --header 'X-Requested-With: XMLHttpRequest' \\\n"
            f"  --header 'x-requestverificationtoken: {token}' \\\n"
            f"  --header 'Referer: {ACCT_URL}' \\\n"
            "  --header 'Accept: application/json, text/javascript, */*; q=0.01' \\\n"
            f"  --cookie '{cookie}' \\\n"
            f"  --data-raw '{body}'"
        )
        print("[curl] one-shot cURL generated from THIS session's live values:", file=sys.stderr)
        print("  * valid only WITH these cookies and token, and only while the session lasts", file=sys.stderr)
        print("  * SENDING IT = ONE REAL RECHARGE. Send it ONCE; sending twice = double recharge.", file=sys.stderr)
        print("  * if you wait long, customer/plan state may change - prefer re-generating", file=sys.stderr)
        print(curl_cmd)
        return 0

    # --- 4. submit EXACTLY ONCE (gate marker written BEFORE the POST) -------
    step += 1
    gate_entries = gate_write_attempt(gate_entries, cust_id, js_user, today)
    try:
        r = portal.s.post(
            RENEW_POST, data=body, timeout=60,
            headers={
                "Referer": ACCT_URL, "x-requestverificationtoken": token,
                "X-Requested-With": "XMLHttpRequest",
                "Accept": "application/json, text/javascript, */*; q=0.01",
                "Content-Type": "application/x-www-form-urlencoded; charset=UTF-8",
            },
        )
    except Exception as exc:
        gate_close(gate_entries, cust_id, "unknown", f"submit error: {exc}")
        print(f"[{step}] SUBMIT ERROR: {exc}")
        print("     OUTCOME UNKNOWN. Do NOT retry; verify on Recent Purchase/Renew Report first "
              f"({REPORT_URL}).")
        return 3

    ok = re.search(r'"IsSuccess"\s*:\s*(true|false)', r.text)
    msg = re.search(r'"Message"\s*:\s*"([^"]*)"', r.text)
    col = re.search(r'"MessageColor"\s*:\s*"([^"]*)"', r.text)
    ok = ok.group(1) if ok else "?"
    print(f"[{step}] SUBMIT DoPurchaseRenewPlan -> HTTP {r.status_code}, IsSuccess={ok}, "
          f"color={col.group(1) if col else ''}, message='{ihtml.unescape(msg.group(1)) if msg else ''}'")
    if ok != "true":
        gate_close(gate_entries, cust_id, "failed", ihtml.unescape(msg.group(1)) if msg else "")
        print("     SUBMISSION FAILED-FLAGGED. Verify on Recent Purchase/Renew Report before any retry.")
        return 2
    gate_close(gate_entries, cust_id, "success", ihtml.unescape(msg.group(1)) if msg else "")

    # --- 5b. record the recharge in the log (limited fields only) -----------
    act = re.search(r'"ActivationDate"\s*:\s*"([^"]*)"', r.text)
    exp = re.search(r'"NewExpirayDateString"\s*:\s*"([^"]*)"', r.text)
    log_recharge(
        cust_id,
        model.get("CustomerName") or js_lit(d, "Name"),
        model.get("PaymentDate", ""),
        act.group(1) if act else model.get("ActivationDate", ""),
        exp.group(1) if exp else model.get("NewExpirayDateString", ""),
    )
    print(f"[{step}] LOG: recharge recorded in recharge_log.csv "
          "(customer name/id, recharge date, activation date, expiry date)")

    # --- 5. verify on the recent purchase report (independent authority) ----
    step += 1
    rp = portal.s.get(REPORT_URL, headers={"Referer": ACCT_URL}, timeout=60)
    txt = re.sub(r"(?is)<script.*?</script>", " ", rp.text)
    txt = re.sub(r"<[^>]+>", "\n", txt)
    lines = [re.sub(r"\s+", " ", ihtml.unescape(l).replace("&nbsp;", " ")).strip()
             for l in txt.split("\n")]
    t = "\n".join(l for l in lines if l)
    hit = re.search(re.escape(js_user) + r"\r?\n.{0,200}", t)
    if hit:
        print(f"[{step}] VERIFY report: found entry:")
        print("     " + hit.group(0)[:220].replace("\n", " | "))
    else:
        print(f"[{step}] VERIFY: entry for {js_user} not on page 1 of the report "
              "(check month/date or pages) - verify manually before treating this as complete")
    print("DONE.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
