#!/usr/bin/env python3
"""
run_billing.py — LOCAL-ONLY GST bill runner (no tunnel, no cron, no server).

You run this by hand whenever you want to bill. It:

  1. reads .env  (appscript_url, appscript_secret, pdf_dir, optional bills_dir)
  2. calls the Apps Script Web App (action "get_due") which returns every
     UPCOMING BILLS row whose Expire Date is YESTERDAY OR EARLIER
     (default; --days N goes further back, --include-today adds today).
  3. for each due customer it edits the bill .docx in place (Invoice Date =
     today, Invoice No = the exact number Apps Script computed, Bill Period
     advanced by the bill's own duration) and writes the PDF to the folder
     in .env "pdf_dir" (default: Desktop\\gst_bills).
  4. calls Apps Script again (action "complete") which appends the GST
     RECORDS row, moves the customer's Expire Date forward to the new
     period end and re-sorts UPCOMING BILLS by Expire Date (nearest first).
     Nothing is ever written to the Bill File / Bill Status columns.

Idempotent end-to-end: the same run can be re-executed safely. Python's
bills_state.json dedupes by key, Apps Script only appends a GST row when
that invoice number is not already in GST RECORDS, and Python refuses to
edit a bill whose invoice number doesn't match what the sheet expected.

Setup (one time):
  * Deploy Code.gs as a Web App (Deploy > New deployment > Web app;
    Execute as: Me; Who has access: Anyone) and paste its /exec URL into
    .env  appscript_url :-
  * Put the value of .env  appscript_secret :-  into SECRET in Code.gs.
  * Test:  python run_billing.py --dry-run

Windows: always  python run_billing.py ...   (never .\\run_billing.py)
"""
from __future__ import annotations

import argparse
import json
import re
import sys
import urllib.error
import urllib.request
from datetime import date, timedelta
from pathlib import Path

SCRIPT_DIR = Path(__file__).resolve().parent
sys.path.insert(0, str(SCRIPT_DIR))
import update_bills as U                     # noqa: E402  (reuses the editor)

ENV_FILE = SCRIPT_DIR / ".env"


# --------------------------------------------------------------------- .env
def setting(key: str, default: str | None = None) -> str | None:
    """Read one 'key :- value' line from .env (same format as the credentials)."""
    for line in ENV_FILE.read_text(encoding="utf-8", errors="ignore").splitlines():
        m = re.match(r"\s*" + re.escape(key) + r"\s*:-\s*(.+)", line)
        if m:
            return m.group(1).strip()
    return default


def call_apps_script(url: str, payload: dict) -> dict:
    data = json.dumps(payload, ensure_ascii=False).encode("utf-8")
    req = urllib.request.Request(url, data=data,
                                 headers={"Content-Type": "application/json"},
                                 method="POST")
    try:
        with urllib.request.urlopen(req, timeout=300) as resp:
            return json.loads(resp.read().decode("utf-8"))
    except urllib.error.HTTPError as exc:
        body = (exc.read() or b"").decode("utf-8", "replace")[:300]
        raise RuntimeError(f"Apps Script returned HTTP {exc.code}: {body}") from exc


# --------------------------------------------------------------------- main
def main() -> int:
    ap = argparse.ArgumentParser(
        description="Bill every UPCOMING BILLS row that is due (local run, no server needed).")
    ap.add_argument("--dry-run", action="store_true",
                    help="show what would be billed; edit nothing, write nothing to the sheet")
    ap.add_argument("--days", type=int, default=1,
                    help="bill rows whose Expire Date is this many days back or older "
                         "(default 1 = yesterday and before)")
    ap.add_argument("--include-today", action="store_true",
                    help="also include rows that expire TODAY")
    ap.add_argument("--period-shift", choices=("auto", "next", "year"), default="auto",
                    help="auto = advance by the bill's own duration (default)")
    ap.add_argument("--pdf-dir", help="override PDF output folder (default: .env pdf_dir)")
    ap.add_argument("--bills-dir", help="override bills folder (default: .env bills_dir)")
    ap.add_argument("--url", help="override the Apps Script web app URL (default: .env appscript_url)")
    args = ap.parse_args()

    # ---- config -----------------------------------------------------------
    url = args.url or setting("appscript_url") or ""
    if not url or "PASTE" in url.upper():
        print("ERROR: .env has no real 'appscript_url :-' yet.\n"
              "Deploy Code.gs as a Web App (Deploy > New deployment > Web app,\n"
              "Execute as: Me, access: Anyone), copy the /exec URL and put it in .env:\n"
              "    appscript_url :- https://script.google.com/macros/s/...../exec",
              file=sys.stderr)
        return 2
    secret = setting("appscript_secret") or ""
    pdf_dir = Path(args.pdf_dir or setting("pdf_dir") or U.PDF_DIR)
    bills_dir = Path(args.bills_dir or setting("bills_dir") or U.BILLS_DIR)

    today = date.today()
    cutoff = today if args.include_today else today - timedelta(days=max(0, args.days))

    # ---- 1. ask the sheet who is due --------------------------------------
    print(f"Asking Apps Script for UPCOMING BILLS with Expire Date <= {cutoff:%d/%m/%y} ...")
    resp = call_apps_script(url, {"action": "get_due", "secret": secret,
                                  "cutoff": f"{cutoff:%d/%m/%y}", "dry": args.dry_run})
    if not resp.get("ok"):
        print(f"ERROR from Apps Script: {resp.get('error')}", file=sys.stderr)
        return 2
    due = resp.get("due") or []
    if not due:
        print("Nothing is due. Nothing to do.")
        return 0
    print(f"{len(due)} bill(s) due"
          + ("" if not args.dry_run else "  [DRY RUN — nothing will be edited]"))
    for item in due:
        last = item.get("last_invoice_no")
        print(f"  - {item['customer']}  (expires {item['expire']}, last GST record: {last or 'NONE'})")

    # ---- 2. edit each bill locally ----------------------------------------
    results = []
    failed = 0
    for item in due:
        rec = {"customer": item["customer"], "new_invoice_no": item["new_invoice_no"],
               "key": item["key"]}
        if item.get("bill_file"):
            rec["bill_file"] = item["bill_file"]
        if item.get("last_invoice_no"):
            rec["expect_invoice_no"] = item["last_invoice_no"]   # double-bill guard
        try:
            res = U.process_record(rec, dry_run=args.dry_run, period_shift=args.period_shift,
                                   bills_dir=bills_dir, pdf_dir=pdf_dir)
            dup = res.get("duplicate", False)
            pdf_note = "" if res.get("pdf_created", True) else f"  [PDF FAILED: {res.get('pdf_error')}]"
            print(f"  {item['customer']}: {'(already done, skipped)' if dup else 'billed'}"
                  f" {res['new_invoice_no']}  period {res['old_bill_period']} -> {res['new_bill_period']}"
                  f"{pdf_note}")
            results.append({"key": item["key"], "ok": True, "duplicate": dup,
                            "new_invoice_no": res["new_invoice_no"],
                            "invoice_date": res["invoice_date"],
                            "new_expire": str(res["new_bill_period"]).split(" to ")[-1],
                            "pdf_error": res.get("pdf_error")})
        except Exception as exc:
            failed += 1
            print(f"  {item['customer']}: ERROR {exc.__class__.__name__}: {exc}", file=sys.stderr)
            results.append({"key": item["key"], "ok": False,
                            "error": f"{exc.__class__.__name__}: {exc}"})

    if args.dry_run:
        print("Dry run finished — nothing was saved, the sheet was not touched.")
        return 0

    # ---- 3. report back: append GST rows, set BILLED/ERROR -----------------
    resp2 = call_apps_script(url, {"action": "complete", "secret": secret, "results": results})
    if not resp2.get("ok"):
        print(f"ERROR while updating the sheet: {resp2.get('error')}\n"
              f"Your bills WERE edited and PDFs created; run this same command again to "
              f"finish updating the sheet (it is safe to re-run).",
              file=sys.stderr)
        return 3
    for line in resp2.get("log", []):
        print("  sheet: " + line)

    if failed:
        print(f"Done with {failed} error(s) — those customers were NOT changed in the sheet "
              f"and will be retried automatically on the next run.")
        return 3
    print("Done: GST RECORDS updated, Expire Dates advanced + tab sorted, PDFs in " + str(pdf_dir))
    return 0


if __name__ == "__main__":
    sys.exit(main())
