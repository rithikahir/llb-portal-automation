#!/usr/bin/env python3
"""
test_backend.py — read-only tester for the mobile-app backend
(gst-mobile-app\\appscript\\Code.js deployed as a Web App).

Only calls actions that NEVER write to the sheet:
  ping       — is the web app alive and the secret correct?
  list       — fetch customers + expiries + invoice history (as the phone does)
  pdf N      — regenerate the PDF for invoice no N (e.g. ST/053/2026-27)
               and save it to  test_pdfs\\  here (no sheet change, no Drive)
  template D — register the Google Doc id D as the bill WORD template
               (created by uploading GST_BILL_TEMPLATE.docx to Drive and
               opening it with Google Docs; read-only for the sheet)

Usage:
  python test_backend.py ping
  python test_backend.py list
  python test_backend.py list --limit 10
  python test_backend.py pdf ST/053/2026-27

Credentials come from the project .env (appscript_url / appscript_secret)
— the same values the phone app will ask you to type in.
"""
from __future__ import annotations

import argparse
import json
import re
import sys
import urllib.error
import urllib.request
from datetime import date, datetime
from pathlib import Path

HERE = Path(__file__).resolve().parent
ENV_CANDIDATES = [HERE / ".env", HERE.parent / ".env"]


def setting(key: str) -> str:
    for env in ENV_CANDIDATES:
        if not env.exists():
            continue
        for line in env.read_text(encoding="utf-8", errors="ignore").splitlines():
            m = re.match(r"\s*" + re.escape(key) + r"\s*:-\s*(.+)", line)
            if m:
                return m.group(1).strip()
    return ""


def call(action: str, extra: dict | None = None, url_override: str = "") -> dict:
    url = url_override or setting("appscript_url")
    secret = setting("appscript_secret")
    if not url or not secret:
        sys.exit(".env with appscript_url / appscript_secret not found next to or above this script")
    payload = json.dumps({"action": action, "secret": secret, **(extra or {})}).encode()
    req = urllib.request.Request(url, data=payload,
                                 headers={"Content-Type": "application/json"}, method="POST")
    try:
        with urllib.request.urlopen(req, timeout=120) as resp:
            out = json.loads(resp.read().decode())
    except urllib.error.HTTPError as exc:
        sys.exit(f"HTTP {exc.code}: {(exc.read() or b'').decode('utf-8', 'replace')[:400]}")
    if not out.get("ok"):
        sys.exit("server error: " + str(out.get("error")))
    return out


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("command", choices=["ping", "list", "pdf", "template",
                                        "folder", "docs", "checkdocs", "rebuild", "pdfs"])
    ap.add_argument("invoice_no", nargs="?", default="")
    ap.add_argument("--limit", type=int, default=0)
    ap.add_argument("--url", default="",
                    help="web app /exec URL to test (default: .env appscript_url — "
                         "pass the MOBILE app URL here to test the phone backend)")
    args = ap.parse_args()

    if args.command == "ping":
        out = call("ping", url_override=args.url)
        print(f"OK — service '{out.get('service')}' alive at {out.get('time')}")
        return

    if args.command == "list":
        out = call("list_all", url_override=args.url)
        today = date.today()
        print(f"today: {out['today']}   customers: {len(out['customers'])}\n")
        rows = out["customers"]
        if args.limit:
            rows = rows[: args.limit]
        print(f"{'#':>3}  {'customer':34} {'expiry':10} {'days':>5}  {'duration':9} {'last invoice':>14} {'total':>10} {'status':>8}")
        for i, c in enumerate(rows, 1):
            days = ""
            if c["expire_iso"]:
                y, m, d = (int(x) for x in c["expire_iso"].split("-"))
                days = str((date(y, m, d) - today).days)
            last = c["last_invoice"] or {}
            print(f"{i:>3}  {c['name'][:34]:34} {c['expire']:10} {days:>5}  "
                  f"{(c['duration'] or '')[:9]:9} {str(last.get('invoice_no', '')):>14} "
                  f"{(last.get('total') or 0):>10.2f} {str(last.get('status', ''))[:8]:>8}")
        print(f"\nbills expiring today: "
              f"{sum(1 for c in out['customers'] if c['expire_iso'] and c['expire_iso'] == today.isoformat())}")
        return

    if args.command == "pdf":
        if not args.invoice_no:
            sys.exit("usage:  python test_backend.py pdf ST/053/2026-27")
        out = call("make_pdf", {"invoice_no": args.invoice_no}, url_override=args.url)   # archived PDF, no sheet writes
        import base64
        dest = HERE / "test_pdfs"
        dest.mkdir(exist_ok=True)
        name = (out.get("pdf_name") or (args.invoice_no.replace("/", "_") + ".pdf")).replace("/", "_")
        path = dest / name
        path.write_bytes(base64.b64decode(out["pdf_base64"]))
        print(f"saved {path}  ({path.stat().st_size} bytes)")
        return

    if args.command == "template":
        if not args.invoice_no:
            sys.exit('usage:  python test_backend.py template <googleDocId>   ("clear" removes it)')
        doc_id = "" if args.invoice_no.lower() == "clear" else args.invoice_no
        out = call("set_props", {"props": {"BILL_TEMPLATE_DOC_ID": doc_id}}, url_override=args.url)
        if doc_id:
            print(f"word template registered: doc {doc_id} (set: {out.get('set')})")
            print("rebuild_doc can now create customer docs from this template")
        else:
            print("word template cleared")
        return

    if args.command == "folder":
        if not args.invoice_no:
            sys.exit('usage:  python test_backend.py folder <driveFolderId>   ("clear" removes it)')
        fid = "" if args.invoice_no.lower() == "clear" else args.invoice_no
        out = call("set_props", {"props": {"DOCS_FOLDER_ID": fid}}, url_override=args.url)
        if fid:
            print(f"docs folder registered: {fid} (set: {out.get('set')})")
            print("next: python test_backend.py docs   (build the customer->doc mapping)")
        else:
            print("docs folder cleared")
        return

    if args.command == "docs":
        out = call("sync_docs", url_override=args.url)
        print(f"docs found: {out['docs_found']}   customers: {out['customers']}   mapped: {out['mapped']}\n")
        for e in out["report"]:
            print(f"  {e['doc'][:36]:36} {'GoogleDoc' if e['isGoogleDoc'] else 'docx/blob ':9} "
                  f"{str(e['match'])[:30]:30} {e['status']}")
        return

    if args.command == "checkdocs":
        out = call("check_docs", {"export_test": True}, url_override=args.url)
        ok = 0
        for e in out["results"]:
            print(f"  [{e['status'][:1]}] {e['customer'][:32]:32} {e['status']}")
            if e["status"] == "OK":
                ok += 1
        print(f"\nOK: {ok}/{len(out['results'])}")
        return

    if args.command == "rebuild":
        if not args.invoice_no:
            sys.exit('usage:  python test_backend.py rebuild "<exact customer name>"')
        out = call("rebuild_doc", {"customer": args.invoice_no}, url_override=args.url)
        print(f"rebuilt doc for {out['customer']}: id {out['doc_id']}")
        print(f"export test: {out.get('export_test')}   mapped: {out.get('mapped')}")
        return

    if args.command == "pdfs":
        if not args.invoice_no:
            sys.exit('usage:  python test_backend.py pdfs "<customer name>"')
        out = call("list_pdfs", {"customer": args.invoice_no}, url_override=args.url)
        print(f"archived PDFs for {out['customer']}: {len(out['files'])}"
              + (f"   folder: {out['folder_url']}" if out.get("folder_url") else ""))
        for f in out["files"]:
            print(f"  {f['name'][:60]:60} {f['size']:>9} bytes")
        return


if __name__ == "__main__":
    main()
