#!/usr/bin/env python3
"""
crm_customers.py — Limerick Broadband reseller portal (Height8 iCRM) customer fetcher.

Sub-commands:
  all            Fetch every customer row from the Accounts grid (incl. CustomerID),
                 clean it and show it in a structured table. Optional CSV/JSON export.
  get <ID>       Fetch detailed info for ONE customer: grid summary, current plan info
                 (from the renewal dialog - read-only, no recharge) and address info.

Reads credentials from .env next to this script (format used there:
    username :- <value>
    password :- <value>
). Nothing is hard-coded. All tokens/values are taken from live responses.

Examples:
  python crm_customers.py all --limit 20
  python crm_customers.py all --csv customers.csv --json customers.json
  python crm_customers.py get 123456
"""
from __future__ import annotations

import argparse
import csv
import json
import re
import sys
from datetime import datetime
from pathlib import Path

import requests
from bs4 import BeautifulSoup

BASE = "https://reseller.limerickbroadband.in"
LOGIN_PAGE = f"{BASE}/Partner/crmhomepage.aspx"
LOGIN_POST = f"{BASE}/Partner/PortalLogin.aspx?h8=1"
ACCOUNTS = f"{BASE}/Partner/Accounts.aspx?DshID=10026"
RENEW_DIALOG = f"{BASE}/Partner/Account/PurchaseRenewPlan/{{cid}}"
CUSTOMER_DETAIL = f"{BASE}/Partner/Account/CustomerDetail/{{cid}}"

# hard date window: anything with an expiry in this range is returned by the grid
WIDE_START = "01 Jan 2000"
WIDE_END = "31 Dec 2027"

UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/156.0 Safari/537.36"

FIELD_NAMES = ["customer_id", "user_id", "name", "mobile", "status", "expiry", "last_renewal"]


# ---------------------------------------------------------------- utilities
def load_env(path: Path) -> dict:
    creds = {}
    for line in path.read_text(encoding="utf-8", errors="ignore").splitlines():
        m = re.match(r"\s*username\s*:-\s*(.+)", line)
        if m:
            creds["username"] = m.group(1).strip()
        m = re.match(r"\s*password\s*:-\s*(.+)", line)
        if m:
            creds["password"] = m.group(1).strip()
    if "username" not in creds or "password" not in creds:
        raise RuntimeError(f"credentials not found in {path}")
    return creds


def hidden(html: str, name: str) -> str:
    m = re.search(r'name="' + re.escape(name) + r'"[^>]*value="([^"]*)"', html)
    return m.group(1) if m else ""


def clean(text: str | None) -> str | None:
    if text is None:
        return None
    # nbsp, zero-width, replacement char, control chars -> normalise/strip
    t = (text.replace("&nbsp;", " ").replace("\u00a0", " ")
             .replace("\u200b", "").replace("\ufffd", " "))
    t = re.sub(r"[\x00-\x1f\x7f]", "", t)
    t = re.sub(r"\s+", " ", t).strip()
    if t in ("", "-", "--", "N/A", "n/a"):
        return None
    return t


def normalise_mobile(v: str | None) -> str | None:
    if not v:
        return None
    m = re.sub(r"[^\d+]", "", v)
    return m if len(m) >= 7 else None


def to_iso(date_str: str | None) -> str | None:
    if not date_str:
        return None
    try:
        return datetime.strptime(date_str.strip(), "%d %b %Y").strftime("%Y-%m-%d")
    except ValueError:
        return None


# ---------------------------------------------------------------- session
class Portal:
    def __init__(self, env_path: Path):
        self.s = requests.Session()
        self.s.headers.update({"User-Agent": UA})
        self.token = ""
        self.env = load_env(env_path)

    def login(self) -> None:
        r0 = self.s.get(LOGIN_PAGE)
        r0.raise_for_status()
        lbody = {
            "__LASTFOCUS": "", "ToolkitScriptManager1_HiddenField": "",
            "__EVENTTARGET": "", "__EVENTARGUMENT": "",
            "__VIEWSTATE": hidden(r0.text, "__VIEWSTATE"),
            "__VIEWSTATEGENERATOR": hidden(r0.text, "__VIEWSTATEGENERATOR"),
            "__EVENTVALIDATION": hidden(r0.text, "__EVENTVALIDATION"),
            "txtUserName": self.env["username"],
            "txtPassword": self.env["password"],
            "chkRember": "chk",
            "save": "LOGIN",
        }
        lp = self.s.post(LOGIN_POST, data=lbody,
                         headers={"Referer": LOGIN_PAGE})
        lp.raise_for_status()
        has_auth = any(c.name == "AuthToken" for c in self.s.cookies)
        if not has_auth:
            raise RuntimeError("login failed (no AuthToken cookie) — check credentials")
        print(f"[login] ok (AuthToken received)", file=sys.stderr)

    def accounts_search(self) -> str:
        """POST the 'Search' with alpha=ALL, wide dates, 500/page.
        Returns the MS-AJAX delta response text."""
        r = self.s.get(ACCOUNTS, headers={"Referer": f"{BASE}/Partner/Default.aspx"})
        r.raise_for_status()
        html = r.text

        pairs: list[tuple[str, str]] = []
        seen: set[str] = set()

        def add(k: str, v: str):
            if k not in seen:
                seen.add(k)
                pairs.append((k, v))

        pairs += [("__EVENTTARGET", ""), ("__EVENTARGUMENT", ""), ("__LASTFOCUS", ""),
                  ("__VIEWSTATE", hidden(html, "__VIEWSTATE")),
                  ("__VIEWSTATEGENERATOR", hidden(html, "__VIEWSTATEGENERATOR")),
                  ("__EVENTVALIDATION", hidden(html, "__EVENTVALIDATION"))]

        # all hidden/text/password inputs, preserve original values
        for m in re.finditer(r'<input[^>]*type=["\'](hidden|text|password)["\'][^>]*>', html):
            tag = m.group(0)
            nm = re.search(r'name="([^"]*)"', tag)
            if not nm:
                continue
            n = nm.group(1)
            if re.match(r"^(?:__VIEWSTATE|__VIEWSTATEGENERATOR|__EVENTTARGET|__EVENTARGUMENT|"
                        r"__EVENTVALIDATION|__LASTFOCUS|fakeusernameremembered|fakepasswordremembered)$", n):
                continue
            v = re.search(r'value="([^"]*)"', tag)
            val = v.group(1) if v else ""
            if n.endswith("txtserch"):
                val = WIDE_START
            elif n.endswith("txtenddate"):
                val = WIDE_END
            add(n, val)

        for m in re.finditer(r'<input[^>]*type=["\']checkbox["\'][^>]*>', html):
            tag = m.group(0)
            nm = re.search(r'name="([^"]*)"', tag)
            if not nm:
                continue
            if re.search(r"\bchecked\b", tag, re.I):
                v = re.search(r'value="([^"]*)"', tag)
                add(nm.group(1), v.group(1) if v else "on")

        alpha = "ctl00$ContentPlaceHolder1$rdblst"
        for m in re.finditer(r'<input[^>]*type=["\']radio["\'][^>]*>', html):
            tag = m.group(0)
            nm = re.search(r'name="([^"]*)"', tag)
            if not nm:
                continue
            if re.search(r"\bchecked\b", tag, re.I):
                if nm.group(1) == alpha:
                    continue
                v = re.search(r'value="([^"]*)"', tag)
                add(nm.group(1), v.group(1) if v else "")

        for m in re.finditer(r'<select[^>]*name="([^"]*)"[^>]*>(.*?)</select>', html, re.S):
            n, inner = m.group(1), m.group(2)
            if n == alpha:
                continue
            for o in re.finditer(r"<option[^>]*\bselected\b[^>]*>", inner):
                v = re.search(r'value="([^"]*)"', o.group(0))
                val = v.group(1) if v else ""
                if n == "ctl00$ContentPlaceHolder1$ddlPageSize":
                    val = "500"
                pairs.append((n, val))

        add(alpha, "ALL")
        add("ctl00$ContentPlaceHolder1$btnserch", "Search")

        r2 = self.s.post(ACCOUNTS, data=pairs, headers={
            "Referer": ACCOUNTS, "X-MicrosoftAjax": "Delta=true",
            "X-Requested-With": "XMLHttpRequest",
        })
        r2.raise_for_status()
        self._refresh_from_delta(r2.text)
        return r2.text

    def _refresh_from_delta(self, delta: str) -> None:
        for name in ("__VIEWSTATE", "__EVENTVALIDATION"):
            m = re.search(r'hiddenField\|' + name + r'\|([^\|]+)\|', delta)
            if m:
                if name == "__VIEWSTATE":
                    self.new_viewstate = m.group(1)
                else:
                    self.new_eventvalidation = m.group(1)

    def next_page(self, delta: str, page_no: int) -> str:
        """Click grid pager page -> returns next delta (uses fresh ViewState)."""
        pairs = [("__EVENTTARGET", "ctl00$ContentPlaceHolder1$gdcomp"),
                 ("__EVENTARGUMENT", f"Page${page_no}"),
                 ("__LASTFOCUS", ""),
                 ("__VIEWSTATE", self.new_viewstate),
                 ("__EVENTVALIDATION", self.new_eventvalidation)]
        r = self.s.post(ACCOUNTS, data=pairs, headers={
            "Referer": ACCOUNTS, "X-MicrosoftAjax": "Delta=true",
            "X-Requested-With": "XMLHttpRequest",
        })
        r.raise_for_status()
        self._refresh_from_delta(r.text)
        return r.text


# ---------------------------------------------------------------- parsing
def update_panel(delta: str, panel_id: str = "ContentPlaceHolder1_updt") -> str:
    m = re.search(r"\|(\d+)\|updatePanel\|" + re.escape(panel_id) + r"\|", delta)
    if not m:
        return ""
    start = m.end()
    return delta[start:start + int(m.group(1))]


def parse_rows(panel: str) -> list[dict]:
    """Extract one dict per customer row found in the update-panel fragment."""
    out = []
    for tr_m in re.finditer(r"(?s)<tr[^>]*>.*?</tr>", panel):
        tr = tr_m.group(0)
        if "gdcomp_Exdate_" not in tr:
            continue
        cust = None
        m = re.search(r'PurchaseRenewPlanById\((\d+)\)', tr)
        if m:
            cust = m.group(1)
        else:
            m = re.search(r'hdncustID[^>]*value="(\d+)"', tr)
            if m:
                cust = m.group(1)
        uid = re.search(r'gdcomp_LinkButton1_\d+[^>]*>([^<]+)</a>', tr)
        name = re.search(r'gdcomp_lblName_\d+[^>]*>([^<]+)</span>', tr)
        status = re.search(r'gdcomp_LBLStatus_\d+[^>]*>([^<]+)</span>', tr)
        expiry = re.search(r'gdcomp_Exdate_\d+[^>]*>([^<]+)</span>', tr)
        renewal = re.search(r'gdcomp_lblRenewalDate_\d+[^>]*>([^<]+)</span>', tr)
        mob = re.search(r'</span>\s*</td>\s*<td[^>]*>\s*([0-9+\-() ]{6,20})\s*</td>', tr)
        row = {
            "customer_id": int(cust) if cust else None,
            "user_id": clean(uid.group(1)) if uid else None,
            "name": clean(name.group(1)) if name else None,
            "mobile": clean(mob.group(1)) if mob else None,
            "status": clean(status.group(1)) if status else None,
            "expiry": clean(expiry.group(1)) if expiry else None,
            "last_renewal": clean(renewal.group(1)) if renewal else None,
        }
        if row["customer_id"] is not None:
            out.append(row)
    return out


def clean_records(rows: list[dict]) -> list[dict]:
    seen, out = set(), []
    for r in rows:
        if r["customer_id"] in seen:
            continue
        seen.add(r["customer_id"])
        rec = dict(r)
        rec["expiry_iso"] = to_iso(r["expiry"])
        rec["last_renewal_iso"] = to_iso(r["last_renewal"])
        rec["status"] = (r["status"] or "").title() if r["status"] else None
        rec["mobile"] = normalise_mobile(r["mobile"])
        out.append(rec)
    # Active customers first (soonest expiry first), then the rest
    out.sort(key=lambda r: (0 if (r["status"] or "").lower().startswith("active") else 1,
                            r["expiry_iso"] or "9999-99-99"))
    return out


def parse_renew_dialog(html: str) -> dict:
    """Read-only info from the renewal dialog (no recharge performed)."""
    def hval(id_attr: str) -> str | None:
        m = re.search(r'<input[^>]*id="' + re.escape(id_attr) + r'"[^>]*value="([^"]*)"', html)
        if not m:
            m = re.search(r'<input[^>]*value="([^"]*)"[^>]*id="' + re.escape(id_attr) + r'"', html)
        return clean(m.group(1)) if m else None

    info: dict = {"customer_id": hval("hdnCustomerId") or re.search(r"CustomerId:\s*'(\d+)'", html) and re.search(r"CustomerId:\s*'(\d+)'", html).group(1)}
    text = re.sub(r"<[^>]+>", "\n", html)
    text = re.sub(r"\n{2,}", "\n", text)
    labels = ["User ID", "Partner", "Partner Bal", "Plan Name", "Plan MRP"]
    for lab in labels:
        m = re.search(lab + r"\s*:\s*\n\s*([^\n]+)", text)
        if m:
            info[lab.lower().replace(" ", "_")] = clean(m.group(1))
    return info


def parse_customer_address(html: str) -> dict:
    soup = BeautifulSoup(html, "lxml")
    data: dict = {}
    get = lambda sid: soup.find(id=sid)
    if get("CustomerId"):
        data["customer_id"] = int(get("CustomerId")["value"])
    if get("Name"):
        data["full_name"] = clean(get("Name").get("value"))
    for hid, key in (("hdnTitleId", "title"), ("txtNewAddress", "present_address"),
                     ("txtZone1", "present_city"), ("txtPincode", "present_pincode"),
                     ("ddlDistrict", "district_id"), ("hdnStateName", "state"),
                     ("txtNation", "country"), ("txtPNewAddress", "payment_address"),
                     ("txtPayPincode", "payment_pincode"), ("txtPLandmark", "payment_landmark")):
        el = get(hid)
        if el is not None and el.get("value"):
            data[key] = clean(el["value"])
    for sid, key in (("ddlWard", "ward"),):
        el = get(sid)
        if el is not None:
            sel = el.find("option", selected=True)
            if sel:
                data[key] = clean(sel.get_text())
    return {k: v for k, v in data.items() if v is not None}


# ---------------------------------------------------------------- display
def print_table(rows: list[dict], limit: int | None = 20) -> None:
    if not rows:
        print("no rows")
        return
    widths = {k: max(len(k.replace('_', ' ')), *(len(str(r[k] or "—")) for r in rows)) for k in FIELD_NAMES}
    widths = {k: min(w, 28) for k, w in widths.items()}
    head = [k.replace('_', ' ').ljust(widths[k]) for k in FIELD_NAMES]
    print("  ".join(head))
    print("-" * (sum(widths.values()) + 2 * (len(widths) - 1)))
    shown = rows[:limit] if limit else rows
    for r in shown:
        line = []
        for k in FIELD_NAMES:
            v = str(r.get(k) or "-")
            line.append(v[:widths[k]].ljust(widths[k]))
        print("  ".join(line))
    if limit and len(rows) > limit:
        print(f"... {len(rows) - limit} more rows not shown")


def status_summary(rows: list[dict]) -> dict:
    s: dict[str, int] = {}
    for r in rows:
        st = r["status"] or "Unknown"
        s[st] = s.get(st, 0) + 1
    return s


# ---------------------------------------------------------------- subcommands
def cmd_all(portal: Portal, args) -> list[dict]:
    delta = portal.accounts_search()
    panel = update_panel(delta)
    rows = parse_rows(panel)
    if not rows:
        head = delta[:60].replace("\n", " ")
        print(f"[warn] grid returned no rows (delta head: {head!r}, panel bytes: {len(panel)})", file=sys.stderr)
    seen_pages = 1
    while True:
        page_links = sorted({int(x) for x in re.findall(r"Page\$(\d+)", panel)})
        nxt = [p for p in page_links if p > seen_pages]
        if not nxt:
            break
        try:
            delta = portal.next_page(delta, min(nxt))
        except Exception as exc:
            print(f"[warn] pager page {min(nxt)} failed: {exc}", file=sys.stderr)
            break
        panel = update_panel(delta)
        got = parse_rows(panel)
        rows.extend(got)
        seen_pages += 1
    records = clean_records(rows)
    print(f"[fetch] customers on page(s): {len(set(r['customer_id'] for r in rows))} "
          f"(pager pages processed: {seen_pages})", file=sys.stderr)
    return records


def cmd_get(portal: Portal, args) -> dict:
    record = cmd_all(portal, argparse.Namespace(limit=None, csv=None, json=None, json_out=None))
    cust_id = int(args.cid)
    summary = next((r for r in record if r["customer_id"] == cust_id), None)
    if summary is None:
        raise RuntimeError(f"customer {cust_id} not found in Accounts grid")

    token_r = portal.s.get(ACCOUNTS, headers={"Referer": f"{BASE}/Partner/Default.aspx"})
    token = hidden(token_r.text, "__RequestVerificationToken")
    dlg = portal.s.get(RENEW_DIALOG.format(cid=cust_id), headers={
        "Referer": ACCOUNTS, "X-Requested-With": "XMLHttpRequest",
        "x-requestverificationtoken": token, "Accept": "text/html, */*; q=0.01"})
    plan_info = parse_renew_dialog(dlg.text) if dlg.status_code == 200 else {}

    det = portal.s.get(CUSTOMER_DETAIL.format(cid=cust_id), headers={
        "Referer": ACCOUNTS, "X-Requested-With": "XMLHttpRequest",
        "x-requestverificationtoken": token, "Accept": "text/html, */*; q=0.01"})
    address = parse_customer_address(det.text) if det.status_code == 200 else {}

    profile = {
        "summary": summary,
        "plan_info_from_renew_dialog": {k: v for k, v in plan_info.items() if k != "customer_id"},
        "address_info": address,
    }
    return profile


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--env", default=None, help="path to .env (default: script dir)")
    sub = ap.add_subparsers(dest="cmd", required=True)
    ap_all = sub.add_parser("all", help="fetch & clean all customers")
    ap_all.add_argument("--limit", type=int, default=20, help="rows to display (0 = all)")
    ap_all.add_argument("--csv", help="write cleaned records to CSV")
    ap_all.add_argument("--json", dest="json_out", help="write cleaned records to JSON")
    ap_get = sub.add_parser("get", help="full info of one customer id")
    ap_get.add_argument("cid")
    ap_get.add_argument("--json", dest="json_out", help="write profile to JSON")
    args = ap.parse_args()

    env_path = Path(args.env) if args.env else Path(__file__).resolve().parent / ".env"
    portal = Portal(env_path)
    portal.login()

    if args.cmd == "all":
        records = cmd_all(portal, args)
        counts = status_summary(records)
        print(f"\n[summary] total customers: {len(records)} | by status: "
              f"{', '.join(f'{k}={v}' for k, v in sorted(counts.items()))}")
        if records:
            exp = [r["expiry_iso"] for r in records if r["expiry_iso"]]
            print(f"[summary] expiry range: {min(exp)} .. {max(exp)}")
        print_table(records, limit=(args.limit if args.limit and args.limit > 0 else None))
        if args.csv:
            with open(args.csv, "w", newline="", encoding="utf-8-sig") as f:
                w = csv.DictWriter(f, fieldnames=list(records[0].keys()))
                w.writeheader()
                w.writerows(records)
            print(f"[saved] {args.csv}")
        if args.json_out:
            Path(args.json_out).write_text(json.dumps(records, indent=2, ensure_ascii=False), encoding="utf-8")
            print(f"[saved] {args.json_out}")
        if args.csv or args.json_out:
            print("[note] output files contain personal data - store & share them carefully.")
    else:
        profile = cmd_get(portal, args)
        print(json.dumps(profile, indent=2, ensure_ascii=False))
        if args.json_out:
            Path(args.json_out).write_text(json.dumps(profile, indent=2, ensure_ascii=False), encoding="utf-8")
            print(f"[saved] {args.json_out}")


if __name__ == "__main__":
    main()
