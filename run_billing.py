#!/usr/bin/env python3
"""
run_billing.py — THE one GST billing script (driver + docx/PDF engine).

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
     period end, updates the Duration column and re-sorts UPCOMING BILLS
     by Expire Date (nearest first).
     Nothing is ever written to the Bill File / Bill Status columns.

Idempotent end-to-end: the same run can be re-executed safely. The local
logs\\bills_state.json dedupes by key, Apps Script only appends a GST row
when that invoice number is not already in GST RECORDS, and the script
refuses to edit a bill whose invoice number doesn't match what the sheet
expected (double-bill guard).

ONE-SHOT MODE (bill one customer by hand, no sheet contact):
  python run_billing.py --customer "SKIN CITY AUNDH" --dry-run
  python run_billing.py --customer "SKIN CITY AUNDH" --invoice-no "ST/048/2026-27"
  Flags: --bill-file, --new-invoice-no, --invoice-date, --period-shift
  {auto,next,year}, --pdf-dir, --bills-dir. (The old standalone HTTP server
  mode of update_bills.py was retired — Apps Script never calls the PC.)

What gets updated inside a bill (label-anchored, so nothing else is touched):
  * Invoice Date : - <old>          -> today's date (same format as found)
  * Invoice No  : - ST/009/2026-27  -> ST/010/2026-27   (number +1, same width;
                                       or the exact "new_invoice_no" sent)
  * Bill Period:-<start> to <end>   -> advanced by the bill's OWN duration in
                                       whole calendar months (default "auto"):
                                       a 1-month bill moves one month, a
                                       3-month bill three, a 13-month bill
                                       thirteen. Alternatives: next (day
                                       count), year (both dates +1 year).

Setup (one time):
  * Deploy Code.js as a Web App (or run  python deploy_apps_script.py)
    and paste its /exec URL into .env  appscript_url :-
  * Secrets live in .env + Script Properties (never in code).
  * Test:  python run_billing.py --dry-run

Windows: always  python run_billing.py ...   (never .\\run_billing.py)
"""
from __future__ import annotations

import argparse
import calendar
import csv
import hashlib
import json
import re
import sys
import urllib.error
import urllib.request
from datetime import date, datetime, timedelta
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile

from lxml import etree

SCRIPT_DIR = Path(__file__).resolve().parent
BILLS_DIR = SCRIPT_DIR / "Shreeganesh_Telecom_GST_BILLS"
PDF_DIR = Path.home() / "Desktop" / "gst_bills"
ENV_FILE = SCRIPT_DIR / ".env"

# every log/state file lives in the logs\ folder next to this script
LOGS_DIR = SCRIPT_DIR / "logs"
LOGS_DIR.mkdir(parents=True, exist_ok=True)
LOG_FILE = LOGS_DIR / "bills_update_log.csv"
STATE_FILE = LOGS_DIR / "bills_state.json"   # idempotency cache: key -> last response


def log(msg: str) -> None:
    print(msg, file=sys.stderr)


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


# ------------------------------------------------------------------ date utils
DATE_FORMATS = [
    ("%d/%m/%Y", r"/"), ("%d-%m-%Y", r"-"), ("%d.%m.%Y", r"\."),
    ("%d/%m/%y", r"/"), ("%d-%m-%y", r"-"), ("%d.%m.%y", r"\."),
]


def parse_date(s: str):
    s = re.sub(r"\s", "", s)
    for fmt, _sep in DATE_FORMATS:
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            continue
    return None


def detect_out_format(old: str) -> str:
    """Same format the bill already used, e.g. dd-mm-yy stays dd-mm-yy."""
    s = re.sub(r"\s", "", old)
    if re.fullmatch(r"\d{1,2}[-/.]\d{1,2}[-/.]\d{4}", s):
        sep = "-" if "-" in s else ("/" if "/" in s else ".")
        return "%d" + sep + "%m" + sep + "%Y"
    return "%d-%m-%y"


def add_one_year(d: date) -> date:
    try:
        return d.replace(year=d.year + 1)
    except ValueError:                       # 29 Feb -> 28 Feb
        return d.replace(year=d.year + 1, day=28)


def add_months(d: date, n: int) -> date:
    """Same day-of-month, n calendar months later (day clamped to month end)."""
    y = d.year + (d.month - 1 + n) // 12
    m = (d.month - 1 + n) % 12 + 1
    last = calendar.monthrange(y, m)[1]
    return d.replace(year=y, month=m, day=min(d.day, last))


def infer_months(start: date, end: date) -> int | None:
    """Whole number of calendar months the bill period spans, or None when the
    period is not a whole number of months (then the caller falls back to
    day-based arithmetic)."""
    dm = (end.year - start.year) * 12 + (end.month - start.month)
    for m in (dm, dm + 1, dm - 1):           # e.g. 01/01->31/01 has dm == 0
        if m and m > 0:
            if add_months(start, m) - timedelta(days=1) == end or add_months(start, m) == end:
                return m
    return None


def shift_period(start: date, end: date, mode: str) -> tuple[date, date]:
    if mode == "year":                       # BOTH dates +1 year
        return add_one_year(start), add_one_year(end)
    if mode == "auto":                       # DEFAULT: the bill's own duration
        m = infer_months(start, end)
        if m:
            return add_months(start, m), add_months(end, m)
        # not a whole number of months -> day-based consecutive period
    # "next": the immediately following period of the same day-count
    duration = (end - start).days + 1
    new_start = end + timedelta(days=1)
    return new_start, new_start + timedelta(days=duration - 1)


def bump_invoice(inv: str) -> str:
    """'ST/009/2026-27' -> 'ST/010/2026-27'  (number +1, same digit width)."""
    m = re.fullmatch(r"\s*([A-Za-z]{1,6})\s*/\s*(\d+)\s*/\s*(\d{4})\s*-\s*(\d{1,2})\s*", inv or "")
    if not m:
        raise ValueError(f"invoice number not in PREFIX/NNNN/FY-FY form: {inv!r}")
    n = int(m.group(2)) + 1
    return f"{m.group(1)}/{str(n).zfill(len(m.group(2)))}/{m.group(3)}-{m.group(4)}"


# ------------------------------------------------------- docx editing machinery
W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"

RE_INVOICE_LABEL = re.compile(r"Invoice\s*No", re.I)
RE_INVOICE_NO = re.compile(r"[A-Za-z]{1,6}\s*/\s*\d{1,6}\s*/\s*\d{4}\s*-\s*\d{1,2}")
RE_INVOICE_DATE_LABEL = re.compile(r"Invoice\s*Date", re.I)
RE_DATE_VALUE = re.compile(r"\d{1,2}\s*[-/.]\s*\d{1,2}\s*[-/.]\s*\d{2,4}")
RE_PERIOD_LABEL = re.compile(r"Bill\s*Period", re.I)
RE_PERIOD = re.compile(
    r"(\d{1,2}\s*[-/.]\s*\d{1,2}\s*[-/.]\s*\d{4})\s*(to)\s*(\d{1,2}\s*[-/.]\s*\d{1,2}\s*[-/.]\s*\d{4})", re.I)

LOG_FIELDS = ["ts", "key", "bill_file", "old_invoice_no", "new_invoice_no",
              "invoice_date", "old_bill_period", "new_bill_period", "pdf_path", "status", "error"]


def _wt_nodes(par) -> list:
    return [t for t in par.iter(W + "t")]


def _offsets(ts) -> tuple[str, list[int]]:
    combined, offs, pos = "", [], 0
    for t in ts:
        offs.append(pos)
        combined += t.text or ""
        pos += len(t.text or "")
    return combined, offs


def _rewrite_span(ts, offs, span_s: int, span_e: int, new_sub: str) -> None:
    """Replace combined[span_s:span_e] with new_sub, preserving every run's
    text outside the span. Formatting of untouched text is untouched."""
    first_done = False
    for t, o in zip(ts, offs):
        s = t.text or ""
        a, b = o, o + len(s)
        if b <= span_s or a >= span_e:       # no overlap -> keep as-is
            continue
        pre = s[: max(0, span_s - a)]
        post = s[max(0, min(len(s), span_e - a)):]
        if not first_done:
            t.text = pre + new_sub + post
            first_done = True
        else:
            t.text = pre + post
        _preserve_space(t)


def _preserve_space(t) -> None:
    t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")


def _para_text(par) -> str:
    return "".join(t.text or "" for t in par.iter(W + "t"))


def _para_replace(par, pattern: re.Pattern, new_sub: str) -> bool:
    """Replace the first match of `pattern` inside one paragraph's runs."""
    ts = _wt_nodes(par)
    if not ts:
        return False
    combined, offs = _offsets(ts)
    m = pattern.search(combined)
    if not m:
        return False
    _rewrite_span(ts, offs, m.start(), m.end(), new_sub)
    return True


def _flat_pass(root, want: dict) -> dict:
    """One ordered walk over every paragraph of the part. When a label is seen
    ('Invoice No', 'Invoice Date', 'Bill Period'), the value may sit in the
    same paragraph, the next one, or (table layout) the following cell — all
    are just the next paragraphs in document order, so a short lookahead
    covers all layouts. Repeats (merged cells) are each updated."""
    hits = {k: 0 for k in want}
    pars = list(root.iter(W + "p"))
    texts = [_para_text(p) for p in pars]

    jobs = [
        ("invoice_no", RE_INVOICE_LABEL, RE_INVOICE_NO, 4),
        ("invoice_date", RE_INVOICE_DATE_LABEL, RE_DATE_VALUE, 4),
        ("period", RE_PERIOD_LABEL, RE_PERIOD, 2),
    ]
    for kind, label_re, value_re, lookahead in jobs:
        new_val = want.get(kind)
        if not new_val:
            continue
        i = 0
        while i < len(pars):
            if label_re.search(texts[i]):
                done = False
                for j in range(i, min(i + lookahead, len(pars))):
                    m = value_re.search(texts[j])
                    if m:
                        if kind == "period":
                            ts = _wt_nodes(pars[j])
                            combined, offs = _offsets(ts)
                            pm = RE_PERIOD.search(combined)
                            if pm:
                                new_comb = (combined[:pm.start(1)] + new_val[0]
                                            + combined[pm.end(1):pm.start(3)]
                                            + new_val[1] + combined[pm.end(3):])
                                head = pm.start(1)
                                tail = len(combined) - pm.end(3)
                                _rewrite_span(ts, offs, head, pm.end(3),
                                              new_comb[head: head + len(new_comb) - head - tail])
                                hits[kind] += 1
                        else:
                            if _para_replace(pars[j], value_re, new_val):
                                hits[kind] += 1
                        i = j + 1                   # continue after the value paragraph
                        done = True
                        break
                if not done:
                    i += 1                         # label with no value nearby
            else:
                i += 1
    return hits


def scan_values(xml_bytes: bytes) -> dict:
    """Detect current invoice no / invoice date / bill period in a part."""
    root = etree.fromstring(xml_bytes)
    pars = list(root.iter(W + "p"))
    texts = [_para_text(p) for p in pars]
    found: dict = {}
    jobs = [
        ("period", RE_PERIOD_LABEL, RE_PERIOD, 2),
        ("invoice_date", RE_INVOICE_DATE_LABEL, RE_DATE_VALUE, 4),
        ("invoice_no", RE_INVOICE_LABEL, RE_INVOICE_NO, 4),
    ]
    for kind, label_re, value_re, lookahead in jobs:
        for i, txt in enumerate(texts):
            if label_re.search(txt):
                for j in range(i, min(i + lookahead, len(pars))):
                    m = value_re.search(texts[j])
                    if m:
                        if kind == "period":
                            s, e = parse_date(m.group(1)), parse_date(m.group(3))
                            if s and e:
                                found[kind] = {"start": m.group(1), "end": m.group(3),
                                               "start_d": s.isoformat(), "end_d": e.isoformat()}
                        else:
                            found[kind] = re.sub(r"\s", "", m.group(0))
                        break
                if kind in found:
                    break
    return found


def apply_updates(xml_bytes: bytes, new_invoice_no: str | None, new_invoice_date: str | None,
                  new_period: tuple[str, str] | None) -> tuple[bytes, int]:
    """Rewrite label-anchored values in one XML part. Returns (new_xml, hits)."""
    root = etree.fromstring(xml_bytes)
    want = {"invoice_no": new_invoice_no, "invoice_date": new_invoice_date,
            "period": new_period}
    hits = _flat_pass(root, want)
    out = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
    return out, sum(hits.values())


def update_docx(path: Path, new_invoice_no: str, new_invoice_date: str,
                new_period: tuple[str, str]) -> dict:
    """Rewrite the docx in place (all document/header/footer parts)."""
    changed: dict[str, int] = {}
    tmp = path.with_suffix(".docx.updating")
    with ZipFile(path) as zin, ZipFile(tmp, "w", ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if re.fullmatch(r"word/(document|header\d*|footer\d*)\.xml", item.filename):
                new_data, hits = apply_updates(data, new_invoice_no, new_invoice_date, new_period)
                data, changed[item.filename] = new_data, hits
            zout.writestr(item, data)
    tmp.replace(path)
    return changed


# ------------------------------------------------------------------- PDF via Word
def _find_soffice() -> str | None:
    import os
    import shutil
    exe = shutil.which("soffice")
    if exe:
        return exe
    for base in (os.environ.get("PROGRAMFILES", r"C:\Program Files"),
                 os.environ.get("PROGRAMFILES(X86)", r"C:\Program Files (x86)")):
        p = Path(base) / "LibreOffice" / "program" / "soffice.exe"
        if p.exists():
            return str(p)
    return None


def docx_to_pdf(docx_path: Path, pdf_path: Path) -> None:
    soffice = _find_soffice()
    if soffice:
        import subprocess
        import tempfile
        prof = (Path(tempfile.gettempdir()) / "lo_bill_pdf_profile").as_uri()
        cmd = [soffice, f"-env:UserInstallation={prof}", "--headless", "--norestore",
               "--convert-to", "pdf", "--outdir", str(pdf_path.parent), str(docx_path)]
        cp = subprocess.run(cmd, capture_output=True, text=True, timeout=300)
        if not pdf_path.exists():
            raise RuntimeError(f"LibreOffice PDF conversion failed: {(cp.stderr or cp.stdout).strip()[:300]}")
        return
    # fallback: MS Word via COM
    import pythoncom
    import win32com.client
    pythoncom.CoInitialize()
    word = None
    try:
        word = win32com.client.DispatchEx("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0
        doc = word.Documents.Open(str(docx_path), ReadOnly=True)
        doc.SaveAs2(str(pdf_path), FileFormat=17)          # 17 = wdFormatPDF
        doc.Close(False)
    finally:
        if word is not None:
            word.Quit()
        pythoncom.CoUninitialize()


# ------------------------------------------------------------------ file lookup
def norm_name(s: str) -> str:
    s = re.sub(r"[^A-Za-z0-9()&._\- ]", "", (s or "").upper())
    return re.sub(r"\s+", "_", s).strip("_")


STOPWORDS_FILE_MATCH = {"PVT", "LTD", "LLP", "PRIVATE", "LIMITED", "CO", "INC", "PC",
                        "THE", "AND", "SOLUTIONS", "ENTERPRISES", "INDIA", "STORES", "COMPANY"}

# explicit customer -> bill file pins (checked BEFORE any fuzzy matching).
# key = norm_name(customer), value = bill file stem in BILLS_DIR.
# Add a line here whenever a customer name is ambiguous or the file name
# differs in spelling from the customer name.
BILL_FILE_ALIASES = {
    "SKIN_CITY_AUNDH": "SKINCITY_AUNDH_BILL",   # 'Skin City Aundh'
    "SKINCITY_AUNDH": "SKINCITY_AUNDH_BILL",    # spelled without the space
}


def _tokens(s: str) -> list[str]:
    """Meaningful tokens for file matching: drop corporate suffixes and
    anything shorter than 3 chars (single letters otherwise match ANY file
    containing that letter and destroy the score)."""
    toks = [t for t in re.sub(r"[^A-Za-z0-9]+", " ", (s or "").upper()).split() if t]
    return [t for t in toks
            if (len(t) >= 3 or t.isdigit()) and t not in STOPWORDS_FILE_MATCH]


def resolve_bill(bills_dir: Path, customer: str | None, bill_file: str | None) -> Path:
    if bill_file:
        p = bills_dir / bill_file
        if p.exists():
            return p
        raise FileNotFoundError(f"bill file {bill_file!r} not found in {bills_dir}")
    if not customer:
        raise ValueError("record needs 'customer' (or 'bill_file')")
    n = norm_name(customer)
    alias = BILL_FILE_ALIASES.get(n)
    if alias:                                # pinned mapping: no fuzzy matching
        p = bills_dir / f"{alias}.docx"
        if p.exists():
            return p
        raise FileNotFoundError(f"pinned bill file {p.name!r} for customer {customer!r} "
                                f"is missing in {bills_dir}")
    toks = _tokens(customer)
    files = {p.stem.upper(): p for p in bills_dir.glob("*.docx") if not p.name.startswith("~$")}
    # 1) exact "<NAME>_BILL"  2) stem starts with name  3) stem contains name
    for cand in (f"{n}_BILL", n):
        if cand in files:
            return files[cand]
    starts = [p for s, p in files.items() if s.startswith(n + "_") or s == n]
    if len(starts) == 1:
        return starts[0]
    contains = [p for s, p in files.items() if n in s]
    if len(contains) == 1:
        return contains[0]
    # 4) token-coverage scoring: fraction of the customer's words that appear
    #    (as substrings) in the file stem minus its trailing "_BILL"
    scored = []
    for stem, p in files.items():
        core = re.sub(r"[_\-]?BILL$", "", stem)
        flat = core.replace("_", " ").replace("-", " ")
        cov = (sum(1 for t in toks if t and t in flat) / len(toks)) if toks else 0.0
        scored.append((round(cov, 3), stem, p))
    scored.sort(key=lambda x: (-x[0], x[1]))
    if scored and scored[0][0] > 0:
        if len(scored) == 1 or scored[0][0] - scored[1][0] > 0.05:
            return scored[0][2]
    raise FileNotFoundError(
        f"no unique bill for customer {customer!r} in {bills_dir.name} — add the exact "
        f"file name to the sheet's 'Bill File' column. Closest: "
        f"{[(s, c) for c, s, _p in scored[:4]]}")


# ------------------------------------------------------------------ persistence
def state_load() -> dict:
    if STATE_FILE.exists():
        try:
            return json.loads(STATE_FILE.read_text(encoding="utf-8"))
        except Exception:
            return {}
    return {}


def state_save(state: dict) -> None:
    STATE_FILE.write_text(json.dumps(state, indent=2, ensure_ascii=False), encoding="utf-8")


def log_append(row: dict) -> None:
    new = LOG_FILE.exists()
    with LOG_FILE.open("a", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=LOG_FIELDS)
        if not new:
            w.writeheader()
        w.writerow({k: row.get(k, "") for k in LOG_FIELDS})


# ------------------------------------------------------------------ core update
def process_record(rec: dict, *, dry_run: bool = False, period_shift: str = "auto",
                   bills_dir: Path | None = None, pdf_dir: Path | None = None) -> dict:
    pdf_dir = Path(pdf_dir) if pdf_dir else PDF_DIR
    bills_dir = Path(bills_dir) if bills_dir else BILLS_DIR
    key = str(rec.get("key") or "").strip() or hashlib.sha256(
        json.dumps(rec, sort_keys=True, default=str).encode()).hexdigest()[:16]

    state = state_load()
    if key in state and state[key].get("status") == "ok":
        return {**state[key], "duplicate": True, "note": "this key was already processed"}

    bill = resolve_bill(bills_dir, rec.get("customer"), rec.get("bill_file"))

    with ZipFile(bill) as z:
        old = scan_values(z.read("word/document.xml"))
    if not old.get("invoice_no"):
        raise ValueError(f"could not find 'Invoice No' in {bill.name} (template changed?)")

    old_inv = old["invoice_no"]
    expect = str(rec.get("expect_invoice_no") or "").strip()
    if expect and old_inv != expect:
        raise ValueError(
            f"bill already shows invoice no {old_inv} but the sheet expected {expect} — "
            f"refusing to update again (double-bill guard). If this update is intentional, "
            f"clear logs\\bills_state.json or check the sheet's GST RECORDS.")

    new_inv = str(rec["new_invoice_no"]).strip() if rec.get("new_invoice_no") else bump_invoice(
        str(rec.get("invoice_no") or old_inv))

    fmt = detect_out_format(old.get("invoice_date", ""))
    inv_date = str(rec.get("invoice_date") or "").strip() or date.today().strftime(fmt)

    per = old.get("period")
    if not per:
        raise ValueError(f"could not find 'Bill Period' in {bill.name} (template changed?)")
    s, e = parse_date(per["start"]), parse_date(per["end"])
    ns, ne = shift_period(s, e, period_shift)
    sep = "/" if "/" in per["start"] else ("-" if "-" in per["start"] else ".")
    pfmt = "%d" + sep + "%m" + sep + "%Y"
    old_period_str = f"{per['start']} to {per['end']}"
    new_period_str = f"{ns.strftime(pfmt)} to {ne.strftime(pfmt)}"

    result = {
        "key": key, "bill": bill.name, "old_invoice_no": old_inv, "new_invoice_no": new_inv,
        "invoice_date": inv_date, "old_bill_period": old_period_str, "new_bill_period": new_period_str,
        "period_shift": period_shift,
        "docx": str(bill), "pdf": str(pdf_dir / (bill.stem + ".pdf")),
    }

    if dry_run:
        result.update(status="dry-run", note="nothing was saved (dry-run)")
        return result

    update_docx(bill, new_inv, inv_date, (ns.strftime(pfmt), ne.strftime(pfmt)))

    pdf_dir.mkdir(parents=True, exist_ok=True)
    pdf_path = pdf_dir / (bill.stem + ".pdf")
    try:
        docx_to_pdf(bill, pdf_path)
        result["pdf_created"] = True
    except Exception as exc:                 # docx is already updated; report pdf separately
        result["pdf_created"] = False
        result["pdf_error"] = f"{exc.__class__.__name__}: {exc}"

    result["status"] = "ok"
    state[key] = result
    state_save(state)
    log_append(result)
    return result


# ------------------------------------------------------------------------- main
def main() -> int:
    ap = argparse.ArgumentParser(
        description="GST billing: bill every UPCOMING BILLS row that is due "
                    "(or bill one customer by hand with --customer/--bill-file).")
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
    # ---- one-shot mode (bill a single customer by hand) -------------------
    ap.add_argument("--customer", help="bill this ONE customer only (no sheet contact)")
    ap.add_argument("--bill-file", help="bill file name directly (overrides --customer)")
    ap.add_argument("--invoice-no", help="CURRENT invoice no, e.g. ST/048/2026-27 (default: read from the bill)")
    ap.add_argument("--new-invoice-no", help="write this EXACT invoice no instead of +1 (testing)")
    ap.add_argument("--invoice-date", help="invoice date to write, e.g. 13-11-26 (default: today, bill's own format)")
    args = ap.parse_args()

    pdf_dir = Path(args.pdf_dir or setting("pdf_dir") or PDF_DIR)
    bills_dir = Path(args.bills_dir or setting("bills_dir") or BILLS_DIR)

    # ---- one-shot mode ------------------------------------------------------
    if args.customer or args.bill_file:
        rec = {k: v for k, v in {
            "customer": args.customer, "bill_file": args.bill_file,
            "invoice_no": args.invoice_no, "new_invoice_no": args.new_invoice_no,
            "invoice_date": args.invoice_date,
        }.items() if v}
        try:
            result = process_record(rec, dry_run=args.dry_run, period_shift=args.period_shift,
                                    bills_dir=bills_dir, pdf_dir=pdf_dir)
        except Exception as exc:
            log(f"[error] {exc.__class__.__name__}: {exc}")
            print(json.dumps({"ok": False, "error": f"{exc.__class__.__name__}: {exc}"},
                             ensure_ascii=False))
            return 1
        print(json.dumps({"ok": True, **result}, ensure_ascii=False, indent=2))
        return 0

    # ---- full billing flow --------------------------------------------------
    url = args.url or setting("appscript_url") or ""
    if not url or "PASTE" in url.upper():
        print("ERROR: .env has no real 'appscript_url :-' yet.\n"
              "Deploy Code.js as a Web App (Deploy > New deployment > Web app,\n"
              "Execute as: Me, access: Anyone), copy the /exec URL and put it in .env:\n"
              "    appscript_url :- https://script.google.com/macros/s/...../exec",
              file=sys.stderr)
        return 2
    secret = setting("appscript_secret") or ""

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
            res = process_record(rec, dry_run=args.dry_run, period_shift=args.period_shift,
                                 bills_dir=bills_dir, pdf_dir=pdf_dir)
            dup = res.get("duplicate", False)
            pdf_note = "" if res.get("pdf_created", True) else f"  [PDF FAILED: {res.get('pdf_error')}]"
            print(f"  {item['customer']}: {'(already done, skipped)' if dup else 'billed'}"
                  f" {res['new_invoice_no']}  period {res['old_bill_period']} -> {res['new_bill_period']}"
                  f"{pdf_note}")
            # plan duration in whole calendar months, from the NEW bill period
            # (e.g. 02/01/2027 to 01/01/2028 -> 12) for the sheet's Duration column
            _s, _e = [parse_date(x) for x in str(res["new_bill_period"]).split(" to ")]
            _dm = (_e.year - _s.year) * 12 + (_e.month - _s.month) if (_s and _e) else None
            results.append({"key": item["key"], "ok": True, "duplicate": dup,
                            "new_invoice_no": res["new_invoice_no"],
                            "invoice_date": res["invoice_date"],
                            "new_expire": str(res["new_bill_period"]).split(" to ")[-1],
                            "duration_months": _dm,
                            "pdf_error": res.get("pdf_error")})
        except Exception as exc:
            failed += 1
            print(f"  {item['customer']}: ERROR {exc.__class__.__name__}: {exc}", file=sys.stderr)
            results.append({"key": item["key"], "ok": False,
                            "error": f"{exc.__class__.__name__}: {exc}"})

    if args.dry_run:
        print("Dry run finished — nothing was saved, the sheet was not touched.")
        return 0

    # ---- 3. report back: append GST rows, advance Expire Dates + Duration ---
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
