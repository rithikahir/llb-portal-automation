#!/usr/bin/env python3
"""
update_bills.py — GST bill updater for the Shreeganesh Telecom bills folder.

Flow (the reason this exists):
  Google Apps Script picks up a record (customer + current invoice no) and
  POSTs it here; this script picks the matching bill .docx from
  Shreeganesh_Telecom_GST_BILLS/, updates it and produces a PDF.

What gets updated inside the bill (label-anchored, so nothing else is touched):
  * Invoice Date : - <old>          -> today's date (same format as found)
  * Invoice No  : - ST/009/2026-27  -> ST/010/2026-27   (number +1, same width)
                                       (or the exact value sent as
                                        "new_invoice_no", used for testing)
  * Bill Period:-<start> to <end>   -> shifted period (default: BOTH dates
                                       +1 year; alternative: --period-shift next)

Outputs:
  * the .docx is saved in place (a backup copy is made first, in
    Shreeganesh_Telecom_GST_BILLS_backup/)
  * a PDF is written to the Desktop folder  Desktop\\gst_bills\\<bill>.pdf

Two ways to run:
  1. SERVER (default)  — python update_bills.py
       Listens for POST /update with JSON from Apps Script:
         {"customer": "SKIN CITY", "invoice_no": "ST/023/2026-27",
          "key": "row-123"}          # key = idempotency (recommended)
       Optional fields: bill_file, new_invoice_no, invoice_date, pdf_dir.
       Duplicate keys are detected and answered from cache without editing.
       Health check: GET /health
  2. ONE-SHOT (CLI)    — python update_bills.py --customer "SKIN CITY" ...
       See --help.  Add --dry-run to preview without saving anything.

Windows quirks:
  * run as  python update_bills.py ...   (never .\\update_bills.py)
  * PDF conversion uses LibreOffice (headless) when installed, otherwise
    MS Word via COM.
"""
from __future__ import annotations

import argparse
import csv
import hashlib
import json
import re
import sys
import threading
from datetime import date, datetime, timedelta
from http.server import BaseHTTPRequestHandler, HTTPServer
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile

from lxml import etree

SCRIPT_DIR = Path(__file__).resolve().parent
BILLS_DIR = SCRIPT_DIR / "Shreeganesh_Telecom_GST_BILLS"
BACKUP_DIR = SCRIPT_DIR / "Shreeganesh_Telecom_GST_BILLS_backup"
PDF_DIR = Path.home() / "Desktop" / "gst_bills"
LOG_FILE = SCRIPT_DIR / "bills_update_log.csv"
STATE_FILE = SCRIPT_DIR / "bills_state.json"      # idempotency cache: key -> last response

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


def log(msg: str) -> None:
    print(msg, file=sys.stderr)


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


def shift_period(start: date, end: date, mode: str) -> tuple[date, date]:
    if mode == "year":                       # BOTH dates +1 year (the user's rule)
        return add_one_year(start), add_one_year(end)
    # "next": the immediately following period of the same duration
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


def resolve_bill(bills_dir: Path, customer: str | None, bill_file: str | None) -> Path:
    if bill_file:
        p = bills_dir / bill_file
        if p.exists():
            return p
        raise FileNotFoundError(f"bill file {bill_file!r} not found in {bills_dir}")
    if not customer:
        raise ValueError("record needs 'customer' (or 'bill_file')")
    n = norm_name(customer)
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
    raise FileNotFoundError(
        f"no unique bill for customer {customer!r} in {bills_dir.name} "
        f"(exact={bool(starts or contains)}, candidates={[p.name for p in (starts + contains)][:8]})")


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
def process_record(rec: dict, *, dry_run: bool = False, period_shift: str = "year",
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

    # backup, then rewrite the docx in place
    BACKUP_DIR.mkdir(parents=True, exist_ok=True)
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    bill_bak = BACKUP_DIR / f"{bill.stem}_{stamp}{bill.suffix}"
    bill_bak.write_bytes(bill.read_bytes())
    result["backup"] = str(bill_bak)

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


# ---------------------------------------------------------------------- server
class Handler(BaseHTTPRequestHandler):
    server_version = "BillUpdater/1.0"

    def _json(self, code: int, obj: dict) -> None:
        body = json.dumps(obj, ensure_ascii=False).encode("utf-8")
        self.send_response(code)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def do_GET(self):                        # health check for Apps Script / browser
        if self.path.startswith("/health"):
            self._json(200, {"ok": True, "service": "update_bills", "bills_dir": str(BILLS_DIR)})
        else:
            self._json(404, {"ok": False, "error": "POST /update with JSON"})

    def do_POST(self):
        if self.path.rstrip("/") not in ("/update", ""):
            self._json(404, {"ok": False, "error": "unknown endpoint"})
            return
        try:
            length = int(self.headers.get("Content-Length") or 0)
            rec = json.loads(self.rfile.read(length) or b"{}")
        except Exception as exc:
            self._json(400, {"ok": False, "error": f"bad JSON body: {exc}"})
            return
        try:
            result = process_record(rec if isinstance(rec, dict) else {},
                                    period_shift=(rec.get("period_shift") or "year") if isinstance(rec, dict) else "year",
                                    pdf_dir=rec.get("pdf_dir") if isinstance(rec, dict) else None)
            self._json(200, {"ok": True, **result})
        except Exception as exc:
            log(f"[error] {exc.__class__.__name__}: {exc}")
            self._json(500, {"ok": False, "error": f"{exc.__class__.__name__}: {exc}"})

    def log_message(self, fmt, *args):       # quieter console
        log(f"[http] {self.address_string()} {fmt % args}")


def serve(host: str, port: int) -> None:
    httpd = HTTPServer((host, port), Handler)   # single-threaded: Word COM stays safe
    log(f"[serve] listening on http://{host}:{port}  (POST /update, GET /health)")
    log(f"[serve] bills folder: {BILLS_DIR}")
    log(f"[serve] pdf output   : {PDF_DIR}")
    log("[serve] Ctrl+C to stop")
    try:
        httpd.serve_forever()
    except KeyboardInterrupt:
        log("[serve] stopped")


# ------------------------------------------------------------------------- CLI
def main() -> int:
    ap = argparse.ArgumentParser(description="Update Shreeganesh Telecom GST bills (docx) and export PDF.")
    ap.add_argument("--serve", action="store_true", help="run as HTTP server for Apps Script (default when no --customer/--bill-file)")
    ap.add_argument("--host", default="0.0.0.0", help="server bind address (default 0.0.0.0)")
    ap.add_argument("--port", type=int, default=8765, help="server port (default 8765)")
    ap.add_argument("--customer", help="customer the record is for; picks <NAME>_BILL.docx")
    ap.add_argument("--bill-file", help="bill file name directly (overrides --customer)")
    ap.add_argument("--invoice-no", help="CURRENT invoice no from Apps Script, e.g. ST/023/2026-27 (default: read from the bill)")
    ap.add_argument("--new-invoice-no", help="write this EXACT invoice no instead of +1 (testing)")
    ap.add_argument("--invoice-date", help="invoice date to write, e.g. 11-11-26 (default: today, bill's own format)")
    ap.add_argument("--period-shift", choices=("year", "next"), default="year",
                    help="year = both bill-period dates +1 year (default); next = the immediately following period")
    ap.add_argument("--pdf-dir", help="override PDF output folder (default: Desktop\\gst_bills)")
    ap.add_argument("--bills-dir", help="override bills folder (default: Shreeganesh_Telecom_GST_BILLS next to script)")
    ap.add_argument("--dry-run", action="store_true", help="show what WOULD change; save nothing")
    args = ap.parse_args()

    if args.bills_dir:
        global BILLS_DIR, BACKUP_DIR
        BILLS_DIR = Path(args.bills_dir)
        BACKUP_DIR = BILLS_DIR.parent / (BILLS_DIR.name + "_backup")

    one_shot = args.customer or args.bill_file
    if not one_shot and not args.serve:
        args.serve = True                    # default mode: server

    if args.serve and not one_shot:
        serve(args.host, args.port)
        return 0

    rec = {k: v for k, v in {
        "customer": args.customer, "bill_file": args.bill_file,
        "invoice_no": args.invoice_no, "new_invoice_no": args.new_invoice_no,
        "invoice_date": args.invoice_date,
    }.items() if v}
    try:
        result = process_record(rec, dry_run=args.dry_run, period_shift=args.period_shift,
                                pdf_dir=Path(args.pdf_dir) if args.pdf_dir else None)
    except Exception as exc:
        log(f"[error] {exc.__class__.__name__}: {exc}")
        print(json.dumps({"ok": False, "error": f"{exc.__class__.__name__}: {exc}"}, ensure_ascii=False))
        return 1
    print(json.dumps({"ok": True, **result}, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    sys.exit(main())
